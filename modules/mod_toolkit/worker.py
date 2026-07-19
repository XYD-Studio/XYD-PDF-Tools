# -*- coding: utf-8 -*-
import os
import hashlib
import shutil
import tempfile
import concurrent.futures
from PyQt5.QtCore import QThread, pyqtSignal
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
import fitz

# 引入核心底层的通用函数
from core.pdf_engine import merge_pdf_with_smart_toc, get_unique_filepath


def _parse_page_positions(text):
    positions = set()
    for part in str(text or "").replace("，", ",").split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            start_text, end_text = part.split("-", 1)
            start, end = int(start_text.strip()), int(end_text.strip())
            if start <= 0 or end <= 0 or end < start:
                raise ValueError("预留页位必须是正整数，范围格式如 3-5")
            positions.update(range(start, end + 1))
        else:
            pos = int(part)
            if pos <= 0:
                raise ValueError("预留页位必须是正整数")
            positions.add(pos)
    return positions


def _build_interleave_sequence(page_count, reserved_text, reverse=False):
    reserved = _parse_page_positions(reserved_text)
    source_pages = list(range(page_count))
    if reverse:
        source_pages.reverse()

    sequence = []
    source_idx = 0
    slot = 1
    max_reserved_slot = max(reserved) if reserved else 0

    while source_idx < len(source_pages) or slot <= max_reserved_slot:
        if slot in reserved:
            sequence.append(None)
        elif source_idx < len(source_pages):
            sequence.append(source_pages[source_idx])
            source_idx += 1
        else:
            sequence.append(None)
        slot += 1

    return sequence


def _slot_page_rect(doc, sequence, slot_idx):
    if not doc or len(doc) == 0:
        return None

    if slot_idx < len(sequence) and sequence[slot_idx] is not None:
        return doc[sequence[slot_idx]].rect

    for idx in range(slot_idx + 1, len(sequence)):
        page_idx = sequence[idx]
        if page_idx is not None:
            return doc[page_idx].rect

    for idx in range(min(slot_idx - 1, len(sequence) - 1), -1, -1):
        page_idx = sequence[idx]
        if page_idx is not None:
            return doc[page_idx].rect

    return doc[0].rect


def _apply_page_rotation(page, rotate_degrees):
    rotate_degrees = int(rotate_degrees or 0) % 360
    if rotate_degrees:
        page.set_rotation((page.rotation + rotate_degrees) % 360)


# ================= 必须放在顶层的多进程 Worker 函数 =================
def _worker_render_temp_img(args):
    pdf_path, page_index, dpi, temp_dir, img_fmt, is_transparent = args
    try:
        import pypdfium2 as pdfium
        from PIL import Image
        pdf = pdfium.PdfDocument(pdf_path)
        scale = dpi / 72.0
        bg_color = (0, 0, 0, 0) if (img_fmt == "png" and is_transparent) else (255, 255, 255, 255)

        pil_image = pdf[page_index].render(scale=scale, rotation=0, fill_color=bg_color).to_pil()
        pdf.close()

        if img_fmt == "jpg" and pil_image.mode != 'RGB':
            pil_image = pil_image.convert('RGB')

        source_key = hashlib.sha1(os.path.abspath(pdf_path).encode('utf-8')).hexdigest()[:12]
        temp_name = f"temp_{source_key}_{page_index}.{img_fmt}"
        temp_path = os.path.join(temp_dir, temp_name)
        pil_image.save(temp_path, quality=95, dpi=(dpi, dpi))
        del pil_image
        return (page_index, temp_path)
    except Exception as e:
        return (page_index, f"ERROR: {str(e)}")

def _worker_pdf2img_save(args):
    pdf_path, page_index, dpi, output_dir, base_name, img_fmt, is_transparent = args
    try:
        import pypdfium2 as pdfium
        from PIL import Image
        pdf = pdfium.PdfDocument(pdf_path)
        scale = dpi / 72.0
        bg_color = (0, 0, 0, 0) if (img_fmt == "png" and is_transparent) else (255, 255, 255, 255)

        pil_image = pdf[page_index].render(scale=scale, rotation=0, fill_color=bg_color).to_pil()
        pdf.close()

        if img_fmt == "jpg" and pil_image.mode != 'RGB':
            pil_image = pil_image.convert('RGB')

        save_name = f"{base_name}_page_{page_index + 1:03d}.{img_fmt}"
        save_path = os.path.join(output_dir, save_name)
        pil_image.save(save_path, quality=95)
        del pil_image
        return True
    except Exception:
        return False

# ================= 工具箱主干处理线程 =================
class ToolkitWorker(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, paths, mode, out_dir, dpi, fmt, is_trans, split_config=None, interleave_config=None):
        super().__init__()
        self.paths = paths
        self.mode = mode
        self.out_dir = out_dir
        self.dpi = dpi
        self.fmt = fmt
        self.is_trans = is_trans
        self.split_config = split_config if split_config else {}
        self.interleave_config = interleave_config if interleave_config else {}

    def _blank_page_size(self, primary_doc, primary_seq, secondary_doc, secondary_seq, slot_idx):
        rect = _slot_page_rect(primary_doc, primary_seq, slot_idx) or _slot_page_rect(secondary_doc, secondary_seq, slot_idx)
        if rect:
            return rect.width, rect.height
        return 595, 842

    def _append_source_page(self, output_doc, source_doc, page_idx, rotate_degrees):
        output_doc.insert_pdf(source_doc, from_page=page_idx, to_page=page_idx)
        _apply_page_rotation(output_doc[-1], rotate_degrees)

    def _append_blank_page(self, output_doc, width, height, rotate_degrees):
        page = output_doc.new_page(width=width, height=height)
        _apply_page_rotation(page, rotate_degrees)

    def _run_interleave_merge(self):
        pdf_paths = [p for p in self.paths if p.lower().endswith('.pdf')]
        if len(pdf_paths) != 2:
            raise Exception("PDF交叉合并需要且仅需要两个PDF文件。")

        cfg = self.interleave_config
        rotate_a = int(cfg.get('rotate_a', 0) or 0)
        rotate_b = int(cfg.get('rotate_b', 0) or 0)
        fill_blank = bool(cfg.get('fill_blank', True))

        doc_a = fitz.open(pdf_paths[0])
        doc_b = fitz.open(pdf_paths[1])
        output_doc = fitz.Document()

        try:
            seq_a = _build_interleave_sequence(len(doc_a), cfg.get('reserved_a', ""), reverse=False)
            seq_b = _build_interleave_sequence(len(doc_b), cfg.get('reserved_b', ""), bool(cfg.get('reverse_b', False)))
            total_slots = max(len(seq_a), len(seq_b))
            if total_slots == 0:
                raise Exception("两个PDF都没有可合并的页面。")

            total_output_units = total_slots * 2
            done_units = 0
            toc = []

            for slot_idx in range(total_slots):
                for label, doc, seq, other_doc, other_seq, rotate in [
                    ("A", doc_a, seq_a, doc_b, seq_b, rotate_a),
                    ("B", doc_b, seq_b, doc_a, seq_a, rotate_b),
                ]:
                    has_slot = slot_idx < len(seq)
                    page_idx = seq[slot_idx] if has_slot else None
                    should_write_blank = (has_slot and page_idx is None) or (not has_slot and fill_blank)

                    if page_idx is not None:
                        self._append_source_page(output_doc, doc, page_idx, rotate)
                        toc.append([1, f"{label} PDF 第 {page_idx + 1} 页", len(output_doc)])
                    elif should_write_blank:
                        width, height = self._blank_page_size(doc, seq, other_doc, other_seq, slot_idx)
                        self._append_blank_page(output_doc, width, height, rotate)
                        blank_title = "预留空白页" if has_slot else "自动补齐空白页"
                        toc.append([1, f"{label} PDF {blank_title} {slot_idx + 1}", len(output_doc)])

                    done_units += 1
                    self.progress.emit(
                        int(done_units / total_output_units * 95),
                        f"正在交叉合并: 第 {slot_idx + 1}/{total_slots} 组"
                    )

            output_doc.set_toc(toc)
            output_doc.save(self.out_dir)

        finally:
            output_doc.close()
            doc_a.close()
            doc_b.close()

    def run(self):
        temp_dir = tempfile.mkdtemp(prefix="pdf_tool_tmp_")
        try:
            max_workers = max(1, (os.cpu_count() or 2) - 1)
            result_message = None

            if self.mode == "PDF交叉合并":
                self.progress.emit(5, "正在准备两个PDF的交叉合并队列...")
                self._run_interleave_merge()

            elif "合并" in self.mode:
                self.progress.emit(50, "正在进行超高速 PDF 合并与智能大纲重建...")
                merged_doc = fitz.Document()
                toc_list = []
                total = len(self.paths)
                for idx, p in enumerate(self.paths):
                    self.progress.emit(int(idx / total * 100), f"正在合并: {os.path.basename(p)}...")
                    if p.lower().endswith('.pdf'):
                        doc = fitz.open(p)
                        merge_pdf_with_smart_toc(doc, os.path.basename(p), merged_doc, toc_list)
                        doc.close()
                merged_doc.set_toc(toc_list)
                merged_doc.save(self.out_dir)
                merged_doc.close()

            elif self.mode == "PDF拆分为单页":
                total = len(self.paths)
                for idx, p in enumerate(self.paths):
                    self.progress.emit(int(idx / total * 100), f"正在拆分 {os.path.basename(p)}...")
                    if p.lower().endswith('.pdf'):
                        reader = PdfReader(p)
                        base = os.path.splitext(os.path.basename(p))[0]
                        for i, page in enumerate(reader.pages):
                            writer = PdfWriter()
                            writer.add_page(page)
                            with open(os.path.join(self.out_dir, f"{base}_p{i + 1}.pdf"), "wb") as f:
                                writer.write(f)

            elif "按书签拆分" in self.mode:
                import re
                total_files = len(self.paths)
                for f_idx, p in enumerate(self.paths):
                    if not p.lower().endswith('.pdf'): continue
                    self.progress.emit(int(f_idx / total_files * 100), f"正在按书签拆分 {os.path.basename(p)}...")

                    doc = fitz.open(p)
                    toc = doc.get_toc()
                    page_to_title = {item[2]: item[1] for item in toc}
                    total_pages = len(doc)
                    base_name = os.path.splitext(os.path.basename(p))[0]
                    current_bookmark = base_name

                    for i in range(total_pages):
                        page_num = i + 1
                        if page_num in page_to_title:
                            current_bookmark = page_to_title[page_num]

                        clean_title = re.sub(r'[\\/*?:"<>|]', '_', current_bookmark).strip()
                        naming_mode = self.split_config.get('mode', 1)
                        segments = self.split_config.get('segments', [])

                        prefix, suffix = "", ""
                        for seg in segments:
                            if seg[0] <= page_num <= seg[1]:
                                prefix, suffix = seg[2], seg[3]
                                break

                        if naming_mode == 2:
                            final_name = f"{page_num:03d}_{prefix}{clean_title}{suffix}.pdf"
                        else:
                            final_name = f"{prefix}{clean_title}{suffix}.pdf"

                        final_path = get_unique_filepath(self.out_dir, final_name)
                        new_doc = fitz.Document()
                        new_doc.insert_pdf(doc, from_page=i, to_page=i)
                        new_doc.save(final_path)
                        new_doc.close()
                    doc.close()

            elif "多图转PDF" in self.mode:
                from PIL import Image
                self.progress.emit(50, "正在拼接图片为PDF...")
                imgs = []
                for p in self.paths:
                    try:
                        img = Image.open(p)
                        if img.mode != 'RGB': img = img.convert('RGB')
                        imgs.append(img)
                    except:
                        pass
                if imgs:
                    imgs[0].save(self.out_dir, "PDF", resolution=100.0, save_all=True, append_images=imgs[1:])

            elif "图片型PDF" in self.mode:
                import pypdfium2 as pdfium
                pdf_paths = [path for path in self.paths if path.lower().endswith('.pdf')]
                if not pdf_paths:
                    raise Exception("没有找到有效的PDF页面")
                total_pages = 0
                page_counts = {}
                for path in pdf_paths:
                    pdfium_doc = pdfium.PdfDocument(path)
                    page_counts[path] = len(pdfium_doc)
                    total_pages += len(pdfium_doc)
                    pdfium_doc.close()

                successes = []
                failures = []
                completed = 0
                for file_index, path in enumerate(pdf_paths):
                    count = page_counts[path]
                    page_files = {}
                    tasks = [(path, page_index, self.dpi, temp_dir, "jpg", False)
                             for page_index in range(count)]
                    with concurrent.futures.ProcessPoolExecutor(max_workers=max_workers) as executor:
                        futures = {executor.submit(_worker_render_temp_img, task): task[1] for task in tasks}
                        for future in concurrent.futures.as_completed(futures):
                            page_index, result = future.result()
                            if isinstance(result, str) and not result.startswith("ERROR"):
                                page_files[page_index] = result
                            completed += 1
                            self.progress.emit(
                                int((completed / max(1, total_pages)) * 85),
                                f"正在转换 {os.path.basename(path)}: {len(page_files)}/{count}",
                            )

                    missing_pages = [index + 1 for index in range(count) if index not in page_files]
                    if missing_pages:
                        failures.append(f"{os.path.basename(path)}: 第 {missing_pages} 页渲染失败")
                        continue

                    if len(pdf_paths) == 1:
                        final_path = self.out_dir
                    else:
                        base_name = os.path.splitext(os.path.basename(path))[0]
                        final_path = get_unique_filepath(self.out_dir, f"{base_name}_图片型.pdf")
                    source_doc = fitz.open(path)
                    output_doc = fitz.Document()
                    try:
                        for page_index in range(count):
                            source_page = source_doc[page_index]
                            output_page = output_doc.new_page(
                                width=source_page.rect.width,
                                height=source_page.rect.height,
                            )
                            output_page.insert_image(
                                output_page.rect,
                                filename=page_files[page_index],
                                keep_proportion=False,
                            )
                        toc = source_doc.get_toc()
                        if toc:
                            output_doc.set_toc(toc)
                        output_doc.save(final_path, garbage=4, deflate=True)
                        successes.append(final_path)
                    except Exception as exc:
                        if os.path.exists(final_path):
                            os.remove(final_path)
                        failures.append(f"{os.path.basename(path)}: {exc}")
                    finally:
                        output_doc.close()
                        source_doc.close()

                if not successes:
                    raise Exception("所有文件转换失败:\n" + "\n".join(failures))
                summary = [f"成功 {len(successes)} 个文件"]
                if failures:
                    summary.append(f"失败 {len(failures)} 个文件:\n" + "\n".join(failures))
                result_message = "图片型 PDF 批处理完成，" + "；".join(summary)

            elif "导出图片" in self.mode:
                import pypdfium2 as pdfium
                tasks = []
                total_pages = 0
                for f in self.paths:
                    if f.lower().endswith('.pdf'):
                        doc = pdfium.PdfDocument(f)
                        count = len(doc)
                        base = os.path.splitext(os.path.basename(f))[0]
                        doc.close()
                        for i in range(count):
                            tasks.append((f, i, self.dpi, self.out_dir, base, self.fmt, self.is_trans))
                        total_pages += count

                completed = 0
                with concurrent.futures.ProcessPoolExecutor(max_workers=max_workers) as executor:
                    futures = [executor.submit(_worker_pdf2img_save, t) for t in tasks]
                    for f in concurrent.futures.as_completed(futures):
                        completed += 1
                        self.progress.emit(int((completed / total_pages) * 95), f"导出图片中: {completed}/{total_pages}")

            self.progress.emit(100, "处理完成！")
            self.finished.emit(result_message or f"任务执行成功，文件已保存至：\n{self.out_dir}")

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.error.emit(str(e))
        finally:
            try:
                shutil.rmtree(temp_dir)
            except:
                pass
