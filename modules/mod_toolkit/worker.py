# -*- coding: utf-8 -*-
import os
import shutil
import tempfile
import concurrent.futures
from PyQt5.QtCore import QThread, pyqtSignal
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
import fitz

# 引入核心底层的通用函数
from core.pdf_engine import merge_pdf_with_smart_toc, get_unique_filepath

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

        temp_name = f"temp_{os.path.basename(pdf_path)}_{page_index}.{img_fmt}"
        temp_path = os.path.join(temp_dir, temp_name)
        pil_image.save(temp_path, quality=95)
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

    def __init__(self, paths, mode, out_dir, dpi, fmt, is_trans, split_config=None):
        super().__init__()
        self.paths = paths
        self.mode = mode
        self.out_dir = out_dir
        self.dpi = dpi
        self.fmt = fmt
        self.is_trans = is_trans
        self.split_config = split_config if split_config else {}

    def run(self):
        temp_dir = tempfile.mkdtemp(prefix="pdf_tool_tmp_")
        try:
            max_workers = max(1, os.cpu_count() - 1)

            if "合并" in self.mode:
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
                from PIL import Image
                tasks = []
                total_pages = 0
                for f in self.paths:
                    if f.lower().endswith('.pdf'):
                        doc = pdfium.PdfDocument(f)
                        count = len(doc)
                        doc.close()
                        for i in range(count):
                            tasks.append((f, i, self.dpi, temp_dir, "jpg", False))
                        total_pages += count

                if not tasks: raise Exception("没有找到有效的PDF页面")
                temp_files_map = {}
                completed = 0

                with concurrent.futures.ProcessPoolExecutor(max_workers=max_workers) as executor:
                    futures = {executor.submit(_worker_render_temp_img, task): i for i, task in enumerate(tasks)}
                    for future in concurrent.futures.as_completed(futures):
                        task_idx = futures[future]
                        _, result = future.result()
                        if not isinstance(result, str) or not result.startswith("ERROR"):
                            temp_files_map[task_idx] = result
                        completed += 1
                        self.progress.emit(int((completed / total_pages) * 80), f"多进程渲染中: {completed}/{total_pages}")

                self.progress.emit(90, "正在合并生成最终PDF...")
                valid_temps = [temp_files_map[i] for i in range(len(tasks)) if i in temp_files_map]
                if valid_temps:
                    img_first = Image.open(valid_temps[0])
                    img_others = [Image.open(p) for p in valid_temps[1:]]
                    img_first.save(self.out_dir, "PDF", resolution=self.dpi, save_all=True, append_images=img_others)
                    img_first.close()
                    for img in img_others: img.close()
                else:
                    raise Exception("渲染全部失败，无法生成PDF")

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
            self.finished.emit(f"任务执行成功，文件已保存至：\n{self.out_dir}")

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.error.emit(str(e))
        finally:
            try:
                shutil.rmtree(temp_dir)
            except:
                pass