# -*- coding: utf-8 -*-
import os
import io
import math
import shutil
import tempfile
import uuid
import threading
import concurrent.futures
import fitz
from PIL import Image
from core.pdf_engine import (BaseFakeProgressWorker, run_ghostscript, merge_pdf_with_smart_toc,
                             get_sub_toc, reinject_toc_after_gs, get_unique_filepath)
from core.utils import MM_TO_PTS

try:
    from pyhanko.sign import signers
    from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
    from pyhanko.sign.fields import SigFieldSpec, append_signature_field
    from cryptography.hazmat.primitives.serialization import pkcs12
    from cryptography.hazmat.primitives import serialization

    PYHANKO_AVAILABLE = True
except ImportError:
    PYHANKO_AVAILABLE = False


class StamperWorker(BaseFakeProgressWorker):
    def __init__(self, file_paths, page_positions, export_config, gs_config, output_path, prefer_filename,
                 pre_flatten=False, clean_annots=False, stamp_dpi=300, gs_path=None, gs_lib_path=None,
                 use_pki=False, pfx_path="", pfx_pwd="", pki_target_id="first"):
        super().__init__()
        self.file_paths = file_paths
        self.page_positions = page_positions
        self.export_mode = export_config['mode']
        self.prefix = export_config['prefix']
        self.suffix = export_config['suffix']
        self.use_gs_compress = gs_config['use_gs']
        self.gs_quality = gs_config['quality']
        self.output_path = output_path
        self.prefer_filename = prefer_filename
        self.pre_flatten = pre_flatten
        self.clean_annots = clean_annots
        self.stamp_dpi = stamp_dpi
        self.gs_path = gs_path
        self.gs_lib_path = gs_lib_path
        self.use_pki = use_pki
        self.pfx_path = pfx_path
        self.pfx_pwd = pfx_pwd
        self.pki_target_id = pki_target_id

        self._image_cache = {}
        self.cache_lock = threading.Lock()
        self.progress_lock = threading.Lock()
        self.completed_tasks = 0
        self.total_tasks = 1

    def _update_progress(self):
        with self.progress_lock:
            self.completed_tasks += 1
            val = int(self.completed_tasks / self.total_tasks * 95)
            self.progress.emit(val)

    def _get_processed_stamp_bytes(self, stamp_path, w_mm, h_mm, effective_angle):
        cache_key = (stamp_path, w_mm, h_mm, effective_angle)
        with self.cache_lock:
            if cache_key in self._image_cache: return self._image_cache[cache_key]

        target_px_w = int((w_mm / 25.4) * self.stamp_dpi)
        target_px_h = int((h_mm / 25.4) * self.stamp_dpi)

        with Image.open(stamp_path) as img:
            img = img.convert("RGBA")
            resample_filter = getattr(Image, 'Resampling', Image).LANCZOS
            img = img.resize((target_px_w, target_px_h), resample=resample_filter)
            if effective_angle != 0:
                img = img.rotate(-effective_angle, expand=True, resample=resample_filter)
            stream = io.BytesIO()
            img.save(stream, format="PNG", optimize=True)
            img_bytes = stream.getvalue()

        with self.cache_lock:
            self._image_cache[cache_key] = img_bytes
        return img_bytes

    def _apply_all_stamps_to_page(self, page, global_page_num):
        stamps_list = self.page_positions.get(global_page_num, [])
        page_rot = page.rotation
        pki_info = None

        for stamp in stamps_list:
            if not os.path.exists(stamp['path']): continue
            w_mm, h_mm = stamp['w'], stamp['h']
            w_pts, h_pts = w_mm * MM_TO_PTS, h_mm * MM_TO_PTS
            pdf_x, pdf_y = stamp['pdf_x'], stamp['pdf_y']

            visual_angle = stamp.get('angle', 0) % 360
            phys_angle = (visual_angle - page_rot) % 360
            img_bytes = self._get_processed_stamp_bytes(stamp['path'], w_mm, h_mm, phys_angle)

            cx = pdf_x + w_pts / 2
            cy = pdf_y + h_pts / 2
            phys_center = fitz.Point(cx, cy) * page.derotation_matrix

            rad = math.radians(phys_angle)
            new_w_pts = abs(w_pts * math.cos(rad)) + abs(h_pts * math.sin(rad))
            new_h_pts = abs(w_pts * math.sin(rad)) + abs(h_pts * math.cos(rad))

            a_rect = fitz.Rect(phys_center.x - new_w_pts / 2, phys_center.y - new_h_pts / 2,
                               phys_center.x + new_w_pts / 2, phys_center.y + new_h_pts / 2)

            try:
                page.insert_image(a_rect, stream=img_bytes, keep_proportion=False)
            except TypeError:
                page.insert_image(a_rect, stream=img_bytes)

            is_target = False
            if self.pki_target_id == "first" and pki_info is None:
                is_target = True
            elif self.pki_target_id == stamp.get('id') and pki_info is None:
                is_target = True

            if self.use_pki and is_target:
                pki_info = {'page_idx': global_page_num, 'rect': a_rect, 'page_height': page.rect.height}

        return pki_info

    def _apply_pki_signature(self, input_pdf, output_pdf, pki_target_info):
        if not PYHANKO_AVAILABLE: shutil.copy2(input_pdf, output_pdf); return
        try:
            with open(self.pfx_path, "rb") as f:
                pfx_bytes = f.read()
            private_key, cert, _ = pkcs12.load_key_and_certificates(pfx_bytes, self.pfx_pwd.encode('utf-8'))
        except Exception as e:
            raise ValueError(f"证书提取失败: {e}")

        with tempfile.TemporaryDirectory() as tmpdir:
            key_path = os.path.join(tmpdir, "tmp_key.pem")
            cert_path = os.path.join(tmpdir, "tmp_cert.pem")
            with open(key_path, "wb") as f:
                f.write(private_key.private_bytes(encoding=serialization.Encoding.PEM,
                                                  format=serialization.PrivateFormat.PKCS8,
                                                  encryption_algorithm=serialization.NoEncryption()))
            with open(cert_path, "wb") as f:
                f.write(cert.public_bytes(serialization.Encoding.PEM))
            signer = signers.SimpleSigner.load(key_file=key_path, cert_file=cert_path, key_passphrase=None)

        shutil.copy2(input_pdf, output_pdf)
        with open(output_pdf, 'r+b') as f:
            w = IncrementalPdfFileWriter(f)
            if pki_target_info:
                target = pki_target_info
                box = (target['rect'].x0, target['page_height'] - target['rect'].y1, target['rect'].x1,
                       target['page_height'] - target['rect'].y0)
                sig_field = SigFieldSpec('DocumentSecurityLock', on_page=target['page_idx'], box=box)
                append_signature_field(w, sig_field)
            meta = signers.PdfSignatureMetadata(field_name='DocumentSecurityLock', location="System",
                                                reason="文档防篡改")
            try:
                from pyhanko.stamp import TextStampStyle
                style = TextStampStyle(stamp_text=' ', border_width=0)
                pdf_signer = signers.PdfSigner(signature_meta=meta, signer=signer, stamp_style=style)
                pdf_signer.sign_pdf(w, in_place=True)
            except Exception:
                signers.sign_pdf(w, meta, signer=signer, in_place=True,
                                 appearance_text_params={'stamp_text': ' ', 'border_width': 0})

    def _task_prep_and_stamp(self, pdf_path, global_page_start):
        """线程独立任务：清理、预拍平、盖章"""
        working_pdf = pdf_path
        temps = []
        try:
            if self.clean_annots:
                doc = fitz.open(working_pdf)
                for page in doc:
                    for w in page.widgets(): page.delete_widget(w)
                    for a in page.annots(): page.delete_annot(a)
                tmp = os.path.join(tempfile.gettempdir(), f"clean_{uuid.uuid4().hex}.pdf")
                doc.save(tmp)
                doc.close()
                working_pdf = tmp
                temps.append(tmp)

            if self.pre_flatten and self.gs_path:
                tmp = os.path.join(tempfile.gettempdir(), f"flat_{uuid.uuid4().hex}.pdf")
                run_ghostscript(self.gs_path, self.gs_lib_path, working_pdf, tmp, quality="/printer")
                working_pdf = tmp
                temps.append(tmp)

            doc = fitz.open(working_pdf)
            pki_infos = []
            for i in range(len(doc)):
                info = self._apply_all_stamps_to_page(doc[i], global_page_start + i)
                if info: pki_infos.append(info)

            tmp_stamped = os.path.join(tempfile.gettempdir(), f"stamped_{uuid.uuid4().hex}.pdf")
            toc = doc.get_toc(simple=False)
            doc.save(tmp_stamped)
            doc.close()
            return tmp_stamped, pki_infos, toc

        finally:
            for t in temps:
                if os.path.exists(t): os.remove(t)

    def _task_finalize_only(self, tmp_visual, final_out_path, target_toc, pki_info):
        """线程独立任务：GS二次压缩、PKI签名、跨盘覆盖/移动"""
        try:
            current = tmp_visual
            if self.use_gs_compress and self.gs_path:
                tmp_gs = current + ".g.tmp.pdf"
                run_ghostscript(self.gs_path, self.gs_lib_path, current, tmp_gs, self.gs_quality)
                reinject_toc_after_gs(tmp_gs, target_toc)
                current = tmp_gs

            if self.use_pki and PYHANKO_AVAILABLE:
                tmp_pki = current + ".pki.tmp.pdf"
                self._apply_pki_signature(current, tmp_pki, pki_info)
                if current != tmp_visual: os.remove(current)
                current = tmp_pki

            # 💡 [修复核心：绝对安全的跨盘覆盖/转移]
            shutil.copy2(current, final_out_path)

            self._update_progress()
            return True
        finally:
            # 清理本方法的上游临时文件 current (有可能是 GS 压好的或 PKI 签好的临时文件)
            if current and current != tmp_visual and os.path.exists(current):
                os.remove(current)
            # 清理传进来的源临时文件
            if tmp_visual and os.path.exists(tmp_visual):
                os.remove(tmp_visual)

    def _task_full_batch(self, pdf_path, global_page_start, final_out_path):
        """完整管线：用于 Batch 和 Overwrite 模式"""
        tmp_stamped = None
        try:
            tmp_stamped, pki_infos, toc = self._task_prep_and_stamp(pdf_path, global_page_start)
            pki_info = pki_infos[0] if pki_infos else None
            return self._task_finalize_only(tmp_stamped, final_out_path, toc, pki_info)
        except Exception as e:
            if tmp_stamped and os.path.exists(tmp_stamped): os.remove(tmp_stamped)
            raise e

    def run(self):
        try:
            self._image_cache.clear()
            self.status.emit("🔍 正在分析图纸页码与队列分配...")

            file_infos = []
            global_page_counter = 0
            for path in self.file_paths:
                doc = fitz.open(path)
                count = len(doc)
                file_infos.append({
                    'path': path, 'start': global_page_counter, 'count': count,
                    'basename': os.path.splitext(os.path.basename(path))[0],
                    'toc': doc.get_toc(simple=False)
                })
                doc.close()
                global_page_counter += count

            max_workers = max(1, (os.cpu_count() or 2) - 1)

            with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:

                # === 合并模式 ===
                if self.export_mode == 'merged':
                    self.status.emit("🚀 [多核并发] 正在进行图章渲染与预拍平...")
                    self.total_tasks = len(file_infos)
                    futures = [executor.submit(self._task_prep_and_stamp, fi['path'], fi['start']) for fi in file_infos]

                    stamped_docs = []
                    for f in concurrent.futures.as_completed(futures):
                        stamped_docs.append(f.result())
                        self._update_progress()

                    self.status.emit("🔄 正在顺序拼接主文件...")
                    export_doc = fitz.Document()
                    merged_toc = []
                    for (tmp_stamped, _, _), fi in zip(stamped_docs, file_infos):
                        doc = fitz.open(tmp_stamped)
                        merge_pdf_with_smart_toc(doc, fi['basename'], export_doc, merged_toc, self.prefer_filename)
                        doc.close()
                        os.remove(tmp_stamped)

                    tmp_merged = os.path.join(tempfile.gettempdir(), f"merged_{uuid.uuid4().hex}.pdf")
                    export_doc.set_toc(merged_toc)
                    export_doc.save(tmp_merged)
                    export_doc.close()

                    self.status.emit("🗜️ 正在执行最终的全局二次压缩...")
                    self.trigger_fake_progress()
                    final_path = get_unique_filepath(os.path.dirname(self.output_path),
                                                     os.path.basename(self.output_path))
                    self._task_finalize_only(tmp_merged, final_path, merged_toc, None)
                    self.stop_fake_progress()

                # === 批量独立 / 原位覆盖 ===
                elif self.export_mode in ['batch', 'overwrite']:
                    self.status.emit("🚀 [多核并发] 正在满载压缩与导出图纸队列...")
                    self.total_tasks = len(file_infos)
                    futures = []
                    for fi in file_infos:
                        final_path = fi['path'] if self.export_mode == 'overwrite' else get_unique_filepath(
                            self.output_path, f"{self.prefix}{fi['basename']}{self.suffix}.pdf")
                        futures.append(executor.submit(self._task_full_batch, fi['path'], fi['start'], final_path))

                    for f in concurrent.futures.as_completed(futures):
                        f.result()

                        # === 单页拆分 ===
                elif self.export_mode == 'split':
                    self.status.emit("🚀 [多核并发] 正在渲染图章并准备拆分页面...")
                    self.total_tasks = len(file_infos)
                    futures = [executor.submit(self._task_prep_and_stamp, fi['path'], fi['start']) for fi in file_infos]

                    self.status.emit("🗜️ [多核并发] 正在并行压缩每一页图纸...")
                    finalize_tasks = []
                    for (tmp_stamped, pki_infos, _), fi in zip([f.result() for f in futures], file_infos):
                        doc = fitz.open(tmp_stamped)
                        for i in range(len(doc)):
                            single_doc = fitz.Document()
                            single_doc.insert_pdf(doc, from_page=i, to_page=i)
                            tmp_visual = os.path.join(tempfile.gettempdir(), f"split_{uuid.uuid4().hex}.pdf")
                            single_doc.save(tmp_visual)
                            single_doc.close()

                            single_toc = [[1, fi['basename'], 1]] if self.prefer_filename else get_sub_toc(fi['toc'], i,
                                                                                                           i)
                            final_name = f"{self.prefix}{fi['basename']}_第{i + 1}页{self.suffix}.pdf"
                            final_path = get_unique_filepath(self.output_path, final_name)
                            page_pki = next((p for p in pki_infos if p['page_idx'] == fi['start'] + i), None)
                            if page_pki: page_pki['page_idx'] = 0

                            finalize_tasks.append((tmp_visual, final_path, single_toc, page_pki))
                        doc.close()
                        os.remove(tmp_stamped)

                    self.total_tasks = len(finalize_tasks)
                    self.completed_tasks = 0
                    fin_futures = [executor.submit(self._task_finalize_only, *t) for t in finalize_tasks]
                    for f in concurrent.futures.as_completed(fin_futures):
                        f.result()

            self._image_cache.clear()
            self.progress.emit(100)
            self.finished.emit("🎉 所有加盖、并发压缩、防伪签名操作已完美执行完毕！")
        except Exception as e:
            import traceback;
            traceback.print_exc()
            self.error.emit(str(e))