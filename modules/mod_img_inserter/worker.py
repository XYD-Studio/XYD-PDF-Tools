# -*- coding: utf-8 -*-
import os
import shutil
import tempfile
import uuid
import threading
import concurrent.futures
import fitz
from core.pdf_engine import (BaseFakeProgressWorker, run_ghostscript, get_unique_filepath,
                             merge_pdf_with_smart_toc, get_sub_toc, reinject_toc_after_gs)
from core.utils import MM_TO_PTS


class ImageInserterWorker(BaseFakeProgressWorker):
    def __init__(self, paired_data, page_configs, export_config, gs_config, output_path, prefer_filename,
                 gs_path=None, gs_lib_path=None):
        super().__init__()
        self.paired_data = paired_data
        self.page_configs = page_configs
        self.export_mode = export_config['mode']
        self.prefix = export_config['prefix']
        self.suffix = export_config['suffix']
        self.use_gs = gs_config['use_gs']
        self.gs_quality = gs_config['quality']
        self.output_path = output_path
        self.prefer_filename = prefer_filename
        self.gs_path = gs_path
        self.gs_lib_path = gs_lib_path

        self.progress_lock = threading.Lock()
        self.completed_tasks = 0
        self.total_tasks = 1

    def _update_progress(self):
        with self.progress_lock:
            self.completed_tasks += 1
            val = int(self.completed_tasks / self.total_tasks * 95)
            self.progress.emit(val)

    def _task_stamp_only(self, pdf_path, img_path, global_page_start, apply_all):
        """线程独立任务：插图"""
        doc = fitz.open(pdf_path)
        for local_idx in range(len(doc)):
            global_page_num = global_page_start + local_idx
            images_list = self.page_configs.get(global_page_num, [])
            page = doc[local_idx]
            for img_info in images_list:
                img_p = img_info.get('path')
                if not img_p or not os.path.exists(img_p): continue
                rect = fitz.Rect(img_info['pdf_x'], img_info['pdf_y'],
                                 img_info['pdf_x'] + img_info['w'] * MM_TO_PTS,
                                 img_info['pdf_y'] + img_info['h'] * MM_TO_PTS)
                try:
                    page.insert_image(rect, filename=img_p, keep_proportion=False)
                except Exception:
                    pass
        tmp_stamped = os.path.join(tempfile.gettempdir(), f"insert_{uuid.uuid4().hex}.pdf")
        doc.save(tmp_stamped)
        doc.close()
        return tmp_stamped

    def _task_finalize_only(self, tmp_visual, final_out_path, target_toc):
        """线程独立任务：GS二次压缩与移动"""
        try:
            current = tmp_visual
            if self.use_gs and self.gs_path:
                tmp_gs = current + ".g.tmp.pdf"
                run_ghostscript(self.gs_path, self.gs_lib_path, current, tmp_gs, self.gs_quality)
                reinject_toc_after_gs(tmp_gs, target_toc)
                current = tmp_gs

            # 💡 [修复核心：绝对安全的跨盘覆盖/转移]
            shutil.copy2(current, final_out_path)

            self._update_progress()
            return True
        finally:
            if current and current != tmp_visual and os.path.exists(current):
                os.remove(current)
            if tmp_visual and os.path.exists(tmp_visual):
                os.remove(tmp_visual)

    def _task_full_batch(self, pdf_path, img_path, global_page_start, apply_all, final_out_path, target_toc):
        tmp_stamped = None
        try:
            tmp_stamped = self._task_stamp_only(pdf_path, img_path, global_page_start, apply_all)
            return self._task_finalize_only(tmp_stamped, final_out_path, target_toc)
        except Exception as e:
            if tmp_stamped and os.path.exists(tmp_stamped): os.remove(tmp_stamped)
            raise e

    def run(self):
        try:
            self.status.emit("🔍 正在分析配对队列...")

            file_infos = []
            global_page_counter = 0
            for pdf_path, img_path in self.paired_data:
                doc = fitz.open(pdf_path)
                count = len(doc)
                file_infos.append({
                    'pdf_path': pdf_path, 'img_path': img_path, 'start': global_page_counter,
                    'basename': os.path.splitext(os.path.basename(pdf_path))[0],
                    'toc': doc.get_toc(simple=False)
                })
                doc.close()
                global_page_counter += count

            max_workers = max(1, (os.cpu_count() or 2) - 1)

            with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:

                # === 合并模式 ===
                if self.export_mode == 'merged':
                    self.status.emit("🚀 [多核并发] 正在批量配对插图...")
                    self.total_tasks = len(file_infos)
                    futures = [executor.submit(self._task_stamp_only, fi['pdf_path'], fi['img_path'], fi['start'], True)
                               for fi in file_infos]

                    stamped_docs = []
                    for f in concurrent.futures.as_completed(futures):
                        stamped_docs.append(f.result())
                        self._update_progress()

                    self.status.emit("🔄 正在顺序拼接长图纸...")
                    export_doc = fitz.Document()
                    merged_toc = []
                    for tmp_stamped, fi in zip(stamped_docs, file_infos):
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
                    self._task_finalize_only(tmp_merged, final_path, merged_toc)
                    self.stop_fake_progress()

                # === 批量独立 / 原位覆盖 ===
                elif self.export_mode in ['batch', 'overwrite']:
                    self.status.emit("🚀 [多核并发] 正在火力全开处理与压缩图纸...")
                    self.total_tasks = len(file_infos)
                    futures = []
                    for fi in file_infos:
                        final_path = fi['pdf_path'] if self.export_mode == 'overwrite' else get_unique_filepath(
                            self.output_path, f"{self.prefix}{fi['basename']}{self.suffix}.pdf")
                        futures.append(
                            executor.submit(self._task_full_batch, fi['pdf_path'], fi['img_path'], fi['start'], True,
                                            final_path, fi['toc']))

                    for f in concurrent.futures.as_completed(futures):
                        f.result()

                        # === 单页拆分 ===
                elif self.export_mode == 'split':
                    self.status.emit("🚀 [多核并发] 正在配对插图...")
                    self.total_tasks = len(file_infos)
                    futures = [executor.submit(self._task_stamp_only, fi['pdf_path'], fi['img_path'], fi['start'], True)
                               for fi in file_infos]

                    self.status.emit("🗜️ [多核并发] 正在拆分并压缩每一页图纸...")
                    finalize_tasks = []
                    for tmp_stamped, fi in zip([f.result() for f in futures], file_infos):
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

                            finalize_tasks.append((tmp_visual, final_path, single_toc))
                        doc.close()
                        os.remove(tmp_stamped)

                    self.total_tasks = len(finalize_tasks)
                    self.completed_tasks = 0
                    fin_futures = [executor.submit(self._task_finalize_only, *t) for t in finalize_tasks]
                    for f in concurrent.futures.as_completed(fin_futures):
                        f.result()

            self.progress.emit(100)
            self.finished.emit("🎉 所有批量配对、并发压缩处理已完美执行完毕！")
        except Exception as e:
            import traceback;
            traceback.print_exc()
            self.error.emit(str(e))