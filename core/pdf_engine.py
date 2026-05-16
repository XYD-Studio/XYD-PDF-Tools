# -*- coding: utf-8 -*-
"""
core/pdf_engine.py
统一的 PDF 底层处理引擎与通用线程基类
绝对保证原有逻辑 0 变动
"""
import os
import subprocess
import fitz
from PyQt5.QtCore import QThread, pyqtSignal


# ================= 核心 PDF 与 GS 函数 =================
def run_ghostscript(gs_path, gs_lib_path, input_pdf, output_pdf, quality="/ebook"):
    """纯净阻塞调用引擎，绝不跨线程传递信号"""
    cmd = [gs_path, "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4", f"-dPDFSETTINGS={quality}",
           "-dNOPAUSE", "-dQUIET", "-dBATCH"]
    if gs_lib_path: cmd.insert(1, f"-I{gs_lib_path}")
    cmd.extend([f"-sOutputFile={output_pdf}", input_pdf])
    subprocess.run(cmd, creationflags=subprocess.CREATE_NO_WINDOW, check=True)


def reinject_toc_after_gs(pdf_path, target_toc):
    if not target_toc: return pdf_path
    doc = fitz.open(pdf_path)
    doc.set_toc(target_toc)
    tmp_toc_path = pdf_path + ".toc.tmp.pdf"
    doc.save(tmp_toc_path)
    doc.close()
    os.remove(pdf_path)
    os.rename(tmp_toc_path, pdf_path)
    return pdf_path


def merge_pdf_with_smart_toc(src_doc, filename, merged_doc, merged_toc_list, prefer_filename_for_single=True):
    start_page = len(merged_doc)
    src_toc = src_doc.get_toc(simple=False)
    basename = os.path.splitext(filename)[0]
    page_count = len(src_doc)

    if page_count == 1 and prefer_filename_for_single:
        merged_toc_list.append([1, basename, start_page + 1])
    elif src_toc:
        for item in src_toc:
            merged_toc_list.append([item[0], item[1], item[2] + start_page])
    else:
        if page_count == 1:
            merged_toc_list.append([1, basename, start_page + 1])
        else:
            merged_toc_list.append([1, basename, start_page + 1])
            for i in range(page_count):
                merged_toc_list.append([2, f"{basename}_{i + 1}", start_page + i + 1])
    merged_doc.insert_pdf(src_doc)


def get_sub_toc(full_toc, start_page, end_page):
    sub_toc = []
    for item in full_toc:
        lvl, title, page = item[0], item[1], item[2]
        if (start_page + 1) <= page <= (end_page + 1):
            sub_toc.append([lvl, title, page - start_page])
    return sub_toc


def get_unique_filepath(directory, desired_filename):
    final_path = os.path.join(directory, desired_filename)
    counter = 1
    name_no_ext, ext = os.path.splitext(desired_filename)
    while os.path.exists(final_path):
        final_path = os.path.join(directory, f"{name_no_ext}_{counter}{ext}")
        counter += 1
    return final_path


# ================= 通用的后台处理线程基类 =================
class BaseFakeProgressWorker(QThread):
    """
    专门为需要调用 Ghostscript 且耗时极长的任务提供安全的跑马灯假进度条回调基类。
    """
    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    TRIGGER_FAKE_PROGRESS = -1
    STOP_FAKE_PROGRESS = 99

    def trigger_fake_progress(self):
        self.progress.emit(self.TRIGGER_FAKE_PROGRESS)

    def stop_fake_progress(self):
        self.progress.emit(self.STOP_FAKE_PROGRESS)