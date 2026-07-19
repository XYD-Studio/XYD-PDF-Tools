# -*- coding: utf-8 -*-
"""
core/pdf_engine.py
统一的 PDF 底层处理引擎与通用线程基类
绝对保证原有逻辑 0 变动
"""
import os
import re
import subprocess
import fitz
from PyQt5.QtCore import QThread, pyqtSignal


# ================= 核心 PDF 与 GS 函数 =================
def inspect_pdf_color_resources(pdf_path):
    """Return color-resource markers used to guard spot colors during GS conversion."""
    markers = {"separation": False, "devicen": False, "icc": False, "spot_names": set()}
    doc = fitz.open(pdf_path)
    try:
        objects = []
        for xref in range(1, doc.xref_length()):
            try:
                objects.append(doc.xref_object(xref, compressed=False))
            except Exception:
                continue
        source = "\n".join(objects)
    finally:
        doc.close()

    markers["separation"] = "/Separation" in source
    markers["devicen"] = "/DeviceN" in source
    markers["icc"] = "/ICCBased" in source
    for name in re.findall(r"/Separation\s*/([^\s<>\[\]()]+)", source):
        if name not in {"None", "All"}:
            markers["spot_names"].add(name)
    for colorants in re.findall(r"/DeviceN\s*\[(.*?)\]", source, flags=re.DOTALL):
        for name in re.findall(r"/([^\s<>\[\]()]+)", colorants):
            if name not in {"None", "All", "Black", "Cyan", "Magenta", "Yellow"}:
                markers["spot_names"].add(name)
    return markers


def _validate_ghostscript_output(input_pdf, output_pdf):
    if not os.path.isfile(output_pdf) or os.path.getsize(output_pdf) < 8:
        raise RuntimeError("Ghostscript did not create a valid output file")
    try:
        output_doc = fitz.open(output_pdf)
        page_count = len(output_doc)
        output_doc.close()
    except Exception as exc:
        raise RuntimeError(f"Ghostscript output cannot be opened: {exc}") from exc
    if page_count <= 0:
        raise RuntimeError("Ghostscript output contains no pages")
    input_doc = fitz.open(input_pdf)
    input_page_count = len(input_doc)
    input_doc.close()
    if page_count != input_page_count:
        raise RuntimeError(
            f"Ghostscript page-count mismatch: input {input_page_count}, output {page_count}"
        )

    before = inspect_pdf_color_resources(input_pdf)
    if not (before["separation"] or before["devicen"]):
        return
    after = inspect_pdf_color_resources(output_pdf)
    missing_types = []
    if before["separation"] and not after["separation"]:
        missing_types.append("Separation")
    if before["devicen"] and not after["devicen"]:
        missing_types.append("DeviceN")
    missing_names = sorted(before["spot_names"] - after["spot_names"])
    if before["icc"] and not after["icc"]:
        missing_types.append("ICCBased")
    if missing_types or missing_names:
        details = []
        if missing_types:
            details.append("resources: " + ", ".join(missing_types))
        if missing_names:
            details.append("spot names: " + ", ".join(missing_names))
        raise RuntimeError("Spot-color safety check failed; missing " + "; ".join(details))


def run_ghostscript(gs_path, gs_lib_path, input_pdf, output_pdf, quality="/ebook"):
    """Run Ghostscript with Unicode-safe paths and actionable diagnostics."""
    if not gs_path or not os.path.exists(gs_path):
        raise FileNotFoundError(f"Ghostscript executable not found: {gs_path}")
    input_pdf = os.path.abspath(input_pdf)
    output_pdf = os.path.abspath(output_pdf)
    os.makedirs(os.path.dirname(output_pdf), exist_ok=True)
    if os.path.exists(output_pdf):
        os.remove(output_pdf)

    cmd = [gs_path, "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4", f"-dPDFSETTINGS={quality}",
           "-dNOPAUSE", "-dQUIET", "-dBATCH"]
    if quality == "/printer":
        cmd.extend([
            "-sColorConversionStrategy=LeaveColorUnchanged",
            "-dPreserveSeparation=true",
            "-dConvertCMYKImagesToRGB=false",
        ])
    if gs_lib_path:
        cmd.append(f"-I{os.path.abspath(gs_lib_path)}")
    cmd.extend([f"-sOutputFile={output_pdf}", input_pdf])

    env = os.environ.copy()
    if gs_lib_path:
        old_gs_lib = env.get("GS_LIB", "")
        env["GS_LIB"] = os.path.abspath(gs_lib_path) + (os.pathsep + old_gs_lib if old_gs_lib else "")
    creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
    result = subprocess.run(
        cmd,
        cwd=os.path.dirname(os.path.abspath(gs_path)),
        env=env,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        errors="replace",
        creationflags=creationflags,
        check=False,
    )
    if result.returncode != 0:
        detail = (result.stderr or result.stdout or "No diagnostic output").strip()
        if len(detail) > 1800:
            detail = detail[-1800:]
        raise RuntimeError(
            f"Ghostscript failed for '{os.path.basename(input_pdf)}' "
            f"(exit code {result.returncode}):\n{detail}"
        )
    try:
        _validate_ghostscript_output(input_pdf, output_pdf)
    except Exception:
        if os.path.exists(output_pdf):
            os.remove(output_pdf)
        raise
    return output_pdf


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
