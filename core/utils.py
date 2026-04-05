import os
import sys
import shutil
import subprocess
import fitz
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QTableWidget, QTableWidgetItem,
                             QPushButton, QHBoxLayout, QHeaderView, QMessageBox, QInputDialog)

# ================= 公共物理换算常量 =================
MM_TO_PTS = 72 / 25.4

# ================= 公共 UI 样式常量 =================
BTN_BLUE = "background-color: #3498DB; color: white; font-weight: bold; padding: 8px; border-radius: 4px;"
BTN_GREEN = "background-color: #2ECC71; color: white; font-weight: bold; padding: 10px; border-radius: 4px;"
BTN_PURPLE = "background-color: #9B59B6; color: white; font-weight: bold; padding: 8px; border-radius: 4px;"
BTN_RED = "background-color: #E74C3C; color: white; font-weight: bold; padding: 10px; border-radius: 4px;"
BTN_GRAY = "background-color: #ECF0F1; color: #2C3E50; font-weight: bold; padding: 6px; border-radius: 4px; border: 1px solid #BDC3C7;"
BTN_ORANGE = "background-color: #E67E22; color: white; font-weight: bold; padding: 8px; border-radius: 4px;"


def get_base_path(relative_path=""):
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def find_ghostscript():
    base_path = get_base_path()
    bundled_gs = os.path.join(base_path, "gs_portable", "bin", "gswin64c.exe")
    bundled_lib = os.path.join(base_path, "gs_portable", "lib")
    if os.path.exists(bundled_gs) and os.path.exists(bundled_lib):
        return bundled_gs, bundled_lib
    possible_paths = [r"C:\Program Files\gs\gs10.04.0\bin\gswin64c.exe",
                      r"D:\Program Files\gs\gs10.04.0\bin\gswin64c.exe"]
    if shutil.which("gswin64c"): return "gswin64c", None
    for p in possible_paths:
        if os.path.exists(p): return p, None
    base_dir = r"C:\Program Files\gs"
    if os.path.exists(base_dir):
        for item in os.listdir(base_dir):
            bin_path = os.path.join(base_dir, item, "bin", "gswin64c.exe")
            if os.path.exists(bin_path): return bin_path, None
    return None, None


def run_ghostscript(gs_path, gs_lib_path, input_pdf, output_pdf, quality="/ebook"):
    """纯净阻塞调用引擎，绝不跨线程传递信号"""
    cmd = [gs_path, "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4", f"-dPDFSETTINGS={quality}",
           "-dNOPAUSE", "-dQUIET", "-dBATCH"]
    if gs_lib_path: cmd.insert(1, f"-I{gs_lib_path}")
    cmd.extend([f"-sOutputFile={output_pdf}", input_pdf])
    subprocess.run(cmd, creationflags=subprocess.CREATE_NO_WINDOW, check=True)


def get_unique_filepath(directory, desired_filename):
    final_path = os.path.join(directory, desired_filename)
    counter = 1
    name_no_ext, ext = os.path.splitext(desired_filename)
    while os.path.exists(final_path):
        final_path = os.path.join(directory, f"{name_no_ext}_{counter}{ext}")
        counter += 1
    return final_path


# ================= 智能书签系统 =================
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


# ================= 原有分段算法 =================
def detect_smart_segments(pdf_doc):
    segments = []
    TOLERANCE = 5.0
    for i in range(len(pdf_doc)):
        page = pdf_doc[i]
        w_mm, h_mm = page.rect.width / MM_TO_PTS, page.rect.height / MM_TO_PTS
        matched = False
        for seg in segments:
            if abs(seg['w_mm'] - w_mm) < TOLERANCE and abs(seg['h_mm'] - h_mm) < TOLERANCE:
                seg['pages'].append(i)
                matched = True
                break
        if not matched:
            segments.append({'pages': [i], 'w_mm': w_mm, 'h_mm': h_mm, 'pos_set': False})
    return segments


def format_page_ranges(pages):
    if not pages: return ""
    pages = sorted(list(set(pages)))
    ranges = []
    start = end = pages[0]
    for p in pages[1:]:
        if p == end + 1:
            end = p
        else:
            ranges.append(f"{start + 1}-{end + 1}" if start != end else f"{start + 1}")
            start = end = p
    ranges.append(f"{start + 1}-{end + 1}" if start != end else f"{start + 1}")
    return ", ".join(ranges)


def parse_page_ranges(text):
    pages = set()
    for part in text.split(','):
        part = part.strip()
        if not part: continue
        if '-' in part:
            s, e = part.split('-')
            pages.update(range(int(s) - 1, int(e)))
        else:
            pages.add(int(part) - 1)
    return sorted(list(pages))


class UniversalSegmentDialog(QDialog):
    def __init__(self, segments, mode_title="智能分段设置", parent=None):
        super().__init__(parent)
        self.setWindowTitle(mode_title)
        self.resize(800, 450)
        self.segments = segments
        self.parent_app = parent
        layout = QVBoxLayout(self)
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["页码范围", "尺寸(mm)", "状态", "操作"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.table)

        btn_layout = QHBoxLayout()
        btn_extract = QPushButton("✂️ 提取选中组的页面成新组 (分离封面)")
        btn_extract.setStyleSheet("background-color: #3498DB; color: white; padding: 6px; border-radius: 4px;")
        btn_extract.clicked.connect(self.extract_pages)
        btn_delete = QPushButton("❌ 删除选中组")
        btn_delete.setStyleSheet("background-color: #E74C3C; color: white; padding: 6px; border-radius: 4px;")
        btn_delete.clicked.connect(self.delete_group)
        btn_layout.addWidget(btn_extract);
        btn_layout.addWidget(btn_delete)
        layout.addLayout(btn_layout)

        btn_confirm = QPushButton("✅ 保存所有分段并返回")
        btn_confirm.setStyleSheet(
            "background-color: #2ECC71; color: white; font-weight: bold; padding: 10px; border-radius: 4px;")
        btn_confirm.clicked.connect(self.accept)
        layout.addWidget(btn_confirm)
        self.refresh_table()

    def refresh_table(self):
        self.table.setRowCount(len(self.segments))
        for i, seg in enumerate(self.segments):
            self.table.setItem(i, 0, QTableWidgetItem(format_page_ranges(seg['pages'])))
            self.table.setItem(i, 1, QTableWidgetItem(f"{int(seg['w_mm'])} x {int(seg['h_mm'])}"))
            status = "✅ 已设置" if seg.get('pos_set') else "❌ 未设置"
            self.table.setItem(i, 2, QTableWidgetItem(status))
            btn = QPushButton("设置位置")
            btn.clicked.connect(lambda checked, idx=i: self.set_position(idx))
            self.table.setCellWidget(i, 3, btn)

    def extract_pages(self):
        row = self.table.currentRow()
        if row < 0: return QMessageBox.warning(self, "提示", "请先选中组！")
        text, ok = QInputDialog.getText(self, "提取页面", "输入页码(如 1):")
        if ok and text:
            try:
                pages_to_extract = parse_page_ranges(text)
                current_pages = set(self.segments[row]['pages'])
                extract_set = set(pages_to_extract)
                if not extract_set.issubset(current_pages): return QMessageBox.warning(self, "错误", "页码不在组内！")
                if extract_set == current_pages: return QMessageBox.warning(self, "错误", "不能提取全部页码！")
                self.segments[row]['pages'] = sorted(list(current_pages - extract_set))
                new_group = {'pages': sorted(list(extract_set)), 'w_mm': self.segments[row]['w_mm'],
                             'h_mm': self.segments[row]['h_mm'], 'pos_set': False}
                if 'pos_pct' in self.segments[row]:
                    if isinstance(self.segments[row]['pos_pct'], dict):
                        new_group['pos_pct'] = {'no': (0.8, 0.9, 150, 40), 'name': (0.8, 0.85, 150, 40)}
                    else:
                        new_group['pos_pct'] = (0.5, 0.5)
                self.segments.append(new_group)
                self.refresh_table()
                QMessageBox.information(self, "成功", "提取成功！")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"输入格式错误: {str(e)}")

    def delete_group(self):
        row = self.table.currentRow()
        if row >= 0:
            del self.segments[row]
            self.refresh_table()

    def set_position(self, idx):
        self.hide()
        self.parent_app.enter_setting_mode_from_dialog(self.segments[idx], idx, self)