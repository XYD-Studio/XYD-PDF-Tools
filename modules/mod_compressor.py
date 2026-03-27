import os
import shutil
import subprocess
import tempfile
from PIL import Image
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                             QGroupBox, QComboBox, QRadioButton, QLineEdit, QFileDialog, QMessageBox, QProgressBar)
from PyQt5.QtCore import QThread, pyqtSignal
import fitz

from core.ui_components import FileListManagerWidget


def get_base_path():
    import sys
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    return os.path.abspath(".")


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


# ================= 专属后台处理线程 =================
class PDFCompressWorker(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(str)

    def __init__(self, paths, mode, output_path, prefix, suffix, quality, gs_path, gs_lib_path):
        super().__init__()
        self.paths = paths
        self.mode = mode
        self.output_path = output_path
        self.prefix = prefix
        self.suffix = suffix
        self.quality = quality
        self.gs_path = gs_path
        self.gs_lib_path = gs_lib_path

    def run(self):
        try:
            # === 模式 1：合并压缩 ===
            if self.mode == 'merged':
                self.progress.emit(10, "正在进行前置合并处理...")
                temp_dir = tempfile.mkdtemp(prefix="pdf_comp_")
                try:
                    # 先用 fitz 极速合并，防止命令行超长
                    temp_merged = os.path.join(temp_dir, "merged_temp.pdf")
                    doc = fitz.Document()
                    for p in self.paths:
                        src = fitz.open(p)
                        doc.insert_pdf(src)
                        src.close()
                    doc.save(temp_merged)
                    doc.close()

                    self.progress.emit(50, "正在启动引擎进行深度强力压缩...")
                    cmd = [self.gs_path, "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4",
                           f"-dPDFSETTINGS={self.quality}",
                           "-dNOPAUSE", "-dQUIET", "-dBATCH"]
                    if self.gs_lib_path: cmd.insert(1, f"-I{self.gs_lib_path}")
                    cmd.append(f"-sOutputFile={self.output_path}")
                    cmd.append(temp_merged)

                    subprocess.run(cmd, creationflags=subprocess.CREATE_NO_WINDOW, check=True)
                finally:
                    shutil.rmtree(temp_dir, ignore_errors=True)

                self.progress.emit(100, "合并压缩完成！")

            # === 模式 2：批量原样压缩 ===
            elif self.mode == 'batch':
                total = len(self.paths)
                for i, input_path in enumerate(self.paths):
                    self.progress.emit(int(i / total * 100), f"正在压缩: {os.path.basename(input_path)}...")
                    base_name = os.path.splitext(os.path.basename(input_path))[0]
                    final_name = f"{self.prefix}{base_name}{self.suffix}.pdf"
                    out_path = os.path.join(self.output_path, final_name)

                    cmd = [self.gs_path, "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4",
                           f"-dPDFSETTINGS={self.quality}",
                           "-dNOPAUSE", "-dQUIET", "-dBATCH"]
                    if self.gs_lib_path: cmd.insert(1, f"-I{self.gs_lib_path}")
                    cmd.append(f"-sOutputFile={out_path}")
                    cmd.append(input_path)

                    subprocess.run(cmd, creationflags=subprocess.CREATE_NO_WINDOW, check=True)
                self.progress.emit(100, "批量压缩完成！")

            # === 模式 3：压缩并拆分为单页 ===
            elif self.mode == 'split':
                total = len(self.paths)
                temp_dir = tempfile.mkdtemp(prefix="pdf_comp_split_")
                try:
                    for i, input_path in enumerate(self.paths):
                        self.progress.emit(int(i / total * 90), f"正在压缩并拆分: {os.path.basename(input_path)}...")
                        base_name = os.path.splitext(os.path.basename(input_path))[0]

                        # 1. 先把原文件压缩输出到临时文件
                        temp_pdf = os.path.join(temp_dir, f"temp_{i}.pdf")
                        cmd = [self.gs_path, "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4",
                               f"-dPDFSETTINGS={self.quality}",
                               "-dNOPAUSE", "-dQUIET", "-dBATCH"]
                        if self.gs_lib_path: cmd.insert(1, f"-I{self.gs_lib_path}")
                        cmd.append(f"-sOutputFile={temp_pdf}")
                        cmd.append(input_path)
                        subprocess.run(cmd, creationflags=subprocess.CREATE_NO_WINDOW, check=True)

                        # 2. 用 fitz 把压缩好的 PDF 拆成单页
                        doc = fitz.open(temp_pdf)
                        for page_num in range(len(doc)):
                            new_doc = fitz.Document()
                            new_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
                            final_name = f"{self.prefix}{base_name}_第{page_num + 1}页{self.suffix}.pdf"
                            new_doc.save(os.path.join(self.output_path, final_name))
                            new_doc.close()
                        doc.close()
                finally:
                    shutil.rmtree(temp_dir, ignore_errors=True)
                self.progress.emit(100, "拆分压缩完成！")

            self.finished.emit("处理完毕，请查看输出目录！")
        except Exception as e:
            import traceback
            print(traceback.format_exc())
            self.finished.emit(f"发生异常: {str(e)}")


class ImgCompressWorker(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(str)

    def __init__(self, paths, save_dir, target_val, fmt_sel, is_dpi, prefix, suffix):
        super().__init__()
        self.paths = paths
        self.save_dir = save_dir
        self.target_val = target_val
        self.fmt_sel = fmt_sel
        self.is_dpi = is_dpi
        self.prefix = prefix
        self.suffix = suffix

    def run(self):
        total = len(self.paths)
        for i, input_path in enumerate(self.paths):
            self.progress.emit(int(i / total * 100), f"正在转换: {os.path.basename(input_path)}...")
            name, ext = os.path.splitext(os.path.basename(input_path))

            target_ext = ext
            if "JPG" in self.fmt_sel:
                target_ext = ".jpg"
            elif "PNG" in self.fmt_sel:
                target_ext = ".png"
            elif "WEBP" in self.fmt_sel:
                target_ext = ".webp"
            elif "PDF" in self.fmt_sel:
                target_ext = ".pdf"

            # 加上前后缀的自定义逻辑
            final_name = f"{self.prefix}{name}{self.suffix}{target_ext}"
            out_path = os.path.join(self.save_dir, final_name)

            try:
                with Image.open(input_path) as img:
                    orig_dpi = img.info.get('dpi', (96, 96))
                    save_dpi = orig_dpi
                    if self.is_dpi:
                        if self.target_val < orig_dpi[0]:
                            scale = self.target_val / orig_dpi[0]
                            img = img.resize((int(img.width * scale), int(img.height * scale)),
                                             Image.Resampling.LANCZOS)
                        save_dpi = (self.target_val, self.target_val)
                    else:
                        w, h = img.size
                        if max(w, h) > self.target_val:
                            ratio = self.target_val / max(w, h)
                            img = img.resize((int(w * ratio), int(h * ratio)), Image.Resampling.LANCZOS)

                    if target_ext in ['.jpg', '.pdf'] and img.mode in ('RGBA', 'P'):
                        bg = Image.new('RGB', img.size, (255, 255, 255))
                        if img.mode == 'RGBA':
                            bg.paste(img, mask=img.split()[3])
                        else:
                            bg.paste(img.convert('RGB'))
                        img = bg

                    if target_ext == '.jpg':
                        img.save(out_path, quality=80, dpi=save_dpi, optimize=True)
                    elif target_ext == '.pdf':
                        img.save(out_path, "PDF", resolution=save_dpi[0])
                    else:
                        img.save(out_path, dpi=save_dpi)
            except Exception as e:
                print(f"处理图片异常: {e}")

        self.progress.emit(100, "图片批量压缩完成！")
        self.finished.emit("压缩转换成功，请查看输出目录！")


# ================= UI 主体 =================
class CompressorWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.gs_path, self.gs_lib_path = find_ghostscript()
        self.init_ui()

    def init_ui(self):
        layout = QHBoxLayout(self)

        # === PDF 压缩区 ===
        pdf_panel = QVBoxLayout()
        group_pdf = QGroupBox("1. PDF 强力压缩引擎 (Ghostscript)")
        pl = QVBoxLayout()

        if self.gs_lib_path:
            status_text = "✅ 已启用内置绿色极速压缩引擎";
            color = "green"
        elif self.gs_path:
            status_text = "✅ 已检测到系统版 GS 引擎";
            color = "green"
        else:
            status_text = "❌ 未检测到压缩引擎，功能不可用";
            color = "red"

        lbl_gs = QLabel(f"引擎状态: {status_text}")
        lbl_gs.setStyleSheet(f"color: {color}; font-weight: bold;")
        pl.addWidget(lbl_gs)

        hz = QHBoxLayout()
        hz.addWidget(QLabel("压缩质量:"))
        self.cmb_pdf_quality = QComboBox()
        self.cmb_pdf_quality.addItems(
            ["/screen (极致压缩 72dpi)", "/ebook (推荐平衡 150dpi)", "/printer (高清画质 300dpi)"])
        hz.addWidget(self.cmb_pdf_quality)
        pl.addLayout(hz)

        # 【核心】：导出模式单选框
        hz_mode = QHBoxLayout()
        hz_mode.addWidget(QLabel("导出模式:"))
        self.radio_pdf_merged = QRadioButton("合并导出 (单一文件)")
        self.radio_pdf_batch = QRadioButton("批量导出 (默认独立文件)");
        self.radio_pdf_batch.setChecked(True)
        self.radio_pdf_split = QRadioButton("拆分导出 (单页独立)")
        hz_mode.addWidget(self.radio_pdf_merged)
        hz_mode.addWidget(self.radio_pdf_batch)
        hz_mode.addWidget(self.radio_pdf_split)
        pl.addLayout(hz_mode)

        # 【核心】：前后缀输入框
        hz_fix = QHBoxLayout()
        self.input_pdf_prefix = QLineEdit();
        self.input_pdf_prefix.setPlaceholderText("导出的文件前缀 (选填)")
        self.input_pdf_suffix = QLineEdit();
        self.input_pdf_suffix.setPlaceholderText("导出的文件后缀 (选填)")
        hz_fix.addWidget(self.input_pdf_prefix)
        hz_fix.addWidget(self.input_pdf_suffix)
        pl.addLayout(hz_fix)

        self.file_manager_pdf = FileListManagerWidget(accept_exts=['.pdf'], title_desc="PDF Files (*.pdf)")
        pl.addWidget(self.file_manager_pdf)

        btn_pdf = QPushButton("🚀 开始批量处理 PDF")
        btn_pdf.setStyleSheet("background-color: #E91E63; color: white; padding: 10px; font-weight: bold;")
        btn_pdf.clicked.connect(self.run_pdf_compress)
        pl.addWidget(btn_pdf)

        self.prog_pdf = QProgressBar()
        self.lbl_pdf_stat = QLabel("就绪")
        pl.addWidget(self.prog_pdf)
        pl.addWidget(self.lbl_pdf_stat)

        group_pdf.setLayout(pl)
        pdf_panel.addWidget(group_pdf)
        layout.addLayout(pdf_panel, 1)

        # === 图片 压缩区 ===
        img_panel = QVBoxLayout()
        group_img = QGroupBox("2. 图片批量压缩与转换")
        il = QVBoxLayout()

        hz2 = QHBoxLayout()
        self.radio_dpi = QRadioButton("按 DPI (物理比例)");
        self.radio_dpi.setChecked(True)
        self.radio_px = QRadioButton("按 像素长边 (屏幕)")
        hz2.addWidget(self.radio_dpi);
        hz2.addWidget(self.radio_px)
        il.addLayout(hz2)

        hz3 = QHBoxLayout()
        hz3.addWidget(QLabel("目标数值:"))
        self.entry_val = QLineEdit("150")
        self.cmb_img_fmt = QComboBox()
        self.cmb_img_fmt.addItems(["原格式", "JPG", "PNG", "WEBP", "PDF"])
        hz3.addWidget(self.entry_val);
        hz3.addWidget(QLabel("输出格式:"));
        hz3.addWidget(self.cmb_img_fmt)
        il.addLayout(hz3)

        # 【核心】：图片同样加上前后缀支持
        hz_img_fix = QHBoxLayout()
        self.input_img_prefix = QLineEdit();
        self.input_img_prefix.setPlaceholderText("导出的文件前缀 (选填)")
        self.input_img_suffix = QLineEdit();
        self.input_img_suffix.setPlaceholderText("导出的文件后缀 (选填)")
        hz_img_fix.addWidget(self.input_img_prefix)
        hz_img_fix.addWidget(self.input_img_suffix)
        il.addLayout(hz_img_fix)

        self.file_manager_img = FileListManagerWidget(accept_exts=['.jpg', '.png', '.jpeg', '.webp', '.bmp'],
                                                      title_desc="Images")
        il.addWidget(self.file_manager_img)

        btn_img = QPushButton("🎨 开始批量转换压缩图片")
        btn_img.setStyleSheet("background-color: #00BCD4; color: white; padding: 10px; font-weight: bold;")
        btn_img.clicked.connect(self.run_img_compress)
        il.addWidget(btn_img)

        self.prog_img = QProgressBar()
        self.lbl_img_stat = QLabel("就绪")
        il.addWidget(self.prog_img)
        il.addWidget(self.lbl_img_stat)

        group_img.setLayout(il)
        img_panel.addWidget(group_img)
        layout.addLayout(img_panel, 1)

    def run_pdf_compress(self):
        if not self.gs_path: return QMessageBox.warning(self, "错误", "缺少压缩引擎模块。")
        paths = self.file_manager_pdf.get_all_filepaths()
        if not paths: return QMessageBox.warning(self, "提示", "请先添加要压缩的 PDF 文件。")

        mode = 'batch'
        if self.radio_pdf_merged.isChecked():
            mode = 'merged'
        elif self.radio_pdf_split.isChecked():
            mode = 'split'

        if mode == 'merged':
            save_path, _ = QFileDialog.getSaveFileName(self, "保存合并文件", "合并压缩文件.pdf", "PDF (*.pdf)")
            if not save_path: return
        else:
            save_path = QFileDialog.getExistingDirectory(self, "选择保存目录")
            if not save_path: return

        quality = self.cmb_pdf_quality.currentText().split(" ")[0]
        prefix = self.input_pdf_prefix.text()
        suffix = self.input_pdf_suffix.text()

        self.pdf_worker = PDFCompressWorker(paths, mode, save_path, prefix, suffix, quality, self.gs_path,
                                            self.gs_lib_path)
        self.pdf_worker.progress.connect(lambda v, txt: (self.prog_pdf.setValue(v), self.lbl_pdf_stat.setText(txt)))
        self.pdf_worker.finished.connect(lambda msg: QMessageBox.information(self, "完成", msg))
        self.pdf_worker.start()

    def run_img_compress(self):
        paths = self.file_manager_img.get_all_filepaths()
        if not paths: return QMessageBox.warning(self, "提示", "请先添加图片文件。")
        save_dir = QFileDialog.getExistingDirectory(self, "选择保存目录")
        if not save_dir: return

        target_val = int(self.entry_val.text())
        fmt_sel = self.cmb_img_fmt.currentText()
        is_dpi = self.radio_dpi.isChecked()
        prefix = self.input_img_prefix.text()
        suffix = self.input_img_suffix.text()

        self.img_worker = ImgCompressWorker(paths, save_dir, target_val, fmt_sel, is_dpi, prefix, suffix)
        self.img_worker.progress.connect(lambda v, txt: (self.prog_img.setValue(v), self.lbl_img_stat.setText(txt)))
        self.img_worker.finished.connect(lambda msg: QMessageBox.information(self, "完成", msg))
        self.img_worker.start()