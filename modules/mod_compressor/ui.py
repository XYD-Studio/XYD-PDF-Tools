# -*- coding: utf-8 -*-
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                             QGroupBox, QComboBox, QRadioButton, QLineEdit, QFileDialog, QMessageBox, QProgressBar)

from core.ui_components import FileListManagerWidget, ExportSettingsPanel, GSSettingsPanel
from core.utils import find_ghostscript
# 引入独立的 Worker
from .worker import PDFCompressWorker, ImgCompressWorker


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
            status_text = "✅ 已启用内置绿色极速压缩引擎"
            color = "green"
        elif self.gs_path:
            status_text = "✅ 已检测到系统版 GS 引擎"
            color = "green"
        else:
            status_text = "❌ 未检测到压缩引擎，功能不可用"
            color = "red"

        lbl_gs = QLabel(f"引擎状态: {status_text}")
        lbl_gs.setStyleSheet(f"color: {color}; font-weight: bold;")
        pl.addWidget(lbl_gs)

        self.gs_panel = GSSettingsPanel(bool(self.gs_path))
        self.gs_panel.chk_gs.hide()  # 纯压缩工具不需要开关，强制启用
        pl.addWidget(self.gs_panel)

        self.export_panel = ExportSettingsPanel("导出模式设置:", default_mode='batch')
        pl.addWidget(self.export_panel)

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

        # === 图片压缩区 ===
        img_panel = QVBoxLayout()
        group_img = QGroupBox("2. 图片批量压缩与转换")
        il = QVBoxLayout()

        hz2 = QHBoxLayout()
        self.radio_dpi = QRadioButton("按 DPI (物理比例)")
        self.radio_dpi.setChecked(True)
        self.radio_px = QRadioButton("按 像素长边 (屏幕)")
        hz2.addWidget(self.radio_dpi)
        hz2.addWidget(self.radio_px)
        il.addLayout(hz2)

        hz3 = QHBoxLayout()
        hz3.addWidget(QLabel("目标数值:"))
        self.entry_val = QLineEdit("150")
        self.cmb_img_fmt = QComboBox()
        self.cmb_img_fmt.addItems(["原格式", "JPG", "PNG", "WEBP", "PDF"])
        hz3.addWidget(self.entry_val)
        hz3.addWidget(QLabel("输出格式:"))
        hz3.addWidget(self.cmb_img_fmt)
        il.addLayout(hz3)

        hz_img_fix = QHBoxLayout()
        self.input_img_prefix = QLineEdit()
        self.input_img_prefix.setPlaceholderText("导出的文件前缀 (选填)")
        self.input_img_suffix = QLineEdit()
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

        export_cfg = self.export_panel.get_config()
        gs_cfg = self.gs_panel.get_config()

        if export_cfg['mode'] == 'merged':
            save_path, _ = QFileDialog.getSaveFileName(self, "保存合并文件", "合并压缩文件.pdf", "PDF (*.pdf)")
            if not save_path: return
        else:
            save_path = QFileDialog.getExistingDirectory(self, "选择保存目录")
            if not save_path: return

        self.pdf_worker = PDFCompressWorker(paths, export_cfg, gs_cfg, save_path, self.gs_path, self.gs_lib_path)
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