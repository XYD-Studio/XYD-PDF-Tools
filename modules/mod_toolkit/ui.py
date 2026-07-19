# -*- coding: utf-8 -*-
import os
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                             QGroupBox, QComboBox, QFileDialog, QMessageBox, QProgressBar, QCheckBox, QLineEdit,
                             QDialog, QRadioButton, QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView,
                             QDialogButtonBox)
from core.ui_components import FileListManagerWidget
from core.utils import first_input_directory, suggested_output_path
from .worker import ToolkitWorker


# ================= 书签拆分高级配置弹窗 =================
class BookmarkSplitConfigDialog(QDialog):
    def __init__(self, current_config, parent=None):
        super().__init__(parent)
        self.setWindowTitle("书签拆分命名规则配置")
        self.resize(550, 400)
        layout = QVBoxLayout(self)

        box_base = QGroupBox("1. 基础命名规则")
        l_base = QVBoxLayout()
        self.rb_name = QRadioButton("按【书签名称】命名")
        self.rb_seq_name = QRadioButton("按【序号 + 书签名称】命名")

        if current_config.get('mode', 1) == 2:
            self.rb_seq_name.setChecked(True)
        else:
            self.rb_name.setChecked(True)

        l_base.addWidget(self.rb_name)
        l_base.addWidget(self.rb_seq_name)
        box_base.setLayout(l_base)
        layout.addWidget(box_base)

        box_seg = QGroupBox("2. 分段追加前后缀 (可选，不设则不追加)")
        l_seg = QVBoxLayout()
        hz_btn = QHBoxLayout()
        btn_add = QPushButton("➕ 添加分段规则")
        btn_del = QPushButton("❌ 删除选中行")
        btn_add.clicked.connect(self.add_row)
        btn_del.clicked.connect(self.del_row)
        hz_btn.addWidget(btn_add);
        hz_btn.addWidget(btn_del);
        hz_btn.addStretch()
        l_seg.addLayout(hz_btn)

        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["起始页码", "结束页码", "追加前缀", "追加后缀"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        l_seg.addWidget(self.table)
        box_seg.setLayout(l_seg)
        layout.addWidget(box_seg)

        for seg in current_config.get('segments', []):
            self.add_row_data(seg)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def add_row(self):
        self.add_row_data([1, 10, "前缀_", "_后缀"])

    def add_row_data(self, data):
        r = self.table.rowCount()
        self.table.insertRow(r)
        self.table.setItem(r, 0, QTableWidgetItem(str(data[0])))
        self.table.setItem(r, 1, QTableWidgetItem(str(data[1])))
        self.table.setItem(r, 2, QTableWidgetItem(str(data[2])))
        self.table.setItem(r, 3, QTableWidgetItem(str(data[3])))

    def del_row(self):
        r = self.table.currentRow()
        if r >= 0: self.table.removeRow(r)

    def get_config(self):
        mode = 2 if self.rb_seq_name.isChecked() else 1
        segments = []
        for r in range(self.table.rowCount()):
            try:
                start = int(self.table.item(r, 0).text())
                end = int(self.table.item(r, 1).text())
                pfx = self.table.item(r, 2).text()
                sfx = self.table.item(r, 3).text()
                segments.append([start, end, pfx, sfx])
            except ValueError:
                pass
        return {'mode': mode, 'segments': segments}


# ================= UI 主类 =================
class ToolkitWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.split_bookmark_config = {}
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)

        box_mode = QGroupBox("1. 选择工具模式")
        hl = QHBoxLayout()
        self.cmb_mode = QComboBox()
        self.cmb_mode.addItems([
            "多图转PDF", "多个PDF合并", "PDF交叉合并", "PDF拆分为单页",
            "按书签拆分PDF为单页", "PDF转图片型PDF", "PDF批量导出图片"
        ])
        self.cmb_mode.currentTextChanged.connect(self.update_ui_state)
        hl.addWidget(QLabel("当前模式:"));
        hl.addWidget(self.cmb_mode)

        self.btn_config_split = QPushButton("⚙️ 配置拆分规则")
        self.btn_config_split.setStyleSheet("background-color: #f39c12; color: white; font-weight: bold;")
        self.btn_config_split.clicked.connect(self.open_split_config)
        self.btn_config_split.hide()
        hl.addWidget(self.btn_config_split)

        box_mode.setLayout(hl)
        layout.addWidget(box_mode)

        self.box_params = QGroupBox("2. 转换参数设置 (仅涉图处理时可用)")

        # 💡 [核心新增]：注入专属的禁用状态样式表，覆盖掉全局颜色的干扰
        self.box_params.setStyleSheet("""
            QGroupBox:disabled { color: #A0A0A0; border-color: #E0E0E0; }
            QLabel:disabled { color: #A0A0A0; }
            QLineEdit:disabled { background-color: #F5F6FA; color: #A0A0A0; border: 1px solid #E0E0E0; }
            QComboBox:disabled { background-color: #F5F6FA; color: #A0A0A0; border: 1px solid #E0E0E0; }
            QCheckBox:disabled { color: #A0A0A0; }
        """)

        pl = QHBoxLayout()
        pl.addWidget(QLabel("DPI (清晰度):"))
        self.entry_dpi = QLineEdit("200")
        self.entry_dpi.setFixedWidth(60)
        pl.addWidget(self.entry_dpi)

        pl.addWidget(QLabel("图片格式:"))
        self.cmb_fmt = QComboBox()
        self.cmb_fmt.addItems(["jpg", "png"])
        self.cmb_fmt.currentTextChanged.connect(self.update_ui_state)
        pl.addWidget(self.cmb_fmt)

        self.chk_trans = QCheckBox("保持背景透明 (仅PNG)")
        pl.addWidget(self.chk_trans);
        pl.addStretch(1)
        self.box_params.setLayout(pl)
        layout.addWidget(self.box_params)

        self.box_interleave = QGroupBox("2.1 交叉合并参数")
        il = QVBoxLayout()

        row_rotate = QHBoxLayout()
        row_rotate.addWidget(QLabel("A PDF 批量旋转:"))
        self.cmb_rotate_a = QComboBox()
        self.cmb_rotate_a.addItems(["0°", "90°", "180°", "270°"])
        row_rotate.addWidget(self.cmb_rotate_a)

        row_rotate.addWidget(QLabel("B PDF 批量旋转:"))
        self.cmb_rotate_b = QComboBox()
        self.cmb_rotate_b.addItems(["0°", "90°", "180°", "270°"])
        row_rotate.addWidget(self.cmb_rotate_b)

        self.chk_reverse_b = QCheckBox("B PDF 倒序参与交叉")
        row_rotate.addWidget(self.chk_reverse_b)
        row_rotate.addStretch(1)
        il.addLayout(row_rotate)

        row_blank = QHBoxLayout()
        self.chk_fill_blank = QCheckBox("页数不一致时自动补空白页")
        self.chk_fill_blank.setChecked(True)
        row_blank.addWidget(self.chk_fill_blank)

        row_blank.addWidget(QLabel("A侧预留空白页位:"))
        self.entry_reserved_a = QLineEdit()
        self.entry_reserved_a.setPlaceholderText("例如: 3, 8-10")
        row_blank.addWidget(self.entry_reserved_a)

        row_blank.addWidget(QLabel("B侧预留空白页位:"))
        self.entry_reserved_b = QLineEdit()
        self.entry_reserved_b.setPlaceholderText("例如: 2, 6")
        row_blank.addWidget(self.entry_reserved_b)
        il.addLayout(row_blank)

        tip = QLabel("文件列表中第1个PDF作为A，第2个PDF作为B；输出顺序为 A1、B1、A2、B2。预留页位会插入空白页且不消耗原PDF页面。")
        tip.setStyleSheet("color: #7F8C8D;")
        il.addWidget(tip)

        self.box_interleave.setLayout(il)
        layout.addWidget(self.box_interleave)

        self.file_manager = FileListManagerWidget(accept_exts=['.pdf', '.jpg', '.png', '.jpeg'], title_desc="Files")
        layout.addWidget(self.file_manager)

        hz_run = QHBoxLayout()
        self.lbl_status = QLabel("就绪")
        self.progress = QProgressBar()

        self.btn_run = QPushButton("⚡ 开始执行任务")
        self.btn_run.setStyleSheet(
            "background-color: #673AB7; color: white; padding: 12px; font-weight: bold; font-size: 14px; border-radius: 4px;")
        self.btn_run.clicked.connect(self.run_tool)

        hz_run.addWidget(self.lbl_status);
        hz_run.addWidget(self.progress);
        hz_run.addWidget(self.btn_run)
        layout.addLayout(hz_run)

        self.update_ui_state()

    def update_ui_state(self):
        mode = self.cmb_mode.currentText()
        fmt = self.cmb_fmt.currentText()

        self.btn_config_split.setVisible(mode == "按书签拆分PDF为单页")
        self.box_interleave.setVisible(mode == "PDF交叉合并")

        if mode == "PDF批量导出图片":
            self.box_params.setEnabled(True)
            self.entry_dpi.setEnabled(True)
            self.cmb_fmt.setEnabled(True)

            if fmt == "png":
                self.chk_trans.setEnabled(True)
            else:
                self.chk_trans.setEnabled(False)
                self.chk_trans.setChecked(False)

        elif mode == "PDF转图片型PDF":
            self.box_params.setEnabled(True)
            self.entry_dpi.setEnabled(True)

            self.cmb_fmt.setEnabled(False)
            self.chk_trans.setEnabled(False)
            self.chk_trans.setChecked(False)

        elif mode == "PDF交叉合并":
            self.box_params.setEnabled(False)
            self.entry_dpi.setEnabled(False)
            self.cmb_fmt.setEnabled(False)
            self.chk_trans.setEnabled(False)
            self.chk_trans.setChecked(False)

        else:
            self.box_params.setEnabled(False)
            self.entry_dpi.setEnabled(False)
            self.cmb_fmt.setEnabled(False)
            self.chk_trans.setEnabled(False)
            self.chk_trans.setChecked(False)

    def open_split_config(self):
        dialog = BookmarkSplitConfigDialog(self.split_bookmark_config, self)
        if dialog.exec_() == QDialog.Accepted:
            self.split_bookmark_config = dialog.get_config()

    def get_interleave_config(self):
        return {
            'rotate_a': int(self.cmb_rotate_a.currentText().replace("°", "")),
            'rotate_b': int(self.cmb_rotate_b.currentText().replace("°", "")),
            'reverse_b': self.chk_reverse_b.isChecked(),
            'fill_blank': self.chk_fill_blank.isChecked(),
            'reserved_a': self.entry_reserved_a.text(),
            'reserved_b': self.entry_reserved_b.text()
        }

    def run_tool(self):
        if self.file_manager.count() == 0: return QMessageBox.warning(self, "提示", "请先添加文件")
        mode = self.cmb_mode.currentText()
        paths = self.file_manager.get_all_filepaths()
        pdf_paths = [p for p in paths if p.lower().endswith('.pdf')]

        if mode == "PDF转图片型PDF" and len(pdf_paths) > 1:
            out_dir = QFileDialog.getExistingDirectory(self, "选择保存目录", first_input_directory(pdf_paths))
            if not out_dir: return
        elif "拆分" in mode or "导出图片" in mode:
            out_dir = QFileDialog.getExistingDirectory(self, "选择保存目录", first_input_directory(paths))
            if not out_dir: return
        else:
            if mode == "PDF转图片型PDF" and pdf_paths:
                base = os.path.splitext(os.path.basename(pdf_paths[0]))[0]
                filename = f"{base}_图片型.pdf"
            else:
                filename = "工具箱输出.pdf"
            initial = suggested_output_path(paths, filename)
            out_dir, _ = QFileDialog.getSaveFileName(self, "保存为", initial, "PDF (*.pdf)")
            if not out_dir: return

        try:
            dpi = float(self.entry_dpi.text())
        except:
            dpi = 200.0

        interleave_config = {}
        if mode == "PDF交叉合并":
            pdf_paths = [p for p in paths if p.lower().endswith('.pdf')]
            if len(pdf_paths) != 2:
                return QMessageBox.warning(self, "提示", "PDF交叉合并需要且仅需要添加两个PDF文件，第1个作为A，第2个作为B。")
            paths = pdf_paths
            interleave_config = self.get_interleave_config()

        self.btn_run.setEnabled(False)

        self.worker = ToolkitWorker(paths, mode, out_dir, dpi, self.cmb_fmt.currentText(), self.chk_trans.isChecked(),
                                    self.split_bookmark_config, interleave_config)
        self.worker.progress.connect(lambda v, txt: (self.progress.setValue(v), self.lbl_status.setText(txt)))
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(lambda e: (QMessageBox.critical(self, "错误", e), self.btn_run.setEnabled(True)))
        self.worker.start()

    def on_finished(self, msg):
        self.btn_run.setEnabled(True)
        self.progress.setValue(100)
        QMessageBox.information(self, "成功", msg)
