# -*- coding: utf-8 -*-
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                             QGroupBox, QComboBox, QFileDialog, QMessageBox, QProgressBar, QCheckBox, QLineEdit,
                             QDialog, QRadioButton, QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView,
                             QDialogButtonBox)
from core.ui_components import FileListManagerWidget
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
            "多图转PDF", "多个PDF合并", "PDF拆分为单页",
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

    def run_tool(self):
        if self.file_manager.count() == 0: return QMessageBox.warning(self, "提示", "请先添加文件")
        mode = self.cmb_mode.currentText()

        if "拆分" in mode or "导出图片" in mode:
            out_dir = QFileDialog.getExistingDirectory(self, "选择保存目录")
            if not out_dir: return
        else:
            out_dir, _ = QFileDialog.getSaveFileName(self, "保存为", "工具箱输出.pdf", "PDF (*.pdf)")
            if not out_dir: return

        try:
            dpi = float(self.entry_dpi.text())
        except:
            dpi = 200.0

        paths = self.file_manager.get_all_filepaths()
        self.btn_run.setEnabled(False)

        self.worker = ToolkitWorker(paths, mode, out_dir, dpi, self.cmb_fmt.currentText(), self.chk_trans.isChecked(),
                                    self.split_bookmark_config)
        self.worker.progress.connect(lambda v, txt: (self.progress.setValue(v), self.lbl_status.setText(txt)))
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(lambda e: (QMessageBox.critical(self, "错误", e), self.btn_run.setEnabled(True)))
        self.worker.start()

    def on_finished(self, msg):
        self.btn_run.setEnabled(True)
        self.progress.setValue(100)
        QMessageBox.information(self, "成功", msg)