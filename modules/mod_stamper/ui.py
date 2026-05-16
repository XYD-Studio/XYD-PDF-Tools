# -*- coding: utf-8 -*-
import os
import json
import uuid
import copy
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                             QGroupBox, QLineEdit, QMessageBox, QFileDialog, QProgressBar, QSplitter,
                             QTableWidget, QTableWidgetItem, QHeaderView, QDialog, QFormLayout, QDialogButtonBox,
                             QCheckBox, QComboBox, QApplication)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QPixmap
import fitz

from core.pdf_viewer import PDFGraphicsView
from core.ui_components import FileListManagerWidget, ExportSettingsPanel, GSSettingsPanel
from core.utils import (detect_smart_segments, UniversalSegmentDialog,
                        find_ghostscript, BTN_BLUE, BTN_GREEN, BTN_PURPLE, BTN_RED, BTN_GRAY, BTN_ORANGE)
from core.pdf_engine import merge_pdf_with_smart_toc, run_ghostscript, BaseFakeProgressWorker
from .worker import StamperWorker

try:
    import pyhanko

    PYHANKO_AVAILABLE = True
except ImportError:
    PYHANKO_AVAILABLE = False


class StamperWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.pdf_doc = None
        self.segments = []
        self.page_stamp_positions = {}
        self.original_filenames = []
        self.global_stamps = []
        self.gs_path, self.gs_lib_path = find_ghostscript()
        self.pfx_path = ""
        self.need_clean_annots = False

        self.fake_timer = QTimer(self)
        self.fake_timer.timeout.connect(self._on_fake_timer_tick)
        self.fake_val = 80.0
        self.fake_msgs = [
            "🗜️ 正在启动底层 Ghostscript 引擎进行深度重绘...",
            "⏳ 根据图纸复杂度，此过程可能极其耗时，请耐心等待...",
            "⚙️ 正在进行图像二次降维与光栅化压缩...",
            "🔒 系统正在高速运转中，未发生卡死，请勿关闭程序..."
        ]
        self.fake_msg_idx = 0

        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(10, 10, 10, 10)
        splitter = QSplitter(Qt.Vertical)

        top_widget = QWidget()
        top_layout = QHBoxLayout(top_widget)
        top_layout.setContentsMargins(0, 0, 0, 0)

        box_left = QGroupBox("1. 文件与多图章配置")
        l_left = QVBoxLayout()

        hz_cfg = QHBoxLayout()
        btn_import_cfg = QPushButton("📂 导入图章配置")
        btn_import_cfg.clicked.connect(self.import_config)
        btn_export_cfg = QPushButton("💾 导出图章配置")
        btn_export_cfg.clicked.connect(self.export_config)
        hz_cfg.addWidget(btn_import_cfg);
        hz_cfg.addWidget(btn_export_cfg)
        l_left.addLayout(hz_cfg)

        self.file_manager = FileListManagerWidget(accept_exts=['.pdf'], title_desc="PDF Files (*.pdf)")
        l_left.addWidget(self.file_manager, 1)

        self.chk_pre_gs = QCheckBox("🛠️ 合并前用GS扁平化 (破除Acrobat密码锁及数字签名)")
        self.chk_pre_gs.setStyleSheet("color: #E67E22; font-weight: bold;")
        if not self.gs_path:
            self.chk_pre_gs.setEnabled(False)
            self.chk_pre_gs.setText("🛠️ 合并前用GS扁平化 (未检测到环境)")
        l_left.addWidget(self.chk_pre_gs)

        self.chk_toc_strategy = QCheckBox("🔖 合并时，单页 PDF 强制使用【文件名】作为书签")
        self.chk_toc_strategy.setChecked(True)
        self.chk_toc_strategy.setStyleSheet("color: #2C3E50; font-weight: bold;")
        l_left.addWidget(self.chk_toc_strategy)

        btn_merge = QPushButton("🔄 生成秒开预览 (首选必点)")
        btn_merge.setStyleSheet(BTN_BLUE)
        btn_merge.clicked.connect(self.merge_pdfs)
        l_left.addWidget(btn_merge)

        btn_clean = QPushButton("🧹 快速清除原PDF附带的图章批注")
        btn_clean.setStyleSheet(BTN_ORANGE)
        btn_clean.clicked.connect(self.clean_original_stamps)
        l_left.addWidget(btn_clean)

        hz_stamp = QHBoxLayout()
        btn_add_stamp = QPushButton("➕ 增加图章/签名")
        btn_add_stamp.clicked.connect(self.add_stamp_item)
        btn_del_stamp = QPushButton("❌ 删除选中印章/签名")
        btn_del_stamp.clicked.connect(self.del_stamp_item)
        hz_stamp.addWidget(btn_add_stamp);
        hz_stamp.addWidget(btn_del_stamp)
        l_left.addLayout(hz_stamp)

        self.stamp_table = QTableWidget(0, 3)
        self.stamp_table.setHorizontalHeaderLabels(["标记名称", "宽(mm)", "高(mm)"])
        self.stamp_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        l_left.addWidget(self.stamp_table, 1)

        box_left.setLayout(l_left)
        top_layout.addWidget(box_left, 1)

        box_right = QGroupBox("2. 智能排布、压缩与防伪签名导出")
        l_right = QVBoxLayout()

        btn_detect = QPushButton("① 智能按图纸尺寸分段组合")
        btn_detect.setStyleSheet(BTN_GRAY)
        btn_detect.clicked.connect(self.detect_segments)

        btn_preview = QPushButton("② 进入终极自由微调预览")
        btn_preview.setStyleSheet(BTN_PURPLE)
        btn_preview.clicked.connect(self.enter_final_preview)
        l_right.addWidget(btn_detect);
        l_right.addWidget(btn_preview)

        self.export_panel = ExportSettingsPanel("导出模式设置:", default_mode='merged', allow_overwrite=True)
        l_right.addWidget(self.export_panel)

        hz_stamp_dpi = QHBoxLayout()
        hz_stamp_dpi.addWidget(QLabel("🖼️ 图章画质:"))
        self.cmb_stamp_dpi = QComboBox()
        self.cmb_stamp_dpi.addItems(["150", "300", "72", "600"])
        hz_stamp_dpi.addWidget(self.cmb_stamp_dpi)
        hz_stamp_dpi.addStretch()
        l_right.addLayout(hz_stamp_dpi)

        self.gs_panel = GSSettingsPanel(bool(self.gs_path))
        l_right.addWidget(self.gs_panel)

        self.chk_pki = QCheckBox("🛡️ 附加防篡改保护锁 (防止他人用编辑器删图章)")
        if not PYHANKO_AVAILABLE:
            self.chk_pki.setEnabled(False)
            self.chk_pki.setText("🛡️ 附加防篡改保护锁 (⚠️请安装 pyhanko/cryptography)")
        l_right.addWidget(self.chk_pki)

        self.widget_pki_params = QWidget()
        fl_pki = QFormLayout(self.widget_pki_params)
        fl_pki.setContentsMargins(20, 0, 0, 0)
        hz_pfx = QHBoxLayout()
        self.lbl_pfx = QLabel("未选择")
        btn_pfx = QPushButton("浏览")
        btn_pfx.clicked.connect(self.select_pfx)
        hz_pfx.addWidget(self.lbl_pfx);
        hz_pfx.addWidget(btn_pfx)
        self.entry_pwd = QLineEdit()
        self.entry_pwd.setEchoMode(QLineEdit.Password)
        self.entry_pwd.setPlaceholderText("提取密码")
        hz_pki_target = QHBoxLayout()
        self.cmb_pki_target = QComboBox()
        self.cmb_pki_target.addItem("默认 (遇到第一个图章处)", userData="first")
        hz_pki_target.addWidget(self.cmb_pki_target);
        hz_pki_target.addStretch()
        fl_pki.addRow(".pfx证书:", hz_pfx)
        fl_pki.addRow("证书密码:", self.entry_pwd)
        fl_pki.addRow("锁定关联区域:", hz_pki_target)
        self.widget_pki_params.hide()
        self.chk_pki.toggled.connect(self.widget_pki_params.setVisible)
        l_right.addWidget(self.widget_pki_params)

        self.btn_export = QPushButton("🚀 终极执行：可视化盖章 ➡️ 压缩 ➡️ 防篡改导出")
        self.btn_export.setStyleSheet(BTN_GREEN)
        self.btn_export.clicked.connect(self.start_export)
        l_right.addWidget(self.btn_export)

        self.lbl_status = QLabel("就绪")
        self.progress_bar = QProgressBar()
        l_right.addWidget(self.lbl_status)
        l_right.addWidget(self.progress_bar)

        box_right.setLayout(l_right)
        top_layout.addWidget(box_right, 1)
        splitter.addWidget(top_widget)

        bottom_widget = QWidget()
        bottom_layout = QVBoxLayout(bottom_widget)
        bottom_layout.setContentsMargins(0, 0, 0, 0)

        hz_preview_tools = QHBoxLayout()
        self.btn_confirm_pos = QPushButton("✅ 确认本页位置应用到全部")
        self.btn_confirm_pos.setStyleSheet(BTN_RED)
        self.btn_confirm_pos.hide()
        self.btn_confirm_pos.clicked.connect(self.confirm_stamp_position)

        self.btn_add_missing = QPushButton("➕ 误删找回：向本页补加图章")
        self.btn_add_missing.setStyleSheet(BTN_BLUE)
        self.btn_add_missing.hide()
        self.btn_add_missing.clicked.connect(self.add_stamp_to_current_preview)

        hz_preview_tools.addWidget(self.btn_confirm_pos);
        hz_preview_tools.addWidget(self.btn_add_missing);
        hz_preview_tools.addStretch()
        self.preview_view = PDFGraphicsView()
        bottom_layout.addLayout(hz_preview_tools);
        bottom_layout.addWidget(self.preview_view)

        splitter.addWidget(bottom_widget);
        splitter.setSizes([450, 550])
        main_layout.addWidget(splitter)

    def _on_fake_timer_tick(self):
        if self.fake_val < 98.0:
            self.fake_val += (99.0 - self.fake_val) * 0.05
            self.progress_bar.setValue(int(self.fake_val))
        self.lbl_status.setText(self.fake_msgs[self.fake_msg_idx % len(self.fake_msgs)])
        self.fake_msg_idx += 1

    def _on_worker_progress(self, val):
        if val == BaseFakeProgressWorker.TRIGGER_FAKE_PROGRESS:
            self.fake_val = 80.0
            self.fake_msg_idx = 0
            self._on_fake_timer_tick()
            self.fake_timer.start(2500)
        else:
            self.fake_timer.stop()
            self.progress_bar.setRange(0, 100)
            self.progress_bar.setValue(val)

    def _on_worker_status(self, msg):
        self.lbl_status.setText(msg)

    def _on_worker_finished(self, msg):
        self.fake_timer.stop();
        self.progress_bar.setRange(0, 100);
        self.progress_bar.setValue(100)
        self.lbl_status.setText("✅ 就绪");
        self.btn_export.setEnabled(True)
        QMessageBox.information(self, "大功告成", msg)

    def _on_worker_error(self, err):
        self.fake_timer.stop();
        self.progress_bar.setRange(0, 100)
        self.lbl_status.setText("❌ 发生异常");
        self.btn_export.setEnabled(True)
        QMessageBox.critical(self, "出错了", err)

    def select_pfx(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择数字证书", "", "PFX Certificates (*.pfx *.p12)")
        if path:
            self.pfx_path = path
            self.lbl_pfx.setText(os.path.basename(path))

    def merge_pdfs(self):
        if self.file_manager.count() == 0: return QMessageBox.warning(self, "提示", "请先添加PDF文件！")
        self.pdf_doc = fitz.Document()
        self.original_filenames = []
        toc_list = []
        filepaths = self.file_manager.get_all_filepaths()
        prefer_filename = self.chk_toc_strategy.isChecked()
        self.need_clean_annots = False

        for i, path in enumerate(filepaths):
            self.original_filenames.append(os.path.basename(path))
            doc = fitz.open(path)
            merge_pdf_with_smart_toc(doc, os.path.basename(path), self.pdf_doc, toc_list, prefer_filename)
            doc.close()

        self.pdf_doc.set_toc(toc_list)
        self.preview_view.load_pdf(self.pdf_doc)
        self.progress_bar.setValue(0)
        QMessageBox.information(self, "成功", f"生成混合阵列完毕，共计 {len(self.pdf_doc)} 页。")

    def add_stamp_to_current_preview(self):
        if not self.global_stamps: return QMessageBox.warning(self, "提示", "全局图章库为空！")
        current_page = getattr(self.preview_view, 'current_page', None)
        if current_page is None: return QMessageBox.warning(self, "提示", "请先加载预览！")

        dialog = QDialog(self)
        dialog.setWindowTitle("找回图章")
        layout = QVBoxLayout(dialog)
        cmb = QComboBox()
        for st in self.global_stamps: cmb.addItem(st['name'], userData=st)
        layout.addWidget(QLabel("选择要加回当前画布的图章:"))
        layout.addWidget(cmb)
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(dialog.accept);
        btns.rejected.connect(dialog.reject)
        layout.addWidget(btns)

        if dialog.exec_() == QDialog.Accepted:
            new_stamp = copy.deepcopy(cmb.currentData())
            new_stamp['pdf_x'] = 100;
            new_stamp['pdf_y'] = 100
            if hasattr(self.preview_view, 'save_current_page_state'): self.preview_view.save_current_page_state()
            if current_page not in self.preview_view.page_data_dict: self.preview_view.page_data_dict[current_page] = []
            self.preview_view.page_data_dict[current_page].append(new_stamp)
            mode = getattr(self.preview_view, 'mode', 'stamp_final')
            self.preview_view.load_pdf(self.pdf_doc, target_page=current_page, mode=mode,
                                       data_dict=self.preview_view.page_data_dict)

    def _ensure_segment_stamps(self, seg):
        if not seg.get('pos_set', False):
            first_page = seg['pages'][0]
            if first_page in self.page_stamp_positions and self.page_stamp_positions[first_page]:
                seg['pos_pct'] = copy.deepcopy(self.page_stamp_positions[first_page])
            else:
                default_stamps = copy.deepcopy(self.global_stamps)
                for i, st in enumerate(default_stamps): st['pdf_x'] = 100; st['pdf_y'] = 100 + i * 150
                seg['pos_pct'] = default_stamps
            seg['pos_set'] = True

    def clean_original_stamps(self):
        if not self.pdf_doc: return QMessageBox.warning(self, "提示", "请先合并预览加载文档！")
        deleted_count = 0
        for page_num in range(len(self.pdf_doc)):
            page = self.pdf_doc[page_num]
            for widget in page.widgets(): page.delete_widget(widget); deleted_count += 1
            for annot in page.annots(): page.delete_annot(annot); deleted_count += 1

        if deleted_count > 0:
            self.need_clean_annots = True
            try:
                self.preview_view.load_pdf(self.pdf_doc, target_page=self.preview_view.current_page,
                                           mode=getattr(self.preview_view, 'mode', 'normal'),
                                           data_dict=self.page_stamp_positions)
            except Exception:
                self.preview_view.load_pdf(self.pdf_doc)
            QMessageBox.information(self, "清理成功",
                                    f"UI视图成功清除了 {deleted_count} 个表单控件！导出时将会真正删除。")
        else:
            QMessageBox.warning(self, "无法清理", "未发现可清理的悬浮批注。")

    def export_config(self):
        if not self.pdf_doc or not self.page_stamp_positions: return
        size_configs = {}
        for p_idx, stamp_list in self.page_stamp_positions.items():
            if not stamp_list: continue
            page = self.pdf_doc[p_idx]
            size_key = f"{int(round(page.rect.width))}x{int(round(page.rect.height))}"
            if size_key not in size_configs: size_configs[size_key] = stamp_list
        export_data = {"module": "Stamper", "global_stamps": self.global_stamps, "size_configs": size_configs}
        path, _ = QFileDialog.getSaveFileName(self, "保存图章配置", "图章排版配置.json", "JSON Files (*.json)")
        if path:
            with open(path, 'w', encoding='utf-8') as f: json.dump(export_data, f, ensure_ascii=False, indent=4)
            QMessageBox.information(self, "成功", "已保存！")

    def import_config(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择图章配置文件", "", "JSON Files (*.json)")
        if not path: return
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            self.global_stamps = data.get("global_stamps", [])
            self.refresh_stamp_table()
            self.page_stamp_positions.clear()
            if self.pdf_doc:
                for i in range(len(self.pdf_doc)):
                    page = self.pdf_doc[i]
                    size_key = f"{int(round(page.rect.width))}x{int(round(page.rect.height))}"
                    if size_key in data.get("size_configs", {}): self.page_stamp_positions[i] = copy.deepcopy(
                        data.get("size_configs")[size_key])
                QMessageBox.information(self, "导入成功", "匹配当前PDF尺寸成功。")
        except Exception as e:
            pass

    def add_stamp_item(self):
        file, _ = QFileDialog.getOpenFileName(self, "选择图章", "", "Image Files (*.png *.jpg *.jpeg)")
        if not file: return
        pixmap = QPixmap(file)
        orig_w = max(1, pixmap.width())
        orig_h = max(1, pixmap.height())
        aspect_ratio = orig_w / orig_h
        dialog = QDialog(self)
        dialog.setWindowTitle("属性")
        layout = QFormLayout(dialog)
        name_input = QLineEdit(f"印章_{len(self.global_stamps) + 1}")
        w_input = QLineEdit("50.0")
        h_input = QLineEdit(str(round(50.0 / aspect_ratio, 2)))
        lock_cb = QCheckBox("锁定原始宽高比")
        lock_cb.setChecked(True)

        def update_h(txt):
            if lock_cb.isChecked() and w_input.hasFocus():
                try:
                    h_input.setText(str(round(float(txt) / aspect_ratio, 2)))
                except:
                    pass

        def update_w(txt):
            if lock_cb.isChecked() and h_input.hasFocus():
                try:
                    w_input.setText(str(round(float(txt) * aspect_ratio, 2)))
                except:
                    pass

        w_input.textEdited.connect(update_h);
        h_input.textEdited.connect(update_w)
        layout.addRow("名称:", name_input);
        layout.addRow("宽(mm):", w_input);
        layout.addRow("高(mm):", h_input);
        layout.addRow("", lock_cb)
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(dialog.accept);
        btns.rejected.connect(dialog.reject)
        layout.addWidget(btns)
        if dialog.exec_() == QDialog.Accepted:
            self.global_stamps.append(
                {'id': str(uuid.uuid4()), 'name': name_input.text(), 'path': file, 'w': float(w_input.text()),
                 'h': float(h_input.text()), 'angle': 0})
            self.refresh_stamp_table();
            self._warn_reset_segments()

    def del_stamp_item(self):
        row = self.stamp_table.currentRow()
        if row >= 0: del self.global_stamps[row]; self.refresh_stamp_table(); self._warn_reset_segments()

    def refresh_stamp_table(self):
        self.stamp_table.setRowCount(len(self.global_stamps))
        current_target_id = self.cmb_pki_target.currentData() if hasattr(self, 'cmb_pki_target') else "first"
        self.cmb_pki_target.clear()
        self.cmb_pki_target.addItem("默认 (遇到第一个图章处)", userData="first")
        for i, st in enumerate(self.global_stamps):
            self.stamp_table.setItem(i, 0, QTableWidgetItem(st['name']))
            self.stamp_table.setItem(i, 1, QTableWidgetItem(str(st['w'])))
            self.stamp_table.setItem(i, 2, QTableWidgetItem(str(st['h'])))
            self.cmb_pki_target.addItem(f"绑定到: {st['name']}", userData=st['id'])
        idx = self.cmb_pki_target.findData(current_target_id)
        if idx >= 0: self.cmb_pki_target.setCurrentIndex(idx)

    def _warn_reset_segments(self):
        if self.segments: self.segments = []; self.page_stamp_positions = {}

    def detect_segments(self):
        if not self.pdf_doc: return QMessageBox.warning(self, "错误", "请先生成合并预览！")
        if not self.global_stamps: return QMessageBox.warning(self, "提示", "请先添加图章！")
        self.segments = detect_smart_segments(self.pdf_doc)
        for seg in self.segments:
            self._ensure_segment_stamps(seg)
            for p in seg['pages']:
                if p not in self.page_stamp_positions: self.page_stamp_positions[p] = copy.deepcopy(seg['pos_pct'])
        self.dialog = UniversalSegmentDialog(self.segments, "设置该尺寸页面的默认印章位置", self)
        self.dialog.exec_()
        for seg in self.segments:
            self._ensure_segment_stamps(seg)
            for p in seg['pages']: self.page_stamp_positions[p] = copy.deepcopy(seg['pos_pct'])

    def enter_setting_mode_from_dialog(self, seg_data, idx, dialog):
        self.current_idx = idx;
        self.dialog = dialog
        self._ensure_segment_stamps(seg_data)
        target_page = seg_data['pages'][0]
        self.preview_view.load_pdf(self.pdf_doc, target_page=target_page, mode='stamp_final',
                                   data_dict={target_page: seg_data['pos_pct']})

        # 💡 [修复核心：调用底层的防翻页锁机制]
        self.preview_view.set_nav_locked(True)
        self.btn_confirm_pos.show();
        self.btn_add_missing.show()

    def confirm_stamp_position(self):
        self.preview_view.save_current_page_state()
        new_stamps_list = self.preview_view.page_data_dict[self.preview_view.current_page]
        self.segments[self.current_idx]['pos_pct'] = new_stamps_list
        self.segments[self.current_idx]['pos_set'] = True

        # 💡 确保将该设置深度同步应用到这一组的所有页！
        for p in self.segments[self.current_idx]['pages']:
            self.page_stamp_positions[p] = copy.deepcopy(new_stamps_list)

        # 💡 解除防翻页锁
        self.preview_view.set_nav_locked(False)
        self.btn_confirm_pos.hide()
        self.dialog.refresh_table();
        self.dialog.show()

    def enter_final_preview(self):
        if not self.pdf_doc or not self.global_stamps: return QMessageBox.warning(self, "错误", "缺少PDF或图章数据。")
        for p in range(len(self.pdf_doc)):
            if p not in self.page_stamp_positions:
                seg_stamps = copy.deepcopy(self.global_stamps)
                for i, st in enumerate(seg_stamps): st['pdf_x'] = 100; st['pdf_y'] = 100 + i * 150
                self.page_stamp_positions[p] = seg_stamps

        # 💡 终极预览模式：解锁防翻页
        self.preview_view.set_nav_locked(False)
        self.preview_view.load_pdf(self.pdf_doc, mode='stamp_final', data_dict=self.page_stamp_positions)
        self.btn_confirm_pos.hide();
        self.btn_add_missing.show()
        QMessageBox.information(self, "高级微调模式", "已进入终极预览。可以 右键/Alt键 复制、缩放、旋转图章！")

    def start_export(self):
        if not self.file_manager.count() or not self.global_stamps: return QMessageBox.warning(self, "提示",
                                                                                               "未检测到文件或图章数据！")
        self.preview_view.save_current_page_state()

        # 💡 终极修复：把用户微调后的数据统统吃进来
        self.page_stamp_positions.update(copy.deepcopy(self.preview_view.page_data_dict))

        export_cfg = self.export_panel.get_config()
        gs_cfg = self.gs_panel.get_config()

        if export_cfg['mode'] == 'merged':
            path, _ = QFileDialog.getSaveFileName(self, "保存", "合并加盖图纸.pdf", "PDF (*.pdf)")
            if not path: return
        elif export_cfg['mode'] in ['batch', 'split']:
            path = QFileDialog.getExistingDirectory(self, "选择保存目录")
            if not path: return
        elif export_cfg['mode'] == 'overwrite':
            reply = QMessageBox.question(self, "高危操作确认",
                                         "【原文件直接覆盖】模式将会彻底替换您的原始 PDF 文件且无法恢复！\n强烈建议提前做好备份。\n\n是否确定继续？",
                                         QMessageBox.Yes | QMessageBox.No)
            if reply != QMessageBox.Yes: return
            path = ""

        self.btn_export.setEnabled(False)

        self.worker = StamperWorker(
            file_paths=self.file_manager.get_all_filepaths(),
            page_positions=self.page_stamp_positions,
            export_config=export_cfg,
            gs_config=gs_cfg,
            output_path=path,
            prefer_filename=self.chk_toc_strategy.isChecked(),
            pre_flatten=self.chk_pre_gs.isChecked(),
            clean_annots=self.need_clean_annots,
            stamp_dpi=int(self.cmb_stamp_dpi.currentText()),
            gs_path=self.gs_path,
            gs_lib_path=self.gs_lib_path,
            use_pki=self.chk_pki.isChecked(),
            pfx_path=self.pfx_path,
            pfx_pwd=self.entry_pwd.text(),
            pki_target_id=self.cmb_pki_target.currentData() if hasattr(self, 'cmb_pki_target') else "first"
        )

        self.worker.progress.connect(self._on_worker_progress)
        self.worker.status.connect(self._on_worker_status)
        self.worker.finished.connect(self._on_worker_finished)
        self.worker.error.connect(self._on_worker_error)
        self.worker.start()