import os
import json
import pandas as pd
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                             QGroupBox, QLineEdit, QTableWidget, QTableWidgetItem,
                             QMessageBox, QFileDialog, QProgressBar, QInputDialog, QHeaderView, QSplitter,
                             QAbstractItemView, QDialog, QFormLayout, QCheckBox, QDialogButtonBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
import fitz

from core.pdf_viewer import PDFGraphicsView
from core.ui_components import FileListManagerWidget
from core.utils import detect_smart_segments, UniversalSegmentDialog

BTN_BLUE = "background-color: #3498DB; color: white; font-weight: bold; padding: 8px; border-radius: 4px;"
BTN_GREEN = "background-color: #2ECC71; color: white; font-weight: bold; padding: 8px; border-radius: 4px;"
BTN_PURPLE = "background-color: #9B59B6; color: white; font-weight: bold; padding: 8px; border-radius: 4px;"
BTN_GRAY = "background-color: #ECF0F1; color: #2C3E50; font-weight: bold; padding: 6px; border-radius: 4px; border: 1px solid #BDC3C7;"


# ================== 高级字段配置对话框 ==================
class FieldConfigDialog(QDialog):
    def __init__(self, default_name="", is_ocr=True, static_val="", parent=None):
        super().__init__(parent)
        self.setWindowTitle("字段配置")
        self.resize(350, 150)
        layout = QFormLayout(self)

        self.name_edit = QLineEdit(default_name)
        layout.addRow("字段名称:", self.name_edit)

        self.ocr_checkbox = QCheckBox("使用 OCR 框选识别 (在图纸上画框)")
        self.ocr_checkbox.setChecked(is_ocr)
        layout.addRow("", self.ocr_checkbox)

        self.static_val_edit = QLineEdit(static_val)
        self.static_val_edit.setPlaceholderText("留空，或输入如：A版、张三")
        self.static_val_edit.setEnabled(not is_ocr)
        layout.addRow("固定/默认值:", self.static_val_edit)

        self.ocr_checkbox.toggled.connect(lambda checked: self.static_val_edit.setEnabled(not checked))

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def get_data(self):
        return {
            'name': self.name_edit.text().strip(),
            'is_ocr': self.ocr_checkbox.isChecked(),
            'static_val': self.static_val_edit.text() if not self.ocr_checkbox.isChecked() else ""
        }


# ================== 命名组合配置对话框 ==================
class NameFormatDialog(QDialog):
    def __init__(self, fields, parent=None):
        super().__init__(parent)
        self.setWindowTitle("组合命名规则配置")
        self.resize(300, 300)
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("请勾选需要拼接的字段（导出时将按勾选顺序组合）："))

        self.checkboxes = []
        for f in fields:
            cb = QCheckBox(f)
            if len(self.checkboxes) < 2: cb.setChecked(True)
            self.checkboxes.append(cb)
            layout.addWidget(cb)

        form = QFormLayout()
        self.sep_edit = QLineEdit("-")
        form.addRow("各字段分隔符:", self.sep_edit)
        layout.addLayout(form)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def get_format(self):
        selected = [cb.text() for cb in self.checkboxes if cb.isChecked()]
        return selected, self.sep_edit.text()


# ================== 纯净稳定版 PaddleOCR 线程 ==================
class OCRWorker(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(list)

    def __init__(self, pdf_doc, page_configs, fields_config):
        super().__init__()
        self.pdf_doc = pdf_doc
        self.page_configs = page_configs
        self.fields_config = fields_config

    def run(self):
        try:
            self.progress.emit(0, "正在加载 OCR 引擎核心库 (首次加载可能需要几秒)...")

            # 【回归最纯粹的懒加载】，只要打包脚本配置对，这里不需要任何骚操作
            import logging
            logging.getLogger("ppocr").setLevel(logging.WARNING)
            from paddleocr import PaddleOCR
            import numpy as np
            from PIL import Image

            self.progress.emit(5, "OCR 引擎加载成功，正在初始化识别模型...")
            ocr = PaddleOCR(use_angle_cls=True, lang="ch", use_gpu=False, show_log=False, enable_mkldnn=False)

            results = []
            total = len(self.pdf_doc)

            for i in range(total):
                self.progress.emit(int((i / total) * 100), f"正在识别第 {i + 1}/{total} 页...")
                page = self.pdf_doc[i]
                page_boxes = self.page_configs.get(i, {})

                row_data = {}
                for field in self.fields_config:
                    fname = field['name']
                    if not field['is_ocr']:
                        row_data[fname] = field['static_val']
                        continue

                    pdf_rect_data = page_boxes.get(fname, (100, 100, 150, 40))
                    pdf_x, pdf_y, pdf_w, pdf_h = pdf_rect_data
                    clip_rect = fitz.Rect(pdf_x, pdf_y, pdf_x + pdf_w, pdf_y + pdf_h)

                    pix = page.get_pixmap(matrix=fitz.Matrix(3.0, 3.0), clip=clip_rect, alpha=False)
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                    res = ocr.ocr(np.array(img), cls=True)
                    if not res or res[0] is None:
                        row_data[fname] = ""
                    else:
                        row_data[fname] = " ".join([line[1][0] for line in res[0]]).strip()

                results.append(row_data)

            self.progress.emit(100, "识别完成！")
            self.finished.emit(results)

        except Exception as e:
            import traceback
            error_msg = traceback.format_exc()
            print(error_msg)  # 留着打印，万一报错还能看到
            self.progress.emit(0, f"OCR 提取异常: 请检查环境或依赖。详细报错已输出控制台。")


# ================== 界面控制模块 ==================
class OCRExtractorWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.pdf_doc = None
        self.segments = []
        self.page_configs = {}
        self.extracted_data = []
        self.templates_cache = {}  # 核心：用于存储导出的各个图纸尺寸对应的框坐标

        self.fields_config = [
            {"name": "图号", "is_ocr": True, "static_val": ""},
            {"name": "图纸名称", "is_ocr": True, "static_val": ""}
        ]

        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(10, 10, 10, 10)
        splitter = QSplitter(Qt.Vertical)

        top_widget = QWidget()
        top_layout = QHBoxLayout(top_widget)
        top_layout.setContentsMargins(0, 0, 0, 0)

        box_left = QGroupBox("1. 识别区域配置与生成")
        l_left = QVBoxLayout()

        self.file_manager = FileListManagerWidget(accept_exts=['.pdf'], title_desc="PDF Files (*.pdf)")
        l_left.addWidget(self.file_manager, 1)

        hz_cfg = QHBoxLayout()
        btn_import = QPushButton("📥 导入配置/模板")
        btn_import.clicked.connect(self.import_config)
        btn_export = QPushButton("📤 导出配置/模板")
        btn_export.clicked.connect(self.export_config)
        hz_cfg.addWidget(btn_import)
        hz_cfg.addWidget(btn_export)
        l_left.addLayout(hz_cfg)

        hz_fields = QHBoxLayout()
        btn_add = QPushButton("➕ 添加目标字段")
        btn_add.clicked.connect(self.add_field)
        btn_ren = QPushButton("✏️ 修改字段")
        btn_ren.clicked.connect(self.edit_field)
        btn_del = QPushButton("❌ 删除字段")
        btn_del.clicked.connect(self.delete_field)
        hz_fields.addWidget(btn_add)
        hz_fields.addWidget(btn_ren)
        hz_fields.addWidget(btn_del)
        l_left.addLayout(hz_fields)

        btn_merge = QPushButton("🔄 生成合并大纲预览 (首选必点)")
        btn_merge.setStyleSheet(BTN_BLUE)
        btn_merge.clicked.connect(self.merge_pdfs)
        l_left.addWidget(btn_merge)

        hz_btns = QHBoxLayout()
        btn_detect = QPushButton("① 智能分段设框")
        btn_detect.setStyleSheet(BTN_GRAY)
        btn_detect.clicked.connect(self.detect_segments)

        btn_preview = QPushButton("② 终极微调预览")
        btn_preview.setStyleSheet(BTN_PURPLE)
        btn_preview.clicked.connect(self.enter_final_preview)
        hz_btns.addWidget(btn_detect)
        hz_btns.addWidget(btn_preview)
        l_left.addLayout(hz_btns)

        btn_extract = QPushButton("🚀 开始深度 OCR 提取")
        btn_extract.setStyleSheet(BTN_GREEN)
        btn_extract.clicked.connect(self.start_ocr)
        l_left.addWidget(btn_extract)

        self.lbl_status = QLabel("就绪")
        self.progress = QProgressBar()
        l_left.addWidget(self.lbl_status)
        l_left.addWidget(self.progress)
        box_left.setLayout(l_left)
        top_layout.addWidget(box_left, 1)

        box_right = QGroupBox("2. 数据表格与导出")
        l_right = QVBoxLayout()

        hz_tools = QHBoxLayout()
        btn_up = QPushButton("⬆️ 上移")
        btn_up.clicked.connect(self.move_row_up)
        btn_down = QPushButton("⬇️ 下移")
        btn_down.clicked.connect(self.move_row_down)
        btn_sort = QPushButton("🔤 首列排序")
        btn_sort.clicked.connect(self.sort_data)
        btn_merge_json = QPushButton("🔗 合并外部 JSON")
        btn_merge_json.clicked.connect(self.merge_json_files)
        hz_tools.addWidget(btn_up)
        hz_tools.addWidget(btn_down)
        hz_tools.addWidget(btn_sort)
        hz_tools.addStretch(1)
        hz_tools.addWidget(btn_merge_json)
        l_right.addLayout(hz_tools)

        hz_exp = QHBoxLayout()
        btn_exp_xls = QPushButton("导出 Excel/JSON")
        btn_exp_xls.clicked.connect(self.export_excel)
        btn_bmk = QPushButton("写入定制书签")
        btn_bmk.clicked.connect(self.write_bookmarks)
        btn_split = QPushButton("按定制名称拆分")
        btn_split.clicked.connect(self.split_pdf_by_name)
        hz_exp.addWidget(btn_exp_xls)
        hz_exp.addWidget(btn_bmk)
        hz_exp.addWidget(btn_split)
        l_right.addLayout(hz_exp)

        self.table = QTableWidget(0, 0)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)

        self.table.horizontalHeader().setSectionsMovable(True)
        self.table.horizontalHeader().sectionDoubleClicked.connect(self.rename_column_header)
        self.table.setToolTip("💡 提示：左右拖动表头可调整列的前后顺序，双击表头可直接修改字段名称")

        l_right.addWidget(self.table)
        box_right.setLayout(l_right)
        top_layout.addWidget(box_right, 2)
        splitter.addWidget(top_widget)

        bottom_widget = QWidget()
        bottom_layout = QVBoxLayout(bottom_widget)
        bottom_layout.setContentsMargins(0, 0, 0, 0)

        self.btn_confirm = QPushButton("✅ 确认本页 OCR 框位置并同步到该尺寸组")
        self.btn_confirm.setStyleSheet("background-color: #E74C3C; color: white; font-weight: bold; padding: 10px;")
        self.btn_confirm.hide()
        self.btn_confirm.clicked.connect(self.confirm_box_position)

        self.preview_view = PDFGraphicsView()
        bottom_layout.addWidget(self.btn_confirm)
        bottom_layout.addWidget(self.preview_view)

        splitter.addWidget(bottom_widget)
        splitter.setSizes([450, 550])
        main_layout.addWidget(splitter)
        self.refresh_table_data()

    def get_field_names(self):
        return [f['name'] for f in self.fields_config]

    def save_visual_order(self):
        if self.table.columnCount() == 0: return
        header = self.table.horizontalHeader()
        col_count = self.table.columnCount()

        new_fields = []
        for v_idx in range(col_count):
            l_idx = header.logicalIndex(v_idx)
            if l_idx == 0: continue
            if l_idx - 1 < len(self.fields_config):
                new_fields.append(self.fields_config[l_idx - 1])

        if len(new_fields) == len(self.fields_config):
            self.fields_config = new_fields
            new_ext_data = []
            for row in self.extracted_data:
                new_row = {}
                for f in self.fields_config:
                    fname = f['name']
                    new_row[fname] = row.get(fname, "")
                new_ext_data.append(new_row)
            self.extracted_data = new_ext_data

    def rename_column_header(self, logical_index):
        if logical_index == 0: return
        old_name = self.fields_config[logical_index - 1]['name']
        new_name, ok = QInputDialog.getText(self, "修改列名", f"将字段 '{old_name}' 修改为新名称:", text=old_name)
        if ok and new_name.strip() and new_name.strip() != old_name:
            new_name = new_name.strip()
            if new_name in self.get_field_names(): return QMessageBox.warning(self, "错误", "该字段名已存在！")

            if self.pdf_doc and self.page_configs:
                for p in self.page_configs:
                    if old_name in self.page_configs[p]:
                        self.page_configs[p][new_name] = self.page_configs[p].pop(old_name)

            if self.segments:
                for seg in self.segments:
                    if 'pos_pct' in seg and old_name in seg['pos_pct']:
                        seg['pos_pct'][new_name] = seg['pos_pct'].pop(old_name)

            for row in self.extracted_data:
                if old_name in row: row[new_name] = row.pop(old_name)

            self.fields_config[logical_index - 1]['name'] = new_name
            self.save_visual_order()
            self.refresh_table_data()

            # 触发UI热刷新
            self._trigger_ui_hot_reload()

    def _trigger_ui_hot_reload(self):
        """核心内部方法：无缝热刷新所有的界面（预览器、设置框）"""
        mode = getattr(self.preview_view, 'mode', None)
        curr_page = getattr(self.preview_view, 'current_page', -1)
        if self.pdf_doc is not None and mode == 'ocr_final' and curr_page != -1:
            self.preview_view.load_pdf(self.pdf_doc, target_page=curr_page, mode='ocr_final',
                                       data_dict=self.page_configs)

        # 如果当前正打开着分段设框的弹窗，强行刷新它的表格数据
        if hasattr(self, 'dialog') and self.dialog and self.dialog.isVisible():
            if hasattr(self.dialog, 'refresh_table'):
                self.dialog.refresh_table()

    # --- 升级：带模板坐标导出的配置管理 ---
    def export_config(self):
        # 抓取当前所有页面尺寸及其对应的框坐标，生成模板
        templates = {}
        if self.pdf_doc and self.page_configs:
            for p_idx, boxes in self.page_configs.items():
                rect = self.pdf_doc[p_idx].rect
                size_key = f"{int(rect.width)}x{int(rect.height)}"
                if size_key not in templates:
                    templates[size_key] = boxes

        out_data = {
            "fields_config": self.fields_config,
            "templates": templates
        }

        path, _ = QFileDialog.getSaveFileName(self, "导出字段配置与模板", "OCR_Templates_Config.json", "JSON (*.json)")
        if path:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(out_data, f, ensure_ascii=False, indent=4)
            QMessageBox.information(self, "成功", "配置与尺寸模板已成功导出！")

    def import_config(self):
        path, _ = QFileDialog.getOpenFileName(self, "导入字段配置与模板", "", "JSON (*.json)")
        if path:
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)

                    if "fields_config" in data:
                        self.fields_config = data["fields_config"]
                        self.templates_cache = data.get("templates", {})  # 加载记忆的模板

                        self.page_configs.clear()
                        self.segments.clear()
                        self.extracted_data.clear()
                        self.refresh_table_data()
                        QMessageBox.information(self, "成功",
                                                "配置模板导入成功！\n\n请直接点击【① 智能分段设框】，系统会自动为您套用记忆的框坐标。")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"配置导入失败: {e}")

    def add_field(self):
        self.save_visual_order()
        dialog = FieldConfigDialog(parent=self)
        if dialog.exec_() == QDialog.Accepted:
            data = dialog.get_data()
            new_name = data['name']
            if not new_name or new_name in self.get_field_names():
                return QMessageBox.warning(self, "错误", "字段名无效或已存在！")

            self.fields_config.append(data)

            if data['is_ocr'] and self.pdf_doc is not None and len(self.segments) > 0:
                ocr_count = len([f for f in self.fields_config if f['is_ocr']])
                offset_y = 100 + (ocr_count - 1) * 50

                for seg in self.segments:
                    if 'pos_pct' not in seg: seg['pos_pct'] = {}
                    if new_name not in seg['pos_pct']:
                        seg['pos_pct'][new_name] = (100, offset_y, 150, 40)

                    for p in seg['pages']:
                        if p not in self.page_configs: self.page_configs[p] = {}
                        self.page_configs[p][new_name] = seg['pos_pct'][new_name]

            for row in self.extracted_data: row[new_name] = data['static_val']
            self.refresh_table_data()

            # 【真·热更新触发】
            self._trigger_ui_hot_reload()

    def edit_field(self):
        self.save_visual_order()
        if not self.fields_config: return
        names = self.get_field_names()
        old_name, ok = QInputDialog.getItem(self, "选择字段", "请选择要修改的字段:", names, 0, False)
        if ok and old_name:
            idx = names.index(old_name)
            old_data = self.fields_config[idx]
            dialog = FieldConfigDialog(default_name=old_name, is_ocr=old_data['is_ocr'],
                                       static_val=old_data['static_val'], parent=self)

            if dialog.exec_() == QDialog.Accepted:
                new_data = dialog.get_data()
                new_name = new_data['name']

                if new_name != old_name and new_name in names:
                    return QMessageBox.warning(self, "错误", "新字段名已存在！")

                self.fields_config[idx] = new_data

                if self.pdf_doc is not None and len(self.segments) > 0:
                    for p in self.page_configs:
                        if old_name in self.page_configs[p]:
                            val = self.page_configs[p].pop(old_name)
                            if new_data['is_ocr']: self.page_configs[p][new_name] = val
                        elif new_data['is_ocr']:
                            self.page_configs[p][new_name] = (150, 150, 150, 40)

                    for seg in self.segments:
                        if 'pos_pct' in seg:
                            if old_name in seg['pos_pct']:
                                val = seg['pos_pct'].pop(old_name)
                                if new_data['is_ocr']: seg['pos_pct'][new_name] = val
                            elif new_data['is_ocr']:
                                seg['pos_pct'][new_name] = (150, 150, 150, 40)

                for row in self.extracted_data:
                    if old_name in row:
                        val = row.pop(old_name)
                        row[new_name] = val if new_data['is_ocr'] else new_data['static_val']

                self.refresh_table_data()
                self._trigger_ui_hot_reload()

    def delete_field(self):
        self.save_visual_order()
        if not self.fields_config: return
        names = self.get_field_names()
        name, ok = QInputDialog.getItem(self, "删除字段", "选择要删除的字段:", names, 0, False)
        if ok and name:
            self.fields_config = [f for f in self.fields_config if f['name'] != name]

            if self.pdf_doc is not None and len(self.segments) > 0:
                for p in self.page_configs:
                    if name in self.page_configs[p]: del self.page_configs[p][name]
                for seg in self.segments:
                    if 'pos_pct' in seg and name in seg['pos_pct']: del seg['pos_pct'][name]

            for row in self.extracted_data:
                if name in row: del row[name]

            self.refresh_table_data()
            self._trigger_ui_hot_reload()

    def refresh_table_data(self):
        self.lbl_status.setText("当前字段配置: " + ", ".join(self.get_field_names()))
        headers = ["序号"] + self.get_field_names()
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setColumnWidth(0, 50)
        self.table.setRowCount(len(self.extracted_data))

        for row, item in enumerate(self.extracted_data):
            it_seq = QTableWidgetItem(str(row + 1))
            it_seq.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            it_seq.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 0, it_seq)
            for col, field in enumerate(self.get_field_names()):
                self.table.setItem(row, col + 1, QTableWidgetItem(str(item.get(field, ""))))

    def get_final_data(self):
        self.save_visual_order()
        return [{k: v for k, v in row.items()} for row in self.extracted_data]

    def merge_pdfs(self):
        if self.file_manager.count() == 0: return QMessageBox.warning(self, "提示", "请先添加PDF文件！")

        self.pdf_doc = fitz.Document()
        toc_list = []
        for path in self.file_manager.get_all_filepaths():
            doc = fitz.open(path)
            start_page = len(self.pdf_doc)
            toc_list.append([1, os.path.basename(path), start_page + 1])
            self.pdf_doc.insert_pdf(doc)
            doc.close()
        self.pdf_doc.set_toc(toc_list)
        self.preview_view.load_pdf(self.pdf_doc)

        self.segments.clear()
        self.page_configs.clear()

        QMessageBox.information(self, "成功", f"大纲预览生成，共 {len(self.pdf_doc)} 页")

    # --- 升级：检测分段时自动融合模板中的记忆坐标 ---
    def detect_segments(self):
        if not self.pdf_doc: return QMessageBox.warning(self, "错误", "请先生成大纲预览！")
        if not self.fields_config: return QMessageBox.warning(self, "提示", "请至少配置一个字段！")

        if not self.segments:
            self.segments = detect_smart_segments(self.pdf_doc)

        ocr_fields = [f['name'] for f in self.fields_config if f['is_ocr']]

        for seg in self.segments:
            p_idx = seg['pages'][0]
            rect = self.pdf_doc[p_idx].rect
            w, h = int(rect.width), int(rect.height)

            # --- 智能匹配缓存模板 ---
            matched_template = None
            if self.templates_cache:
                for t_key, t_boxes in self.templates_cache.items():
                    tw, th = map(int, t_key.split('x'))
                    # 容差设为 5 像素，应对各种图纸细微偏差
                    if abs(w - tw) <= 5 and abs(h - th) <= 5:
                        matched_template = t_boxes
                        break

            if 'pos_pct' not in seg:
                seg['pos_pct'] = {}

            current_boxes = seg['pos_pct']
            existing_keys = list(current_boxes.keys())

            for k in existing_keys:
                if k not in ocr_fields:
                    del current_boxes[k]

            for i, field in enumerate(ocr_fields):
                if field not in current_boxes:
                    # 如果模板匹配成功且包含该字段，优先使用模板坐标！否则套用默认。
                    if matched_template and field in matched_template:
                        current_boxes[field] = matched_template[field]
                    else:
                        current_boxes[field] = (100, 100 + i * 50, 150, 40)

            seg['pos_pct'] = current_boxes

            for p in seg['pages']:
                self.page_configs[p] = dict(current_boxes)

        self.dialog = UniversalSegmentDialog(self.segments, "设置各个字段专属提取框", self)
        self.dialog.exec_()

    def enter_setting_mode_from_dialog(self, seg_data, idx, dialog):
        self.current_idx = idx
        self.dialog = dialog
        target_page = seg_data['pages'][0]
        self.preview_view.load_pdf(self.pdf_doc, target_page=target_page, mode='ocr_final',
                                   data_dict={target_page: seg_data['pos_pct']})
        self.btn_confirm.show()

    def confirm_box_position(self):
        self.preview_view.save_current_page_state()
        new_pos_dict = self.preview_view.page_data_dict[self.preview_view.current_page]

        self.segments[self.current_idx]['pos_pct'] = new_pos_dict
        self.segments[self.current_idx]['pos_set'] = True

        for p in self.segments[self.current_idx]['pages']:
            self.page_configs[p] = dict(new_pos_dict)

        self.btn_confirm.hide()
        self.dialog.refresh_table()
        self.dialog.show()

    def enter_final_preview(self):
        if not self.pdf_doc: return QMessageBox.warning(self, "错误", "请先生成预览！")
        self.preview_view.load_pdf(self.pdf_doc, mode='ocr_final', data_dict=self.page_configs)
        QMessageBox.information(self, "提示",
                                "已进入终极预览，您可以拖动右下角调整大小！\n在任意一页的调整都会被独立记录。")

    def start_ocr(self):
        if not self.pdf_doc: return
        self.preview_view.save_current_page_state()
        self.page_configs = self.preview_view.page_data_dict
        self.worker = OCRWorker(self.pdf_doc, self.page_configs, self.fields_config)
        self.worker.progress.connect(lambda v, txt: (self.progress.setValue(v), self.lbl_status.setText(txt)))
        self.worker.finished.connect(self.on_ocr_finished)
        self.worker.start()

    def on_ocr_finished(self, results):
        self.extracted_data = results
        self.refresh_table_data()

    def move_row_up(self):
        row = self.table.currentRow()
        if row > 0:
            data = self.get_final_data()
            data[row], data[row - 1] = data[row - 1], data[row]
            self.extracted_data = data
            self.refresh_table_data()
            self.table.selectRow(row - 1)

    def move_row_down(self):
        row = self.table.currentRow()
        if row >= 0 and row < self.table.rowCount() - 1:
            data = self.get_final_data()
            data[row], data[row + 1] = data[row + 1], data[row]
            self.extracted_data = data
            self.refresh_table_data()
            self.table.selectRow(row + 1)

    def sort_data(self):
        data = self.get_final_data()
        if not data: return
        first_field = list(data[0].keys())[0]
        data.sort(key=lambda x: str(x.get(first_field, "")))
        self.extracted_data = data
        self.refresh_table_data()

    def merge_json_files(self):
        self.save_visual_order()
        files, _ = QFileDialog.getOpenFileNames(self, "选择 JSON", "", "JSON (*.json)")
        if not files: return
        try:
            merged_list = []
            for f in files:
                with open(f, 'r', encoding='utf-8') as file:
                    content = json.load(file)
                    if "data" in content and "headers" in content:
                        headers = [h["name"] for h in content["headers"]]
                        for row_vals in content["data"]:
                            row_dict = {k: v for k, v in zip(headers, row_vals) if k != "序号"}
                            merged_list.append(row_dict)

            existing_names = self.get_field_names()
            for item in merged_list:
                for k in item.keys():
                    if k not in existing_names:
                        self.fields_config.append({"name": k, "is_ocr": False, "static_val": ""})
                        existing_names.append(k)

            data = self.get_final_data()
            data.extend(merged_list)
            self.extracted_data = data
            self.refresh_table_data()
            QMessageBox.information(self, "合并成功", f"成功追加记录")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"解析失败: {e}")

    def export_excel(self):
        data = self.get_final_data()
        if not data: return

        default_name = "图纸提取数据.xlsx"
        if self.file_manager.count() > 0:
            first_file = self.file_manager.get_all_filepaths()[0]
            base_name = os.path.splitext(os.path.basename(first_file))[0]
            default_name = f"{base_name}_目录.xlsx"

        path, _ = QFileDialog.getSaveFileName(self, "保存数据", default_name, "Excel (*.xlsx)")
        if path:
            export_data = [{"序号": i + 1, **row} for i, row in enumerate(data)]
            pd.DataFrame(export_data).to_excel(path, index=False)

            json_path = path.replace('.xlsx', '.json')
            headers = [{"name": c, "width": 40} for c in export_data[0].keys()]
            vals = [list(d.values()) for d in export_data]
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump({"headers": headers, "data": vals}, f, ensure_ascii=False, indent=2)
            QMessageBox.information(self, "成功", "Excel 与 JSON 均已生成！")

    def get_custom_name(self, row, idx, selected_fields, sep):
        parts = []
        for f in selected_fields:
            if f == "序号":
                parts.append(str(idx + 1))
            else:
                parts.append(str(row.get(f, "")))
        return sep.join([p for p in parts if p]).strip()

    def write_bookmarks(self):
        if not self.pdf_doc: return
        data = self.get_final_data()

        fields = ["序号"] + self.get_field_names()
        dialog = NameFormatDialog(fields, self)
        if dialog.exec_() != QDialog.Accepted: return
        selected_fields, sep = dialog.get_format()
        if not selected_fields: return QMessageBox.warning(self, "错误", "至少勾选一项！")

        toc = []
        for i, row in enumerate(data):
            title = self.get_custom_name(row, i, selected_fields, sep)
            if not title: title = f"Page_{i + 1}"
            toc.append([1, title, i + 1])

        self.pdf_doc.set_toc(toc)
        path, _ = QFileDialog.getSaveFileName(self, "保存", "带定制书签的图纸.pdf", "PDF (*.pdf)")
        if path:
            self.pdf_doc.save(path)
            QMessageBox.information(self, "成功", "定制书签已写入！")

    def split_pdf_by_name(self):
        if not self.pdf_doc: return
        data = self.get_final_data()

        fields = ["序号"] + self.get_field_names()
        dialog = NameFormatDialog(fields, self)
        if dialog.exec_() != QDialog.Accepted: return
        selected_fields, sep = dialog.get_format()
        if not selected_fields: return QMessageBox.warning(self, "错误", "至少勾选一项！")

        out_dir = QFileDialog.getExistingDirectory(self, "选择保存目录")
        if not out_dir: return

        for i, row in enumerate(data):
            title = self.get_custom_name(row, i, selected_fields, sep)
            if not title: title = f"Page_{i + 1}"

            import re
            title = re.sub(r'[\\/*?:"<>|]', '_', title)
            final_name = f"{title}.pdf"

            new_doc = fitz.Document()
            new_doc.insert_pdf(self.pdf_doc, from_page=i, to_page=i)
            new_doc.save(os.path.join(out_dir, final_name))
            new_doc.close()

        QMessageBox.information(self, "成功", "按定制名称拆分完成！")
