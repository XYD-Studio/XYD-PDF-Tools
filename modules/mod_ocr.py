import os
import json
import pandas as pd
import copy
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                             QGroupBox, QLineEdit, QTableWidget, QTableWidgetItem,
                             QMessageBox, QFileDialog, QProgressBar, QInputDialog, QHeaderView, QSplitter,
                             QAbstractItemView)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
import fitz

from core.pdf_viewer import PDFGraphicsView
from core.ui_components import FileListManagerWidget
from core.utils import detect_smart_segments, UniversalSegmentDialog, BTN_BLUE, BTN_GREEN, BTN_PURPLE, BTN_GRAY

class OCRWorker(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(list)

    def __init__(self, pdf_doc, page_configs):
        super().__init__()
        self.pdf_doc = pdf_doc
        self.page_configs = page_configs

    def run(self):
        try:
            import logging
            logging.getLogger("ppocr").setLevel(logging.WARNING)

            from paddleocr import PaddleOCR
            import numpy as np
            from PIL import Image

            self.progress.emit(0, "正在加载稳定版 PaddleOCR 引擎...")
            ocr = PaddleOCR(use_angle_cls=True, lang="ch", use_gpu=False, show_log=False, enable_mkldnn=False)

            results = []
            total = len(self.pdf_doc)

            for i in range(total):
                self.progress.emit(int((i / total) * 100), f"正在识别第 {i + 1}/{total} 页...")
                page = self.pdf_doc[i]
                config = self.page_configs.get(i, {})

                row_data = {}
                for field_name, pdf_rect_data in config.items():
                    if len(pdf_rect_data) >= 4:
                        pdf_x, pdf_y, pdf_w, pdf_h = pdf_rect_data[:4]
                    else:
                        pdf_x, pdf_y, pdf_w, pdf_h = 100, 100, 150, 40

                    clip_rect = fitz.Rect(pdf_x, pdf_y, pdf_x + pdf_w, pdf_y + pdf_h)
                    pix = page.get_pixmap(matrix=fitz.Matrix(3.0, 3.0), clip=clip_rect, alpha=False)
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                    res = ocr.ocr(np.array(img), cls=True)

                    if not res or res[0] is None:
                        row_data[field_name] = ""
                    else:
                        row_data[field_name] = " ".join([line[1][0] for line in res[0]]).strip()

                results.append(row_data)

            self.progress.emit(100, "识别完成！")
            self.finished.emit(results)

        except Exception as e:
            import traceback
            print(traceback.format_exc())
            self.progress.emit(0, f"OCR 发生异常: {str(e)}")


class OCRExtractorWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.pdf_doc = None
        self.segments = []
        self.page_configs = {}
        self.extracted_data = []

        self.ocr_fields = ["图号", "图纸名称"]
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

        # 导入导出配置按钮
        hz_cfg = QHBoxLayout()
        btn_import_cfg = QPushButton("📂 导入历史配置")
        btn_import_cfg.clicked.connect(self.import_config)
        btn_export_cfg = QPushButton("💾 导出当前配置")
        btn_export_cfg.clicked.connect(self.export_config)
        hz_cfg.addWidget(btn_import_cfg)
        hz_cfg.addWidget(btn_export_cfg)
        l_left.addLayout(hz_cfg)

        self.file_manager = FileListManagerWidget(accept_exts=['.pdf'], title_desc="PDF Files (*.pdf)")
        l_left.addWidget(self.file_manager, 1)

        hz_fields = QHBoxLayout()
        btn_add = QPushButton("➕ 添加目标字段");
        btn_add.clicked.connect(self.add_field)
        btn_ren = QPushButton("✏️ 改名");
        btn_ren.clicked.connect(self.rename_field)
        btn_del = QPushButton("❌ 删除");
        btn_del.clicked.connect(self.delete_field)
        hz_fields.addWidget(btn_add);
        hz_fields.addWidget(btn_ren);
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
        hz_btns.addWidget(btn_detect);
        hz_btns.addWidget(btn_preview)
        l_left.addLayout(hz_btns)

        btn_extract = QPushButton("🚀 开始深度 OCR 提取")
        btn_extract.setStyleSheet(BTN_GREEN)
        btn_extract.clicked.connect(self.start_ocr)
        l_left.addWidget(btn_extract)

        self.lbl_status = QLabel("当前配置目标字段: " + ", ".join(self.ocr_fields))
        self.progress = QProgressBar()
        l_left.addWidget(self.lbl_status);
        l_left.addWidget(self.progress)
        box_left.setLayout(l_left)
        top_layout.addWidget(box_left, 1)

        box_right = QGroupBox("2. 数据表格与导出")
        l_right = QVBoxLayout()

        hz_tools = QHBoxLayout()
        btn_up = QPushButton("⬆️ 上移");
        btn_up.clicked.connect(self.move_row_up)
        btn_down = QPushButton("⬇️ 下移");
        btn_down.clicked.connect(self.move_row_down)
        btn_sort = QPushButton("🔤 图号排序");
        btn_sort.clicked.connect(self.sort_by_no)
        btn_merge_json = QPushButton("🔗 合并外部 JSON");
        btn_merge_json.clicked.connect(self.merge_json_files)
        hz_tools.addWidget(btn_up);
        hz_tools.addWidget(btn_down);
        hz_tools.addWidget(btn_sort);
        hz_tools.addStretch(1);
        hz_tools.addWidget(btn_merge_json)
        l_right.addLayout(hz_tools)

        hz_exp = QHBoxLayout()
        btn_exp_xls = QPushButton("导出 Excel/JSON");
        btn_exp_xls.clicked.connect(self.export_excel)
        btn_bmk = QPushButton("写入书签");
        btn_bmk.clicked.connect(self.write_bookmarks)
        btn_split = QPushButton("按名称拆分");
        btn_split.clicked.connect(self.split_pdf_by_name)
        hz_exp.addWidget(btn_exp_xls);
        hz_exp.addWidget(btn_bmk);
        hz_exp.addWidget(btn_split)
        l_right.addLayout(hz_exp)

        self.table = QTableWidget(0, 0)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)
        l_right.addWidget(self.table)

        box_right.setLayout(l_right)
        top_layout.addWidget(box_right, 2)
        splitter.addWidget(top_widget)

        bottom_widget = QWidget()
        bottom_layout = QVBoxLayout(bottom_widget)
        bottom_layout.setContentsMargins(0, 0, 0, 0)

        self.btn_confirm = QPushButton("✅ 确认本页 OCR 框位置并应用到该尺寸所有图纸")
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

    # ================= 导入导出配置核心逻辑 =================
    def export_config(self):
        if not self.pdf_doc or not self.page_configs:
            return QMessageBox.warning(self, "错误", "当前没有完成任何框配置，无法导出。请先设置框。")

        # 将各页码的配置，转化为物理尺寸的字典
        size_configs = {}
        for p_idx, config in self.page_configs.items():
            if not config: continue
            page = self.pdf_doc[p_idx]
            # 宽x高 四舍五入，作为该尺寸的唯一指纹
            size_key = f"{int(round(page.rect.width))}x{int(round(page.rect.height))}"
            # 只要这个尺寸还没存过，就存第一份
            if size_key not in size_configs:
                size_configs[size_key] = config

        export_data = {
            "module": "OCR",
            "fields": self.ocr_fields,
            "size_configs": size_configs
        }

        path, _ = QFileDialog.getSaveFileName(self, "保存 OCR 配置", "OCR框配置.json", "JSON Files (*.json)")
        if path:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(export_data, f, ensure_ascii=False, indent=4)
            QMessageBox.information(self, "成功", "当前字段与各尺寸对应的框位置已成功保存！")

    def import_config(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择 OCR 配置文件", "", "JSON Files (*.json)")
        if not path: return

        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)

            if data.get("module") != "OCR":
                return QMessageBox.warning(self, "错误", "这不是有效的 OCR 配置文件！")

            self.ocr_fields = data.get("fields", [])
            self.lbl_status.setText("当前配置目标字段: " + ", ".join(self.ocr_fields))
            self.refresh_table_data()

            size_configs = data.get("size_configs", {})
            self.page_configs.clear()

            # 如果当前已经生成了预览，立即尝试智能匹配当前文档的物理尺寸
            matched_count = 0
            if self.pdf_doc:
                for i in range(len(self.pdf_doc)):
                    page = self.pdf_doc[i]
                    size_key = f"{int(round(page.rect.width))}x{int(round(page.rect.height))}"
                    if size_key in size_configs:
                        self.page_configs[i] = copy.deepcopy(size_configs[size_key])
                        matched_count += 1

                QMessageBox.information(self, "导入成功",
                                        f"配置已加载！目标字段已更新为 {len(self.ocr_fields)} 个。\n"
                                        f"在当前文档的 {len(self.pdf_doc)} 页中，成功匹配到 {matched_count} 页的已知尺寸。\n\n"
                                        f"请务必点击【① 智能分段设框】检查，未匹配的新尺寸需要手动设置一次！")
            else:
                QMessageBox.information(self, "导入成功",
                                        "目标字段已更新。\n由于当前未生成 PDF 预览，尺寸配置将在您生成预览并点击【智能分段】时生效。")

        except Exception as e:
            QMessageBox.critical(self, "读取失败", f"配置文件损坏或解析异常:\n{e}")

    def add_field(self):
        text, ok = QInputDialog.getText(self, "新增", "输入需要提取的目标字段(如 版本号、设计人):")
        if ok and text and text not in self.ocr_fields:
            self.ocr_fields.append(text)
            self.lbl_status.setText("当前配置目标字段: " + ", ".join(self.ocr_fields))
            self.refresh_table_data()
            self._warn_reset_segments()

    def rename_field(self):
        if not self.ocr_fields: return
        old_name, ok = QInputDialog.getItem(self, "选择字段", "选择要修改的字段:", self.ocr_fields, 0, False)
        if ok and old_name:
            new_name, ok2 = QInputDialog.getText(self, "重命名", f"将 '{old_name}' 重命名为:")
            if ok2 and new_name and new_name not in self.ocr_fields:
                idx = self.ocr_fields.index(old_name)
                self.ocr_fields[idx] = new_name
                self.lbl_status.setText("当前配置目标字段: " + ", ".join(self.ocr_fields))
                for row in self.extracted_data:
                    if old_name in row: row[new_name] = row.pop(old_name)
                self.refresh_table_data()
                self._warn_reset_segments()

    def delete_field(self):
        if not self.ocr_fields: return
        name, ok = QInputDialog.getItem(self, "删除字段", "选择要删除的字段:", self.ocr_fields, 0, False)
        if ok and name:
            self.ocr_fields.remove(name)
            self.lbl_status.setText("当前配置目标字段: " + ", ".join(self.ocr_fields))
            for row in self.extracted_data:
                if name in row: del row[name]
            self.refresh_table_data()
            self._warn_reset_segments()

    def _warn_reset_segments(self):
        if self.segments:
            QMessageBox.warning(self, "警告", "由于您修改了提取字段数量，必须重新点击【① 智能分段设框】才能生效！")
            self.segments = []
            self.page_configs = {}

    def refresh_table_data(self):
        headers = ["序号"] + self.ocr_fields
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setColumnWidth(0, 50)
        self.table.setRowCount(len(self.extracted_data))
        for row, item in enumerate(self.extracted_data):
            it_seq = QTableWidgetItem(str(row + 1))
            it_seq.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            it_seq.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 0, it_seq)
            for col, field in enumerate(self.ocr_fields):
                self.table.setItem(row, col + 1, QTableWidgetItem(item.get(field, "")))

    def get_final_data(self):
        data = []
        headers = ["序号"] + self.ocr_fields
        for row in range(self.table.rowCount()):
            row_data = {}
            for col in range(1, len(headers)):
                it = self.table.item(row, col)
                row_data[headers[col]] = it.text() if it else ""
            data.append(row_data)
        return data

    def merge_pdfs(self):
        if self.file_manager.count() == 0: return QMessageBox.warning(self, "提示", "请先添加PDF文件！")
        self.pdf_doc = fitz.Document()
        toc_list = []

        filepaths = self.file_manager.get_all_filepaths()
        for path in filepaths:
            doc = fitz.open(path)
            start_page = len(self.pdf_doc)
            toc_list.append([1, os.path.basename(path), start_page + 1])
            self.pdf_doc.insert_pdf(doc)
            doc.close()

        self.pdf_doc.set_toc(toc_list)
        self.preview_view.load_pdf(self.pdf_doc)
        QMessageBox.information(self, "成功", f"大纲预览生成，共 {len(self.pdf_doc)} 页")

    def detect_segments(self):
        if not self.pdf_doc: return QMessageBox.warning(self, "错误", "请先生成大纲预览！")
        if not self.ocr_fields: return QMessageBox.warning(self, "提示", "请至少添加一个目标字段！")

        self.segments = detect_smart_segments(self.pdf_doc)

        for seg in self.segments:
            first_page_idx = seg['pages'][0]


            if first_page_idx in self.page_configs and self.page_configs[first_page_idx]:
                seg['pos_pct'] = copy.deepcopy(self.page_configs[first_page_idx])
                seg['pos_set'] = True
            else:

                default_dict = {}
                for i, field in enumerate(self.ocr_fields):
                    default_dict[field] = (100, 100 + i * 50, 150, 40)
                seg['pos_pct'] = default_dict
                for p in seg['pages']:
                    self.page_configs[p] = copy.deepcopy(default_dict)

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
            self.page_configs[p] = copy.deepcopy(new_pos_dict)

        self.btn_confirm.hide()
        self.dialog.refresh_table()
        self.dialog.show()

    def enter_final_preview(self):
        if not self.pdf_doc: return QMessageBox.warning(self, "错误", "请先生成预览！")

        for p in range(len(self.pdf_doc)):
            if p not in self.page_configs:
                default_dict = {}
                for i, field in enumerate(self.ocr_fields):
                    default_dict[field] = (100, 100 + i * 50, 150, 40)
                self.page_configs[p] = default_dict

        self.preview_view.load_pdf(self.pdf_doc, mode='ocr_final', data_dict=self.page_configs)
        QMessageBox.information(self, "提示", "已进入终极预览，您可以拖动右下角调整大小！")

    def start_ocr(self):
        if not self.pdf_doc: return
        self.preview_view.save_current_page_state()
        self.page_configs.update(copy.deepcopy(self.preview_view.page_data_dict))

        for p in range(len(self.pdf_doc)):
            if p not in self.page_configs:
                default_dict = {}
                for i, field in enumerate(self.ocr_fields):
                    default_dict[field] = (100, 100 + i * 50, 150, 40)
                self.page_configs[p] = default_dict

        self.worker = OCRWorker(self.pdf_doc, self.page_configs)
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

    def sort_by_no(self):
        data = self.get_final_data()
        data.sort(key=lambda x: str(x.get(self.ocr_fields[0], "")))
        self.extracted_data = data
        self.refresh_table_data()

    def merge_json_files(self):
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

            for item in merged_list:
                for k in item.keys():
                    if k not in self.ocr_fields:
                        self.ocr_fields.append(k)

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
        path, _ = QFileDialog.getSaveFileName(self, "保存数据", "图纸目录.xlsx", "Excel (*.xlsx)")
        if path:
            export_data = [{"序号": i + 1, **row} for i, row in enumerate(data)]
            pd.DataFrame(export_data).to_excel(path, index=False)

            json_path = path.replace('.xlsx', '.json')
            headers = [{"name": c, "width": 40} for c in export_data[0].keys()]
            vals = [list(d.values()) for d in export_data]
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump({"headers": headers, "data": vals}, f, ensure_ascii=False, indent=2)
            QMessageBox.information(self, "成功", "Excel 与 JSON 均已生成！")

    def write_bookmarks(self):
        if not self.pdf_doc: return
        data = self.get_final_data()
        toc = []
        for i, row in enumerate(data):
            title = f"{list(row.values())[0]} {list(row.values())[1] if len(row) > 1 else ''}".strip()
            if not title: title = f"Page_{i + 1}"
            toc.append([1, title, i + 1])
        self.pdf_doc.set_toc(toc)
        path, _ = QFileDialog.getSaveFileName(self, "保存", "带书签的图纸.pdf", "PDF (*.pdf)")
        if path:
            self.pdf_doc.save(path)
            QMessageBox.information(self, "成功", "书签已写入！")

    def split_pdf_by_name(self):
        if not self.pdf_doc: return
        out_dir = QFileDialog.getExistingDirectory(self, "选择保存目录")
        if not out_dir: return
        data = self.get_final_data()

        prefix, _ = QInputDialog.getText(self, "前缀", "增加前缀(选填):")
        suffix, _ = QInputDialog.getText(self, "后缀", "增加后缀(选填):")

        for i, row in enumerate(data):
            title = f"{list(row.values())[0]} {list(row.values())[1] if len(row) > 1 else ''}".strip()
            if not title: title = f"Page_{i + 1}"
            import re
            title = re.sub(r'[\\/*?:"<>|]', '_', title)
            final_name = f"{prefix}{title}{suffix}.pdf"

            new_doc = fitz.Document()
            new_doc.insert_pdf(self.pdf_doc, from_page=i, to_page=i)
            new_doc.save(os.path.join(out_dir, final_name))
            new_doc.close()

        QMessageBox.information(self, "成功", "按名称拆分完成！")