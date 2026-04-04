import os
import copy
import uuid
import subprocess
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                             QGroupBox, QLineEdit, QRadioButton, QMessageBox, QFileDialog, QProgressBar, QSplitter,
                             QTableWidget, QTableWidgetItem, QHeaderView, QCheckBox, QComboBox, QApplication,
                             QAbstractItemView)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
import fitz
from core.pdf_viewer import PDFGraphicsView
from core.ui_components import FileListManagerWidget
from core.utils import detect_smart_segments, UniversalSegmentDialog, find_ghostscript, run_ghostscript, \
    get_unique_filepath, BTN_BLUE, BTN_GREEN, BTN_PURPLE, BTN_RED, BTN_GRAY, BTN_ORANGE

MM_TO_PTS = 72 / 25.4


# ================= 后台处理线程 =================
class ImageInserterWorker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, pdf_doc, page_configs, export_mode, output_path, original_filenames, prefix="", suffix="",
                 use_gs=False, gs_quality="/ebook", gs_path=None, gs_lib_path=None):
        super().__init__()
        self.pdf_doc = pdf_doc
        self.page_configs = page_configs
        self.export_mode = export_mode
        self.output_path = output_path
        self.original_filenames = original_filenames
        self.prefix = prefix
        self.suffix = suffix
        self.use_gs = use_gs
        self.gs_quality = gs_quality
        self.gs_path = gs_path
        self.gs_lib_path = gs_lib_path

    def _process_single_document(self, doc, final_out_path, global_start_page):
        # 1. 逐页插入图片
        for local_page_num in range(len(doc)):
            global_page_num = global_start_page + local_page_num
            images_list = self.page_configs.get(global_page_num, [])
            page = doc[local_page_num]

            for img_info in images_list:
                img_path = img_info.get('path')
                if not img_path or not os.path.exists(img_path): continue

                w_pts = img_info['w'] * MM_TO_PTS
                h_pts = img_info['h'] * MM_TO_PTS
                pdf_x, pdf_y = img_info['pdf_x'], img_info['pdf_y']

                # 构造插入矩形
                rect = fitz.Rect(pdf_x, pdf_y, pdf_x + w_pts, pdf_y + h_pts)

                # 若需要处理旋转，可在 rect 映射前处理，此处沿用图章底层逻辑
                try:
                    page.insert_image(rect, filename=img_path, keep_proportion=False)
                except Exception as e:
                    print(f"插入图片异常: {e}")

        # 2. GS 压缩与保存逻辑
        tmp_visual = final_out_path + ".v.tmp.pdf"
        try:
            doc.save(tmp_visual)
            doc.close()
        except Exception as e:
            if not doc.is_closed: doc.close()
            raise Exception(f"保存中间文件失败: {e}")

        if self.use_gs and self.gs_path:
            try:
                run_ghostscript(self.gs_path, self.gs_lib_path, tmp_visual, final_out_path, quality=self.gs_quality)
                os.remove(tmp_visual)
            except Exception as e:
                print(f"GS压缩失败: {e}")
                os.rename(tmp_visual, final_out_path)
        else:
            os.rename(tmp_visual, final_out_path)

    def run(self):
        try:
            total_pages = len(self.pdf_doc)

            if self.export_mode == 'merged':
                export_doc = fitz.Document()
                export_doc.insert_pdf(self.pdf_doc)
                final_path = get_unique_filepath(os.path.dirname(self.output_path), os.path.basename(self.output_path))
                self._process_single_document(export_doc, final_path, 0)
                self.progress.emit(100)

            elif self.export_mode == 'batch':
                toc = self.pdf_doc.get_toc(simple=False)
                bookmarks = [item for item in toc if item[0] == 1]
                if not bookmarks: bookmarks = [[1, "Document", 1]]

                for i, bm in enumerate(bookmarks):
                    start_page = bm[2] - 1
                    end_page = bookmarks[i + 1][2] - 1 if i + 1 < len(bookmarks) else total_pages
                    new_doc = fitz.Document()
                    new_doc.insert_pdf(self.pdf_doc, from_page=start_page, to_page=end_page - 1)

                    original_name = self.original_filenames[i] if i < len(self.original_filenames) else f"Batch_{i}.pdf"
                    base_name = original_name.replace('.pdf', '')
                    final_name = f"{self.prefix}{base_name}{self.suffix}.pdf"
                    final_path = get_unique_filepath(self.output_path, final_name)

                    self._process_single_document(new_doc, final_path, start_page)
                    self.progress.emit(int((i + 1) / len(bookmarks) * 90))

            elif self.export_mode == 'split':
                for global_page_num in range(total_pages):
                    new_doc = fitz.Document()
                    new_doc.insert_pdf(self.pdf_doc, from_page=global_page_num, to_page=global_page_num)

                    # 尝试寻找其归属的源文件名
                    toc = self.pdf_doc.get_toc(simple=False)
                    bms = [item for item in toc if item[0] == 1]
                    base_name = "Page"
                    for i, bm in enumerate(bms):
                        if bm[2] - 1 <= global_page_num:
                            base_name = bm[1].replace('.pdf', '')

                    final_name = f"{self.prefix}{base_name}_第{global_page_num + 1}页{self.suffix}.pdf"
                    final_path = get_unique_filepath(self.output_path, final_name)

                    self._process_single_document(new_doc, final_path, global_page_num)
                    self.progress.emit(int((global_page_num + 1) / total_pages * 90))

            self.progress.emit(100)
            self.finished.emit("🎉 批量插图与导出任务完美完成！")
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.error.emit(str(e))


# ================= UI 界面 =================
class ImgInserterWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.pdf_doc = None
        self.segments = []
        self.page_configs = {}  # 页面绝对布局配置
        self.paired_data = []  # PDF与图片的匹配对
        self.original_filenames = []
        self.gs_path, self.gs_lib_path = find_ghostscript()
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(10, 10, 10, 10)
        splitter = QSplitter(Qt.Vertical)

        top_widget = QWidget()
        top_layout = QHBoxLayout(top_widget)
        top_layout.setContentsMargins(0, 0, 0, 0)

        # =============== 区域 1：双列文件池与智能配对 ===============
        box_left = QGroupBox("1. 文件池与一对一智能匹配")
        l_left = QVBoxLayout()

        hz_files = QHBoxLayout()
        # PDF 池
        v_pdf = QVBoxLayout()
        v_pdf.addWidget(QLabel("📄 目标 PDF 列表"))
        self.fm_pdf = FileListManagerWidget(accept_exts=['.pdf'], title_desc="PDF Files")
        v_pdf.addWidget(self.fm_pdf)
        hz_files.addLayout(v_pdf)

        # 图片 池
        v_img = QVBoxLayout()
        v_img.addWidget(QLabel("🖼️ 待插入图片池"))
        self.fm_img = FileListManagerWidget(accept_exts=['.png', '.jpg', '.jpeg'], title_desc="Images")
        v_img.addWidget(self.fm_img)
        hz_files.addLayout(v_img)
        l_left.addLayout(hz_files, 1)

        # 配对工具
        hz_pair = QHBoxLayout()
        btn_pair_name = QPushButton("🧠 按相同名称智能配对")
        btn_pair_name.setStyleSheet(BTN_BLUE)
        btn_pair_name.clicked.connect(self.pair_by_name)

        btn_pair_order = QPushButton("⬇️ 按上下顺序强制配对")
        btn_pair_order.setStyleSheet(BTN_ORANGE)
        btn_pair_order.clicked.connect(self.pair_by_order)

        hz_pair.addWidget(btn_pair_name)
        hz_pair.addWidget(btn_pair_order)
        l_left.addLayout(hz_pair)

        self.table_pair = QTableWidget(0, 2)
        self.table_pair.setHorizontalHeaderLabels(["目标 PDF 文件", "将插入的图片"])
        self.table_pair.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_pair.setSelectionBehavior(QAbstractItemView.SelectRows)
        l_left.addWidget(self.table_pair, 1)

        self.chk_all_pages = QCheckBox("☑️ 图片应用到该 PDF 的【所有页面】(不勾选则仅插首页)")
        self.chk_all_pages.setChecked(True)
        l_left.addWidget(self.chk_all_pages)

        btn_merge = QPushButton("🔄 确认配对并生成融合预览大纲")
        btn_merge.setStyleSheet(BTN_GREEN)
        btn_merge.clicked.connect(self.merge_and_build_config)
        l_left.addWidget(btn_merge)

        box_left.setLayout(l_left)
        top_layout.addWidget(box_left, 1)

        # =============== 区域 2：智能排布与导出 ===============
        box_right = QGroupBox("2. 智能分段设位、压缩与导出")
        l_right = QVBoxLayout()

        btn_detect = QPushButton("① 智能按图纸尺寸分段设位")
        btn_detect.setStyleSheet(BTN_GRAY)
        btn_detect.clicked.connect(self.detect_segments)

        btn_preview = QPushButton("② 进入终极自由微调预览")
        btn_preview.setStyleSheet(BTN_PURPLE)
        btn_preview.clicked.connect(self.enter_final_preview)
        l_right.addWidget(btn_detect)
        l_right.addWidget(btn_preview)

        l_right.addWidget(QLabel("导出拆分模式:"))
        self.radio_merged = QRadioButton("合并为单一长文件")
        self.radio_batch = QRadioButton("批量按原 PDF 独立导出")
        self.radio_batch.setChecked(True)
        self.radio_split = QRadioButton("拆分为单页独立文件")
        l_right.addWidget(self.radio_merged)
        l_right.addWidget(self.radio_batch)
        l_right.addWidget(self.radio_split)

        hz_fix = QHBoxLayout()
        self.input_prefix = QLineEdit();
        self.input_prefix.setPlaceholderText("前缀")
        self.input_suffix = QLineEdit();
        self.input_suffix.setPlaceholderText("后缀")
        hz_fix.addWidget(self.input_prefix)
        hz_fix.addWidget(self.input_suffix)
        l_right.addLayout(hz_fix)

        l_right.addStretch()

        hz_gs = QHBoxLayout()
        self.chk_gs = QCheckBox("🗜️ 启用 GS 全局强力压缩")
        self.cmb_gs_quality = QComboBox()
        self.cmb_gs_quality.addItems(["/screen (72dpi)", "/ebook (150dpi)", "/printer (300dpi)"])
        self.cmb_gs_quality.setCurrentIndex(1)
        if not self.gs_path: self.chk_gs.setEnabled(False)
        hz_gs.addWidget(self.chk_gs)
        hz_gs.addWidget(self.cmb_gs_quality)
        l_right.addLayout(hz_gs)

        btn_export = QPushButton("🚀 终极执行：可视化插入 ➡️ 压缩 ➡️ 导出")
        btn_export.setStyleSheet(BTN_GREEN)
        btn_export.clicked.connect(self.start_export)
        l_right.addWidget(btn_export)

        self.progress_bar = QProgressBar()
        l_right.addWidget(self.progress_bar)

        box_right.setLayout(l_right)
        top_layout.addWidget(box_right, 1)
        splitter.addWidget(top_widget)

        # =============== 下方：预览区 ===============
        bottom_widget = QWidget()
        bottom_layout = QVBoxLayout(bottom_widget)
        bottom_layout.setContentsMargins(0, 0, 0, 0)

        hz_preview_tools = QHBoxLayout()
        self.btn_confirm_pos = QPushButton("✅ 确认本页图片位置并应用到该尺寸所有图纸")
        self.btn_confirm_pos.setStyleSheet(BTN_RED)
        self.btn_confirm_pos.hide()
        self.btn_confirm_pos.clicked.connect(self.confirm_stamp_position)

        hz_preview_tools.addWidget(self.btn_confirm_pos)
        hz_preview_tools.addStretch()

        self.preview_view = PDFGraphicsView()
        bottom_layout.addLayout(hz_preview_tools)
        bottom_layout.addWidget(self.preview_view)

        splitter.addWidget(bottom_widget)
        splitter.setSizes([450, 550])
        main_layout.addWidget(splitter)

    # --- 匹配逻辑 ---
    def pair_by_name(self):
        pdfs = self.fm_pdf.get_all_filepaths()
        imgs = self.fm_img.get_all_filepaths()
        if not pdfs or not imgs: return QMessageBox.warning(self, "提示", "PDF 和图片池均不能为空！")

        self.paired_data.clear()
        img_dict = {os.path.splitext(os.path.basename(p))[0]: p for p in imgs}

        matched_count = 0
        for pdf in pdfs:
            pdf_name = os.path.splitext(os.path.basename(pdf))[0]
            if pdf_name in img_dict:
                self.paired_data.append((pdf, img_dict[pdf_name]))
                matched_count += 1
            else:
                self.paired_data.append((pdf, None))  # 未匹配留空

        self.refresh_table()
        QMessageBox.information(self, "智能匹配完成",
                                f"总计 {len(pdfs)} 个 PDF，成功同名匹配 {matched_count} 张图片！\n请在表格中检查。")

    def pair_by_order(self):
        pdfs = self.fm_pdf.get_all_filepaths()
        imgs = self.fm_img.get_all_filepaths()
        if not pdfs or not imgs: return QMessageBox.warning(self, "提示", "PDF 和图片池均不能为空！")

        self.paired_data.clear()
        for i, pdf in enumerate(pdfs):
            img_path = imgs[i] if i < len(imgs) else None
            self.paired_data.append((pdf, img_path))

        self.refresh_table()
        if len(pdfs) != len(imgs):
            QMessageBox.warning(self, "数量不一致",
                                f"PDF数量({len(pdfs)}) 与 图片数量({len(imgs)}) 不匹配！\n超出的部分将不会被插入图片。")

    def refresh_table(self):
        self.table_pair.setRowCount(len(self.paired_data))
        for i, (pdf, img) in enumerate(self.paired_data):
            self.table_pair.setItem(i, 0, QTableWidgetItem(os.path.basename(pdf)))
            img_name = os.path.basename(img) if img else "⚠️ 未匹配到图片"
            item_img = QTableWidgetItem(img_name)
            if not img: item_img.setForeground(Qt.red)
            self.table_pair.setItem(i, 1, item_img)

    # --- 融合与预览 ---
    def merge_and_build_config(self):
        if not self.paired_data: return QMessageBox.warning(self, "提示", "请先完成配对！")

        self.pdf_doc = fitz.Document()
        self.original_filenames = []
        self.page_configs.clear()
        toc_list = []

        apply_all = self.chk_all_pages.isChecked()
        global_page_counter = 0

        for pdf_path, img_path in self.paired_data:
            self.original_filenames.append(os.path.basename(pdf_path))
            doc = fitz.open(pdf_path)
            toc_list.append([1, os.path.basename(pdf_path), global_page_counter + 1])

            for local_idx in range(len(doc)):
                # 如果勾选了全应用，或者仅当前是首页，且匹配到了图片
                if img_path and (apply_all or local_idx == 0):
                    # 💡 核心：为每一页构建一个符合底层 stamp_final 格式的数据结构
                    self.page_configs[global_page_counter] = [{
                        'id': str(uuid.uuid4()),
                        'name': os.path.basename(img_path),
                        'path': img_path,
                        'w': 50,  # 默认初始宽高
                        'h': 50,
                        'pdf_x': 50,
                        'pdf_y': 50,
                        'angle': 0,
                        'lock_ratio': True
                    }]
                else:
                    self.page_configs[global_page_counter] = []

                global_page_counter += 1

            self.pdf_doc.insert_pdf(doc)
            doc.close()

        self.pdf_doc.set_toc(toc_list)
        self.preview_view.load_pdf(self.pdf_doc)
        self.progress_bar.setValue(0)
        QMessageBox.information(self, "成功",
                                f"生成配对大纲成功！共计 {len(self.pdf_doc)} 页。\n请点击右侧【智能分段设位】赋予初始位置。")

    # --- 智能分段复用逻辑 ---
    def detect_segments(self):
        if not self.pdf_doc: return QMessageBox.warning(self, "错误", "请先生成预览！")
        self.segments = detect_smart_segments(self.pdf_doc)

        # 为段落绑定默认位置 (仅同步 X,Y,W,H, 不修改各自关联的图片 Path)
        for seg in self.segments:
            seg['pos_set'] = False
            first_page = seg['pages'][0]
            if first_page in self.page_configs and self.page_configs[first_page]:
                # 记录这一段的通用几何模板
                geom = self.page_configs[first_page][0]
                seg['template_geom'] = {'w': geom['w'], 'h': geom['h'], 'pdf_x': geom['pdf_x'], 'pdf_y': geom['pdf_y']}
            else:
                seg['template_geom'] = {'w': 50, 'h': 50, 'pdf_x': 50, 'pdf_y': 50}

        self.dialog = UniversalSegmentDialog(self.segments, "设置该尺寸页面的图片默认位置", self)
        self.dialog.exec_()

    def enter_setting_mode_from_dialog(self, seg_data, idx, dialog):
        self.current_idx = idx
        self.dialog = dialog
        target_page = seg_data['pages'][0]

        # 即使目标页没有图片，也借用第一页造一个占位符去调位置
        if not self.page_configs.get(target_page):
            return QMessageBox.warning(self, "提示",
                                       "当前代表页恰好没有分配图片，无法调位置。请在左侧列表双击提取其他页！")

        self.preview_view.load_pdf(self.pdf_doc, target_page=target_page, mode='stamp_final',
                                   data_dict=self.page_configs)
        self.btn_confirm_pos.show()

    def confirm_stamp_position(self):
        self.preview_view.save_current_page_state()
        curr_page_cfg = self.page_configs[self.preview_view.current_page]
        if not curr_page_cfg: return

        # 提取用户刚刚调整好的模板几何信息
        new_geom = curr_page_cfg[0]

        self.segments[self.current_idx]['template_geom'] = {
            'w': new_geom['w'], 'h': new_geom['h'],
            'pdf_x': new_geom['pdf_x'], 'pdf_y': new_geom['pdf_y']
        }
        self.segments[self.current_idx]['pos_set'] = True

        # 将这个几何模板硬塞进该段落所有有图片的页面中 (保留各自的独立 path)
        for p in self.segments[self.current_idx]['pages']:
            if self.page_configs.get(p):
                self.page_configs[p][0].update(self.segments[self.current_idx]['template_geom'])

        self.btn_confirm_pos.hide()
        self.dialog.refresh_table()
        self.dialog.show()

    def enter_final_preview(self):
        if not self.pdf_doc: return QMessageBox.warning(self, "错误", "缺少PDF数据。")
        self.preview_view.load_pdf(self.pdf_doc, mode='stamp_final', data_dict=self.page_configs)
        self.btn_confirm_pos.hide()
        QMessageBox.information(self, "微调模式", "您可以滚动查看每一页，如果发现某张图片位置特殊，可直接独立拖拽它！")

    def start_export(self):
        if not self.pdf_doc: return
        self.preview_view.save_current_page_state()

        mode = 'merged' if self.radio_merged.isChecked() else ('batch' if self.radio_batch.isChecked() else 'split')

        if mode == 'merged':
            path, _ = QFileDialog.getSaveFileName(self, "保存", "合并批量加图文件.pdf", "PDF (*.pdf)")
            if not path: return
        else:
            path = QFileDialog.getExistingDirectory(self, "选择保存目录")
            if not path: return

        use_gs = self.chk_gs.isChecked()
        quality_str = self.cmb_gs_quality.currentText().split(" ")[0]

        self.worker = ImageInserterWorker(
            pdf_doc=self.pdf_doc,
            page_configs=self.page_configs,
            export_mode=mode,
            output_path=path,
            original_filenames=self.original_filenames,
            prefix=self.input_prefix.text(),
            suffix=self.input_suffix.text(),
            use_gs=use_gs,
            gs_quality=quality_str,
            gs_path=self.gs_path,
            gs_lib_path=self.gs_lib_path
        )
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.finished.connect(lambda msg: QMessageBox.information(self, "大功告成", msg))
        self.worker.error.connect(lambda e: QMessageBox.critical(self, "出错了", e))
        self.worker.start()