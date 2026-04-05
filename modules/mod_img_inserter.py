import os
import uuid
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                             QGroupBox, QLineEdit, QRadioButton, QMessageBox, QFileDialog, QProgressBar, QSplitter,
                             QTableWidget, QTableWidgetItem, QHeaderView, QCheckBox, QComboBox, QAbstractItemView)
from PyQt5.QtCore import QThread, pyqtSignal, Qt, QTimer
from PyQt5.QtGui import QPixmap
import fitz

from core.pdf_viewer import PDFGraphicsView
from core.ui_components import FileListManagerWidget
from core.utils import (detect_smart_segments, UniversalSegmentDialog, find_ghostscript,
                        run_ghostscript, get_unique_filepath, merge_pdf_with_smart_toc,
                        get_sub_toc, reinject_toc_after_gs,
                        BTN_BLUE, BTN_GREEN, BTN_PURPLE, BTN_RED, BTN_GRAY, BTN_ORANGE)

MM_TO_PTS = 72 / 25.4


# ================= 1. 绝对安全的后台处理线程 (全盘读取物理文件) =================
class ImageInserterWorker(QThread):
    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, paired_data, page_configs, export_mode, output_path, prefer_filename, prefix="", suffix="",
                 use_gs=False, gs_quality="/ebook", gs_path=None, gs_lib_path=None):
        super().__init__()
        self.paired_data = paired_data  # [(pdf_path, img_path), ...]
        self.page_configs = page_configs
        self.export_mode = export_mode
        self.output_path = output_path
        self.prefer_filename = prefer_filename
        self.prefix = prefix
        self.suffix = suffix
        self.use_gs = use_gs
        self.gs_quality = gs_quality
        self.gs_path = gs_path
        self.gs_lib_path = gs_lib_path

    def _finalize_document(self, doc, final_out_path, target_toc):
        tmp_visual = final_out_path + ".v.tmp.pdf"
        doc.save(tmp_visual)
        if self.use_gs and self.gs_path:
            tmp_gs = final_out_path + ".g.tmp.pdf"
            self.progress.emit(-1)  # 触发动画
            try:
                run_ghostscript(self.gs_path, self.gs_lib_path, tmp_visual, tmp_gs, quality=self.gs_quality)
                os.remove(tmp_visual)
                self.progress.emit(99)  # 停止动画
                self.status.emit("✅ 压缩完成！正在回写书签...")
                reinject_toc_after_gs(tmp_gs, target_toc)
                os.rename(tmp_gs, final_out_path)
            except Exception as e:
                os.rename(tmp_visual, final_out_path)
        else:
            os.rename(tmp_visual, final_out_path)

    def run(self):
        try:
            export_doc = fitz.Document() if self.export_mode == 'merged' else None
            toc_list = []
            global_page_counter = 0
            total_pdfs = len(self.paired_data)

            # 遍历物理文件进行组装，彻底与 UI 预览文档隔离
            for idx, (pdf_path, _) in enumerate(self.paired_data):
                self.status.emit(f"正在处理源文件: {os.path.basename(pdf_path)}")

                doc = fitz.open(pdf_path)
                total_local = len(doc)

                for local_page_num in range(total_local):
                    global_page_num = global_page_counter + local_page_num
                    images_list = self.page_configs.get(global_page_num, [])
                    page = doc[local_page_num]
                    for img_info in images_list:
                        img_p = img_info.get('path')
                        if not img_p or not os.path.exists(img_p): continue
                        rect = fitz.Rect(img_info['pdf_x'], img_info['pdf_y'],
                                         img_info['pdf_x'] + img_info['w'] * MM_TO_PTS,
                                         img_info['pdf_y'] + img_info['h'] * MM_TO_PTS)
                        try:
                            page.insert_image(rect, filename=img_p, keep_proportion=False)
                        except Exception:
                            pass

                if self.export_mode == 'merged':
                    merge_pdf_with_smart_toc(doc, os.path.basename(pdf_path), export_doc, toc_list,
                                             self.prefer_filename)

                elif self.export_mode == 'batch':
                    out_name = f"{self.prefix}{os.path.splitext(os.path.basename(pdf_path))[0]}{self.suffix}.pdf"
                    final_out_path = get_unique_filepath(self.output_path, out_name)
                    self._finalize_document(doc, final_out_path, doc.get_toc(simple=False))

                elif self.export_mode == 'split':
                    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
                    for local_page_num in range(total_local):
                        single_doc = fitz.Document()
                        single_doc.insert_pdf(doc, from_page=local_page_num, to_page=local_page_num)
                        single_toc = [[1, base_name, 1]] if self.prefer_filename else get_sub_toc(
                            doc.get_toc(simple=False), local_page_num, local_page_num)
                        single_doc.set_toc(single_toc)
                        out_name = f"{self.prefix}{base_name}_第{local_page_num + 1}页{self.suffix}.pdf"
                        final_out_path = get_unique_filepath(self.output_path, out_name)
                        self._finalize_document(single_doc, final_out_path, single_toc)
                        single_doc.close()

                doc.close()
                global_page_counter += total_local
                self.progress.emit(int((idx + 1) / total_pdfs * 80))

            if self.export_mode == 'merged':
                export_doc.set_toc(toc_list)
                final_out_path = get_unique_filepath(os.path.dirname(self.output_path),
                                                     os.path.basename(self.output_path))
                self._finalize_document(export_doc, final_out_path, toc_list)
                export_doc.close()

            self.progress.emit(100)
            self.finished.emit("🎉 所有任务已完美执行完毕！")

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.error.emit(str(e))


# ================= 2. 商业级 UI 主组件 =================
class ImgInserterWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.pdf_doc = None
        self.segments = []
        self.page_configs = {}
        self.paired_data = []
        self.original_page_img_map = {}
        self.gs_path, self.gs_lib_path = find_ghostscript()

        # 💡 安全的跑马灯动画
        self.fake_timer = QTimer(self)
        self.fake_timer.timeout.connect(self._on_fake_timer_tick)
        self.fake_val = 80.0
        self.fake_msgs = [
            "🗜️ 正在启动底层 Ghostscript 引擎进行深度重绘...",
            "⏳ 此过程可能极其耗时，请耐心等待...",
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

        box_left = QGroupBox("1. 文件池与一对一智能匹配")
        l_left = QVBoxLayout()
        hz_files = QHBoxLayout()
        v_pdf = QVBoxLayout()
        v_pdf.addWidget(QLabel("📄 目标 PDF 图纸列表"))
        self.fm_pdf = FileListManagerWidget(accept_exts=['.pdf'], title_desc="PDF Files")
        v_pdf.addWidget(self.fm_pdf)
        hz_files.addLayout(v_pdf)

        v_img = QVBoxLayout()
        v_img.addWidget(QLabel("🖼️ 待插入图片池 (二维码/防伪图等)"))
        self.fm_img = FileListManagerWidget(accept_exts=['.png', '.jpg', '.jpeg'], title_desc="Images")
        v_img.addWidget(self.fm_img)
        hz_files.addLayout(v_img)
        l_left.addLayout(hz_files, 1)

        hz_pair = QHBoxLayout()
        btn_pair_name = QPushButton("🧠 按相同名称智能配对")
        btn_pair_name.setStyleSheet(BTN_BLUE)
        btn_pair_name.clicked.connect(self.pair_by_name)
        btn_pair_order = QPushButton("⬇️ 按上下顺序强制配对")
        btn_pair_order.setStyleSheet(BTN_ORANGE)
        btn_pair_order.clicked.connect(self.pair_by_order)
        hz_pair.addWidget(btn_pair_name);
        hz_pair.addWidget(btn_pair_order)
        l_left.addLayout(hz_pair)

        self.table_pair = QTableWidget(0, 2)
        self.table_pair.setHorizontalHeaderLabels(["目标 PDF 文件", "专属配对图片"])
        self.table_pair.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_pair.setSelectionBehavior(QAbstractItemView.SelectRows)
        l_left.addWidget(self.table_pair, 1)

        self.chk_all_pages = QCheckBox("☑️ 图片应用到该 PDF 的【所有页面】(不勾选则仅插首页)")
        self.chk_all_pages.setChecked(True)
        l_left.addWidget(self.chk_all_pages)

        self.chk_toc_strategy = QCheckBox("🔖 合并时，单页 PDF 强制使用【文件名】作为书签")
        self.chk_toc_strategy.setChecked(True)
        self.chk_toc_strategy.setStyleSheet("color: #2C3E50; font-weight: bold;")
        l_left.addWidget(self.chk_toc_strategy)

        btn_merge = QPushButton("🔄 确认配对并生成融合预览大纲")
        btn_merge.setStyleSheet(BTN_GREEN)
        btn_merge.clicked.connect(self.merge_and_build_config)
        l_left.addWidget(btn_merge)

        box_left.setLayout(l_left);
        top_layout.addWidget(box_left, 1)

        box_right = QGroupBox("2. 智能分段设位、压缩与导出")
        l_right = QVBoxLayout()

        btn_detect = QPushButton("① 智能按图纸尺寸分段设位")
        btn_detect.setStyleSheet(BTN_GRAY)
        btn_detect.clicked.connect(self.detect_segments)
        btn_preview = QPushButton("② 进入终极自由微调预览")
        btn_preview.setStyleSheet(BTN_PURPLE)
        btn_preview.clicked.connect(self.enter_final_preview)
        l_right.addWidget(btn_detect);
        l_right.addWidget(btn_preview)

        l_right.addWidget(QLabel("导出拆分模式:"))
        self.radio_merged = QRadioButton("合并为单一长文件")
        self.radio_batch = QRadioButton("批量按原 PDF 独立导出")
        self.radio_batch.setChecked(True)
        self.radio_split = QRadioButton("拆分为单页独立文件")
        l_right.addWidget(self.radio_merged);
        l_right.addWidget(self.radio_batch);
        l_right.addWidget(self.radio_split)

        hz_fix = QHBoxLayout()
        self.input_prefix = QLineEdit();
        self.input_prefix.setPlaceholderText("前缀")
        self.input_suffix = QLineEdit();
        self.input_suffix.setPlaceholderText("后缀")
        hz_fix.addWidget(self.input_prefix);
        hz_fix.addWidget(self.input_suffix)
        l_right.addLayout(hz_fix);
        l_right.addStretch()

        hz_gs = QHBoxLayout()
        self.chk_gs = QCheckBox("🗜️ 启用 GS 全局强力压缩")
        self.cmb_gs_quality = QComboBox()
        self.cmb_gs_quality.addItems(["/screen (72dpi)", "/ebook (150dpi)", "/printer (300dpi)"])
        self.cmb_gs_quality.setCurrentIndex(1)
        if not self.gs_path: self.chk_gs.setEnabled(False)
        hz_gs.addWidget(self.chk_gs);
        hz_gs.addWidget(self.cmb_gs_quality)
        l_right.addLayout(hz_gs)

        self.btn_export = QPushButton("🚀 终极执行：后台批量插图 ➡️ 压缩 ➡️ 导出")
        self.btn_export.setStyleSheet(BTN_GREEN)
        self.btn_export.clicked.connect(self.start_export)
        l_right.addWidget(self.btn_export)

        self.lbl_status = QLabel("就绪")
        self.progress_bar = QProgressBar()
        l_right.addWidget(self.lbl_status);
        l_right.addWidget(self.progress_bar)

        box_right.setLayout(l_right);
        top_layout.addWidget(box_right, 1)
        splitter.addWidget(top_widget)

        bottom_widget = QWidget()
        bottom_layout = QVBoxLayout(bottom_widget)
        bottom_layout.setContentsMargins(0, 0, 0, 0)

        hz_preview_tools = QHBoxLayout()
        self.btn_confirm_pos = QPushButton("✅ 确认本页图片位置并应用到该尺寸全部图纸")
        self.btn_confirm_pos.setStyleSheet(BTN_RED)
        self.btn_confirm_pos.hide()
        self.btn_confirm_pos.clicked.connect(self.confirm_stamp_position)

        self.btn_recover_img = QPushButton("➕ 找回本页被误删的配对图片")
        self.btn_recover_img.setStyleSheet(BTN_BLUE)
        self.btn_recover_img.hide()
        self.btn_recover_img.clicked.connect(self.recover_current_page_image)

        hz_preview_tools.addWidget(self.btn_confirm_pos);
        hz_preview_tools.addWidget(self.btn_recover_img)
        hz_preview_tools.addStretch()

        self.preview_view = PDFGraphicsView()
        bottom_layout.addLayout(hz_preview_tools);
        bottom_layout.addWidget(self.preview_view)

        splitter.addWidget(bottom_widget);
        splitter.setSizes([450, 550])
        main_layout.addWidget(splitter)

    # --- 防拉伸算法 ---
    def _calculate_initial_image_geom(self, img_path, pdf_page):
        img_w_mm, img_h_mm = 50.0, 50.0
        px = QPixmap(img_path)
        if not px.isNull() and px.width() > 0 and px.height() > 0:
            img_w_mm = px.width() / MM_TO_PTS
            img_h_mm = px.height() / MM_TO_PTS
            pdf_w_mm = pdf_page.rect.width / MM_TO_PTS
            pdf_h_mm = pdf_page.rect.height / MM_TO_PTS
            if img_w_mm > pdf_w_mm or img_h_mm > pdf_h_mm:
                scale = min(pdf_w_mm / img_w_mm, pdf_h_mm / img_h_mm) * 0.9
                img_w_mm *= scale
                img_h_mm *= scale
        return img_w_mm, img_h_mm

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
                self.paired_data.append((pdf, None))
        self.refresh_table()
        QMessageBox.information(self, "智能匹配完成", f"总计 {len(pdfs)} 个 PDF，成功同名匹配 {matched_count} 张图片！")

    def pair_by_order(self):
        pdfs = self.fm_pdf.get_all_filepaths()
        imgs = self.fm_img.get_all_filepaths()
        if not pdfs or not imgs: return QMessageBox.warning(self, "提示", "PDF 和图片池均不能为空！")
        self.paired_data.clear()
        for i, pdf in enumerate(pdfs):
            img_path = imgs[i] if i < len(imgs) else None
            self.paired_data.append((pdf, img_path))
        self.refresh_table()
        if len(pdfs) != len(imgs): QMessageBox.warning(self, "数量不一致",
                                                       f"PDF数量({len(pdfs)}) 与 图片数量({len(imgs)}) 不匹配！")

    def refresh_table(self):
        self.table_pair.setRowCount(len(self.paired_data))
        for i, (pdf, img) in enumerate(self.paired_data):
            self.table_pair.setItem(i, 0, QTableWidgetItem(os.path.basename(pdf)))
            item_img = QTableWidgetItem(os.path.basename(img) if img else "⚠️ 未匹配")
            if not img: item_img.setForeground(Qt.red)
            self.table_pair.setItem(i, 1, item_img)

    def merge_and_build_config(self):
        if not self.paired_data: return QMessageBox.warning(self, "提示", "请先完成配对！")
        self.pdf_doc = fitz.Document()
        self.page_configs.clear()
        self.original_page_img_map.clear()
        toc_list = []
        apply_all = self.chk_all_pages.isChecked()
        prefer_filename = self.chk_toc_strategy.isChecked()
        global_page_counter = 0

        for pdf_path, img_path in self.paired_data:
            doc = fitz.open(pdf_path)
            for local_idx in range(len(doc)):
                if img_path and (apply_all or local_idx == 0) and os.path.exists(img_path):
                    img_w, img_h = self._calculate_initial_image_geom(img_path, doc[local_idx])
                    self.page_configs[global_page_counter] = [{
                        'id': str(uuid.uuid4()), 'name': os.path.basename(img_path),
                        'path': img_path, 'w': img_w, 'h': img_h,
                        'pdf_x': 10.0, 'pdf_y': 10.0, 'angle': 0, 'lock_ratio': True
                    }]
                    self.original_page_img_map[global_page_counter] = img_path
                else:
                    self.page_configs[global_page_counter] = []
                global_page_counter += 1
            merge_pdf_with_smart_toc(doc, os.path.basename(pdf_path), self.pdf_doc, toc_list, prefer_filename)
            doc.close()

        self.pdf_doc.set_toc(toc_list)
        self.preview_view.load_pdf(self.pdf_doc)
        self.progress_bar.setValue(0)
        QMessageBox.information(self, "成功",
                                f"大纲生成成功！共计 {len(self.pdf_doc)} 页。\n请点击右侧【智能分段设位】赋予初始位置。")

    # --- 智能排版 ---
    def detect_segments(self):
        if not self.pdf_doc: return QMessageBox.warning(self, "错误", "请先生成预览！")
        self.segments = detect_smart_segments(self.pdf_doc)
        for seg in self.segments:
            seg['pos_set'] = False
            first_page = seg['pages'][0]
            if first_page in self.page_configs and self.page_configs[first_page]:
                geom = self.page_configs[first_page][0]
                seg['template_geom'] = {'w': geom['w'], 'h': geom['h'], 'pdf_x': geom['pdf_x'], 'pdf_y': geom['pdf_y']}
            else:
                seg['template_geom'] = {'w': 50, 'h': 50, 'pdf_x': 10, 'pdf_y': 10}
        self.dialog = UniversalSegmentDialog(self.segments, "设置该尺寸页面的图片默认位置", self)
        self.dialog.exec_()

    def enter_setting_mode_from_dialog(self, seg_data, idx, dialog):
        self.current_idx = idx
        self.dialog = dialog
        target_page = seg_data['pages'][0]
        if not self.page_configs.get(target_page): return QMessageBox.warning(self, "提示",
                                                                              "当前页无图，请在列表中双击其他页！")
        self.preview_view.load_pdf(self.pdf_doc, target_page=target_page, mode='stamp_final',
                                   data_dict=self.page_configs)
        self.btn_confirm_pos.show();
        self.btn_recover_img.show()

    def confirm_stamp_position(self):
        self.preview_view.save_current_page_state()
        curr_page_cfg = self.page_configs[self.preview_view.current_page]
        if not curr_page_cfg: return
        new_geom = curr_page_cfg[0]
        self.segments[self.current_idx]['template_geom'] = {
            'w': new_geom['w'], 'h': new_geom['h'], 'pdf_x': new_geom['pdf_x'], 'pdf_y': new_geom['pdf_y']
        }
        self.segments[self.current_idx]['pos_set'] = True
        for p in self.segments[self.current_idx]['pages']:
            if self.page_configs.get(p):
                self.page_configs[p][0].update(self.segments[self.current_idx]['template_geom'])
        self.btn_confirm_pos.hide();
        self.btn_recover_img.hide()
        self.dialog.refresh_table();
        self.dialog.show()

    def enter_final_preview(self):
        if not self.pdf_doc: return QMessageBox.warning(self, "错误", "缺少PDF数据。")
        self.preview_view.load_pdf(self.pdf_doc, mode='stamp_final', data_dict=self.page_configs)
        self.btn_confirm_pos.hide();
        self.btn_recover_img.show()
        QMessageBox.information(self, "微调模式",
                                "可以滚动查看每一页，发现位置偏差可直接独立拖拽修正。\n右键图片可执行替换或删除。")

    def recover_current_page_image(self):
        curr_page = getattr(self.preview_view, 'current_page', -1)
        if curr_page < 0: return QMessageBox.warning(self, "提示", "请先定位到具体的页面。")

        img_path = self.original_page_img_map.get(curr_page)
        if not img_path or not os.path.exists(img_path): return QMessageBox.warning(self, "无法找回",
                                                                                    "该页面原本就无关联图片！")

        self.preview_view.save_current_page_state()
        if len(self.page_configs.get(curr_page, [])) > 0: return QMessageBox.information(self, "提示",
                                                                                         "当前页面已有图！若想换图，请右键选【替换图像】。")

        img_w, img_h = self._calculate_initial_image_geom(img_path, self.pdf_doc[curr_page])
        new_stamp = {
            'id': str(uuid.uuid4()), 'name': os.path.basename(img_path),
            'path': img_path, 'w': img_w, 'h': img_h,
            'pdf_x': 10.0, 'pdf_y': 10.0, 'angle': 0, 'lock_ratio': True
        }

        if curr_page not in self.page_configs: self.page_configs[curr_page] = []
        self.page_configs[curr_page].append(new_stamp)
        self.preview_view.load_pdf(self.pdf_doc, target_page=curr_page, mode='stamp_final', data_dict=self.page_configs)
        QMessageBox.information(self, "找回成功", "已完美恢复该页面专属配对图片！")

    # ================= UI 槽函数与导出 =================
    def _on_fake_timer_tick(self):
        if self.fake_val < 98.0:
            self.fake_val += (99.0 - self.fake_val) * 0.05
            self.progress_bar.setValue(int(self.fake_val))
        self.lbl_status.setText(self.fake_msgs[self.fake_msg_idx % len(self.fake_msgs)])
        self.fake_msg_idx += 1

    def _on_worker_progress(self, val):
        if val == -1:
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

    def start_export(self):
        if not self.paired_data: return QMessageBox.warning(self, "提示", "未检测到数据！")
        self.preview_view.save_current_page_state()

        mode = 'merged' if self.radio_merged.isChecked() else ('batch' if self.radio_batch.isChecked() else 'split')
        if mode == 'merged':
            path, _ = QFileDialog.getSaveFileName(self, "保存", "合并插图图纸.pdf", "PDF (*.pdf)")
            if not path: return
        else:
            path = QFileDialog.getExistingDirectory(self, "选择保存目录")
            if not path: return

        self.btn_export.setEnabled(False)  # 💡 禁用按钮防双击
        self.worker = ImageInserterWorker(
            paired_data=self.paired_data,  # 💡 将源文件列表直接丢给 Worker
            page_configs=self.page_configs,
            export_mode=mode,
            output_path=path,
            prefer_filename=self.chk_toc_strategy.isChecked(),
            prefix=self.input_prefix.text(),
            suffix=self.input_suffix.text(),
            use_gs=self.chk_gs.isChecked(),
            gs_quality=self.cmb_gs_quality.currentText().split(" ")[0],
            gs_path=self.gs_path,
            gs_lib_path=self.gs_lib_path
        )

        self.worker.progress.connect(self._on_worker_progress)
        self.worker.status.connect(self._on_worker_status)
        self.worker.finished.connect(self._on_worker_finished)
        self.worker.error.connect(self._on_worker_error)
        self.worker.start()