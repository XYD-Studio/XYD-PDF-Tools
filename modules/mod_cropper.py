import os
import copy
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                             QGroupBox, QMessageBox, QFileDialog, QProgressBar, QSplitter, QLineEdit)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QRectF
from PyQt5.QtGui import QPixmap, QImage, QPainter, QPen, QColor, QWheelEvent, QCursor
from PyQt5.QtWidgets import QGraphicsView, QGraphicsScene
import fitz

from core.ui_components import FileListManagerWidget
from core.utils import (detect_smart_segments, UniversalSegmentDialog, merge_pdf_with_smart_toc,
                        get_unique_filepath, BTN_BLUE, BTN_GREEN, BTN_PURPLE, BTN_GRAY, BTN_RED)


class CropperGraphicsView(QGraphicsView):
    pageChanged = pyqtSignal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.scene = QGraphicsScene(self)
        self.setScene(self.scene)
        self.setRenderHint(QPainter.Antialiasing)
        self.setRenderHint(QPainter.SmoothPixmapTransform)
        self.setMouseTracking(True)

        self.pdf_doc = None
        self.current_page = -1
        self.zoom_factor = 1.0
        self.page_item = None
        self.page_data_dict = {}
        self.drag_line = None
        self.RENDER_SCALE = 2.0

    def load_pdf(self, doc, target_page=0, data_dict=None):
        self.pdf_doc = doc
        self.page_data_dict = data_dict if data_dict else {}
        self.current_page = -1
        self.show_page(target_page)

    def show_page(self, page_num):
        if not self.pdf_doc or page_num < 0 or page_num >= len(self.pdf_doc): return
        self.current_page = page_num
        page = self.pdf_doc[page_num]

        # 💡 核心：保留透明通道
        pix = page.get_pixmap(matrix=fitz.Matrix(self.RENDER_SCALE, self.RENDER_SCALE), alpha=True)
        img_format = QImage.Format_RGBA8888 if pix.alpha else QImage.Format_RGB888
        img = QImage(pix.samples, pix.width, pix.height, pix.stride, img_format)

        self.scene.clear()
        self.page_item = self.scene.addPixmap(QPixmap.fromImage(img))
        self.page_item.setZValue(-1)
        self.scene.setSceneRect(QRectF(self.page_item.boundingRect()))
        self.resetTransform()
        self.scale(self.zoom_factor, self.zoom_factor)

        if self.current_page not in self.page_data_dict:
            self.page_data_dict[self.current_page] = {'v_lines': [0.5], 'h_lines': [], 'disabled': []}

        self.pageChanged.emit(self.current_page)
        self.viewport().update()

    def drawForeground(self, painter, rect):
        if self.current_page not in self.page_data_dict or not self.page_item: return
        data = self.page_data_dict[self.current_page]
        w = self.page_item.boundingRect().width()
        h = self.page_item.boundingRect().height()

        v_lines = sorted([0.0] + data['v_lines'] + [1.0])
        h_lines = sorted([0.0] + data['h_lines'] + [1.0])

        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor(0, 0, 0, 160))
        for r in range(len(h_lines) - 1):
            for c in range(len(v_lines) - 1):
                if f"{r},{c}" in data['disabled']:
                    x1, y1 = v_lines[c] * w, h_lines[r] * h
                    x2, y2 = v_lines[c + 1] * w, h_lines[r + 1] * h
                    painter.drawRect(QRectF(x1, y1, x2 - x1, y2 - y1))
                    painter.setPen(QPen(QColor(255, 50, 50, 200), 3))
                    cx, cy = (x1 + x2) / 2, (y1 + y2) / 2
                    sz = min(x2 - x1, y2 - y1) / 4
                    if sz > 10:
                        painter.drawLine(int(cx - sz), int(cy - sz), int(cx + sz), int(cy + sz))
                        painter.drawLine(int(cx - sz), int(cy + sz), int(cx + sz), int(cy - sz))
                    painter.setPen(Qt.NoPen)

        painter.setPen(QPen(Qt.red, 2))
        for vx in data['v_lines']:
            painter.drawLine(int(vx * w), 0, int(vx * w), int(h))
        for hy in data['h_lines']:
            painter.drawLine(0, int(hy * h), int(w), int(hy * h))

    def mousePressEvent(self, event):
        if self.current_page not in self.page_data_dict: return
        pos = self.mapToScene(event.pos())
        w = self.page_item.boundingRect().width()
        h = self.page_item.boundingRect().height()
        data = self.page_data_dict[self.current_page]
        hit_range = 10 / self.transform().m11()

        if event.button() == Qt.RightButton:
            for i, vx in enumerate(data['v_lines']):
                if abs(pos.x() - vx * w) < hit_range:
                    data['v_lines'].pop(i);
                    data['disabled'].clear();
                    self.viewport().update();
                    return
            for i, hy in enumerate(data['h_lines']):
                if abs(pos.y() - hy * h) < hit_range:
                    data['h_lines'].pop(i);
                    data['disabled'].clear();
                    self.viewport().update();
                    return
            return

        if event.button() == Qt.LeftButton:
            for i, vx in enumerate(data['v_lines']):
                if abs(pos.x() - vx * w) < hit_range:
                    self.drag_line = ('v', i);
                    return
            for i, hy in enumerate(data['h_lines']):
                if abs(pos.y() - hy * h) < hit_range:
                    self.drag_line = ('h', i);
                    return

            rx, ry = max(0, min(1, pos.x() / w)), max(0, min(1, pos.y() / h))
            v_lines = sorted([0.0] + data['v_lines'] + [1.0])
            h_lines = sorted([0.0] + data['h_lines'] + [1.0])
            col = next((i for i in range(len(v_lines) - 1) if v_lines[i] <= rx <= v_lines[i + 1]), -1)
            row = next((i for i in range(len(h_lines) - 1) if h_lines[i] <= ry <= h_lines[i + 1]), -1)

            if col != -1 and row != -1:
                key = f"{row},{col}"
                if key in data['disabled']:
                    data['disabled'].remove(key)
                else:
                    data['disabled'].append(key)
                self.viewport().update()

    def mouseMoveEvent(self, event):
        if self.drag_line and self.page_item:
            pos = self.mapToScene(event.pos())
            w = self.page_item.boundingRect().width()
            h = self.page_item.boundingRect().height()
            data = self.page_data_dict[self.current_page]
            axis, idx = self.drag_line
            if axis == 'v':
                data['v_lines'][idx] = max(0.01, min(0.99, pos.x() / w))
            else:
                data['h_lines'][idx] = max(0.01, min(0.99, pos.y() / h))
            self.viewport().update()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        self.drag_line = None

    def wheelEvent(self, event: QWheelEvent):
        pass

    def add_v_line(self):
        if self.current_page in self.page_data_dict:
            self.page_data_dict[self.current_page]['v_lines'].append(0.5)
            self.page_data_dict[self.current_page]['disabled'].clear()
            self.viewport().update()

    def add_h_line(self):
        if self.current_page in self.page_data_dict:
            self.page_data_dict[self.current_page]['h_lines'].append(0.5)
            self.page_data_dict[self.current_page]['disabled'].clear()
            self.viewport().update()

    def reset_lines(self):
        if self.current_page in self.page_data_dict:
            self.page_data_dict[self.current_page] = {'v_lines': [0.5], 'h_lines': [], 'disabled': []}
            self.viewport().update()


class CropperPreviewWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.view = CropperGraphicsView()
        self.layout.addWidget(self.view, 1)

        self.nav_layout = QHBoxLayout()
        self.btn_zoom_out = QPushButton("➖ 缩小")
        self.btn_zoom_fit = QPushButton("🔲 适应窗口")
        self.btn_zoom_in = QPushButton("➕ 放大")
        self.btn_prev = QPushButton("◀ 上一页")
        self.btn_next = QPushButton("下一页 ▶")
        self.entry_page = QLineEdit()
        self.entry_page.setFixedWidth(60)
        self.entry_page.setAlignment(Qt.AlignCenter)
        self.lbl_total = QLabel("/ 0 页")

        nav_btn_style = "padding: 5px 15px; font-weight: bold; background-color: #3498DB; color: white; border-radius: 4px;"
        zoom_btn_style = "padding: 5px 12px; font-weight: bold; background-color: #95A5A6; color: white; border-radius: 4px;"
        self.btn_prev.setStyleSheet(nav_btn_style);
        self.btn_next.setStyleSheet(nav_btn_style)
        self.btn_zoom_out.setStyleSheet(zoom_btn_style);
        self.btn_zoom_fit.setStyleSheet(zoom_btn_style);
        self.btn_zoom_in.setStyleSheet(zoom_btn_style)
        self.entry_page.setStyleSheet("padding: 5px; border: 1px solid #BDC3C7; border-radius: 3px; font-weight: bold;")
        self.lbl_total.setStyleSheet("font-weight: bold; color: #2C3E50;")

        self.nav_layout.addWidget(self.btn_zoom_out);
        self.nav_layout.addWidget(self.btn_zoom_fit);
        self.nav_layout.addWidget(self.btn_zoom_in)
        self.nav_layout.addStretch(1)
        self.nav_layout.addWidget(self.btn_prev);
        self.nav_layout.addWidget(QLabel("当前第"));
        self.nav_layout.addWidget(self.entry_page)
        self.nav_layout.addWidget(self.lbl_total);
        self.nav_layout.addWidget(self.btn_next);
        self.nav_layout.addStretch(1)
        self.layout.addLayout(self.nav_layout)

        self.btn_prev.clicked.connect(self._go_prev);
        self.btn_next.clicked.connect(self._go_next)
        self.entry_page.returnPressed.connect(self._jump_page)
        self.view.pageChanged.connect(self._on_page_changed)
        self.btn_zoom_in.clicked.connect(self._zoom_in);
        self.btn_zoom_out.clicked.connect(self._zoom_out);
        self.btn_zoom_fit.clicked.connect(self._zoom_fit)

    @property
    def page_data_dict(self):
        return self.view.page_data_dict

    @property
    def current_page(self):
        return self.view.current_page

    def add_v_line(self):
        self.view.add_v_line()

    def add_h_line(self):
        self.view.add_h_line()

    def reset_lines(self):
        self.view.reset_lines()

    def load_pdf(self, doc, target_page=0, data_dict=None):
        self.view.load_pdf(doc, target_page, data_dict)
        self._zoom_fit()

    def _go_prev(self):
        if self.view.pdf_doc and self.view.current_page > 0: self.view.show_page(self.view.current_page - 1)

    def _go_next(self):
        if self.view.pdf_doc and self.view.current_page < len(self.view.pdf_doc) - 1: self.view.show_page(
            self.view.current_page + 1)

    def _jump_page(self):
        if not self.view.pdf_doc: return
        try:
            p = int(self.entry_page.text()) - 1
            if 0 <= p < len(self.view.pdf_doc):
                self.view.show_page(p)
            else:
                self.entry_page.setText(str(self.view.current_page + 1))
        except ValueError:
            self.entry_page.setText(str(self.view.current_page + 1))

    def _on_page_changed(self, page_idx):
        self.entry_page.setText(str(page_idx + 1))
        self.lbl_total.setText(f"/ 共 {len(self.view.pdf_doc)} 页" if self.view.pdf_doc else "/ 0 页")

    def _zoom_in(self):
        if not self.view.page_item: return
        self.view.zoom_factor *= 1.15;
        self.view.scale(1.15, 1.15)

    def _zoom_out(self):
        if not self.view.page_item: return
        self.view.zoom_factor /= 1.15;
        self.view.scale(1 / 1.15, 1 / 1.15)

    def _zoom_fit(self):
        if not self.view.page_item: return
        rect = self.view.page_item.boundingRect()
        view_rect = self.view.viewport().rect()
        if rect.width() == 0 or rect.height() == 0: return
        ratio = min(view_rect.width() / rect.width(), view_rect.height() / rect.height()) * 0.95
        self.view.resetTransform();
        self.view.zoom_factor = ratio;
        self.view.scale(ratio, ratio)
        self.view.centerOn(self.view.page_item)


class CropWorker(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    # 💡 接收源文件名映射字典
    def __init__(self, pdf_doc, page_configs, mode, output_path, page_to_filename):
        super().__init__()
        self.pdf_doc = pdf_doc
        self.page_configs = page_configs
        self.mode = mode
        self.output_path = output_path
        self.page_to_filename = page_to_filename

    def run(self):
        try:
            total_pages = len(self.pdf_doc)
            from PIL import Image
            results_images = []

            for p_idx in range(total_pages):
                self.progress.emit(int(p_idx / total_pages * 90), f"正在精细裁剪第 {p_idx + 1} 页...")
                page = self.pdf_doc[p_idx]
                w, h = page.rect.width, page.rect.height
                cfg = self.page_configs.get(p_idx, {'v_lines': [0.5], 'h_lines': [], 'disabled': []})

                v_lines = sorted([0.0] + cfg['v_lines'] + [1.0])
                h_lines = sorted([0.0] + cfg['h_lines'] + [1.0])

                # 提取属于这一页的原文件名
                base_name = self.page_to_filename.get(p_idx, f"Page_{p_idx + 1}")
                piece_counter = 1

                for r in range(len(h_lines) - 1):
                    for c in range(len(v_lines) - 1):
                        if f"{r},{c}" not in cfg['disabled']:
                            rect = fitz.Rect(v_lines[c] * w, h_lines[r] * h, v_lines[c + 1] * w, h_lines[r + 1] * h)
                            if rect.width < 5 or rect.height < 5: continue

                            # 💡 核心修复：带透明通道高清切割
                            pix = page.get_pixmap(matrix=fitz.Matrix(4.0, 4.0), clip=rect, alpha=True)
                            img_mode = "RGBA" if pix.alpha else "RGB"
                            img = Image.frombytes(img_mode, [pix.width, pix.height], pix.samples)

                            if self.mode == 'images':
                                # 💡 核心：原文件名+编号，并采用 png 无损带透明导出
                                save_name = f"{base_name}_切片_{piece_counter:02d}.png"
                                final_path = get_unique_filepath(self.output_path, save_name)
                                img.save(final_path, format="PNG")
                                piece_counter += 1
                            else:
                                results_images.append(img)

            if self.mode == 'pdf' and results_images:
                self.progress.emit(95, "正在打包生成全新 PDF...")
                # 统一转为 RGB 保存 PDF (PDF底层不支持纯 RGBA 图片直接合并包)
                rgb_images = []
                for img in results_images:
                    if img.mode == 'RGBA':
                        bg = Image.new('RGB', img.size, (255, 255, 255))
                        bg.paste(img, mask=img.split()[3])
                        rgb_images.append(bg)
                    else:
                        rgb_images.append(img.convert('RGB'))

                rgb_images[0].save(self.output_path, "PDF", resolution=150.0, save_all=True,
                                   append_images=rgb_images[1:])

            self.progress.emit(100, "裁剪完毕！")
            self.finished.emit("批量超级裁剪已完成！")
        except Exception as e:
            import traceback;
            print(traceback.format_exc())
            self.error.emit(str(e))


class CropperWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.pdf_doc = None
        self.segments = []
        self.page_configs = {}
        self.page_to_filename = {}  # 💡 记录页码所属的原文件名
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(10, 10, 10, 10)
        splitter = QSplitter(Qt.Vertical)

        top_widget = QWidget()
        top_layout = QHBoxLayout(top_widget)
        top_layout.setContentsMargins(0, 0, 0, 0)

        box_left = QGroupBox("1. 载入图片与 PDF 混合列阵")
        l_left = QVBoxLayout()
        self.file_manager = FileListManagerWidget(accept_exts=['.pdf', '.jpg', '.png', '.jpeg', '.bmp'],
                                                  title_desc="Files")
        l_left.addWidget(self.file_manager, 1)

        btn_merge = QPushButton("🔄 生成合并预览 (首选必点)")
        btn_merge.setStyleSheet(BTN_BLUE)
        btn_merge.clicked.connect(self.merge_files)
        l_left.addWidget(btn_merge)
        box_left.setLayout(l_left);
        top_layout.addWidget(box_left, 1)

        box_right = QGroupBox("2. 智能分段设线与极速导出")
        l_right = QVBoxLayout()
        btn_detect = QPushButton("① 智能按图纸尺寸分段设线")
        btn_detect.setStyleSheet(BTN_GRAY)
        btn_detect.clicked.connect(self.detect_segments)
        btn_preview = QPushButton("② 终极自由单页微调预览")
        btn_preview.setStyleSheet(BTN_PURPLE)
        btn_preview.clicked.connect(self.enter_final_preview)
        l_right.addWidget(btn_detect);
        l_right.addWidget(btn_preview)

        tool_frame = QGroupBox("单页手工切线工具 (微调状态可用)")
        tool_l = QHBoxLayout()
        btn_add_v = QPushButton("➕ 加竖线");
        btn_add_v.clicked.connect(lambda: self.preview_view.add_v_line())
        btn_add_h = QPushButton("➕ 加横线");
        btn_add_h.clicked.connect(lambda: self.preview_view.add_h_line())
        btn_reset = QPushButton("↺ 重置清空");
        btn_reset.clicked.connect(lambda: self.preview_view.reset_lines())
        tool_l.addWidget(btn_add_v);
        tool_l.addWidget(btn_add_h);
        tool_l.addWidget(btn_reset)
        tool_frame.setLayout(tool_l);
        l_right.addWidget(tool_frame);
        l_right.addStretch(1)

        hz_exp = QHBoxLayout()
        btn_exp_img = QPushButton("📸 导出为高清切片图")
        btn_exp_img.setStyleSheet("background-color: #f39c12; color: white; padding: 10px; font-weight: bold;")
        btn_exp_img.clicked.connect(lambda: self.start_crop('images'))
        btn_exp_pdf = QPushButton("📄 导出拼接为新 PDF")
        btn_exp_pdf.setStyleSheet(BTN_GREEN)
        btn_exp_pdf.clicked.connect(lambda: self.start_crop('pdf'))
        hz_exp.addWidget(btn_exp_img);
        hz_exp.addWidget(btn_exp_pdf)
        l_right.addLayout(hz_exp)

        self.progress_bar = QProgressBar()
        l_right.addWidget(self.progress_bar)

        box_right.setLayout(l_right);
        top_layout.addWidget(box_right, 1)
        splitter.addWidget(top_widget)

        bottom_widget = QWidget()
        bottom_layout = QVBoxLayout(bottom_widget)
        bottom_layout.setContentsMargins(0, 0, 0, 0)

        self.btn_confirm_pos = QPushButton("✅ 确认本页分割线排布并应用到该尺寸所有页面")
        self.btn_confirm_pos.setStyleSheet(BTN_RED)
        self.btn_confirm_pos.hide()
        self.btn_confirm_pos.clicked.connect(self.confirm_crop_position)

        self.preview_view = CropperPreviewWidget()
        bottom_layout.addWidget(self.btn_confirm_pos);
        bottom_layout.addWidget(self.preview_view)

        splitter.addWidget(bottom_widget);
        splitter.setSizes([400, 600])
        main_layout.addWidget(splitter)

    def merge_files(self):
        if self.file_manager.count() == 0: return QMessageBox.warning(self, "提示", "请先添加文件！")
        self.pdf_doc = fitz.Document()
        filepaths = self.file_manager.get_all_filepaths()
        toc_list = []
        self.page_to_filename.clear()  # 💡 初始化映射
        global_page_idx = 0

        for path in filepaths:
            ext = os.path.splitext(path)[1].lower()
            base_name = os.path.splitext(os.path.basename(path))[0]
            if ext == '.pdf':
                doc = fitz.open(path)
                merge_pdf_with_smart_toc(doc, os.path.basename(path), self.pdf_doc, toc_list,
                                         prefer_filename_for_single=True)
                for _ in range(len(doc)):
                    self.page_to_filename[global_page_idx] = base_name
                    global_page_idx += 1
                doc.close()
            else:
                img = fitz.open(path)
                pdfbytes = img.convert_to_pdf()
                img.close()
                imgPDF = fitz.open("pdf", pdfbytes)
                merge_pdf_with_smart_toc(imgPDF, os.path.basename(path), self.pdf_doc, toc_list,
                                         prefer_filename_for_single=True)
                self.page_to_filename[global_page_idx] = base_name
                global_page_idx += 1
                imgPDF.close()

        self.pdf_doc.set_toc(toc_list)
        self.preview_view.load_pdf(self.pdf_doc)
        QMessageBox.information(self, "成功", f"生成混合阵列完毕，共计 {len(self.pdf_doc)} 页。")

    def detect_segments(self):
        if not self.pdf_doc: return QMessageBox.warning(self, "错误", "请先生成合并预览！")
        self.segments = detect_smart_segments(self.pdf_doc)
        for seg in self.segments:
            default_dict = {'v_lines': [0.5], 'h_lines': [], 'disabled': []}
            seg['pos_pct'] = default_dict
            for p in seg['pages']:
                if p not in self.page_configs: self.page_configs[p] = copy.deepcopy(default_dict)
        self.dialog = UniversalSegmentDialog(self.segments, "设置该尺寸页面的默认分割线", self)
        self.dialog.exec_()

    def enter_setting_mode_from_dialog(self, seg_data, idx, dialog):
        self.current_idx = idx;
        self.dialog = dialog
        target_page = seg_data['pages'][0]
        self.preview_view.load_pdf(self.pdf_doc, target_page=target_page, data_dict={target_page: seg_data['pos_pct']})
        self.btn_confirm_pos.show()

    def confirm_crop_position(self):
        new_cfg = self.preview_view.page_data_dict[self.preview_view.current_page]
        self.segments[self.current_idx]['pos_pct'] = new_cfg;
        self.segments[self.current_idx]['pos_set'] = True
        for p in self.segments[self.current_idx]['pages']: self.page_configs[p] = copy.deepcopy(new_cfg)
        self.btn_confirm_pos.hide();
        self.dialog.refresh_table();
        self.dialog.show()

    def enter_final_preview(self):
        if not self.pdf_doc: return QMessageBox.warning(self, "错误", "缺少数据。")
        for p in range(len(self.pdf_doc)):
            if p not in self.page_configs: self.page_configs[p] = {'v_lines': [0.5], 'h_lines': [], 'disabled': []}
        self.preview_view.load_pdf(self.pdf_doc, data_dict=self.page_configs)
        QMessageBox.information(self, "高级微调模式",
                                "已进入终极预览。\n您可以滚动翻页查看每一页；\n左键拖拽红线，右键红线删除，点击格子剔除区域！")

    def start_crop(self, mode):
        if not self.pdf_doc: return
        self.page_configs.update(copy.deepcopy(self.preview_view.page_data_dict))
        if mode == 'pdf':
            path, _ = QFileDialog.getSaveFileName(self, "保存", "切片拼接图纸.pdf", "PDF (*.pdf)")
            if not path: return
        else:
            path = QFileDialog.getExistingDirectory(self, "选择保存目录")
            if not path: return

        # 💡 传入原文件名称映射
        self.worker = CropWorker(self.pdf_doc, self.page_configs, mode, path, self.page_to_filename)
        self.worker.progress.connect(lambda v, txt: self.progress_bar.setValue(v))
        self.worker.finished.connect(lambda msg: QMessageBox.information(self, "大功告成", msg))
        self.worker.error.connect(lambda e: QMessageBox.critical(self, "严重错误", e))
        self.worker.start()