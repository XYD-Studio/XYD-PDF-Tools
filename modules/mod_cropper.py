import os
import copy
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                             QGroupBox, QMessageBox, QFileDialog, QProgressBar, QSplitter)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QRectF
from PyQt5.QtGui import QPixmap, QImage, QPainter, QPen, QColor, QWheelEvent, QCursor
from PyQt5.QtWidgets import QGraphicsView, QGraphicsScene
import fitz

from core.ui_components import FileListManagerWidget
from core.utils import detect_smart_segments, UniversalSegmentDialog

BTN_BLUE = "background-color: #3498DB; color: white; font-weight: bold; padding: 8px; border-radius: 4px;"
BTN_GREEN = "background-color: #2ECC71; color: white; font-weight: bold; padding: 8px; border-radius: 4px;"
BTN_PURPLE = "background-color: #9B59B6; color: white; font-weight: bold; padding: 8px; border-radius: 4px;"
BTN_GRAY = "background-color: #ECF0F1; color: #2C3E50; font-weight: bold; padding: 6px; border-radius: 4px; border: 1px solid #BDC3C7;"
BTN_RED = "background-color: #E74C3C; color: white; font-weight: bold; padding: 10px; border-radius: 4px;"


# ================= 自定义裁剪专用视图引擎 =================
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

        # 交互状态
        self.drag_line = None  # ('v', index) 或 ('h', index)
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

        pix = page.get_pixmap(matrix=fitz.Matrix(self.RENDER_SCALE, self.RENDER_SCALE))
        img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)

        self.scene.clear()
        self.page_item = self.scene.addPixmap(QPixmap.fromImage(img))
        self.page_item.setZValue(-1)
        self.scene.setSceneRect(QRectF(self.page_item.boundingRect()))
        self.resetTransform()
        self.scale(self.zoom_factor, self.zoom_factor)

        # 确保初始化数据
        if self.current_page not in self.page_data_dict:
            self.page_data_dict[self.current_page] = {'v_lines': [0.5], 'h_lines': [], 'disabled': []}

        self.pageChanged.emit(self.current_page)
        self.viewport().update()  # 触发 drawForeground重绘

    # 专门的前景绘制：画红线与阴影格子
    def drawForeground(self, painter, rect):
        if self.current_page not in self.page_data_dict or not self.page_item: return
        data = self.page_data_dict[self.current_page]
        w = self.page_item.boundingRect().width()
        h = self.page_item.boundingRect().height()

        v_lines = sorted([0.0] + data['v_lines'] + [1.0])
        h_lines = sorted([0.0] + data['h_lines'] + [1.0])

        # 1. 绘制废弃区域网格遮罩
        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor(0, 0, 0, 160))
        for r in range(len(h_lines) - 1):
            for c in range(len(v_lines) - 1):
                if f"{r},{c}" in data['disabled']:
                    x1, y1 = v_lines[c] * w, h_lines[r] * h
                    x2, y2 = v_lines[c + 1] * w, h_lines[r + 1] * h
                    painter.drawRect(QRectF(x1, y1, x2 - x1, y2 - y1))

                    # 画个红叉
                    painter.setPen(QPen(QColor(255, 50, 50, 200), 3))
                    cx, cy = (x1 + x2) / 2, (y1 + y2) / 2
                    sz = min(x2 - x1, y2 - y1) / 4
                    if sz > 10:
                        painter.drawLine(int(cx - sz), int(cy - sz), int(cx + sz), int(cy + sz))
                        painter.drawLine(int(cx - sz), int(cy + sz), int(cx + sz), int(cy - sz))
                    painter.setPen(Qt.NoPen)

        # 2. 绘制分割红线
        painter.setPen(QPen(Qt.red, 2))
        for vx in data['v_lines']:
            painter.drawLine(int(vx * w), 0, int(vx * w), int(h))
        for hy in data['h_lines']:
            painter.drawLine(0, int(hy * h), int(w), int(hy * h))

    # --- 交互核心 ---
    def mousePressEvent(self, event):
        if self.current_page not in self.page_data_dict: return
        pos = self.mapToScene(event.pos())
        w = self.page_item.boundingRect().width()
        h = self.page_item.boundingRect().height()
        data = self.page_data_dict[self.current_page]

        # 判定是否点在红线上 (吸附范围 10像素)
        hit_range = 10 / self.transform().m11()

        # 右键点击删除线
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
            # 抓取竖线
            for i, vx in enumerate(data['v_lines']):
                if abs(pos.x() - vx * w) < hit_range:
                    self.drag_line = ('v', i);
                    return
            # 抓取横线
            for i, hy in enumerate(data['h_lines']):
                if abs(pos.y() - hy * h) < hit_range:
                    self.drag_line = ('h', i);
                    return

            # 没点到线，判定点到了哪个格子，切换禁用状态
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
        if event.modifiers() == Qt.ControlModifier:
            zoom = 1.15 if event.angleDelta().y() > 0 else 1 / 1.15
            self.zoom_factor *= zoom
            self.scale(zoom, zoom)
        else:
            if event.angleDelta().y() < 0:
                self.show_page(self.current_page + 1)
            else:
                self.show_page(self.current_page - 1)

    # 对外暴露的控制接口
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


# ================= 裁剪执行后台线程 =================
class CropWorker(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, pdf_doc, page_configs, mode, output_path):
        super().__init__()
        self.pdf_doc = pdf_doc
        self.page_configs = page_configs
        self.mode = mode
        self.output_path = output_path

    def run(self):
        try:
            total_pages = len(self.pdf_doc)
            from PIL import Image

            results_images = []  # 缓存 PIL Image 用于合并为 PDF
            img_counter = 1

            for p_idx in range(total_pages):
                self.progress.emit(int(p_idx / total_pages * 90), f"正在精细裁剪第 {p_idx + 1} 页...")
                page = self.pdf_doc[p_idx]
                w, h = page.rect.width, page.rect.height
                cfg = self.page_configs.get(p_idx, {'v_lines': [0.5], 'h_lines': [], 'disabled': []})

                v_lines = sorted([0.0] + cfg['v_lines'] + [1.0])
                h_lines = sorted([0.0] + cfg['h_lines'] + [1.0])

                for r in range(len(h_lines) - 1):
                    for c in range(len(v_lines) - 1):
                        if f"{r},{c}" not in cfg['disabled']:
                            # 计算绝对矩形 (为了保证高清，采用 4.0 放大倍率取图)
                            rect = fitz.Rect(v_lines[c] * w, h_lines[r] * h, v_lines[c + 1] * w, h_lines[r + 1] * h)

                            # 过滤极小面积防止切出坏图
                            if rect.width < 5 or rect.height < 5: continue

                            pix = page.get_pixmap(matrix=fitz.Matrix(4.0, 4.0), clip=rect, alpha=False)
                            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                            if self.mode == 'images':
                                save_name = f"Cropped_{img_counter:04d}.jpg"
                                img.save(os.path.join(self.output_path, save_name), quality=95)
                                img_counter += 1
                            else:
                                results_images.append(img)

            if self.mode == 'pdf' and results_images:
                self.progress.emit(95, "正在打包生成全新 PDF...")
                results_images[0].save(self.output_path, "PDF", resolution=150.0, save_all=True,
                                       append_images=results_images[1:])

            self.progress.emit(100, "裁剪完毕！")
            self.finished.emit("批量超级裁剪已完成！")

        except Exception as e:
            import traceback
            print(traceback.format_exc())
            self.error.emit(str(e))


# ================= UI 界面 =================
class CropperWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.pdf_doc = None
        self.segments = []
        self.page_configs = {}
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
        box_left.setLayout(l_left)
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

        # 视图工具栏
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
        tool_frame.setLayout(tool_l)
        l_right.addWidget(tool_frame)

        l_right.addStretch(1)

        hz_exp = QHBoxLayout()
        btn_exp_img = QPushButton("📸 导出为多张碎图")
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

        box_right.setLayout(l_right)
        top_layout.addWidget(box_right, 1)
        splitter.addWidget(top_widget)

        bottom_widget = QWidget()
        bottom_layout = QVBoxLayout(bottom_widget)
        bottom_layout.setContentsMargins(0, 0, 0, 0)

        self.btn_confirm_pos = QPushButton("✅ 确认本页分割线排布并应用到该尺寸所有页面")
        self.btn_confirm_pos.setStyleSheet(BTN_RED)
        self.btn_confirm_pos.hide()
        self.btn_confirm_pos.clicked.connect(self.confirm_crop_position)

        self.preview_view = CropperGraphicsView()
        bottom_layout.addWidget(self.btn_confirm_pos)
        bottom_layout.addWidget(self.preview_view)

        splitter.addWidget(bottom_widget)
        splitter.setSizes([400, 600])
        main_layout.addWidget(splitter)

    def merge_files(self):
        if self.file_manager.count() == 0: return QMessageBox.warning(self, "提示", "请先添加文件！")
        self.pdf_doc = fitz.Document()

        filepaths = self.file_manager.get_all_filepaths()
        for path in filepaths:
            ext = os.path.splitext(path)[1].lower()
            if ext == '.pdf':
                doc = fitz.open(path)
                self.pdf_doc.insert_pdf(doc)
                doc.close()
            else:
                # 把图片转成一页极清 PDF 塞进去
                img = fitz.open(path)
                rect = img[0].rect
                pdfbytes = img.convert_to_pdf()
                img.close()
                imgPDF = fitz.open("pdf", pdfbytes)
                self.pdf_doc.insert_pdf(imgPDF)
                imgPDF.close()

        self.preview_view.load_pdf(self.pdf_doc)
        QMessageBox.information(self, "成功", f"生成混合阵列完毕，共计 {len(self.pdf_doc)} 页。")

    def detect_segments(self):
        if not self.pdf_doc: return QMessageBox.warning(self, "错误", "请先生成合并预览！")

        self.segments = detect_smart_segments(self.pdf_doc)
        for seg in self.segments:
            default_dict = {'v_lines': [0.5], 'h_lines': [], 'disabled': []}
            seg['pos_pct'] = default_dict
            for p in seg['pages']:
                if p not in self.page_configs:
                    self.page_configs[p] = copy.deepcopy(default_dict)

        self.dialog = UniversalSegmentDialog(self.segments, "设置该尺寸页面的默认分割线", self)
        self.dialog.exec_()

    def enter_setting_mode_from_dialog(self, seg_data, idx, dialog):
        self.current_idx = idx
        self.dialog = dialog
        target_page = seg_data['pages'][0]

        self.preview_view.load_pdf(self.pdf_doc, target_page=target_page, data_dict={target_page: seg_data['pos_pct']})
        self.btn_confirm_pos.show()

    def confirm_crop_position(self):
        new_cfg = self.preview_view.page_data_dict[self.preview_view.current_page]
        self.segments[self.current_idx]['pos_pct'] = new_cfg
        self.segments[self.current_idx]['pos_set'] = True

        for p in self.segments[self.current_idx]['pages']:
            self.page_configs[p] = copy.deepcopy(new_cfg)

        self.btn_confirm_pos.hide()
        self.dialog.refresh_table()
        self.dialog.show()

    def enter_final_preview(self):
        if not self.pdf_doc: return QMessageBox.warning(self, "错误", "缺少数据。")
        for p in range(len(self.pdf_doc)):
            if p not in self.page_configs:
                self.page_configs[p] = {'v_lines': [0.5], 'h_lines': [], 'disabled': []}

        self.preview_view.load_pdf(self.pdf_doc, data_dict=self.page_configs)
        QMessageBox.information(self, "高级微调模式",
                                "已进入终极预览。\n您可以滚动翻页查看每一页；\n用上方加线按钮排线，左键拖拽红线，右键红线删除，点击格子切换保留/丢弃！")

    def start_crop(self, mode):
        if not self.pdf_doc: return
        self.page_configs.update(copy.deepcopy(self.preview_view.page_data_dict))

        if mode == 'pdf':
            path, _ = QFileDialog.getSaveFileName(self, "保存", "切片拼接图纸.pdf", "PDF (*.pdf)")
            if not path: return
        else:
            path = QFileDialog.getExistingDirectory(self, "选择保存目录")
            if not path: return

        self.worker = CropWorker(self.pdf_doc, self.page_configs, mode, path)
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.finished.connect(lambda msg: QMessageBox.information(self, "大功告成", msg))
        self.worker.error.connect(lambda e: QMessageBox.critical(self, "严重错误", e))
        self.worker.start()