import copy
import uuid
import os
from PyQt5.QtWidgets import (QGraphicsView, QGraphicsScene, QGraphicsRectItem,
                             QGraphicsTextItem, QMenu, QInputDialog, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QLabel, QLineEdit, QFileDialog)
from PyQt5.QtCore import Qt, pyqtSignal, QRectF
from PyQt5.QtGui import QPixmap, QImage, QPainter, QWheelEvent, QPen, QColor, QFont
import fitz

MM_TO_PTS = 72 / 25.4
RENDER_SCALE = 2.0


class ResizableRectItem(QGraphicsRectItem):
    def __init__(self, w, h, name, color):
        super().__init__(0, 0, w, h)
        self.setAcceptHoverEvents(True)
        self.setFlag(QGraphicsRectItem.ItemIsMovable, True)
        self.setFlag(QGraphicsRectItem.ItemSendsGeometryChanges, True)
        self.setZValue(999)
        self.resizing = False
        self.setPen(QPen(QColor(color.red(), color.green(), color.blue(), 255), 3))
        self.setBrush(QColor(color.red(), color.green(), color.blue(), 50))
        self.label_bg = QGraphicsRectItem(self)
        self.label_bg.setBrush(QColor(255, 255, 255, 220))
        self.label_bg.setPen(QPen(Qt.NoPen))
        self.label = QGraphicsTextItem(name, self)
        self.label.setFont(QFont("Microsoft YaHei", 11, QFont.Bold))
        self.label.setDefaultTextColor(QColor(255, 0, 0))
        self.label.setPos(2, -25)
        self.label_bg.setRect(2, -25, self.label.boundingRect().width(), 25)

    def hoverMoveEvent(self, event):
        rect = self.rect()
        self.setCursor(
            Qt.SizeFDiagCursor if event.pos().x() > rect.right() - 15 and event.pos().y() > rect.bottom() - 15 else Qt.SizeAllCursor)
        super().hoverMoveEvent(event)

    def mousePressEvent(self, event):
        rect = self.rect()
        if event.pos().x() > rect.right() - 15 and event.pos().y() > rect.bottom() - 15: self.resizing = True
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self.resizing:
            rect = self.rect()
            new_w = max(20.0, event.pos().x() - rect.left())
            new_h = max(20.0, event.pos().y() - rect.top())
            self.setRect(rect.left(), rect.top(), new_w, new_h)
        else:
            super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        self.resizing = False
        super().mouseReleaseEvent(event)


class ResizableStampItem(QGraphicsRectItem):
    def __init__(self, stamp_info, view):
        w_scene = float(stamp_info.get('w', 50)) * MM_TO_PTS * RENDER_SCALE
        h_scene = float(stamp_info.get('h', 30)) * MM_TO_PTS * RENDER_SCALE
        super().__init__(0, 0, w_scene, h_scene)
        self.stamp_info = stamp_info
        self.view = view
        self.setAcceptHoverEvents(True)
        self.setFlag(QGraphicsRectItem.ItemIsMovable, True)
        self.setFlag(QGraphicsRectItem.ItemIsSelectable, True)
        self.setFlag(QGraphicsRectItem.ItemSendsGeometryChanges, True)
        self.setZValue(999)
        self.orig_pixmap = QPixmap()
        try:
            with open(stamp_info.get('path', ''), 'rb') as f:
                self.orig_pixmap.loadFromData(f.read())
        except Exception:
            self.orig_pixmap = QPixmap(int(w_scene), int(h_scene));
            self.orig_pixmap.fill(Qt.lightGray)
        self.setPen(QPen(Qt.NoPen))
        self.label_bg = QGraphicsRectItem(self)
        self.label_bg.setBrush(QColor(255, 255, 255, 200))
        self.label_bg.setPen(QPen(Qt.NoPen))
        self.label = QGraphicsTextItem(str(stamp_info.get('name', '未命名图章')), self)
        self.label.setFont(QFont("Microsoft YaHei", 11, QFont.Bold))
        self.label.setDefaultTextColor(QColor(255, 0, 0))
        self.label.setPos(2, -25)
        self.label_bg.setRect(2, -25, self.label.boundingRect().width(), 25)
        self.setTransformOriginPoint(self.rect().center())
        self.setRotation(stamp_info.get('angle', 0))
        self.lock_ratio = stamp_info.get('lock_ratio', True)
        self.orig_ratio = w_scene / h_scene if h_scene != 0 else 1.0
        self.resizing = False

    def paint(self, painter, option, widget):
        rect = self.rect()
        if not self.orig_pixmap.isNull():
            scaled_pixmap = self.orig_pixmap.scaled(int(rect.width()), int(rect.height()), Qt.IgnoreAspectRatio,
                                                    Qt.SmoothTransformation)
            painter.drawPixmap(rect.toRect(), scaled_pixmap)
        if self.isSelected() or self.isUnderMouse():
            painter.setPen(QPen(Qt.red, 2, Qt.DashLine));
            painter.setBrush(Qt.NoBrush);
            painter.drawRect(rect)
            painter.setBrush(Qt.red);
            painter.drawRect(int(rect.right() - 10), int(rect.bottom() - 10), 10, 10)

    def hoverMoveEvent(self, event):
        rect = self.rect()
        self.setCursor(
            Qt.SizeFDiagCursor if event.pos().x() > rect.right() - 15 and event.pos().y() > rect.bottom() - 15 else Qt.SizeAllCursor)
        super().hoverMoveEvent(event)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            if event.modifiers() == Qt.AltModifier:
                self.view.duplicate_stamp(self.stamp_info);
                event.accept();
                return
            rect = self.rect()
            if event.pos().x() > rect.right() - 15 and event.pos().y() > rect.bottom() - 15:
                self.resizing = True;
                event.accept();
                return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self.resizing:
            rect = self.rect()
            new_w = max(20.0, event.pos().x() - rect.left())
            new_h = max(20.0, event.pos().y() - rect.top())
            if self.lock_ratio: new_h = new_w / self.orig_ratio
            self.setRect(rect.left(), rect.top(), new_w, new_h)
            self.setTransformOriginPoint(self.rect().center())
            self.update()
        else:
            super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if self.resizing:
            self.resizing = False
            self.stamp_info['w'] = self.rect().width() / RENDER_SCALE / MM_TO_PTS
            self.stamp_info['h'] = self.rect().height() / RENDER_SCALE / MM_TO_PTS
            event.accept()
        else:
            super().mouseReleaseEvent(event)

    def contextMenuEvent(self, event):
        menu = QMenu()
        menu.setStyleSheet("QMenu { background-color: white; border: 1px solid #ccc; font-size: 14px; }")
        action_copy = menu.addAction("➕ 复制图像 (也可按住 Alt 拖动)")
        action_replace = menu.addAction("🔄 替换图像 (自动等比匹配最大边)")
        action_lock = menu.addAction("🔓 解锁等比例缩放" if self.lock_ratio else "🔒 锁定等比例缩放")
        action_rot_90 = menu.addAction("↻ 顺时针旋转90°")
        action_rot_cus = menu.addAction("⟳ 自定义旋转角度...")
        action_del = menu.addAction("❌ 删除该图像")

        action = menu.exec_(event.screenPos())
        if action == action_copy:
            self.view.duplicate_stamp(self.stamp_info)
        elif action == action_replace:
            file_path, _ = QFileDialog.getOpenFileName(None, "选择要替换的图片", "",
                                                       "Images (*.png *.jpg *.jpeg *.bmp)")
            if file_path:
                new_pixmap = QPixmap(file_path)
                if not new_pixmap.isNull():
                    max_dim = max(self.stamp_info['w'], self.stamp_info['h'])
                    new_ratio = new_pixmap.width() / new_pixmap.height() if new_pixmap.height() != 0 else 1.0
                    if new_pixmap.width() >= new_pixmap.height():
                        final_w, final_h = max_dim, max_dim / new_ratio
                    else:
                        final_h, final_w = max_dim, max_dim * new_ratio
                    self.stamp_info.update(
                        {'path': file_path, 'w': final_w, 'h': final_h, 'name': os.path.basename(file_path)})
                    with open(file_path, 'rb') as f:
                        self.orig_pixmap.loadFromData(f.read())
                    w_scene = final_w * MM_TO_PTS * RENDER_SCALE
                    h_scene = final_h * MM_TO_PTS * RENDER_SCALE
                    self.setRect(self.rect().left(), self.rect().top(), w_scene, h_scene)
                    self.setTransformOriginPoint(self.rect().center())
                    self.label.setPlainText(self.stamp_info['name'])
                    self.update()
        elif action == action_lock:
            self.lock_ratio = not self.lock_ratio;
            self.stamp_info['lock_ratio'] = self.lock_ratio;
            self.update()
        elif action == action_rot_90:
            new_angle = (self.rotation() + 90) % 360;
            self.setRotation(new_angle);
            self.stamp_info['angle'] = new_angle
        elif action == action_rot_cus:
            angle, ok = QInputDialog.getDouble(None, "自定义旋转", "输入角度:", self.rotation(), 0, 360, 1)
            if ok: self.setRotation(angle); self.stamp_info['angle'] = angle
        elif action == action_del:
            self.view.delete_stamp(self.stamp_info['id'])


class _PDFGraphicsView(QGraphicsView):
    pageChanged = pyqtSignal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.scene = QGraphicsScene(self);
        self.setScene(self.scene)
        self.setRenderHint(QPainter.Antialiasing);
        self.setRenderHint(QPainter.SmoothPixmapTransform)
        self.setDragMode(QGraphicsView.ScrollHandDrag)
        self.pdf_doc = None;
        self.current_page = -1;
        self.zoom_factor = 1.0
        self.page_item = None;
        self.mode = 'view'
        self.page_data_dict = {};
        self.dynamic_items = {}
        self.color_palette = [QColor(255, 0, 0), QColor(0, 0, 255), QColor(0, 200, 0), QColor(255, 140, 0),
                              QColor(128, 0, 128), QColor(0, 150, 255)]

    def load_pdf(self, doc, target_page=0, mode='view', data_dict=None, **kwargs):
        self.pdf_doc = doc;
        self.mode = mode;
        self.page_data_dict = data_dict if data_dict is not None else {}
        self.current_page = -1;
        self.page_item = None;
        self.dynamic_items.clear()
        self.show_page(target_page)

    def show_page(self, page_num):
        if not self.pdf_doc or page_num < 0 or page_num >= len(self.pdf_doc): return
        if self.current_page != -1: self.save_current_page_state()
        self.current_page = page_num
        page = self.pdf_doc[page_num]

        # 💡 核心修复：保留透明通道 alpha=True
        pix = page.get_pixmap(matrix=fitz.Matrix(RENDER_SCALE, RENDER_SCALE), alpha=True)
        # 根据是否存在 alpha 通道，智能选择 PyQt 渲染格式
        img_format = QImage.Format_RGBA8888 if pix.alpha else QImage.Format_RGB888
        img = QImage(pix.samples, pix.width, pix.height, pix.stride, img_format)

        self.scene.clear();
        self.dynamic_items.clear()
        self.page_item = self.scene.addPixmap(QPixmap.fromImage(img))
        self.page_item.setZValue(-1);
        self.scene.setSceneRect(QRectF(self.page_item.boundingRect()))
        self.resetTransform();
        self.scale(self.zoom_factor, self.zoom_factor)

        if self.current_page in self.page_data_dict:
            if self.mode == 'stamp_final':
                for stamp in self.page_data_dict[self.current_page]: self._draw_stamp(stamp)
            elif self.mode == 'ocr_final':
                self._draw_ocr_boxes(self.page_data_dict[self.current_page])
        self.pageChanged.emit(self.current_page);
        self.scene.update()

    def save_current_page_state(self):
        if not self.page_item or self.current_page not in self.page_data_dict: return
        if self.mode == 'stamp_final':
            new_stamps = []
            for stamp_id, item in self.dynamic_items.items():
                st_info = item.stamp_info
                st_info['pdf_x'] = item.pos().x() / RENDER_SCALE
                st_info['pdf_y'] = item.pos().y() / RENDER_SCALE
                new_stamps.append(st_info)
            self.page_data_dict[self.current_page] = new_stamps
        elif self.mode == 'ocr_final':
            new_pos = {}
            for field_name, item in self.dynamic_items.items():
                new_pos[field_name] = (
                item.pos().x() / RENDER_SCALE, item.pos().y() / RENDER_SCALE, item.rect().width() / RENDER_SCALE,
                item.rect().height() / RENDER_SCALE)
            self.page_data_dict[self.current_page] = new_pos

    def _draw_stamp(self, stamp_info):
        try:
            if not isinstance(stamp_info, dict): return
            item = ResizableStampItem(stamp_info, self)
            item.setPos(float(stamp_info.get('pdf_x', 100)) * RENDER_SCALE,
                        float(stamp_info.get('pdf_y', 100)) * RENDER_SCALE)
            self.scene.addItem(item)
            self.dynamic_items[stamp_info.get('id', str(uuid.uuid4()))] = item
        except Exception:
            pass

    def _draw_ocr_boxes(self, pos_dict):
        try:
            if not isinstance(pos_dict, dict): return
            for i, (field_name, data) in enumerate(pos_dict.items()):
                color = self.color_palette[i % len(self.color_palette)]
                pdf_x, pdf_y, pdf_w, pdf_h = map(float, data[:4]) if isinstance(data, (list, tuple)) and len(
                    data) >= 4 else (100.0, 100.0, 150.0, 40.0)
                item = ResizableRectItem(max(20.0, pdf_w * RENDER_SCALE), max(20.0, pdf_h * RENDER_SCALE),
                                         str(field_name), color)
                item.setPos(pdf_x * RENDER_SCALE, pdf_y * RENDER_SCALE)
                self.scene.addItem(item)
                self.dynamic_items[field_name] = item
        except Exception:
            pass

    def duplicate_stamp(self, original_info):
        new_stamp = copy.deepcopy(original_info)
        new_stamp['id'] = str(uuid.uuid4())
        current_item = self.dynamic_items.get(original_info['id'])
        if current_item:
            new_stamp['pdf_x'] = (current_item.pos().x() + 30) / RENDER_SCALE
            new_stamp['pdf_y'] = (current_item.pos().y() + 30) / RENDER_SCALE
            self._draw_stamp(new_stamp)
            if self.current_page in self.page_data_dict: self.page_data_dict[self.current_page].append(new_stamp)

    def delete_stamp(self, stamp_id):
        item = self.dynamic_items.get(stamp_id)
        if item:
            self.scene.removeItem(item);
            del self.dynamic_items[stamp_id]
            if self.current_page in self.page_data_dict:
                self.page_data_dict[self.current_page] = [s for s in self.page_data_dict[self.current_page] if
                                                          s.get('id') != stamp_id]

    def wheelEvent(self, event: QWheelEvent):
        if event.modifiers() == Qt.ControlModifier:
            zoom = 1.15 if event.angleDelta().y() > 0 else 1 / 1.15
            self.zoom_factor *= zoom;
            self.scale(zoom, zoom)
        else:
            self.show_page(self.current_page + 1 if event.angleDelta().y() < 0 else self.current_page - 1)


class PDFGraphicsView(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.layout = QVBoxLayout(self);
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.view = _PDFGraphicsView();
        self.layout.addWidget(self.view, 1)
        self.nav_layout = QHBoxLayout()
        self.btn_zoom_out = QPushButton("➖ 缩小");
        self.btn_zoom_fit = QPushButton("🔲 适应窗口");
        self.btn_zoom_in = QPushButton("➕ 放大")
        self.btn_prev = QPushButton("◀ 上一页");
        self.btn_next = QPushButton("下一页 ▶")
        self.entry_page = QLineEdit();
        self.entry_page.setFixedWidth(60);
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
        self.btn_next.clicked.connect(self._go_next);
        self.entry_page.returnPressed.connect(self._jump_page)
        self.view.pageChanged.connect(self._on_page_changed)
        self.btn_zoom_in.clicked.connect(lambda: self._do_zoom(1.2));
        self.btn_zoom_out.clicked.connect(lambda: self._do_zoom(1 / 1.2));
        self.btn_zoom_fit.clicked.connect(self._zoom_fit)

    def _do_zoom(self, factor):
        if self.view.page_item: self.view.zoom_factor *= factor; self.view.scale(factor, factor)

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

    def _zoom_fit(self):
        if not self.view.page_item: return
        rect, view_rect = self.view.page_item.boundingRect(), self.view.viewport().rect()
        ratio = min(view_rect.width() / rect.width(), view_rect.height() / rect.height()) * 0.95
        self.view.resetTransform();
        self.view.zoom_factor = ratio;
        self.view.scale(ratio, ratio)
        self.view.centerOn(self.view.page_item)

    def load_pdf(self, doc, target_page=0, mode='view', data_dict=None, **kwargs):
        self.view.load_pdf(doc, target_page, mode, data_dict, **kwargs);
        self._zoom_fit()

    def save_current_page_state(self):
        self.view.save_current_page_state()

    @property
    def page_data_dict(self):
        return self.view.page_data_dict

    @page_data_dict.setter
    def page_data_dict(self, val):
        self.view.page_data_dict = val

    @property
    def current_page(self):
        return self.view.current_page