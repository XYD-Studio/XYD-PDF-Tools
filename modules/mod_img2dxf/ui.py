# -*- coding: utf-8 -*-
import os
import cv2
import numpy as np
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                             QGroupBox, QSlider, QCheckBox, QFileDialog, QMessageBox, QProgressBar, QSplitter,
                             QComboBox, QGraphicsView, QGraphicsScene)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QImage, QPixmap, QPainter, QWheelEvent

from core.ui_components import FileListManagerWidget
from core.utils import BTN_GREEN
# === 引入分离的算法和后台处理 ===
from algorithms.cv_engine import (extract_skeleton, convert_contour_to_centerline,
                                  smart_simplify_path, orthogonalize_and_snap_path)
from .worker import DxfWorker


# ================= UI 辅助组件 =================
def create_slider_row(label_text, min_val, max_val, default_val, callback):
    """创建带有动态数值显示的滑块行"""
    layout = QHBoxLayout()
    lbl_title = QLabel(label_text)
    slider = QSlider(Qt.Horizontal)
    slider.setRange(min_val, max_val)
    slider.setValue(default_val)

    lbl_val = QLabel(str(default_val))
    lbl_val.setFixedWidth(30)
    lbl_val.setStyleSheet("color: #3498DB; font-weight: bold;")

    layout.addWidget(lbl_title)
    layout.addWidget(slider)
    layout.addWidget(lbl_val)

    def on_change(val):
        lbl_val.setText(str(val))
        callback()

    slider.valueChanged.connect(on_change)
    return layout, slider


# ================= 高级交互视图组件 =================
class InteractiveImageView(QGraphicsView):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.scene = QGraphicsScene(self)
        self.setScene(self.scene)
        self.setRenderHint(QPainter.Antialiasing)
        self.setRenderHint(QPainter.SmoothPixmapTransform)
        self.setDragMode(QGraphicsView.ScrollHandDrag)
        self.setStyleSheet("background-color: #E2E2E2; border: 1px dashed #A0A0A0; border-radius: 8px;")
        self.pixmap_item = None

    def set_image(self, q_img, reset_view=False):
        pixmap = QPixmap.fromImage(q_img)
        if self.pixmap_item is None:
            self.pixmap_item = self.scene.addPixmap(pixmap)
        else:
            self.pixmap_item.setPixmap(pixmap)
        self.scene.setSceneRect(self.pixmap_item.boundingRect())
        if reset_view:
            self.fit_to_window()

    def fit_to_window(self):
        if not self.pixmap_item: return
        rect = self.pixmap_item.boundingRect()
        view_rect = self.viewport().rect()
        if rect.width() == 0 or rect.height() == 0: return
        ratio = min(view_rect.width() / rect.width(), view_rect.height() / rect.height()) * 0.95
        self.resetTransform()
        self.scale(ratio, ratio)
        self.centerOn(self.pixmap_item)

    def wheelEvent(self, event: QWheelEvent):
        zoom_in_factor = 1.15
        if event.angleDelta().y() > 0:
            self.scale(zoom_in_factor, zoom_in_factor)
        else:
            self.scale(1 / zoom_in_factor, 1 / zoom_in_factor)


# ================= 主体界面组件 =================
class ImgToDxfWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.preview_cv_img = None
        self.is_loading_new_image = False
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(10, 10, 10, 10)
        splitter = QSplitter(Qt.Horizontal)

        left_widget = QWidget()
        l_layout = QVBoxLayout(left_widget)
        l_layout.setContentsMargins(0, 0, 0, 0)

        box_files = QGroupBox("1. 导入图片")
        lf_layout = QVBoxLayout()
        self.file_manager = FileListManagerWidget(accept_exts=['.jpg', '.png', '.jpeg', '.bmp'], title_desc="Images")
        self.file_manager.list_widget.itemClicked.connect(self.load_preview_image)
        lf_layout.addWidget(self.file_manager)
        box_files.setLayout(lf_layout)
        l_layout.addWidget(box_files, 2)

        box_params = QGroupBox("2. 矢量化引擎调优 (实时渲染)")
        lp_layout = QVBoxLayout()

        hz_mode = QHBoxLayout()
        hz_mode.addWidget(QLabel("⚙️ CAD 提取策略:"))
        self.cmb_mode = QComboBox()
        self.cmb_mode.addItems(["策略1：双线闭合轮廓提取 (适合照片、手写签名)",
                                "策略2：手绘级单线 (适合简单线条画)",
                                "策略3：保持正交 (适合建筑线稿图)"])
        self.cmb_mode.setCurrentIndex(0)
        self.cmb_mode.currentIndexChanged.connect(self._on_mode_changed)
        hz_mode.addWidget(self.cmb_mode)
        lp_layout.addLayout(hz_mode)

        self.chk_invert = QCheckBox("🔄 黑白反转 (必须保证要提取的线是白色的)")
        self.chk_invert.setChecked(True)
        self.chk_invert.stateChanged.connect(self.update_live_preview)
        lp_layout.addWidget(self.chk_invert)

        row1, self.slider_thresh = create_slider_row("🎚️ 二值化容差 (过滤背景):", 10, 245, 127, self.update_live_preview)
        row2, self.slider_curve = create_slider_row("📉 曲线保真度 (低=保留细节 高=圆滑):", 1, 30, 8, self.update_live_preview)
        row3, self.slider_straight = create_slider_row("📏 直线硬化容差 (将抖动变为直线):", 1, 20, 5, self.update_live_preview)

        lp_layout.addLayout(row1); lp_layout.addLayout(row2); lp_layout.addLayout(row3)

        self.group_ortho = QWidget()
        ortho_layout = QVBoxLayout(self.group_ortho)
        ortho_layout.setContentsMargins(0, 0, 0, 0)

        row4, self.slider_angle = create_slider_row("📐 正交捕捉角度 (±度数内强制拉直):", 1, 20, 5, self.update_live_preview)
        row5, self.slider_snap = create_slider_row("🧲 共线对齐阈值 (相近线条合并):", 0, 30, 8, self.update_live_preview)

        ortho_layout.addLayout(row4); ortho_layout.addLayout(row5)
        lp_layout.addWidget(self.group_ortho)
        lp_layout.addStretch()
        box_params.setLayout(lp_layout)
        l_layout.addWidget(box_params, 1)

        box_run = QGroupBox("3. 批量导出")
        lr_layout = QVBoxLayout()
        self.lbl_status = QLabel("就绪")
        self.progress = QProgressBar()
        btn_run = QPushButton("🚀 批量生成高级 CAD DXF 文件")
        btn_run.setStyleSheet(BTN_GREEN)
        btn_run.clicked.connect(self.run_conversion)
        lr_layout.addWidget(self.lbl_status); lr_layout.addWidget(self.progress); lr_layout.addWidget(btn_run)
        box_run.setLayout(lr_layout)
        l_layout.addWidget(box_run)

        splitter.addWidget(left_widget)

        right_widget = QWidget()
        r_layout = QVBoxLayout(right_widget)
        r_layout.setContentsMargins(0, 0, 0, 0)

        r_tools = QHBoxLayout()
        r_tools.addWidget(QLabel("👀 <b>实时视觉追踪预览</b> (滚轮缩放、右侧查看红线提取效果)"))
        r_tools.addStretch()
        btn_fit = QPushButton("🔲 适应窗口")
        btn_fit.clicked.connect(lambda: self.preview_view.fit_to_window() if hasattr(self, 'preview_view') else None)
        r_tools.addWidget(btn_fit)

        self.preview_view = InteractiveImageView()
        r_layout.addLayout(r_tools); r_layout.addWidget(self.preview_view)

        splitter.addWidget(right_widget); splitter.setSizes([400, 600])
        main_layout.addWidget(splitter)
        self._on_mode_changed()

    def _on_mode_changed(self):
        is_ortho = self.cmb_mode.currentIndex() == 2
        self.group_ortho.setVisible(is_ortho)
        self.update_live_preview()

    def load_preview_image(self, item):
        path = item.text()
        if not os.path.exists(path): return

        img_data = np.fromfile(path, np.uint8)
        img = cv2.imdecode(img_data, cv2.IMREAD_GRAYSCALE)
        if img is None: return

        max_dim = 1500
        h, w = img.shape
        if max(h, w) > max_dim:
            scale = max_dim / max(h, w)
            self.preview_cv_img = cv2.resize(img, (int(w * scale), int(h * scale)), interpolation=cv2.INTER_AREA)
        else:
            self.preview_cv_img = img

        self.is_loading_new_image = True
        self.update_live_preview()
        self.is_loading_new_image = False

    def update_live_preview(self):
        if self.preview_cv_img is None: return

        img = self.preview_cv_img.copy()
        if self.chk_invert.isChecked(): img = cv2.bitwise_not(img)
        _, thresh = cv2.threshold(img, self.slider_thresh.value(), 255, cv2.THRESH_BINARY)

        display_img = cv2.cvtColor(self.preview_cv_img, cv2.COLOR_GRAY2RGB)
        display_img = cv2.addWeighted(display_img, 0.4, np.zeros_like(display_img), 0.6, 0)

        mode_idx = self.cmb_mode.currentIndex()

        if mode_idx >= 1:
            thresh = extract_skeleton(thresh)

        contours, _ = cv2.findContours(thresh, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)

        curve_eps = self.slider_curve.value() / 10.0
        straight_tol = self.slider_straight.value()

        for cnt in contours:
            if mode_idx >= 1:
                approx = cv2.approxPolyDP(cnt, 1.0, True)
                center_path = convert_contour_to_centerline(approx)
                smart_path = smart_simplify_path(center_path, curve_eps, straight_tol)

                if mode_idx == 2:
                    smart_path = orthogonalize_and_snap_path(smart_path, self.slider_angle.value(),
                                                             self.slider_snap.value())

                for i in range(len(smart_path) - 1):
                    pt1 = (int(smart_path[i][0]), int(smart_path[i][1]))
                    pt2 = (int(smart_path[i + 1][0]), int(smart_path[i + 1][1]))
                    cv2.line(display_img, pt1, pt2, (255, 0, 0), 1, cv2.LINE_AA)
            else:
                epsilon = (self.slider_curve.value() / 10000.0) * cv2.arcLength(cnt, True)
                approx = cv2.approxPolyDP(cnt, epsilon, True)
                cv2.drawContours(display_img, [approx], -1, (255, 0, 0), 1, cv2.LINE_AA)

        h, w, ch = display_img.shape
        q_img = QImage(display_img.data, w, h, ch * w, QImage.Format_RGB888)
        self.preview_view.set_image(q_img, reset_view=self.is_loading_new_image)

    def run_conversion(self):
        paths = self.file_manager.get_all_filepaths()
        if not paths: return QMessageBox.warning(self, "提示", "请先添加图片文件。")
        out_dir = QFileDialog.getExistingDirectory(self, "选择 DXF 保存目录")
        if not out_dir: return

        thresh_val = self.slider_thresh.value()
        curve_eps = self.slider_curve.value() / 10.0 if self.cmb_mode.currentIndex() >= 1 else (self.slider_curve.value() / 10000.0)
        straight_tol = self.slider_straight.value()
        invert = self.chk_invert.isChecked()
        mode_idx = self.cmb_mode.currentIndex()
        ang_tol = self.slider_angle.value()
        snap_tol = self.slider_snap.value()

        self.worker = DxfWorker(paths, out_dir, thresh_val, curve_eps, invert, mode_idx, ang_tol, snap_tol, straight_tol)
        self.worker.progress.connect(lambda v, txt: (self.progress.setValue(v), self.lbl_status.setText(txt)))
        self.worker.finished.connect(lambda msg: QMessageBox.information(self, "完成", msg))
        self.worker.error.connect(lambda e: QMessageBox.critical(self, "错误", e))
        self.worker.start()