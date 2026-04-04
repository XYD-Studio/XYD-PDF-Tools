import os
import cv2
import math
import ezdxf
import numpy as np
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                             QGroupBox, QSlider, QCheckBox, QFileDialog, QMessageBox, QProgressBar, QSplitter,
                             QComboBox, QGraphicsView, QGraphicsScene)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QImage, QPixmap, QPainter, QWheelEvent
from core.ui_components import FileListManagerWidget
from core.utils import get_unique_filepath, BTN_BLUE, BTN_GREEN


# ================= 核心几何引擎算法 =================

def extract_skeleton(img):
    """提取1像素中心骨架线"""
    # 闭运算填补断层
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
    closed = cv2.morphologyEx(img, cv2.MORPH_CLOSE, kernel, iterations=1)

    skel = np.zeros(closed.shape, np.uint8)
    element = cv2.getStructuringElement(cv2.MORPH_CROSS, (3, 3))
    while True:
        eroded = cv2.erode(closed, element)
        temp = cv2.dilate(eroded, element)
        temp = cv2.subtract(closed, temp)
        skel = cv2.bitwise_or(skel, temp)
        closed = eroded.copy()
        if cv2.countNonZero(closed) == 0:
            break
    return skel


def convert_contour_to_centerline(approx_contour):
    """回路截断：将骨架的双线闭环截断为单条长线"""
    pts = [p[0] for p in approx_contour]
    if len(pts) <= 2: return pts

    max_dist = 0
    end1_idx, end2_idx = 0, 0
    n = len(pts)
    for i in range(n):
        for j in range(i + 1, n):
            dist = (pts[i][0] - pts[j][0]) ** 2 + (pts[i][1] - pts[j][1]) ** 2
            if dist > max_dist:
                max_dist = dist
                end1_idx = i
                end2_idx = j

    if end1_idx > end2_idx: end1_idx, end2_idx = end2_idx, end1_idx
    path1 = pts[end1_idx:end2_idx + 1]
    path2 = pts[end2_idx:] + pts[:end1_idx + 1]
    return path1 if len(path1) > len(path2) else path2


def smart_simplify_path(points, curve_eps, straight_angle_tol):
    """
    智能平滑算法：
    区分直线和曲线。曲线保留密集节点，直线强行剔除中间抖动节点。
    """
    if len(points) <= 2: return points

    # 1. 轻度基础平滑，滤除像素级毛刺
    pts = np.array(points).reshape((-1, 1, 2)).astype(np.float32)
    approx = cv2.approxPolyDP(pts, curve_eps, False)
    base_pts = [p[0] for p in approx]

    if len(base_pts) <= 2: return base_pts

    # 2. 角度智能剔除：如果三个点接近一条直线（夹角偏差极小），删掉中间的控制点
    final_pts = [base_pts[0]]
    for i in range(1, len(base_pts) - 1):
        p_prev = final_pts[-1]
        p_curr = base_pts[i]
        p_next = base_pts[i + 1]

        v1 = (p_curr[0] - p_prev[0], p_curr[1] - p_prev[1])
        v2 = (p_next[0] - p_curr[0], p_next[1] - p_curr[1])

        mag1 = math.hypot(*v1)
        mag2 = math.hypot(*v2)

        if mag1 == 0 or mag2 == 0: continue

        # 计算向量夹角偏差 (180度直线对应偏差为0)
        dot = v1[0] * v2[0] + v1[1] * v2[1]
        cos_theta = max(-1.0, min(1.0, dot / (mag1 * mag2)))
        angle_diff = math.degrees(math.acos(cos_theta))

        # 如果角度偏差大于阈值，说明是在拐弯或者是曲线，保留节点！
        if angle_diff > straight_angle_tol:
            final_pts.append(p_curr)

    final_pts.append(base_pts[-1])
    return final_pts


def orthogonalize_and_snap_path(path, angle_tol_deg, snap_dist):
    """
    建筑级正交捕捉：只拉直接近水平/垂直的线段，绝不破坏斜线和曲线！
    """
    if len(path) <= 1: return path

    # 转换为列表以方便修改坐标
    pts = [list(p) for p in path]
    angle_rad = math.radians(angle_tol_deg)
    tan_tol = math.tan(angle_rad) if angle_rad < math.pi / 2 else 999

    # 阶段1：局部拉直
    for i in range(len(pts) - 1):
        dx = pts[i + 1][0] - pts[i][0]
        dy = pts[i + 1][1] - pts[i][1]

        if dx != 0 and abs(dy) / abs(dx) <= tan_tol:
            # 近似水平，拉平Y轴
            avg_y = (pts[i][1] + pts[i + 1][1]) / 2.0
            pts[i][1] = avg_y
            pts[i + 1][1] = avg_y
        elif dy != 0 and abs(dx) / abs(dy) <= tan_tol:
            # 近似垂直，拉平X轴
            avg_x = (pts[i][0] + pts[i + 1][0]) / 2.0
            pts[i][0] = avg_x
            pts[i + 1][0] = avg_x

    # 阶段2：全局共线合并 (将相近厚度的墙体双线合并，或对齐齐平的窗框)
    # 此处为简化处理，针对同一条多段线内部的临近水平/垂直线进行对齐
    for i in range(len(pts) - 1):
        for j in range(i + 2, len(pts) - 1):
            # 对齐水平段
            if pts[i][1] == pts[i + 1][1] and pts[j][1] == pts[j + 1][1]:
                if abs(pts[i][1] - pts[j][1]) <= snap_dist:
                    avg_y = (pts[i][1] + pts[j][1]) / 2.0
                    pts[i][1] = pts[i + 1][1] = pts[j][1] = pts[j + 1][1] = avg_y
            # 对齐垂直段
            if pts[i][0] == pts[i + 1][0] and pts[j][0] == pts[j + 1][0]:
                if abs(pts[i][0] - pts[j][0]) <= snap_dist:
                    avg_x = (pts[i][0] + pts[j][0]) / 2.0
                    pts[i][0] = pts[i + 1][0] = pts[j][0] = pts[j + 1][0] = avg_x

    return [tuple(p) for p in pts]


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


# ================= 后台处理线程 =================
class DxfWorker(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, file_paths, out_dir, thresh, eps, invert, mode_idx, ang_tol, snap_tol, str_tol):
        super().__init__()
        self.file_paths = file_paths
        self.output_dir = out_dir
        self.threshold = thresh
        self.curve_eps = eps
        self.invert = invert
        self.mode_idx = mode_idx
        self.angle_tol = ang_tol
        self.snap_tol = snap_tol
        self.straight_tol = str_tol

    def run(self):
        try:
            total = len(self.file_paths)
            for i, path in enumerate(self.file_paths):
                filename = os.path.basename(path)
                self.progress.emit(int(i / total * 100), f"正在矢量化提取: {filename}...")

                img_data = np.fromfile(path, np.uint8)
                img = cv2.imdecode(img_data, cv2.IMREAD_GRAYSCALE)
                if img is None: continue

                if self.invert: img = cv2.bitwise_not(img)
                _, thresh = cv2.threshold(img, self.threshold, 255, cv2.THRESH_BINARY)

                if self.mode_idx >= 1:
                    thresh = extract_skeleton(thresh)

                contours, _ = cv2.findContours(thresh, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)

                doc = ezdxf.new('R2010')
                msp = doc.modelspace()
                height = img.shape[0]

                for cnt in contours:
                    if self.mode_idx >= 1:
                        # 策略2与3：单线骨架智能提取
                        approx = cv2.approxPolyDP(cnt, 1.0, True)  # 基础微降噪
                        center_path = convert_contour_to_centerline(approx)

                        # 应用智能平滑 (曲线保留，直线硬化)
                        smart_path = smart_simplify_path(center_path, self.curve_eps, self.straight_tol)

                        # 策略3应用正交拉直和对齐，保留斜线与曲线
                        if self.mode_idx == 2:
                            smart_path = orthogonalize_and_snap_path(smart_path, self.angle_tol, self.snap_tol)

                        points = [(p[0], height - p[1]) for p in smart_path]
                        if len(points) > 1: msp.add_lwpolyline(points, close=False)
                    else:
                        # 策略1：闭合双线面状轮廓
                        epsilon = self.curve_eps * cv2.arcLength(cnt, True)
                        approx = cv2.approxPolyDP(cnt, epsilon, True)
                        points = [(p[0][0], height - p[0][1]) for p in approx]
                        if len(points) > 2:
                            msp.add_lwpolyline(points, close=True)
                        elif len(points) == 2:
                            msp.add_line(points[0], points[1])

                mode_dict = {0: "实心轮廓", 1: "纯净单线", 2: "正交混排线"}
                out_path = get_unique_filepath(self.output_dir,
                                               f"{os.path.splitext(filename)[0]}_{mode_dict[self.mode_idx]}.dxf")
                doc.saveas(out_path)

            self.progress.emit(100, "全部处理完成！")
            self.finished.emit(f"成功将 {total} 个文件转换为 DXF！\n请前往输出目录查看。")
        except Exception as e:
            self.error.emit(str(e))


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


# ================= 商业级 UI 主件 =================
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

        # 基础参数 (带数值显示)
        row1, self.slider_thresh = create_slider_row("🎚️ 二值化容差 (过滤背景):", 10, 245, 127,
                                                     self.update_live_preview)
        row2, self.slider_curve = create_slider_row("📉 曲线保真度 (低=保留细节 高=圆滑):", 1, 30, 8,
                                                    self.update_live_preview)
        row3, self.slider_straight = create_slider_row("📏 直线硬化容差 (将微弱抖动变为绝对直线):", 1, 20, 5,
                                                       self.update_live_preview)

        lp_layout.addLayout(row1)
        lp_layout.addLayout(row2)
        lp_layout.addLayout(row3)

        # 策略3 专属参数区
        self.group_ortho = QWidget()
        ortho_layout = QVBoxLayout(self.group_ortho)
        ortho_layout.setContentsMargins(0, 0, 0, 0)

        row4, self.slider_angle = create_slider_row("📐 正交捕捉角度 (±度数内强制拉直):", 1, 20, 5,
                                                    self.update_live_preview)
        row5, self.slider_snap = create_slider_row("🧲 共线对齐阈值 (相近线条合并):", 0, 30, 8, self.update_live_preview)

        ortho_layout.addLayout(row4)
        ortho_layout.addLayout(row5)
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
        lr_layout.addWidget(self.lbl_status)
        lr_layout.addWidget(self.progress)
        lr_layout.addWidget(btn_run)
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
        r_layout.addLayout(r_tools)
        r_layout.addWidget(self.preview_view)

        splitter.addWidget(right_widget)
        splitter.setSizes([400, 600])
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

        # 预览分辨率提高，以看清曲线平滑效果
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

        # 参数映射
        curve_eps = self.slider_curve.value() / 10.0
        straight_tol = self.slider_straight.value()

        for cnt in contours:
            if mode_idx >= 1:
                # 单线提取
                approx = cv2.approxPolyDP(cnt, 1.0, True)
                center_path = convert_contour_to_centerline(approx)
                smart_path = smart_simplify_path(center_path, curve_eps, straight_tol)

                if mode_idx == 2:
                    smart_path = orthogonalize_and_snap_path(smart_path, self.slider_angle.value(),
                                                             self.slider_snap.value())

                # 绘制单线
                for i in range(len(smart_path) - 1):
                    pt1 = (int(smart_path[i][0]), int(smart_path[i][1]))
                    pt2 = (int(smart_path[i + 1][0]), int(smart_path[i + 1][1]))
                    cv2.line(display_img, pt1, pt2, (255, 0, 0), 1, cv2.LINE_AA)
            else:
                # 策略1：面状轮廓
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
        curve_eps = self.slider_curve.value() / 10.0 if self.cmb_mode.currentIndex() >= 1 else (
                    self.slider_curve.value() / 10000.0)
        straight_tol = self.slider_straight.value()
        invert = self.chk_invert.isChecked()
        mode_idx = self.cmb_mode.currentIndex()
        ang_tol = self.slider_angle.value()
        snap_tol = self.slider_snap.value()

        self.worker = DxfWorker(paths, out_dir, thresh_val, curve_eps, invert, mode_idx, ang_tol, snap_tol,
                                straight_tol)
        self.worker.progress.connect(lambda v, txt: (self.progress.setValue(v), self.lbl_status.setText(txt)))
        self.worker.finished.connect(lambda msg: QMessageBox.information(self, "完成", msg))
        self.worker.error.connect(lambda e: QMessageBox.critical(self, "错误", e))
        self.worker.start()