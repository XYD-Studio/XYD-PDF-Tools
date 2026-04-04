import os
import cv2
import ezdxf
import numpy as np
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                             QGroupBox, QSlider, QCheckBox, QFileDialog, QMessageBox, QProgressBar, QSplitter)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QImage, QPixmap
from core.ui_components import FileListManagerWidget
from core.utils import get_unique_filepath, BTN_BLUE, BTN_GREEN


# ================= 后台批量处理线程 =================
class DxfWorker(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, file_paths, output_dir, threshold, epsilon_factor, invert):
        super().__init__()
        self.file_paths = file_paths
        self.output_dir = output_dir
        self.threshold = threshold
        self.epsilon_factor = epsilon_factor
        self.invert = invert

    def run(self):
        try:
            total = len(self.file_paths)
            for i, path in enumerate(self.file_paths):
                filename = os.path.basename(path)
                self.progress.emit(int(i / total * 100), f"正在矢量化提取: {filename}...")

                # 1. OpenCV 图像处理
                # 使用 imdecode 支持中文路径
                img_data = np.fromfile(path, np.uint8)
                img = cv2.imdecode(img_data, cv2.IMREAD_GRAYSCALE)

                if img is None:
                    continue

                if self.invert:
                    img = cv2.bitwise_not(img)

                _, thresh = cv2.threshold(img, self.threshold, 255, cv2.THRESH_BINARY)

                # 寻找轮廓
                contours, _ = cv2.findContours(thresh, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)

                # 2. 写入 DXF
                doc = ezdxf.new('R2010')
                msp = doc.modelspace()

                height = img.shape[0]

                for cnt in contours:
                    # 多边形拟合平滑 (epsilon 越大约平滑，节点越少)
                    epsilon = self.epsilon_factor * cv2.arcLength(cnt, True)
                    approx = cv2.approxPolyDP(cnt, epsilon, True)

                    # 坐标系转换 (图像左上角原点 -> CAD左下角原点)
                    points = [(p[0][0], height - p[0][1]) for p in approx]

                    if len(points) > 1:
                        # 写入 CAD 多段线
                        msp.add_lwpolyline(points, close=True)

                # 3. 保存文件
                base_name = os.path.splitext(filename)[0]
                out_path = get_unique_filepath(self.output_dir, f"{base_name}_矢量提取.dxf")
                doc.saveas(out_path)

            self.progress.emit(100, "全部处理完成！")
            self.finished.emit(f"成功将 {total} 个文件转换为 DXF！\n请前往输出目录查看。")
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.error.emit(str(e))


# ================= 商业级 UI 组件 =================
class ImgToDxfWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.preview_cv_img = None  # 用于缓存用于预览的低分辨率 OpenCV 图像
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(10, 10, 10, 10)
        splitter = QSplitter(Qt.Horizontal)

        # ====== 左侧：文件管理与参数调优 ======
        left_widget = QWidget()
        l_layout = QVBoxLayout(left_widget)
        l_layout.setContentsMargins(0, 0, 0, 0)

        box_files = QGroupBox("1. 导入线稿图片")
        lf_layout = QVBoxLayout()
        self.file_manager = FileListManagerWidget(accept_exts=['.jpg', '.png', '.jpeg', '.bmp'], title_desc="Images")
        # 绑定列表点击事件，用于触发实时预览
        self.file_manager.list_widget.itemClicked.connect(self.load_preview_image)
        lf_layout.addWidget(self.file_manager)
        box_files.setLayout(lf_layout)
        l_layout.addWidget(box_files, 2)

        box_params = QGroupBox("2. 矢量化核心引擎参数 (实时预览)")
        lp_layout = QVBoxLayout()

        # 反转选项 (白底黑线必须反转，黑底白线不反转)
        self.chk_invert = QCheckBox("🔄 黑白反转 (通常白底黑线的线稿必须勾选)")
        self.chk_invert.setChecked(True)
        self.chk_invert.stateChanged.connect(self.update_live_preview)
        lp_layout.addWidget(self.chk_invert)

        # 阈值滑块
        lp_layout.addWidget(QLabel("🎚️ 二值化阈值 (剔除浅色阴影):"))
        self.slider_thresh = QSlider(Qt.Horizontal)
        self.slider_thresh.setRange(10, 245)
        self.slider_thresh.setValue(127)
        self.slider_thresh.valueChanged.connect(self.update_live_preview)
        lp_layout.addWidget(self.slider_thresh)

        # 平滑度滑块
        lp_layout.addWidget(QLabel("📉 节点降噪平滑度 (越大越平滑，DXF文件越小):"))
        self.slider_smooth = QSlider(Qt.Horizontal)
        self.slider_smooth.setRange(1, 50)  # 映射为 0.0001 到 0.0050
        self.slider_smooth.setValue(10)
        self.slider_smooth.valueChanged.connect(self.update_live_preview)
        lp_layout.addWidget(self.slider_smooth)

        lp_layout.addStretch()
        box_params.setLayout(lp_layout)
        l_layout.addWidget(box_params, 1)

        # 执行面板
        box_run = QGroupBox("3. 批量执行")
        lr_layout = QVBoxLayout()
        self.lbl_status = QLabel("就绪")
        self.progress = QProgressBar()
        btn_run = QPushButton("🚀 批量生成高精度 DXF 文件")
        btn_run.setStyleSheet(BTN_GREEN)
        btn_run.clicked.connect(self.run_conversion)

        lr_layout.addWidget(self.lbl_status)
        lr_layout.addWidget(self.progress)
        lr_layout.addWidget(btn_run)
        box_run.setLayout(lr_layout)
        l_layout.addWidget(box_run)

        splitter.addWidget(left_widget)

        # ====== 右侧：实时视觉预览区 ======
        right_widget = QGroupBox("👀 提取轮廓实时预览 (红线为提取结果)")
        r_layout = QVBoxLayout(right_widget)
        self.lbl_preview = QLabel("👈 请在左侧列表中点击任意一张图片以加载预览")
        self.lbl_preview.setAlignment(Qt.AlignCenter)
        self.lbl_preview.setStyleSheet("background-color: #E2E2E2; border: 1px dashed #A0A0A0; border-radius: 8px;")
        r_layout.addWidget(self.lbl_preview)
        splitter.addWidget(right_widget)

        splitter.setSizes([350, 650])
        main_layout.addWidget(splitter)

    def load_preview_image(self, item):
        """加载图片并降低分辨率存入缓存，供滑动滑块时毫秒级处理"""
        path = item.text()
        if not os.path.exists(path): return

        # 使用 OpenCV 读取，防止中文路径报错
        img_data = np.fromfile(path, np.uint8)
        img = cv2.imdecode(img_data, cv2.IMREAD_GRAYSCALE)
        if img is None: return

        # 核心商业级优化：预览图强制降采样到最大 800 像素，保证极速预览不卡主线程
        max_dim = 800
        h, w = img.shape
        if max(h, w) > max_dim:
            scale = max_dim / max(h, w)
            self.preview_cv_img = cv2.resize(img, (int(w * scale), int(h * scale)), interpolation=cv2.INTER_AREA)
        else:
            self.preview_cv_img = img

        self.update_live_preview()

    def update_live_preview(self):
        """核心视觉反馈：基于缓存的低分辨率图实时计算轮廓并渲染到 QLabel 上"""
        if self.preview_cv_img is None: return

        img = self.preview_cv_img.copy()

        # 1. 算法处理
        if self.chk_invert.isChecked():
            img = cv2.bitwise_not(img)

        thresh_val = self.slider_thresh.value()
        _, thresh = cv2.threshold(img, thresh_val, 255, cv2.THRESH_BINARY)

        contours, _ = cv2.findContours(thresh, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)

        # 2. 渲染结果展示图 (转为彩色以绘制红线)
        # 用暗化处理的原图作为背景，让红线更明显
        display_img = cv2.cvtColor(self.preview_cv_img, cv2.COLOR_GRAY2BGR)
        display_img = cv2.addWeighted(display_img, 0.4, np.zeros_like(display_img), 0.6, 0)  # 调暗背景

        eps_factor = self.slider_smooth.value() / 10000.0

        drawn_contours = []
        for cnt in contours:
            epsilon = eps_factor * cv2.arcLength(cnt, True)
            approx = cv2.approxPolyDP(cnt, epsilon, True)
            drawn_contours.append(approx)

        # 在暗色背景上画鲜艳的红色线条，模拟 CAD 感觉
        cv2.drawContours(display_img, drawn_contours, -1, (0, 0, 255), 1, cv2.LINE_AA)

        # 3. 转为 Qt 图像并显示适应窗口
        h, w, ch = display_img.shape
        bytes_per_line = ch * w
        q_img = QImage(display_img.data, w, h, bytes_per_line, QImage.Format_RGB888)

        # 保证图片在 QLabel 中保持比例自适应
        pixmap = QPixmap.fromImage(q_img)
        scaled_pixmap = pixmap.scaled(self.lbl_preview.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self.lbl_preview.setPixmap(scaled_pixmap)

    # 监听控件大小变化，实时重绘保持自适应
    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.preview_cv_img is not None:
            self.update_live_preview()

    def run_conversion(self):
        paths = self.file_manager.get_all_filepaths()
        if not paths:
            return QMessageBox.warning(self, "提示", "请先添加要转换的图片文件。")

        out_dir = QFileDialog.getExistingDirectory(self, "选择 DXF 保存目录")
        if not out_dir: return

        thresh_val = self.slider_thresh.value()
        eps_factor = self.slider_smooth.value() / 10000.0
        invert = self.chk_invert.isChecked()

        self.worker = DxfWorker(paths, out_dir, thresh_val, eps_factor, invert)
        self.worker.progress.connect(lambda v, txt: (self.progress.setValue(v), self.lbl_status.setText(txt)))
        self.worker.finished.connect(lambda msg: QMessageBox.information(self, "完成", msg))
        self.worker.error.connect(lambda e: QMessageBox.critical(self, "错误", e))
        self.worker.start()