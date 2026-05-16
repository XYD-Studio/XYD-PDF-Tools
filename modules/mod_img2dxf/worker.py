# -*- coding: utf-8 -*-
import os
import cv2
import ezdxf
import numpy as np
from PyQt5.QtCore import QThread, pyqtSignal

from core.utils import get_unique_filepath
# === 从刚刚抽离的独立算法层导入纯数学逻辑 ===
from algorithms.cv_engine import (extract_skeleton, convert_contour_to_centerline,
                                  smart_simplify_path, orthogonalize_and_snap_path)

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
                        approx = cv2.approxPolyDP(cnt, 1.0, True)
                        center_path = convert_contour_to_centerline(approx)

                        # 应用智能平滑
                        smart_path = smart_simplify_path(center_path, self.curve_eps, self.straight_tol)

                        # 策略3应用正交拉直和对齐
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