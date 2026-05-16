# -*- coding: utf-8 -*-
"""
algorithms/cv_engine.py
纯粹的计算机视觉与几何处理引擎。
【注意】此文件绝对不可导入任何 PyQt5/UI 相关的库，以保证算法能在服务器端独立运行。
"""
import cv2
import math
import numpy as np

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