# -*- coding: utf-8 -*-
"""
algorithms/ocr_engine.py
纯粹的 OCR 识别引擎封装，完全脱离 PyQt5/UI 库。
采用懒加载模式，只有在实际调用时才去加载沉重的 PaddleOCR 模型。
"""
import logging
import numpy as np


class LocalOCREngine:
    def __init__(self, use_gpu=False):
        # 屏蔽 PaddleOCR 内部冗长的控制台 debug 日志
        logging.getLogger("ppocr").setLevel(logging.WARNING)

        # 懒加载导入
        from paddleocr import PaddleOCR

        # 初始化模型 (单例)
        self.ocr = PaddleOCR(use_angle_cls=True, lang="ch", use_gpu=use_gpu, show_log=False, enable_mkldnn=False)

    def recognize_image(self, pil_img):
        """
        传入 PIL Image，返回识别拼接后的文本字符串
        """
        img_array = np.array(pil_img)
        res = self.ocr.ocr(img_array, cls=True)

        if not res or res[0] is None:
            return ""

        # 拼接识别出来的所有文本行
        return " ".join([line[1][0] for line in res[0]]).strip()