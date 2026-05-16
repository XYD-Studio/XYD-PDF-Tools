# -*- coding: utf-8 -*-
from PyQt5.QtCore import QThread, pyqtSignal
import fitz
from PIL import Image

# === 引入我们刚才剥离出来的纯算法引擎 ===
from algorithms.ocr_engine import LocalOCREngine


class OCRWorker(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(list)

    def __init__(self, pdf_doc, page_configs, fields_config):
        super().__init__()
        self.pdf_doc = pdf_doc
        self.page_configs = page_configs
        self.fields_config = fields_config

    def run(self):
        try:
            self.progress.emit(0, "正在加载 OCR 引擎核心库 (首次加载可能需要几秒)...")

            # 实例化纯算法引擎
            ocr_engine = LocalOCREngine(use_gpu=False)

            self.progress.emit(5, "OCR 引擎加载成功，正在初始化识别模型...")

            results = []
            total = len(self.pdf_doc)

            for i in range(total):
                self.progress.emit(int((i / total) * 100), f"正在识别第 {i + 1}/{total} 页...")
                page = self.pdf_doc[i]
                page_boxes = self.page_configs.get(i, {})

                row_data = {}
                for field in self.fields_config:
                    fname = field['name']
                    if not field['is_ocr']:
                        row_data[fname] = field['static_val']
                        continue

                    pdf_rect_data = page_boxes.get(fname, (100, 100, 150, 40))
                    pdf_x, pdf_y, pdf_w, pdf_h = pdf_rect_data
                    clip_rect = fitz.Rect(pdf_x, pdf_y, pdf_x + pdf_w, pdf_y + pdf_h)

                    # 截取框选区域的高清图
                    pix = page.get_pixmap(matrix=fitz.Matrix(3.0, 3.0), clip=clip_rect, alpha=False)
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                    # 丢给算法引擎识别
                    row_data[fname] = ocr_engine.recognize_image(img)

                results.append(row_data)

            self.progress.emit(100, "识别完成！")
            self.finished.emit(results)

        except Exception as e:
            import traceback
            print(traceback.format_exc())
            self.progress.emit(0, f"OCR 提取异常: 请检查环境或依赖。")