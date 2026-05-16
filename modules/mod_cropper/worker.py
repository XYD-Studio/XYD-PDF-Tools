# -*- coding: utf-8 -*-
import fitz
from PyQt5.QtCore import QThread, pyqtSignal
from core.pdf_engine import get_unique_filepath

class CropWorker(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

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

                base_name = self.page_to_filename.get(p_idx, f"Page_{p_idx + 1}")
                piece_counter = 1

                for r in range(len(h_lines) - 1):
                    for c in range(len(v_lines) - 1):
                        if f"{r},{c}" not in cfg['disabled']:
                            rect = fitz.Rect(v_lines[c] * w, h_lines[r] * h, v_lines[c + 1] * w, h_lines[r + 1] * h)
                            if rect.width < 5 or rect.height < 5: continue

                            pix = page.get_pixmap(matrix=fitz.Matrix(4.0, 4.0), clip=rect, alpha=True)
                            img_mode = "RGBA" if pix.alpha else "RGB"
                            img = Image.frombytes(img_mode, [pix.width, pix.height], pix.samples)

                            if self.mode == 'images':
                                save_name = f"{base_name}_切片_{piece_counter:02d}.png"
                                final_path = get_unique_filepath(self.output_path, save_name)
                                img.save(final_path, format="PNG")
                                piece_counter += 1
                            else:
                                results_images.append(img)

            if self.mode == 'pdf' and results_images:
                self.progress.emit(95, "正在打包生成全新 PDF...")
                rgb_images = []
                for img in results_images:
                    if img.mode == 'RGBA':
                        bg = Image.new('RGB', img.size, (255, 255, 255))
                        bg.paste(img, mask=img.split()[3])
                        rgb_images.append(bg)
                    else:
                        rgb_images.append(img.convert('RGB'))

                rgb_images[0].save(self.output_path, "PDF", resolution=150.0, save_all=True, append_images=rgb_images[1:])

            self.progress.emit(100, "裁剪完毕！")
            self.finished.emit("批量超级裁剪已完成！")
        except Exception as e:
            import traceback
            print(traceback.format_exc())
            self.error.emit(str(e))