# -*- coding: utf-8 -*-
import fitz
import io
from PyQt5.QtCore import QThread, pyqtSignal
from core.pdf_engine import get_unique_filepath

class CropWorker(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, pdf_doc, page_configs, mode, output_path, page_to_filename, page_render_dpi=None):
        super().__init__()
        self.pdf_doc = pdf_doc
        self.page_configs = page_configs
        self.mode = mode
        self.output_path = output_path
        self.page_to_filename = page_to_filename
        self.page_render_dpi = page_render_dpi or {}

    @staticmethod
    def _normalise_lines(values):
        clean = {round(max(0.0, min(1.0, float(value))), 6) for value in values}
        return sorted(value for value in clean if 0.0001 < value < 0.9999)

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

                v_lines = [0.0] + self._normalise_lines(cfg.get('v_lines', [])) + [1.0]
                h_lines = [0.0] + self._normalise_lines(cfg.get('h_lines', [])) + [1.0]
                disabled = set(cfg.get('disabled', []))
                render_dpi = max(72.0, min(300.0, float(self.page_render_dpi.get(p_idx, 300.0))))
                render_scale = render_dpi / 72.0

                base_name = self.page_to_filename.get(p_idx, f"Page_{p_idx + 1}")
                piece_counter = 1

                for r in range(len(h_lines) - 1):
                    for c in range(len(v_lines) - 1):
                        if f"{r},{c}" not in disabled:
                            rect = fitz.Rect(v_lines[c] * w, h_lines[r] * h, v_lines[c + 1] * w, h_lines[r + 1] * h)
                            if rect.width < 5 or rect.height < 5: continue

                            pix = page.get_pixmap(matrix=fitz.Matrix(render_scale, render_scale), clip=rect, alpha=True)
                            img_mode = "RGBA" if pix.alpha else "RGB"
                            img = Image.frombytes(img_mode, [pix.width, pix.height], pix.samples)

                            if self.mode == 'images':
                                save_name = f"{base_name}_切片_{piece_counter:02d}.png"
                                final_path = get_unique_filepath(self.output_path, save_name)
                                img.save(final_path, format="PNG", dpi=(render_dpi, render_dpi), optimize=True)
                                piece_counter += 1
                            else:
                                results_images.append((img, rect.width, rect.height))

            if self.mode == 'pdf' and results_images:
                self.progress.emit(95, "正在打包生成全新 PDF...")
                output_doc = fitz.Document()
                try:
                    for img, width, height in results_images:
                        stream = io.BytesIO()
                        img.save(stream, format='PNG', optimize=True)
                        output_page = output_doc.new_page(width=width, height=height)
                        output_page.insert_image(output_page.rect, stream=stream.getvalue(), keep_proportion=False)
                    output_doc.save(self.output_path, garbage=4, deflate=True)
                finally:
                    output_doc.close()

            self.progress.emit(100, "裁剪完毕！")
            self.finished.emit("批量超级裁剪已完成！")
        except Exception as e:
            import traceback
            print(traceback.format_exc())
            self.error.emit(str(e))
