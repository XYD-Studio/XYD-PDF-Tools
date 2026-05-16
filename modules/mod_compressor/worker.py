# -*- coding: utf-8 -*-
import os
import shutil
import tempfile
from PIL import Image
from PyQt5.QtCore import QThread, pyqtSignal
import fitz

from core.pdf_engine import run_ghostscript


class PDFCompressWorker(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(str)

    def __init__(self, paths, export_config, gs_config, output_path, gs_path, gs_lib_path):
        super().__init__()
        self.paths = paths
        self.mode = export_config['mode']
        self.prefix = export_config['prefix']
        self.suffix = export_config['suffix']
        self.quality = gs_config['quality']
        self.output_path = output_path
        self.gs_path = gs_path
        self.gs_lib_path = gs_lib_path

    def run(self):
        try:
            if self.mode == 'merged':
                self.progress.emit(10, "正在进行前置合并处理...")
                temp_dir = tempfile.mkdtemp(prefix="pdf_comp_")
                try:
                    temp_merged = os.path.join(temp_dir, "merged_temp.pdf")
                    doc = fitz.Document()
                    for p in self.paths:
                        src = fitz.open(p)
                        doc.insert_pdf(src)
                        src.close()
                    doc.save(temp_merged)
                    doc.close()

                    self.progress.emit(50, "正在启动引擎进行深度强力压缩...")
                    run_ghostscript(self.gs_path, self.gs_lib_path, temp_merged, self.output_path, quality=self.quality)
                finally:
                    shutil.rmtree(temp_dir, ignore_errors=True)
                self.progress.emit(100, "合并压缩完成！")

            elif self.mode == 'batch':
                total = len(self.paths)
                for i, input_path in enumerate(self.paths):
                    self.progress.emit(int(i / total * 100), f"正在压缩: {os.path.basename(input_path)}...")
                    base_name = os.path.splitext(os.path.basename(input_path))[0]
                    final_name = f"{self.prefix}{base_name}{self.suffix}.pdf"
                    out_path = os.path.join(self.output_path, final_name)

                    run_ghostscript(self.gs_path, self.gs_lib_path, input_path, out_path, quality=self.quality)
                self.progress.emit(100, "批量压缩完成！")

            elif self.mode == 'split':
                total = len(self.paths)
                temp_dir = tempfile.mkdtemp(prefix="pdf_comp_split_")
                try:
                    for i, input_path in enumerate(self.paths):
                        self.progress.emit(int(i / total * 90), f"正在压缩并拆分: {os.path.basename(input_path)}...")
                        base_name = os.path.splitext(os.path.basename(input_path))[0]

                        temp_pdf = os.path.join(temp_dir, f"temp_{i}.pdf")
                        run_ghostscript(self.gs_path, self.gs_lib_path, input_path, temp_pdf, quality=self.quality)

                        doc = fitz.open(temp_pdf)
                        for page_num in range(len(doc)):
                            new_doc = fitz.Document()
                            new_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
                            final_name = f"{self.prefix}{base_name}_第{page_num + 1}页{self.suffix}.pdf"
                            new_doc.save(os.path.join(self.output_path, final_name))
                            new_doc.close()
                        doc.close()
                finally:
                    shutil.rmtree(temp_dir, ignore_errors=True)
                self.progress.emit(100, "拆分压缩完成！")

            self.finished.emit("处理完毕，请查看输出目录！")
        except Exception as e:
            import traceback
            print(traceback.format_exc())
            self.finished.emit(f"发生异常: {str(e)}")


class ImgCompressWorker(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(str)

    def __init__(self, paths, save_dir, target_val, fmt_sel, is_dpi, prefix, suffix):
        super().__init__()
        self.paths = paths
        self.save_dir = save_dir
        self.target_val = target_val
        self.fmt_sel = fmt_sel
        self.is_dpi = is_dpi
        self.prefix = prefix
        self.suffix = suffix

    def run(self):
        total = len(self.paths)
        for i, input_path in enumerate(self.paths):
            self.progress.emit(int(i / total * 100), f"正在转换: {os.path.basename(input_path)}...")
            name, ext = os.path.splitext(os.path.basename(input_path))

            target_ext = ext
            if "JPG" in self.fmt_sel:
                target_ext = ".jpg"
            elif "PNG" in self.fmt_sel:
                target_ext = ".png"
            elif "WEBP" in self.fmt_sel:
                target_ext = ".webp"
            elif "PDF" in self.fmt_sel:
                target_ext = ".pdf"

            final_name = f"{self.prefix}{name}{self.suffix}{target_ext}"
            out_path = os.path.join(self.save_dir, final_name)

            try:
                with Image.open(input_path) as img:
                    orig_dpi = img.info.get('dpi', (96, 96))
                    save_dpi = orig_dpi
                    if self.is_dpi:
                        if self.target_val < orig_dpi[0]:
                            scale = self.target_val / orig_dpi[0]
                            img = img.resize((int(img.width * scale), int(img.height * scale)),
                                             Image.Resampling.LANCZOS)
                        save_dpi = (self.target_val, self.target_val)
                    else:
                        w, h = img.size
                        if max(w, h) > self.target_val:
                            ratio = self.target_val / max(w, h)
                            img = img.resize((int(w * ratio), int(h * ratio)), Image.Resampling.LANCZOS)

                    if target_ext in ['.jpg', '.pdf'] and img.mode in ('RGBA', 'P'):
                        bg = Image.new('RGB', img.size, (255, 255, 255))
                        if img.mode == 'RGBA':
                            bg.paste(img, mask=img.split()[3])
                        else:
                            bg.paste(img.convert('RGB'))
                        img = bg

                    if target_ext == '.jpg':
                        img.save(out_path, quality=80, dpi=save_dpi, optimize=True)
                    elif target_ext == '.pdf':
                        img.save(out_path, "PDF", resolution=save_dpi[0])
                    else:
                        img.save(out_path, dpi=save_dpi)
            except Exception as e:
                print(f"处理图片异常: {e}")

        self.progress.emit(100, "图片批量压缩完成！")
        self.finished.emit("压缩转换成功，请查看输出目录！")