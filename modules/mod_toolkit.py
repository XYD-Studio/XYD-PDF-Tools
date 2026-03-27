import os
import shutil
import tempfile
import concurrent.futures
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                             QGroupBox, QComboBox, QFileDialog, QMessageBox, QProgressBar, QCheckBox, QLineEdit,
                             QDialog, QRadioButton, QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView,
                             QDialogButtonBox)
from PyQt5.QtCore import QThread, pyqtSignal

# 【核心更新】：引入高级文件管理器
from core.ui_components import FileListManagerWidget
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
import fitz


# ================= 书签拆分高级配置弹窗 =================
class BookmarkSplitConfigDialog(QDialog):
    def __init__(self, current_config, parent=None):
        super().__init__(parent)
        self.setWindowTitle("书签拆分命名规则配置")
        self.resize(550, 400)
        layout = QVBoxLayout(self)

        box_base = QGroupBox("1. 基础命名规则")
        l_base = QVBoxLayout()
        self.rb_name = QRadioButton("按【书签名称】命名")
        self.rb_seq_name = QRadioButton("按【序号 + 书签名称】命名")

        if current_config.get('mode', 1) == 2:
            self.rb_seq_name.setChecked(True)
        else:
            self.rb_name.setChecked(True)

        l_base.addWidget(self.rb_name)
        l_base.addWidget(self.rb_seq_name)
        box_base.setLayout(l_base)
        layout.addWidget(box_base)

        box_seg = QGroupBox("2. 分段追加前后缀 (可选，不设则不追加)")
        l_seg = QVBoxLayout()

        hz_btn = QHBoxLayout()
        btn_add = QPushButton("➕ 添加分段规则")
        btn_del = QPushButton("❌ 删除选中行")
        btn_add.clicked.connect(self.add_row)
        btn_del.clicked.connect(self.del_row)
        hz_btn.addWidget(btn_add)
        hz_btn.addWidget(btn_del)
        hz_btn.addStretch()
        l_seg.addLayout(hz_btn)

        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["起始页码", "结束页码", "追加前缀", "追加后缀"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        l_seg.addWidget(self.table)
        box_seg.setLayout(l_seg)
        layout.addWidget(box_seg)

        for seg in current_config.get('segments', []):
            self.add_row_data(seg)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def add_row(self):
        self.add_row_data([1, 10, "前缀_", "_后缀"])

    def add_row_data(self, data):
        r = self.table.rowCount()
        self.table.insertRow(r)
        self.table.setItem(r, 0, QTableWidgetItem(str(data[0])))
        self.table.setItem(r, 1, QTableWidgetItem(str(data[1])))
        self.table.setItem(r, 2, QTableWidgetItem(str(data[2])))
        self.table.setItem(r, 3, QTableWidgetItem(str(data[3])))

    def del_row(self):
        r = self.table.currentRow()
        if r >= 0:
            self.table.removeRow(r)

    def get_config(self):
        mode = 2 if self.rb_seq_name.isChecked() else 1
        segments = []
        for r in range(self.table.rowCount()):
            try:
                start = int(self.table.item(r, 0).text())
                end = int(self.table.item(r, 1).text())
                pfx = self.table.item(r, 2).text()
                sfx = self.table.item(r, 3).text()
                segments.append([start, end, pfx, sfx])
            except ValueError:
                pass
        return {'mode': mode, 'segments': segments}


def _worker_render_temp_img(args):
    pdf_path, page_index, dpi, temp_dir, img_fmt, is_transparent = args
    try:
        import pypdfium2 as pdfium
        from PIL import Image
        pdf = pdfium.PdfDocument(pdf_path)
        scale = dpi / 72.0
        bg_color = (0, 0, 0, 0) if (img_fmt == "png" and is_transparent) else (255, 255, 255, 255)

        pil_image = pdf[page_index].render(scale=scale, rotation=0, fill_color=bg_color).to_pil()
        pdf.close()

        if img_fmt == "jpg" and pil_image.mode != 'RGB':
            pil_image = pil_image.convert('RGB')

        temp_name = f"temp_{os.path.basename(pdf_path)}_{page_index}.{img_fmt}"
        temp_path = os.path.join(temp_dir, temp_name)
        pil_image.save(temp_path, quality=95)
        del pil_image
        return (page_index, temp_path)
    except Exception as e:
        return (page_index, f"ERROR: {str(e)}")


def _worker_pdf2img_save(args):
    pdf_path, page_index, dpi, output_dir, base_name, img_fmt, is_transparent = args
    try:
        import pypdfium2 as pdfium
        from PIL import Image
        pdf = pdfium.PdfDocument(pdf_path)
        scale = dpi / 72.0
        bg_color = (0, 0, 0, 0) if (img_fmt == "png" and is_transparent) else (255, 255, 255, 255)

        pil_image = pdf[page_index].render(scale=scale, rotation=0, fill_color=bg_color).to_pil()
        pdf.close()

        if img_fmt == "jpg" and pil_image.mode != 'RGB':
            pil_image = pil_image.convert('RGB')

        save_name = f"{base_name}_page_{page_index + 1:03d}.{img_fmt}"
        save_path = os.path.join(output_dir, save_name)
        pil_image.save(save_path, quality=95)
        del pil_image
        return True
    except Exception:
        return False


class ToolkitWorker(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, paths, mode, out_dir, dpi, fmt, is_trans, split_config=None):
        super().__init__()
        self.paths = paths
        self.mode = mode
        self.out_dir = out_dir
        self.dpi = dpi
        self.fmt = fmt
        self.is_trans = is_trans
        self.split_config = split_config if split_config else {}

    def run(self):
        temp_dir = tempfile.mkdtemp(prefix="pdf_tool_tmp_")
        try:
            max_workers = max(1, os.cpu_count() - 1)

            if "合并" in self.mode:
                self.progress.emit(50, "正在合并PDF...")
                merger = PdfMerger()
                for p in self.paths:
                    if p.lower().endswith('.pdf'): merger.append(p)
                merger.write(self.out_dir)
                merger.close()

            elif self.mode == "PDF拆分为单页":
                total = len(self.paths)
                for idx, p in enumerate(self.paths):
                    self.progress.emit(int(idx / total * 100), f"正在拆分 {os.path.basename(p)}...")
                    if p.lower().endswith('.pdf'):
                        reader = PdfReader(p)
                        base = os.path.splitext(os.path.basename(p))[0]
                        for i, page in enumerate(reader.pages):
                            writer = PdfWriter()
                            writer.add_page(page)
                            with open(os.path.join(self.out_dir, f"{base}_p{i + 1}.pdf"), "wb") as f:
                                writer.write(f)

            elif "按书签拆分" in self.mode:
                import re
                total_files = len(self.paths)
                for f_idx, p in enumerate(self.paths):
                    if not p.lower().endswith('.pdf'): continue
                    self.progress.emit(int(f_idx / total_files * 100), f"正在按书签拆分 {os.path.basename(p)}...")

                    doc = fitz.open(p)
                    toc = doc.get_toc()

                    page_to_title = {item[2]: item[1] for item in toc}

                    total_pages = len(doc)
                    base_name = os.path.splitext(os.path.basename(p))[0]
                    current_bookmark = base_name

                    for i in range(total_pages):
                        page_num = i + 1
                        if page_num in page_to_title:
                            current_bookmark = page_to_title[page_num]

                        clean_title = re.sub(r'[\\/*?:"<>|]', '_', current_bookmark).strip()

                        naming_mode = self.split_config.get('mode', 1)
                        segments = self.split_config.get('segments', [])

                        prefix = ""
                        suffix = ""
                        for seg in segments:
                            if seg[0] <= page_num <= seg[1]:
                                prefix = seg[2]
                                suffix = seg[3]
                                break

                        if naming_mode == 2:
                            final_name = f"{page_num:03d}_{prefix}{clean_title}{suffix}.pdf"
                        else:
                            final_name = f"{prefix}{clean_title}{suffix}.pdf"

                        final_path = os.path.join(self.out_dir, final_name)

                        counter = 1
                        while os.path.exists(final_path):
                            name_no_ext, ext = os.path.splitext(final_name)
                            final_path = os.path.join(self.out_dir, f"{name_no_ext}_{counter}{ext}")
                            counter += 1

                        new_doc = fitz.Document()
                        new_doc.insert_pdf(doc, from_page=i, to_page=i)
                        new_doc.save(final_path)
                        new_doc.close()

                    doc.close()

            elif "多图转PDF" in self.mode:
                from PIL import Image
                self.progress.emit(50, "正在拼接图片为PDF...")
                imgs = []
                for p in self.paths:
                    try:
                        img = Image.open(p)
                        if img.mode != 'RGB': img = img.convert('RGB')
                        imgs.append(img)
                    except:
                        pass
                if imgs:
                    imgs[0].save(self.out_dir, "PDF", resolution=100.0, save_all=True, append_images=imgs[1:])

            elif "图片型PDF" in self.mode:
                import pypdfium2 as pdfium
                from PIL import Image
                tasks = []
                total_pages = 0
                for f in self.paths:
                    if f.lower().endswith('.pdf'):
                        doc = pdfium.PdfDocument(f)
                        count = len(doc)
                        doc.close()
                        for i in range(count):
                            tasks.append((f, i, self.dpi, temp_dir, "jpg", False))
                        total_pages += count

                if not tasks: raise Exception("没有找到有效的PDF页面")

                temp_files_map = {}
                completed = 0

                with concurrent.futures.ProcessPoolExecutor(max_workers=max_workers) as executor:
                    futures = {executor.submit(_worker_render_temp_img, task): i for i, task in enumerate(tasks)}
                    for future in concurrent.futures.as_completed(futures):
                        task_idx = futures[future]
                        _, result = future.result()
                        if not isinstance(result, str) or not result.startswith("ERROR"):
                            temp_files_map[task_idx] = result
                        completed += 1
                        self.progress.emit(int((completed / total_pages) * 80),
                                           f"多进程渲染中: {completed}/{total_pages}")

                self.progress.emit(90, "正在合并生成最终PDF...")
                valid_temps = [temp_files_map[i] for i in range(len(tasks)) if i in temp_files_map]
                if valid_temps:
                    img_first = Image.open(valid_temps[0])
                    img_others = [Image.open(p) for p in valid_temps[1:]]
                    img_first.save(self.out_dir, "PDF", resolution=self.dpi, save_all=True, append_images=img_others)
                    img_first.close()
                    for img in img_others: img.close()
                else:
                    raise Exception("渲染全部失败，无法生成PDF")

            elif "导出图片" in self.mode:
                import pypdfium2 as pdfium
                tasks = []
                total_pages = 0
                for f in self.paths:
                    if f.lower().endswith('.pdf'):
                        doc = pdfium.PdfDocument(f)
                        count = len(doc)
                        base = os.path.splitext(os.path.basename(f))[0]
                        doc.close()
                        for i in range(count):
                            tasks.append((f, i, self.dpi, self.out_dir, base, self.fmt, self.is_trans))
                        total_pages += count

                completed = 0
                with concurrent.futures.ProcessPoolExecutor(max_workers=max_workers) as executor:
                    futures = [executor.submit(_worker_pdf2img_save, t) for t in tasks]
                    for f in concurrent.futures.as_completed(futures):
                        completed += 1
                        self.progress.emit(int((completed / total_pages) * 95),
                                           f"导出图片中: {completed}/{total_pages}")

            self.progress.emit(100, "处理完成！")
            self.finished.emit(f"任务执行成功，文件已保存至：\n{self.out_dir}")

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.error.emit(str(e))
        finally:
            try:
                shutil.rmtree(temp_dir)
            except:
                pass


# ================= UI =================
class ToolkitWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.split_bookmark_config = {}
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)

        box_mode = QGroupBox("1. 选择工具模式")
        hl = QHBoxLayout()
        self.cmb_mode = QComboBox()
        self.cmb_mode.addItems([
            "多图转PDF",
            "多个PDF合并",
            "PDF拆分为单页",
            "按书签拆分PDF为单页",
            "PDF转图片型PDF",
            "PDF批量导出图片"
        ])
        self.cmb_mode.currentTextChanged.connect(self.update_ui_state)
        hl.addWidget(QLabel("当前模式:"))
        hl.addWidget(self.cmb_mode)

        self.btn_config_split = QPushButton("⚙️ 配置拆分规则")
        self.btn_config_split.setStyleSheet("background-color: #f39c12; color: white; font-weight: bold;")
        self.btn_config_split.clicked.connect(self.open_split_config)
        self.btn_config_split.hide()
        hl.addWidget(self.btn_config_split)

        box_mode.setLayout(hl)
        layout.addWidget(box_mode)

        box_params = QGroupBox("2. 转换参数设置 (转图片/栅格化可用)")
        pl = QHBoxLayout()
        pl.addWidget(QLabel("DPI (清晰度):"))
        self.entry_dpi = QLineEdit("200");
        self.entry_dpi.setFixedWidth(60)
        pl.addWidget(self.entry_dpi)

        pl.addWidget(QLabel("图片格式:"))
        self.cmb_fmt = QComboBox();
        self.cmb_fmt.addItems(["jpg", "png"])
        pl.addWidget(self.cmb_fmt)

        self.chk_trans = QCheckBox("保持背景透明 (仅PNG)")
        pl.addWidget(self.chk_trans)
        pl.addStretch(1)
        box_params.setLayout(pl)
        layout.addWidget(box_params)

        # 【核心更新】：采用高级的文件管理组件，自动支持按钮添加、清空、删除
        self.file_manager = FileListManagerWidget(accept_exts=['.pdf', '.jpg', '.png', '.jpeg'], title_desc="Files")
        layout.addWidget(self.file_manager)

        hz_run = QHBoxLayout()
        self.lbl_status = QLabel("就绪")
        self.progress = QProgressBar()

        self.btn_run = QPushButton("⚡ 开始执行任务")
        self.btn_run.setStyleSheet(
            "background-color: #673AB7; color: white; padding: 12px; font-weight: bold; font-size: 14px; border-radius: 4px;")
        self.btn_run.clicked.connect(self.run_tool)

        hz_run.addWidget(self.lbl_status);
        hz_run.addWidget(self.progress);
        hz_run.addWidget(self.btn_run)
        layout.addLayout(hz_run)

        self.update_ui_state()

    def update_ui_state(self):
        mode = self.cmb_mode.currentText()
        self.entry_dpi.setEnabled(mode in ["PDF转图片型PDF (修复乱码)", "PDF批量导出图片"])
        self.cmb_fmt.setEnabled(mode == "PDF批量导出图片")
        self.chk_trans.setEnabled(mode == "PDF批量导出图片")

        self.btn_config_split.setVisible(mode == "按书签拆分PDF为单页")

    def open_split_config(self):
        dialog = BookmarkSplitConfigDialog(self.split_bookmark_config, self)
        if dialog.exec_() == QDialog.Accepted:
            self.split_bookmark_config = dialog.get_config()

    def run_tool(self):
        # 使用 FileListManagerWidget 专属的方法检查数据
        if self.file_manager.count() == 0: return QMessageBox.warning(self, "提示", "请先添加文件")
        mode = self.cmb_mode.currentText()

        if "拆分" in mode or "导出图片" in mode:
            out_dir = QFileDialog.getExistingDirectory(self, "选择保存目录")
            if not out_dir: return
        else:
            out_dir, _ = QFileDialog.getSaveFileName(self, "保存为", "工具箱输出.pdf", "PDF (*.pdf)")
            if not out_dir: return

        try:
            dpi = float(self.entry_dpi.text())
        except:
            dpi = 200.0

        paths = self.file_manager.get_all_filepaths()
        self.btn_run.setEnabled(False)

        self.worker = ToolkitWorker(paths, mode, out_dir, dpi, self.cmb_fmt.currentText(), self.chk_trans.isChecked(),
                                    self.split_bookmark_config)
        self.worker.progress.connect(lambda v, txt: (self.progress.setValue(v), self.lbl_status.setText(txt)))
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(lambda e: (QMessageBox.critical(self, "错误", e), self.btn_run.setEnabled(True)))
        self.worker.start()

    def on_finished(self, msg):
        self.btn_run.setEnabled(True)
        self.progress.setValue(100)
        QMessageBox.information(self, "成功", msg)