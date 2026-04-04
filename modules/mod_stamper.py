import os
import sys
import json
import uuid
import copy
import io
import math
import shutil
import subprocess
import tempfile
from PIL import Image
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                             QGroupBox, QLineEdit, QRadioButton, QMessageBox, QFileDialog, QProgressBar, QSplitter,
                             QTableWidget, QTableWidgetItem, QHeaderView, QDialog, QFormLayout, QDialogButtonBox,
                             QCheckBox, QComboBox, QApplication)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtGui import QPixmap
import fitz
from core.pdf_viewer import PDFGraphicsView

from core.utils import (detect_smart_segments, UniversalSegmentDialog,
                        find_ghostscript, run_ghostscript, MM_TO_PTS,
                        BTN_BLUE, BTN_GREEN, BTN_PURPLE, BTN_RED, BTN_GRAY, BTN_ORANGE)

# ================= 尝试导入工业级数字签名库 =================
try:
    from pyhanko.sign import signers
    from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
    from pyhanko.sign.fields import SigFieldSpec, append_signature_field
    from cryptography.hazmat.primitives.serialization import pkcs12
    from cryptography.hazmat.primitives import serialization

    PYHANKO_AVAILABLE = True
except ImportError:
    PYHANKO_AVAILABLE = False




class StamperWorker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, pdf_doc, page_positions, export_mode, output_path, file_names, prefix="", suffix="",
                 stamp_dpi=300, use_gs_compress=False, gs_quality="/ebook", gs_path=None, gs_lib_path=None,
                 use_pki=False, pfx_path="", pfx_pwd="", pki_target_id="first"):
        super().__init__()
        self.pdf_doc = pdf_doc
        self.page_positions = page_positions
        self.export_mode = export_mode
        self.output_path = output_path
        self.file_names = file_names
        self.prefix = prefix
        self.suffix = suffix

        self.stamp_dpi = stamp_dpi
        self.use_gs_compress = use_gs_compress
        self.gs_quality = gs_quality
        self.gs_path = gs_path
        self.gs_lib_path = gs_lib_path

        self.use_pki = use_pki
        self.pfx_path = pfx_path
        self.pfx_pwd = pfx_pwd
        self.pki_target_id = pki_target_id

        self._image_cache = {}
        self._pki_target_info = None

    def _get_processed_stamp_bytes(self, stamp_path, w_mm, h_mm, effective_angle):
        cache_key = (stamp_path, w_mm, h_mm, effective_angle)
        if cache_key in self._image_cache:
            return self._image_cache[cache_key]

        target_px_w = int((w_mm / 25.4) * self.stamp_dpi)
        target_px_h = int((h_mm / 25.4) * self.stamp_dpi)

        with Image.open(stamp_path) as img:
            img = img.convert("RGBA")
            resample_filter = getattr(Image, 'Resampling', Image).LANCZOS

            img = img.resize((target_px_w, target_px_h), resample=resample_filter)

            if effective_angle != 0:
                img = img.rotate(-effective_angle, expand=True, resample=resample_filter)

            stream = io.BytesIO()
            img.save(stream, format="PNG", optimize=True)
            img_bytes = stream.getvalue()

            self._image_cache[cache_key] = img_bytes
            return img_bytes

    def _apply_all_stamps_to_page(self, page, global_page_num):
        stamps_list = self.page_positions.get(global_page_num, [])
        page_rot = page.rotation

        for stamp in stamps_list:
            if not os.path.exists(stamp['path']): continue

            w_mm, h_mm = stamp['w'], stamp['h']
            w_pts, h_pts = w_mm * MM_TO_PTS, h_mm * MM_TO_PTS
            pdf_x, pdf_y = stamp['pdf_x'], stamp['pdf_y']

            visual_angle = stamp.get('angle', 0) % 360
            phys_angle = (visual_angle - page_rot) % 360

            img_bytes = self._get_processed_stamp_bytes(stamp['path'], w_mm, h_mm, phys_angle)

            cx = pdf_x + w_pts / 2
            cy = pdf_y + h_pts / 2
            visual_center = fitz.Point(cx, cy)

            phys_center = visual_center * page.derotation_matrix

            rad = math.radians(phys_angle)
            new_w_pts = abs(w_pts * math.cos(rad)) + abs(h_pts * math.sin(rad))
            new_h_pts = abs(w_pts * math.sin(rad)) + abs(h_pts * math.cos(rad))

            a_rect = fitz.Rect(
                phys_center.x - new_w_pts / 2,
                phys_center.y - new_h_pts / 2,
                phys_center.x + new_w_pts / 2,
                phys_center.y + new_h_pts / 2
            )

            try:
                page.insert_image(a_rect, stream=img_bytes, keep_proportion=False)
            except TypeError:
                page.insert_image(a_rect, stream=img_bytes)

            is_target = False
            if self.pki_target_id == "first" and self._pki_target_info is None:
                is_target = True
            elif self.pki_target_id == stamp.get('id') and self._pki_target_info is None:
                is_target = True

            # 💡 这里我们不仅记录坐标，还要记录所在页面的高度 (PyHanko的坐标系和PyMuPDF的Y轴是反着的)
            if self.use_pki and is_target:
                self._pki_target_info = {
                    'page_idx': global_page_num,
                    'rect': a_rect,
                    'page_height': page.rect.height
                }

    def _apply_pki_signature(self, input_pdf, output_pdf):
        if not PYHANKO_AVAILABLE:
            shutil.copy2(input_pdf, output_pdf)
            return

        from pyhanko.sign.fields import SigFieldSpec, append_signature_field
        from pyhanko.sign import signers
        from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
        from cryptography.hazmat.primitives.serialization import pkcs12
        from cryptography.hazmat.primitives import serialization

        try:
            with open(self.pfx_path, "rb") as f:
                pfx_bytes = f.read()
            private_key, cert, _ = pkcs12.load_key_and_certificates(
                pfx_bytes, self.pfx_pwd.encode('utf-8')
            )
        except Exception as e:
            raise ValueError(f"证书提取失败 (密码错误或证书损坏): {e}")

        if private_key is None:
            raise ValueError("提取失败！该 .pfx 文件不包含私钥 (Private Key)，无法用于签名！")

        with tempfile.TemporaryDirectory() as tmpdir:
            key_path = os.path.join(tmpdir, "tmp_key.pem")
            cert_path = os.path.join(tmpdir, "tmp_cert.pem")

            with open(key_path, "wb") as f:
                f.write(private_key.private_bytes(
                    encoding=serialization.Encoding.PEM,
                    format=serialization.PrivateFormat.PKCS8,
                    encryption_algorithm=serialization.NoEncryption()
                ))
            with open(cert_path, "wb") as f:
                f.write(cert.public_bytes(serialization.Encoding.PEM))

            signer = signers.SimpleSigner.load(
                key_file=key_path,
                cert_file=cert_path,
                key_passphrase=None
            )

        shutil.copy2(input_pdf, output_pdf)
        with open(output_pdf, 'r+b') as f:
            w = IncrementalPdfFileWriter(f)

            # 画出合法的点击热区坑位
            if self._pki_target_info:
                target = self._pki_target_info
                # 坐标系换算（PyHanko原点在左下角，与PyMuPDF相反）
                box = (
                    target['rect'].x0,
                    target['page_height'] - target['rect'].y1,
                    target['rect'].x1,
                    target['page_height'] - target['rect'].y0
                )
                sig_field = SigFieldSpec('DocumentSecurityLock', on_page=target['page_idx'], box=box)
                append_signature_field(w, sig_field)

            meta = signers.PdfSignatureMetadata(
                field_name='DocumentSecurityLock',
                location="System",
                reason="文档防篡改加密锁定"
            )

            # ================= 🚀 终极真理：正确实例化 PdfSigner 类 =================
            try:
                from pyhanko.stamp import TextStampStyle
                # 创建一个完美的透明样式：文字为一个空格，边框厚度为0
                style = TextStampStyle(stamp_text=' ', border_width=0)

                # 正确用法：样式参数是传给 PdfSigner 类的，绝不是传给 sign_pdf 函数！
                pdf_signer = signers.PdfSigner(
                    signature_meta=meta,
                    signer=signer,
                    stamp_style=style
                )
                # 使用实例化后的专属对象去签名
                pdf_signer.sign_pdf(w, in_place=True)

            except Exception as e:
                # 针对极老版本 pyhanko 的兜底容错（通常不会触发）
                signers.sign_pdf(
                    w, meta, signer=signer, in_place=True,
                    appearance_text_params={'stamp_text': ' ', 'border_width': 0}
                )
            # ========================================================================

    def _process_single_document(self, doc, final_out_path):
        self._pki_target_info = None

        for page_num in range(len(doc)):
            self._apply_all_stamps_to_page(doc[page_num], page_num)

        tmp_visual = final_out_path + ".v.tmp.pdf"
        tmp_gs = final_out_path + ".g.tmp.pdf"

        try:
            doc.save(tmp_visual)
            doc.close()
        except Exception as e:
            if not doc.is_closed: doc.close()
            raise Exception(f"保存视觉图章中间文件失败: {e}")

        current_file = tmp_visual

        if self.use_gs_compress and self.gs_path:
            try:
                run_ghostscript(self.gs_path, self.gs_lib_path, current_file, tmp_gs, quality=self.gs_quality)
                current_file = tmp_gs
            except Exception as e:
                print(f"GS压缩失败，跳过: {e}")


        if self.use_pki and PYHANKO_AVAILABLE:
            self._apply_pki_signature(current_file, final_out_path)
        else:
            shutil.copy2(current_file, final_out_path)

        for t in [tmp_visual, tmp_gs]:
            if os.path.exists(t):
                try:
                    os.remove(t)
                except:
                    pass

    def run(self):
        try:
            self._image_cache.clear()
            total_pages = len(self.pdf_doc)

            if self.export_mode == 'merged':
                export_doc = fitz.Document()
                export_doc.insert_pdf(self.pdf_doc)
                self._process_single_document(export_doc, self.output_path)

            elif self.export_mode == 'batch':
                toc = self.pdf_doc.get_toc(simple=False)
                bookmarks = [item for item in toc if item[0] == 1]
                if not bookmarks: bookmarks = [[1, "Document", 1]]
                for i, bm in enumerate(bookmarks):
                    start_page = bm[2] - 1
                    end_page = bookmarks[i + 1][2] - 1 if i + 1 < len(bookmarks) else total_pages
                    new_doc = fitz.Document()
                    new_doc.insert_pdf(self.pdf_doc, from_page=start_page, to_page=end_page - 1)

                    original_name = self.file_names[i] if i < len(self.file_names) else f"Batch_{i}.pdf"
                    base_name = original_name.replace('.pdf', '')
                    final_name = f"{self.prefix}{base_name}{self.suffix}.pdf"

                    self._process_single_document(new_doc, os.path.join(self.output_path, final_name))
                    self.progress.emit(int((i + 1) / len(bookmarks) * 90))

            elif self.export_mode == 'split':
                toc = self.pdf_doc.get_toc(simple=False)
                bookmarks = [item for item in toc if item[0] == 1]
                if not bookmarks: bookmarks = [[1, "Document", 1]]
                for i, bm in enumerate(bookmarks):
                    start_page = bm[2] - 1
                    end_page = bookmarks[i + 1][2] - 1 if i + 1 < len(bookmarks) else total_pages
                    file_page_count = end_page - start_page
                    base_name = bm[1].replace('.pdf', '')
                    for local_idx, global_page_num in enumerate(range(start_page, end_page)):
                        new_doc = fitz.Document()
                        new_doc.insert_pdf(self.pdf_doc, from_page=global_page_num, to_page=global_page_num)

                        if file_page_count == 1:
                            final_name = f"{self.prefix}{base_name}{self.suffix}.pdf"
                        else:
                            final_name = f"{self.prefix}{base_name}_第{local_idx + 1}页{self.suffix}.pdf"

                        self._process_single_document(new_doc, os.path.join(self.output_path, final_name))
                        self.progress.emit(int((global_page_num + 1) / total_pages * 90))

            self._image_cache.clear()
            self.progress.emit(100)
            self.finished.emit("🎉 所有加盖、压缩、防伪签名操作已完美执行完毕！")
        except Exception as e:
            import traceback;
            print(traceback.format_exc())
            self.error.emit(str(e))


class StamperWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.pdf_doc = None
        self.segments = []
        self.page_stamp_positions = {}
        self.original_filenames = []
        self.global_stamps = []

        self.gs_path, self.gs_lib_path = find_ghostscript()
        self.pfx_path = ""
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(10, 10, 10, 10)
        splitter = QSplitter(Qt.Vertical)

        top_widget = QWidget()
        top_layout = QHBoxLayout(top_widget)
        top_layout.setContentsMargins(0, 0, 0, 0)

        # =============== 区域 1：文件与配置 ===============
        box_left = QGroupBox("1. 文件与多图章配置")
        l_left = QVBoxLayout()

        hz_cfg = QHBoxLayout()
        btn_import_cfg = QPushButton("📂 导入图章配置")
        btn_import_cfg.clicked.connect(self.import_config)
        btn_export_cfg = QPushButton("💾 导出图章配置")
        btn_export_cfg.clicked.connect(self.export_config)
        hz_cfg.addWidget(btn_import_cfg);
        hz_cfg.addWidget(btn_export_cfg)
        l_left.addLayout(hz_cfg)

        from core.ui_components import FileListManagerWidget
        self.file_manager = FileListManagerWidget(accept_exts=['.pdf'], title_desc="PDF Files (*.pdf)")
        l_left.addWidget(self.file_manager, 1)

        self.chk_pre_gs = QCheckBox("🛠️ 合并前用GS扁平化 (破除Acrobat密码锁)")
        self.chk_pre_gs.setStyleSheet("color: #E67E22; font-weight: bold;")
        if not self.gs_path:
            self.chk_pre_gs.setEnabled(False);
            self.chk_pre_gs.setText("🛠️ 合并前用GS扁平化 (未检测到环境)")
        l_left.addWidget(self.chk_pre_gs)

        btn_merge = QPushButton("🔄 生成合并预览 (首选必点)")
        btn_merge.setStyleSheet(BTN_BLUE)
        btn_merge.clicked.connect(self.merge_pdfs)
        l_left.addWidget(btn_merge)

        btn_clean = QPushButton("🧹 快速清除原PDF附带的图章批注")
        btn_clean.setStyleSheet(BTN_ORANGE)
        btn_clean.clicked.connect(self.clean_original_stamps)
        l_left.addWidget(btn_clean)

        hz_stamp = QHBoxLayout()
        btn_add_stamp = QPushButton("➕ 增加图章/签名")
        btn_add_stamp.clicked.connect(self.add_stamp_item)
        btn_del_stamp = QPushButton("❌ 删除选中印章/签名")
        btn_del_stamp.clicked.connect(self.del_stamp_item)
        hz_stamp.addWidget(btn_add_stamp);
        hz_stamp.addWidget(btn_del_stamp)
        l_left.addLayout(hz_stamp)

        self.stamp_table = QTableWidget(0, 3)
        self.stamp_table.setHorizontalHeaderLabels(["标记名称", "宽(mm)", "高(mm)"])
        self.stamp_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        l_left.addWidget(self.stamp_table, 1)

        box_left.setLayout(l_left);
        top_layout.addWidget(box_left, 1)

        # =============== 区域 2：排布与导出 ===============
        box_right = QGroupBox("2. 智能排布、压缩与防伪签名导出")
        l_right = QVBoxLayout()

        btn_detect = QPushButton("① 智能按图纸尺寸分段组合")
        btn_detect.setStyleSheet(BTN_GRAY)
        btn_detect.clicked.connect(self.detect_segments)

        btn_preview = QPushButton("② 进入终极自由微调预览")
        btn_preview.setStyleSheet(BTN_PURPLE)
        btn_preview.clicked.connect(self.enter_final_preview)
        l_right.addWidget(btn_detect);
        l_right.addWidget(btn_preview)

        l_right.addWidget(QLabel("导出拆分模式:"))
        self.radio_merged = QRadioButton("合并为单一文件")
        self.radio_merged.setChecked(True)
        self.radio_batch = QRadioButton("批量按原文件拆分")
        self.radio_split = QRadioButton("拆分为单页独立文件")
        l_right.addWidget(self.radio_merged)
        l_right.addWidget(self.radio_batch)
        l_right.addWidget(self.radio_split)

        hz_fix = QHBoxLayout()
        self.input_prefix = QLineEdit()
        self.input_prefix.setPlaceholderText("前缀")
        self.input_suffix = QLineEdit()
        self.input_suffix.setPlaceholderText("后缀")
        hz_fix.addWidget(self.input_prefix)
        hz_fix.addWidget(self.input_suffix)
        l_right.addLayout(hz_fix)

        l_right.addStretch()

        # —— 画质控制 ——
        hz_stamp_dpi = QHBoxLayout()
        hz_stamp_dpi.addWidget(QLabel("🖼️ 图章画质:"))
        self.cmb_stamp_dpi = QComboBox()
        self.cmb_stamp_dpi.addItems(["150", "300", "72", "600"])
        hz_stamp_dpi.addWidget(self.cmb_stamp_dpi)
        hz_stamp_dpi.addStretch()
        l_right.addLayout(hz_stamp_dpi)

        hz_gs = QHBoxLayout()
        self.chk_gs = QCheckBox("🗜️ 启用 GS 二次全局压缩")
        self.cmb_gs_quality = QComboBox()
        self.cmb_gs_quality.addItems(["/screen (72dpi)", "/ebook (150dpi)", "/printer (300dpi)"])
        self.cmb_gs_quality.setCurrentIndex(1)
        if not self.gs_path: self.chk_gs.setEnabled(False)
        hz_gs.addWidget(self.chk_gs)
        hz_gs.addWidget(self.cmb_gs_quality)
        l_right.addLayout(hz_gs)

        # —— PKI 证书 ——
        self.chk_pki = QCheckBox("🛡️ 附加防篡改保护锁 (防止他人用编辑器删图章)")
        if not PYHANKO_AVAILABLE:
            self.chk_pki.setEnabled(False)
            self.chk_pki.setText("🛡️ 附加防篡改保护锁 (⚠️请安装 pyhanko/cryptography)")
        l_right.addWidget(self.chk_pki)

        self.widget_pki_params = QWidget()
        fl_pki = QFormLayout(self.widget_pki_params)
        fl_pki.setContentsMargins(20, 0, 0, 0)

        hz_pfx = QHBoxLayout()
        self.lbl_pfx = QLabel("未选择")
        btn_pfx = QPushButton("浏览")
        btn_pfx.clicked.connect(self.select_pfx)
        hz_pfx.addWidget(self.lbl_pfx)
        hz_pfx.addWidget(btn_pfx)

        self.entry_pwd = QLineEdit();
        self.entry_pwd.setEchoMode(QLineEdit.Password)
        self.entry_pwd.setPlaceholderText("提取密码")

        hz_pki_target = QHBoxLayout()
        self.cmb_pki_target = QComboBox()
        self.cmb_pki_target.addItem("默认 (遇到第一个图章处)", userData="first")
        hz_pki_target.addWidget(self.cmb_pki_target)
        hz_pki_target.addStretch()

        fl_pki.addRow(".pfx证书:", hz_pfx)
        fl_pki.addRow("证书密码:", self.entry_pwd)
        fl_pki.addRow("锁定关联区域:", hz_pki_target)
        self.widget_pki_params.hide()

        self.chk_pki.toggled.connect(self.widget_pki_params.setVisible)
        l_right.addWidget(self.widget_pki_params)

        btn_export = QPushButton("🚀 终极执行：可视化盖章 ➡️ 压缩 ➡️ 防篡改导出")
        btn_export.setStyleSheet(BTN_GREEN)
        btn_export.clicked.connect(self.start_export)
        l_right.addWidget(btn_export)

        self.progress_bar = QProgressBar()
        l_right.addWidget(self.progress_bar)

        box_right.setLayout(l_right);
        top_layout.addWidget(box_right, 1)
        splitter.addWidget(top_widget)

        # =============== 预览区 ===============
        bottom_widget = QWidget()
        bottom_layout = QVBoxLayout(bottom_widget)
        bottom_layout.setContentsMargins(0, 0, 0, 0)

        hz_preview_tools = QHBoxLayout()
        self.btn_confirm_pos = QPushButton("✅ 确认本页位置应用到全部")
        self.btn_confirm_pos.setStyleSheet(BTN_RED)
        self.btn_confirm_pos.hide()
        self.btn_confirm_pos.clicked.connect(self.confirm_stamp_position)

        self.btn_add_missing = QPushButton("➕ 误删找回：向本页补加图章")
        self.btn_add_missing.setStyleSheet(BTN_BLUE)
        self.btn_add_missing.hide()
        self.btn_add_missing.clicked.connect(self.add_stamp_to_current_preview)

        hz_preview_tools.addWidget(self.btn_confirm_pos)
        hz_preview_tools.addWidget(self.btn_add_missing)
        hz_preview_tools.addStretch()

        self.preview_view = PDFGraphicsView()
        bottom_layout.addLayout(hz_preview_tools)
        bottom_layout.addWidget(self.preview_view)

        splitter.addWidget(bottom_widget);
        splitter.setSizes([450, 550])
        main_layout.addWidget(splitter)

    def select_pfx(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择数字证书", "", "PFX Certificates (*.pfx *.p12)")
        if path:
            self.pfx_path = path
            self.lbl_pfx.setText(os.path.basename(path))

    def merge_pdfs(self):
        if self.file_manager.count() == 0: return QMessageBox.warning(self, "提示", "请先添加PDF文件！")
        self.pdf_doc = fitz.Document()
        self.original_filenames = []
        toc_list = []
        filepaths = self.file_manager.get_all_filepaths()
        use_pre_flatten = self.chk_pre_gs.isChecked() and self.gs_path

        for i, path in enumerate(filepaths):
            self.original_filenames.append(os.path.basename(path))
            start_page = len(self.pdf_doc)
            toc_list.append([1, os.path.basename(path), start_page + 1])
            if use_pre_flatten:
                self.progress_bar.setValue(int(i / len(filepaths) * 100))
                QApplication.processEvents()
                tmp_flatten_path = path + ".flatten.tmp.pdf"
                try:
                    run_ghostscript(self.gs_path, self.gs_lib_path, path, tmp_flatten_path, quality="/printer")
                    doc = fitz.open(tmp_flatten_path)
                # if self.gs_lib_path: cmd.insert(1, f"-I{self.gs_lib_path}")
                # cmd.extend([f"-sOutputFile={tmp_flatten_path}", path])
                # try:
                #     subprocess.run(cmd, creationflags=subprocess.CREATE_NO_WINDOW, check=True)
                #     doc = fitz.open(tmp_flatten_path)
                #     self.pdf_doc.insert_pdf(doc)
                #     doc.close()
                #     os.remove(tmp_flatten_path)
                except Exception as e:
                    print(f"扁平化失败 {path}: {e}")
                    doc = fitz.open(path)
                    self.pdf_doc.insert_pdf(doc)
                    doc.close()
            else:
                doc = fitz.open(path)
                self.pdf_doc.insert_pdf(doc)
                doc.close()

        self.pdf_doc.set_toc(toc_list)
        self.preview_view.load_pdf(self.pdf_doc)
        self.progress_bar.setValue(0)

        msg = f"生成大纲预览成功，共计 {len(self.pdf_doc)} 页。"
        if use_pre_flatten: msg += "\n\n✅ 已在后台完成强力扁平化！您现在可以任意盖章，原文件已解除所有锁定！"
        QMessageBox.information(self, "成功", msg)

    def add_stamp_to_current_preview(self):
        if not self.global_stamps: return QMessageBox.warning(self, "提示", "全局图章库为空！")
        current_page = getattr(self.preview_view, 'current_page', None)
        if current_page is None: return QMessageBox.warning(self, "提示", "请先加载预览！")

        dialog = QDialog(self)
        dialog.setWindowTitle("找回图章")
        layout = QVBoxLayout(dialog)
        cmb = QComboBox()
        for st in self.global_stamps: cmb.addItem(st['name'], userData=st)
        layout.addWidget(QLabel("选择要加回当前画布的图章:"))
        layout.addWidget(cmb)
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(dialog.accept)
        btns.rejected.connect(dialog.reject)
        layout.addWidget(btns)

        if dialog.exec_() == QDialog.Accepted:
            new_stamp = copy.deepcopy(cmb.currentData())
            new_stamp['pdf_x'] = 100
            new_stamp['pdf_y'] = 100
            if hasattr(self.preview_view, 'save_current_page_state'): self.preview_view.save_current_page_state()
            if current_page not in self.preview_view.page_data_dict: self.preview_view.page_data_dict[current_page] = []
            self.preview_view.page_data_dict[current_page].append(new_stamp)
            mode = getattr(self.preview_view, 'mode', 'stamp_final')
            self.preview_view.load_pdf(self.pdf_doc, target_page=current_page, mode=mode,
                                       data_dict=self.preview_view.page_data_dict)

    def _ensure_segment_stamps(self, seg):
        if not seg.get('pos_set', False):
            first_page = seg['pages'][0]
            if first_page in self.page_stamp_positions and self.page_stamp_positions[first_page]:
                seg['pos_pct'] = copy.deepcopy(self.page_stamp_positions[first_page])
            else:
                default_stamps = copy.deepcopy(self.global_stamps)
                for i, st in enumerate(default_stamps): st['pdf_x'] = 100; st['pdf_y'] = 100 + i * 150
                seg['pos_pct'] = default_stamps
            seg['pos_set'] = True

    def clean_original_stamps(self):
        if not self.pdf_doc: return QMessageBox.warning(self, "提示", "请先合并预览加载文档！")
        deleted_count = 0

        for page_num in range(len(self.pdf_doc)):
            page = self.pdf_doc[page_num]

            for widget in page.widgets():
                page.delete_widget(widget)
                deleted_count += 1

            for annot in page.annots():
                page.delete_annot(annot)
                deleted_count += 1

        if deleted_count > 0:
            try:
                self.preview_view.load_pdf(self.pdf_doc, target_page=self.preview_view.current_page,
                                           mode=getattr(self.preview_view, 'mode', 'normal'),
                                           data_dict=self.page_stamp_positions)
            except Exception:
                self.preview_view.load_pdf(self.pdf_doc)
            QMessageBox.information(self, "清理成功", f"成功清除了 {deleted_count} 个表单控件或悬浮批注！")
        else:
            msg = ("未发现任何可悬浮的签名或批注。\n\n"
                   "⚠️ 核心原因科普：\n"
                   "1. 如果图章是本软件之前盖的：我们采用的是【底层像素写入】，它已经融化在纸张上了，没有任何软件能删掉它。\n"
                   "2. 如果你勾选了【合并前用GS扁平化】：GS引擎已经把原先可以删的批注死死压进了纸张里，它变成了普通图片，自然无法清除。\n\n"
                   "👉 如果你想清除别人文件里的签名，请【不要】勾选GS扁平化，直接加载并点击本按钮！")
            QMessageBox.warning(self, "无法清理", msg)

    def export_config(self):
        if not self.pdf_doc or not self.page_stamp_positions: return QMessageBox.warning(self, "错误", "无配置。")
        size_configs = {}
        for p_idx, stamp_list in self.page_stamp_positions.items():
            if not stamp_list: continue
            page = self.pdf_doc[p_idx];
            size_key = f"{int(round(page.rect.width))}x{int(round(page.rect.height))}"
            if size_key not in size_configs: size_configs[size_key] = stamp_list
        export_data = {"module": "Stamper", "global_stamps": self.global_stamps, "size_configs": size_configs}
        path, _ = QFileDialog.getSaveFileName(self, "保存图章配置", "图章排版配置.json", "JSON Files (*.json)")
        if path:
            with open(path, 'w', encoding='utf-8') as f: json.dump(export_data, f, ensure_ascii=False, indent=4)
            QMessageBox.information(self, "成功", "已保存！")

    def import_config(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择图章配置文件", "", "JSON Files (*.json)")
        if not path: return
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            self.global_stamps = data.get("global_stamps", [])
            self.refresh_stamp_table()
            self.page_stamp_positions.clear()
            if self.pdf_doc:
                for i in range(len(self.pdf_doc)):
                    page = self.pdf_doc[i]
                    size_key = f"{int(round(page.rect.width))}x{int(round(page.rect.height))}"
                    if size_key in data.get("size_configs", {}): self.page_stamp_positions[i] = copy.deepcopy(
                        data.get("size_configs")[size_key])
                QMessageBox.information(self, "导入成功", "匹配当前PDF尺寸成功。")
            else:
                QMessageBox.information(self, "导入成功", "图章库已载入。")
        except Exception as e:
            QMessageBox.critical(self, "失败", f"异常:\n{e}")

    def add_stamp_item(self):
        file, _ = QFileDialog.getOpenFileName(self, "选择图章", "", "Image Files (*.png *.jpg *.jpeg)")
        if not file: return
        pixmap = QPixmap(file)
        orig_w = max(1, pixmap.width())
        orig_h = max(1, pixmap.height())
        aspect_ratio = orig_w / orig_h
        dialog = QDialog(self)
        dialog.setWindowTitle("属性")
        layout = QFormLayout(dialog)
        name_input = QLineEdit(f"印章_{len(self.global_stamps) + 1}")
        w_input = QLineEdit("50.0")
        h_input = QLineEdit(str(round(50.0 / aspect_ratio, 2)))
        lock_cb = QCheckBox("锁定原始宽高比")
        lock_cb.setChecked(True)

        def update_h(txt):
            if lock_cb.isChecked() and w_input.hasFocus():
                try:
                    h_input.setText(str(round(float(txt) / aspect_ratio, 2)))
                except:
                    pass

        def update_w(txt):
            if lock_cb.isChecked() and h_input.hasFocus():
                try:
                    w_input.setText(str(round(float(txt) * aspect_ratio, 2)))
                except:
                    pass

        w_input.textEdited.connect(update_h)
        h_input.textEdited.connect(update_w)
        layout.addRow("名称:", name_input)
        layout.addRow("宽(mm):", w_input)
        layout.addRow("高(mm):", h_input)
        layout.addRow("", lock_cb)
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(dialog.accept)
        btns.rejected.connect(dialog.reject)
        layout.addWidget(btns)
        if dialog.exec_() == QDialog.Accepted:
            self.global_stamps.append(
                {'id': str(uuid.uuid4()), 'name': name_input.text(), 'path': file, 'w': float(w_input.text()),
                 'h': float(h_input.text()), 'angle': 0})
            self.refresh_stamp_table()
            self._warn_reset_segments()

    def del_stamp_item(self):
        row = self.stamp_table.currentRow()
        if row >= 0: del self.global_stamps[row]; self.refresh_stamp_table(); self._warn_reset_segments()

    def refresh_stamp_table(self):
        self.stamp_table.setRowCount(len(self.global_stamps))

        current_target_id = self.cmb_pki_target.currentData() if hasattr(self, 'cmb_pki_target') else "first"
        self.cmb_pki_target.clear()
        self.cmb_pki_target.addItem("默认 (遇到第一个图章处)", userData="first")

        for i, st in enumerate(self.global_stamps):
            self.stamp_table.setItem(i, 0, QTableWidgetItem(st['name']))
            self.stamp_table.setItem(i, 1, QTableWidgetItem(str(st['w'])))
            self.stamp_table.setItem(i, 2, QTableWidgetItem(str(st['h'])))
            self.cmb_pki_target.addItem(f"绑定到: {st['name']}", userData=st['id'])

        idx = self.cmb_pki_target.findData(current_target_id)
        if idx >= 0:
            self.cmb_pki_target.setCurrentIndex(idx)

    def _warn_reset_segments(self):
        if self.segments: self.segments = []; self.page_stamp_positions = {}

    def detect_segments(self):
        if not self.pdf_doc: return QMessageBox.warning(self, "错误", "请先生成合并预览！")
        if not self.global_stamps: return QMessageBox.warning(self, "提示", "请先添加图章！")
        self.segments = detect_smart_segments(self.pdf_doc)
        for seg in self.segments:
            self._ensure_segment_stamps(seg)
            for p in seg['pages']:
                if p not in self.page_stamp_positions: self.page_stamp_positions[p] = copy.deepcopy(seg['pos_pct'])
        self.dialog = UniversalSegmentDialog(self.segments, "设置该尺寸页面的默认印章位置", self)
        self.dialog.exec_()
        for seg in self.segments:
            self._ensure_segment_stamps(seg)
            for p in seg['pages']: self.page_stamp_positions[p] = copy.deepcopy(seg['pos_pct'])

    def enter_setting_mode_from_dialog(self, seg_data, idx, dialog):
        self.current_idx = idx;
        self.dialog = dialog
        self._ensure_segment_stamps(seg_data)
        target_page = seg_data['pages'][0]
        self.preview_view.load_pdf(self.pdf_doc, target_page=target_page, mode='stamp_final',
                                   data_dict={target_page: seg_data['pos_pct']})
        self.btn_confirm_pos.show()
        self.btn_add_missing.show()

    def confirm_stamp_position(self):
        self.preview_view.save_current_page_state()
        new_stamps_list = self.preview_view.page_data_dict[self.preview_view.current_page]
        self.segments[self.current_idx]['pos_pct'] = new_stamps_list
        self.segments[self.current_idx]['pos_set'] = True
        for p in self.segments[self.current_idx]['pages']: self.page_stamp_positions[p] = copy.deepcopy(new_stamps_list)
        self.btn_confirm_pos.hide()
        self.dialog.refresh_table()
        self.dialog.show()

    def enter_final_preview(self):
        if not self.pdf_doc or not self.global_stamps: return QMessageBox.warning(self, "错误", "缺少PDF或图章数据。")
        for p in range(len(self.pdf_doc)):
            if p not in self.page_stamp_positions:
                seg_stamps = copy.deepcopy(self.global_stamps)
                for i, st in enumerate(seg_stamps): st['pdf_x'] = 100; st['pdf_y'] = 100 + i * 150
                self.page_stamp_positions[p] = seg_stamps
        self.preview_view.load_pdf(self.pdf_doc, mode='stamp_final', data_dict=self.page_stamp_positions)
        self.btn_confirm_pos.hide()
        self.btn_add_missing.show()
        QMessageBox.information(self, "高级微调模式", "已进入终极预览。可以 右键/Alt键 复制、缩放、旋转图章！")

    def start_export(self):
        if not self.pdf_doc or not self.global_stamps: return
        self.preview_view.save_current_page_state()
        self.page_stamp_positions.update(copy.deepcopy(self.preview_view.page_data_dict))

        if self.chk_pki.isChecked():
            if not self.pfx_path: return QMessageBox.warning(self, "提示", "请选择 .pfx 或 .p12 证书！")
            if not self.entry_pwd.text(): return QMessageBox.warning(self, "提示", "请输入证书提取密码！")

        for p in range(len(self.pdf_doc)):
            if p not in self.page_stamp_positions:
                seg_stamps = copy.deepcopy(self.global_stamps)
                for i, st in enumerate(seg_stamps): st['pdf_x'] = 100; st['pdf_y'] = 100 + i * 150
                self.page_stamp_positions[p] = seg_stamps

        mode = 'merged' if self.radio_merged.isChecked() else ('batch' if self.radio_batch.isChecked() else 'split')

        if mode == 'merged':
            path, _ = QFileDialog.getSaveFileName(self, "保存", "合并加盖图纸.pdf", "PDF (*.pdf)")
            if not path: return
        else:
            path = QFileDialog.getExistingDirectory(self, "选择保存目录")
            if not path: return

        target_dpi = int(self.cmb_stamp_dpi.currentText())
        use_gs = self.chk_gs.isChecked()
        quality_str = self.cmb_gs_quality.currentText().split(" ")[0]
        use_pki = self.chk_pki.isChecked()

        selected_pki_target_id = self.cmb_pki_target.currentData() if hasattr(self, 'cmb_pki_target') else "first"

        self.worker = StamperWorker(
            pdf_doc=self.pdf_doc,
            page_positions=self.page_stamp_positions,
            export_mode=mode,
            output_path=path,
            file_names=self.original_filenames,
            prefix=self.input_prefix.text(),
            suffix=self.input_suffix.text(),
            stamp_dpi=target_dpi,
            use_gs_compress=use_gs,
            gs_quality=quality_str,
            gs_path=self.gs_path,
            gs_lib_path=self.gs_lib_path,
            use_pki=use_pki,
            pfx_path=self.pfx_path,
            pfx_pwd=self.entry_pwd.text(),
            pki_target_id=selected_pki_target_id
        )
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.finished.connect(lambda msg: QMessageBox.information(self, "大功告成", msg))
        self.worker.error.connect(lambda e: QMessageBox.critical(self, "出错了", e))
        self.worker.start()
