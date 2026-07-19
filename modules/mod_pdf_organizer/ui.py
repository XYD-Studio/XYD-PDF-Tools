# -*- coding: utf-8 -*-
import copy
import os
import tempfile
import uuid
from dataclasses import dataclass

import fitz
from PyQt5.QtCore import QObject, QRunnable, QSize, Qt, QThreadPool, QTimer, pyqtSignal
from PyQt5.QtGui import QIcon, QImage, QKeySequence, QPixmap
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QDialog,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QListView,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QShortcut,
    QSlider,
    QSpinBox,
    QToolBar,
    QVBoxLayout,
    QWidget,
)

from core.utils import first_input_directory, suggested_output_path


@dataclass
class PageEntry:
    source_path: str
    source_page: int
    rotation: int = 0
    uid: str = ""

    def __post_init__(self):
        if not self.uid:
            self.uid = uuid.uuid4().hex


class ThumbnailSignals(QObject):
    ready = pyqtSignal(str, QImage)


class ThumbnailTask(QRunnable):
    def __init__(self, entry, width):
        super().__init__()
        self.entry = copy.copy(entry)
        self.width = width
        self.signals = ThumbnailSignals()

    def run(self):
        try:
            doc = fitz.open(self.entry.source_path)
            page = doc[self.entry.source_page]
            scale = max(0.2, self.width / max(1.0, page.rect.width))
            matrix = fitz.Matrix(scale, scale).prerotate(self.entry.rotation)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            image = QImage(
                pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888
            ).copy()
            doc.close()
            self.signals.ready.emit(self.entry.uid, image)
        except Exception:
            return


class PageGrid(QListWidget):
    externalPdfsDropped = pyqtSignal(list, int)
    internalPagesDropped = pyqtSignal(list, int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setViewMode(QListView.IconMode)
        self.setFlow(QListView.LeftToRight)
        self.setWrapping(True)
        self.setResizeMode(QListView.Adjust)
        self.setMovement(QListView.Static)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setDragDropMode(QAbstractItemView.InternalMove)
        self.setDragEnabled(True)
        self.setDefaultDropAction(Qt.MoveAction)
        self.setAcceptDrops(True)
        self.viewport().setAcceptDrops(True)
        self.setDragDropOverwriteMode(False)
        self.setDropIndicatorShown(True)
        self.setSpacing(8)
        self._dragged_uids = []

    def startDrag(self, supported_actions):
        self._dragged_uids = [
            item.data(Qt.UserRole) for item in self.selectedItems()
            if item.data(Qt.UserRole)
        ]
        if not self._dragged_uids:
            return
        super().startDrag(Qt.MoveAction)
        self._dragged_uids = []

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            return
        if event.source() is self:
            super().dragEnterEvent(event)
            event.setDropAction(Qt.MoveAction)
            event.accept()
            return
        super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            return
        if event.source() is self:
            super().dragMoveEvent(event)
            event.setDropAction(Qt.MoveAction)
            event.accept()
            return
        super().dragMoveEvent(event)

    def _drop_row(self, position):
        item = self.itemAt(position)
        if item is None:
            return self.count()
        row = self.row(item)
        rect = self.visualItemRect(item)
        return row + (1 if position.x() >= rect.center().x() else 0)

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            paths = [url.toLocalFile() for url in event.mimeData().urls()]
            paths = [path for path in paths if path.lower().endswith('.pdf') and os.path.isfile(path)]
            if paths:
                self.externalPdfsDropped.emit(paths, self._drop_row(event.pos()))
                event.acceptProposedAction()
            return
        if event.source() is self and self._dragged_uids:
            self.internalPagesDropped.emit(list(self._dragged_uids), self._drop_row(event.pos()))
            event.setDropAction(Qt.MoveAction)
            event.accept()
            return
        super().dropEvent(event)


class ReaderDialog(QDialog):
    def __init__(self, entries, start_index=0, parent=None):
        super().__init__(parent)
        self.entries = copy.deepcopy(entries)
        self.index = max(0, min(start_index, len(entries) - 1))
        self.zoom = 1.0
        self.fit_mode = True
        self.docs = {}
        self.setWindowTitle("PDF 阅读")
        self.resize(1100, 800)

        layout = QVBoxLayout(self)
        controls = QHBoxLayout()
        self.btn_prev = QPushButton("◀")
        self.btn_prev.setToolTip("上一页")
        self.btn_next = QPushButton("▶")
        self.btn_next.setToolTip("下一页")
        self.page_spin = QSpinBox()
        self.page_spin.setRange(1, max(1, len(entries)))
        self.page_spin.setValue(self.index + 1)
        self.total_label = QLabel(f"/ {len(entries)}")
        self.btn_zoom_out = QPushButton("−")
        self.btn_zoom_out.setToolTip("缩小")
        self.btn_zoom_in = QPushButton("+")
        self.btn_zoom_in.setToolTip("放大")
        self.btn_fit = QPushButton("适合窗口")
        self.btn_full = QPushButton("全屏")
        controls.addWidget(self.btn_prev)
        controls.addWidget(self.btn_next)
        controls.addStretch()
        controls.addWidget(self.page_spin)
        controls.addWidget(self.total_label)
        controls.addStretch()
        controls.addWidget(self.btn_zoom_out)
        controls.addWidget(self.btn_zoom_in)
        controls.addWidget(self.btn_fit)
        controls.addWidget(self.btn_full)
        layout.addLayout(controls)

        self.image_label = QLabel()
        self.image_label.setAlignment(Qt.AlignCenter)
        self.image_label.setStyleSheet("background: #3A3A3A;")
        self.scroll = QScrollArea()
        self.scroll.setWidget(self.image_label)
        self.scroll.setWidgetResizable(False)
        self.scroll.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.scroll, 1)

        self.btn_prev.clicked.connect(lambda: self.set_page(self.index - 1))
        self.btn_next.clicked.connect(lambda: self.set_page(self.index + 1))
        self.page_spin.valueChanged.connect(lambda value: self.set_page(value - 1))
        self.btn_zoom_in.clicked.connect(lambda: self.change_zoom(1.25))
        self.btn_zoom_out.clicked.connect(lambda: self.change_zoom(0.8))
        self.btn_fit.clicked.connect(self.fit_page)
        self.btn_full.clicked.connect(self.toggle_fullscreen)
        QShortcut(QKeySequence("F11"), self, self.toggle_fullscreen)
        QShortcut(QKeySequence("Esc"), self, self.exit_fullscreen_or_close)
        self.render_page()

    def _doc(self, path):
        if path not in self.docs:
            self.docs[path] = fitz.open(path)
        return self.docs[path]

    def set_page(self, index):
        if 0 <= index < len(self.entries) and index != self.index:
            self.index = index
            self.page_spin.blockSignals(True)
            self.page_spin.setValue(index + 1)
            self.page_spin.blockSignals(False)
            self.render_page()

    def change_zoom(self, factor):
        self.fit_mode = False
        self.zoom = max(0.2, min(5.0, self.zoom * factor))
        self.render_page()

    def fit_page(self):
        self.fit_mode = True
        self.render_page()

    def render_page(self):
        if not self.entries:
            return
        entry = self.entries[self.index]
        page = self._doc(entry.source_path)[entry.source_page]
        if self.fit_mode:
            viewport = self.scroll.viewport().size()
            page_width, page_height = page.rect.width, page.rect.height
            if entry.rotation % 180:
                page_width, page_height = page_height, page_width
            self.zoom = max(0.2, min(
                (viewport.width() - 24) / max(1.0, page_width),
                (viewport.height() - 24) / max(1.0, page_height),
            ))
        pix = page.get_pixmap(
            matrix=fitz.Matrix(self.zoom, self.zoom).prerotate(entry.rotation),
            alpha=False,
        )
        image = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888).copy()
        self.image_label.setPixmap(QPixmap.fromImage(image))
        self.image_label.resize(image.size())

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.fit_mode:
            QTimer.singleShot(0, self.render_page)

    def toggle_fullscreen(self):
        self.showNormal() if self.isFullScreen() else self.showFullScreen()

    def exit_fullscreen_or_close(self):
        self.showNormal() if self.isFullScreen() else self.close()

    def closeEvent(self, event):
        for doc in self.docs.values():
            doc.close()
        self.docs.clear()
        super().closeEvent(event)


class PDFOrganizerWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.entries = []
        self.undo_stack = []
        self.redo_stack = []
        self.current_path = ""
        self.save_path = ""
        self.dirty = False
        self.thread_pool = QThreadPool(self)
        self.thread_pool.setMaxThreadCount(max(2, min(6, (os.cpu_count() or 2))))
        self.thumbnail_images = {}
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)

        toolbar = QToolBar()
        toolbar.setMovable(False)
        toolbar.setToolButtonStyle(Qt.ToolButtonTextOnly)
        actions = [
            ("打开", self.open_pdf), ("添加", self.add_pdf), ("保存", self.save),
            ("另存为", self.save_as), ("撤销", self.undo), ("重做", self.redo),
            ("全选", self.select_all), ("删除", self.delete_selected),
            ("左转", lambda: self.rotate_selected(-90)),
            ("右转", lambda: self.rotate_selected(90)),
            ("全部左转", lambda: self.rotate_all(-90)),
            ("全部右转", lambda: self.rotate_all(90)),
            ("提取", self.extract_selected),
        ]
        for label, slot in actions:
            action = toolbar.addAction(label)
            action.triggered.connect(slot)
        layout.addWidget(toolbar)

        size_row = QHBoxLayout()
        self.document_label = QLabel("未打开文档")
        self.document_label.setStyleSheet("font-weight: bold;")
        self.count_label = QLabel("0 页")
        self.thumb_slider = QSlider(Qt.Horizontal)
        self.thumb_slider.setRange(100, 280)
        self.thumb_slider.setValue(160)
        self.thumb_slider.setFixedWidth(220)
        self.thumb_slider.setToolTip("缩略图大小")
        size_row.addWidget(self.document_label)
        size_row.addWidget(self.count_label)
        size_row.addStretch()
        size_row.addWidget(QLabel("缩略图"))
        size_row.addWidget(self.thumb_slider)
        layout.addLayout(size_row)

        self.grid = PageGrid()
        self.grid.setStyleSheet(
            "QListWidget { background: #FFFFFF; border: 1px solid #D9DEE5; }"
            "QListWidget::item { background: #F7F8FA; border: 1px solid #D9DEE5; padding: 6px; }"
            "QListWidget::item:selected { border: 2px solid #2980B9; background: #EAF4FB; }"
        )
        layout.addWidget(self.grid, 1)

        self.thumb_slider.valueChanged.connect(self.update_grid_size)
        self.grid.itemDoubleClicked.connect(self.open_reader)
        self.grid.externalPdfsDropped.connect(self.insert_external_pdfs)
        self.grid.internalPagesDropped.connect(self.move_pages)
        QShortcut(QKeySequence.Delete, self, self.delete_selected)
        QShortcut(QKeySequence.Undo, self, self.undo)
        QShortcut(QKeySequence.Redo, self, self.redo)
        QShortcut(QKeySequence.SelectAll, self, self.select_all)
        self.update_grid_size(self.thumb_slider.value())

    def _snapshot(self):
        return copy.deepcopy(self.entries)

    def _commit(self, new_entries):
        self.undo_stack.append(self._snapshot())
        self.redo_stack.clear()
        self.entries = new_entries
        self.set_dirty(True)
        self.refresh_grid()

    def set_dirty(self, value):
        self.dirty = value
        name = os.path.basename(self.current_path) if self.current_path else "未保存文档"
        self.document_label.setText(("* " if value else "") + name)
        self.count_label.setText(f"{len(self.entries)} 页")

    def maybe_discard_changes(self):
        if not self.dirty:
            return True
        answer = QMessageBox.question(
            self, "未保存更改", "当前页面调整尚未保存，是否放弃这些更改？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
        )
        return answer == QMessageBox.Yes

    @staticmethod
    def _has_signature(path):
        try:
            with open(path, 'rb') as source:
                if b'/ByteRange' in source.read():
                    return True
            doc = fitz.open(path)
            found = any(
                widget.field_type == fitz.PDF_WIDGET_TYPE_SIGNATURE
                for page in doc for widget in page.widgets()
            )
            doc.close()
            return found
        except Exception:
            return False

    def open_pdf(self):
        if not self.maybe_discard_changes():
            return
        path, _ = QFileDialog.getOpenFileName(self, "打开 PDF", "", "PDF (*.pdf)")
        if not path:
            return
        if self._has_signature(path):
            QMessageBox.warning(self, "签名提示", "该 PDF 已包含数字签名。编辑并另存后，原签名将失效。")
        self.entries = self._entries_from_pdf(path)
        self.current_path = path
        self.save_path = ""
        self.undo_stack.clear()
        self.redo_stack.clear()
        self.thumbnail_images.clear()
        self.set_dirty(False)
        self.refresh_grid()

    def add_pdf(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "添加 PDF", first_input_directory([self.current_path] if self.current_path else []), "PDF (*.pdf)"
        )
        if paths:
            selected_rows = sorted(index.row() for index in self.grid.selectedIndexes())
            insert_at = selected_rows[0] if selected_rows else len(self.entries)
            self.insert_external_pdfs(paths, insert_at)

    @staticmethod
    def _entries_from_pdf(path):
        doc = fitz.open(path)
        entries = [PageEntry(path, index) for index in range(len(doc))]
        doc.close()
        return entries

    def insert_external_pdfs(self, paths, index):
        inserted = []
        try:
            for path in paths:
                if self._has_signature(path):
                    QMessageBox.warning(
                        self, "签名提示",
                        f"{os.path.basename(path)} 已包含数字签名，插入并另存后原签名将失效。"
                    )
                inserted.extend(self._entries_from_pdf(path))
        except Exception as exc:
            QMessageBox.critical(self, "无法添加 PDF", str(exc))
            return
        if not inserted:
            return
        index = max(0, min(index, len(self.entries)))
        new_entries = self._snapshot()
        new_entries[index:index] = inserted
        if not self.current_path:
            self.current_path = paths[0]
        self._commit(new_entries)

    def move_pages(self, dragged_uids, destination):
        dragged_uids = set(dragged_uids)
        if not dragged_uids:
            return
        destination = max(0, min(destination, len(self.entries)))
        moved = [entry for entry in self.entries if entry.uid in dragged_uids]
        remaining = [entry for entry in self.entries if entry.uid not in dragged_uids]
        destination -= sum(
            1 for index, entry in enumerate(self.entries)
            if index < destination and entry.uid in dragged_uids
        )
        ordered = remaining[:destination] + moved + remaining[destination:]
        if [entry.uid for entry in ordered] != [entry.uid for entry in self.entries]:
            self._commit(ordered)

    def selected_rows(self):
        return sorted({index.row() for index in self.grid.selectedIndexes()})

    def select_all(self):
        self.grid.selectAll()

    def delete_selected(self):
        rows = set(self.selected_rows())
        if rows:
            self._commit([entry for index, entry in enumerate(self.entries) if index not in rows])

    def rotate_selected(self, degrees):
        rows = set(self.selected_rows())
        if not rows:
            return
        updated = self._snapshot()
        for row in rows:
            updated[row].rotation = (updated[row].rotation + degrees) % 360
            self.thumbnail_images.pop(updated[row].uid, None)
        self._commit(updated)

    def rotate_all(self, degrees):
        if not self.entries:
            return
        updated = self._snapshot()
        for entry in updated:
            entry.rotation = (entry.rotation + degrees) % 360
            self.thumbnail_images.pop(entry.uid, None)
        self._commit(updated)

    def undo(self):
        if not self.undo_stack:
            return
        self.redo_stack.append(self._snapshot())
        self.entries = self.undo_stack.pop()
        self.set_dirty(True)
        self.refresh_grid()

    def redo(self):
        if not self.redo_stack:
            return
        self.undo_stack.append(self._snapshot())
        self.entries = self.redo_stack.pop()
        self.set_dirty(True)
        self.refresh_grid()

    def update_grid_size(self, width):
        self.grid.setIconSize(QSize(width, int(width * 1.35)))
        self.grid.setGridSize(QSize(width + 30, int(width * 1.35) + 48))
        self.grid.scheduleDelayedItemsLayout()

    def refresh_grid(self):
        selected_uids = {
            self.grid.item(row).data(Qt.UserRole) for row in self.selected_rows()
            if self.grid.item(row)
        }
        self.grid.clear()
        for page_number, entry in enumerate(self.entries, 1):
            source_name = os.path.splitext(os.path.basename(entry.source_path))[0]
            item = QListWidgetItem(f"{page_number}\n{source_name} · {entry.source_page + 1}")
            item.setTextAlignment(Qt.AlignHCenter | Qt.AlignTop)
            item.setData(Qt.UserRole, entry.uid)
            item.setToolTip(f"{entry.source_path}\n源页 {entry.source_page + 1} · 旋转 {entry.rotation}°")
            cached = self.thumbnail_images.get(entry.uid)
            if cached:
                item.setIcon(QIcon(QPixmap.fromImage(cached)))
            self.grid.addItem(item)
            item.setSelected(entry.uid in selected_uids)
            if not cached:
                task = ThumbnailTask(entry, 300)
                task.signals.ready.connect(self.thumbnail_ready)
                self.thread_pool.start(task)
        self.set_dirty(self.dirty)

    def thumbnail_ready(self, uid, image):
        self.thumbnail_images[uid] = image
        for row in range(self.grid.count()):
            item = self.grid.item(row)
            if item.data(Qt.UserRole) == uid:
                item.setIcon(QIcon(QPixmap.fromImage(image)))
                break

    @staticmethod
    def _compose(entries, output_path):
        source_docs = {}
        output_doc = fitz.Document()
        try:
            for entry in entries:
                if entry.source_path not in source_docs:
                    source_docs[entry.source_path] = fitz.open(entry.source_path)
                source = source_docs[entry.source_path]
                output_doc.insert_pdf(
                    source, from_page=entry.source_page, to_page=entry.source_page,
                    links=True, annots=True,
                )
                if entry.rotation:
                    output_doc[-1].set_rotation((output_doc[-1].rotation + entry.rotation) % 360)

            toc = []
            previous_level = 0
            for source_path, source in source_docs.items():
                page_map = {}
                for output_index, entry in enumerate(entries, 1):
                    if entry.source_path == source_path:
                        page_map.setdefault(entry.source_page + 1, output_index)
                for level, title, source_page, *rest in source.get_toc():
                    if source_page not in page_map:
                        continue
                    level = max(1, min(level, previous_level + 1))
                    toc.append([level, title, page_map[source_page]])
                    previous_level = level
            if toc:
                output_doc.set_toc(toc)

            for source in source_docs.values():
                source.close()
            source_docs.clear()

            output_dir = os.path.dirname(os.path.abspath(output_path))
            os.makedirs(output_dir, exist_ok=True)
            fd, temp_path = tempfile.mkstemp(prefix="organizer_", suffix=".pdf", dir=output_dir)
            os.close(fd)
            try:
                output_doc.save(temp_path, garbage=4, deflate=True)
                output_doc.close()
                output_doc = None
                os.replace(temp_path, output_path)
            finally:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
        finally:
            if output_doc is not None:
                output_doc.close()
            for source in source_docs.values():
                source.close()

    def save(self):
        if not self.entries:
            return
        if not self.save_path:
            return self.save_as()
        self._save_to(self.save_path)

    def save_as(self):
        if not self.entries:
            return
        first_path = self.entries[0].source_path
        base = os.path.splitext(os.path.basename(first_path))[0]
        initial = suggested_output_path([first_path], f"{base}_已组织.pdf")
        path, _ = QFileDialog.getSaveFileName(self, "另存为", initial, "PDF (*.pdf)")
        if path:
            self._save_to(path)

    def _save_to(self, path):
        try:
            self._compose(self.entries, path)
            self.current_path = path
            self.save_path = path
            self.set_dirty(False)
            QMessageBox.information(self, "保存完成", f"已保存至：\n{path}")
        except Exception as exc:
            QMessageBox.critical(self, "保存失败", str(exc))

    def extract_selected(self):
        rows = self.selected_rows()
        if not rows:
            return QMessageBox.warning(self, "提示", "请先选择需要提取的页面。")
        first_path = self.entries[rows[0]].source_path
        initial = suggested_output_path([first_path], "选中页面.pdf")
        path, _ = QFileDialog.getSaveFileName(self, "提取选中页面", initial, "PDF (*.pdf)")
        if not path:
            return
        try:
            self._compose([self.entries[row] for row in rows], path)
            QMessageBox.information(self, "提取完成", f"已提取 {len(rows)} 页。")
        except Exception as exc:
            QMessageBox.critical(self, "提取失败", str(exc))

    def open_reader(self, item):
        if not self.entries:
            return
        row = self.grid.row(item)
        dialog = ReaderDialog(self.entries, row, self)
        dialog.exec_()

    def can_close(self):
        return self.maybe_discard_changes()
