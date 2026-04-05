import os
from PyQt5.QtWidgets import QListWidget, QAbstractItemView, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QFileDialog
from PyQt5.QtCore import Qt, pyqtSignal


class DropListWidget(QListWidget):
    filesDropped = pyqtSignal(list)

    def __init__(self, accept_exts=None, parent=None):
        super().__init__(parent)
        self.accept_exts = accept_exts if accept_exts else ['.pdf']
        self.setDragDropMode(QAbstractItemView.DragDrop)
        self.setDefaultDropAction(Qt.MoveAction)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setStyleSheet("""
            QListWidget { background-color: #FFFFFF; border: 1px solid #DFE4EA; border-radius: 6px; padding: 5px; font-size: 13px; color: #2F3542; }
            QListWidget::item { height: 28px; border-radius: 4px; padding-left: 5px; }
            QListWidget::item:selected { background-color: #ECCC68; color: #2F3542; font-weight: bold; }
            QListWidget::item:hover { background-color: #F1F2F6; }
        """)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls() or event.source() == self:
            event.accept()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls() or event.source() == self:
            event.setDropAction(Qt.MoveAction);
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasUrls() and event.source() != self:
            event.setDropAction(Qt.CopyAction)
            event.accept()
            links = []
            existing_files = [self.item(i).text() for i in range(self.count())]
            for url in event.mimeData().urls():
                if url.isLocalFile():
                    path = url.toLocalFile()
                    ext = os.path.splitext(path)[1].lower()
                    if (not self.accept_exts or ext in self.accept_exts) and path not in existing_files:
                        links.append(path);
                        self.addItem(path)
            if links: self.filesDropped.emit(links)
        else:
            super().dropEvent(event)


class FileListManagerWidget(QWidget):
    def __init__(self, accept_exts=['.pdf'], title_desc="PDF Files (*.pdf)", parent=None):
        super().__init__(parent)
        self.accept_exts = accept_exts
        self.title_desc = title_desc
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        btn_layout = QHBoxLayout()
        btn_add = QPushButton("📂 添加文件")
        btn_add.setStyleSheet(
            "background-color: #3498DB; color: white; border-radius: 4px; padding: 6px; font-weight: bold;")
        btn_add.clicked.connect(self.add_files)
        btn_del = QPushButton("➖ 移除选中")
        btn_del.setStyleSheet(
            "background-color: #E74C3C; color: white; border-radius: 4px; padding: 6px; font-weight: bold;")
        btn_del.clicked.connect(self.delete_selected)
        btn_clear = QPushButton("🗑️ 清空")
        btn_clear.setStyleSheet("background-color: #95A5A6; color: white; border-radius: 4px; padding: 6px;")
        btn_clear.clicked.connect(self.clear_all)

        btn_layout.addWidget(btn_add);
        btn_layout.addWidget(btn_del);
        btn_layout.addWidget(btn_clear)
        layout.addLayout(btn_layout)

        tool_layout = QHBoxLayout()
        btn_up = QPushButton("⬆️ 上移")
        btn_down = QPushButton("⬇️ 下移")
        btn_sort_az = QPushButton("🔤 正序 (A-Z)")
        btn_sort_za = QPushButton("🔤 倒序 (Z-A)")

        btn_up.clicked.connect(self.move_up);
        btn_down.clicked.connect(self.move_down)
        btn_sort_az.clicked.connect(lambda: self.list_widget.sortItems(Qt.AscendingOrder))
        btn_sort_za.clicked.connect(lambda: self.list_widget.sortItems(Qt.DescendingOrder))

        tool_style = "background-color: #ECF0F1; color: #2C3E50; border: 1px solid #BDC3C7; border-radius: 4px; padding: 4px;"
        for btn in [btn_up, btn_down, btn_sort_az, btn_sort_za]:
            btn.setStyleSheet(tool_style);
            tool_layout.addWidget(btn)
        layout.addLayout(tool_layout)

        self.list_widget = DropListWidget(accept_exts=self.accept_exts)
        layout.addWidget(self.list_widget)

    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择文件", "", self.title_desc)
        if files:
            existing_files = [self.list_widget.item(i).text() for i in range(self.list_widget.count())]
            for path in files:
                if path not in existing_files: self.list_widget.addItem(path)
            self.list_widget.filesDropped.emit(files)

    def delete_selected(self):
        for item in self.list_widget.selectedItems(): self.list_widget.takeItem(self.list_widget.row(item))

    def clear_all(self):
        self.list_widget.clear()

    def move_up(self):
        row_list = sorted([self.list_widget.row(item) for item in self.list_widget.selectedItems()])
        if not row_list or row_list[0] == 0: return
        for row in row_list:
            item = self.list_widget.takeItem(row)
            self.list_widget.insertItem(row - 1, item)
            item.setSelected(True)
            self.list_widget.scrollToItem(item)

    def move_down(self):
        row_list = sorted([self.list_widget.row(item) for item in self.list_widget.selectedItems()], reverse=True)
        count = self.list_widget.count()
        if not row_list or row_list[0] == count - 1: return
        for row in row_list:
            item = self.list_widget.takeItem(row)
            self.list_widget.insertItem(row + 1, item)
            item.setSelected(True)
            self.list_widget.scrollToItem(item)

    def get_all_filepaths(self):
        return [self.list_widget.item(i).text() for i in range(self.list_widget.count())]

    def count(self):
        return self.list_widget.count()