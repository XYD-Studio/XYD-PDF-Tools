from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QTableWidget, QTableWidgetItem,
                             QPushButton, QHBoxLayout, QHeaderView, QMessageBox, QInputDialog)

MM_TO_PTS = 72 / 25.4


def detect_smart_segments(pdf_doc):
    """智能按照页面尺寸进行分段"""
    segments = []
    TOLERANCE = 5.0
    for i in range(len(pdf_doc)):
        page = pdf_doc[i]
        w_mm, h_mm = page.rect.width / MM_TO_PTS, page.rect.height / MM_TO_PTS
        matched = False
        for seg in segments:
            if abs(seg['w_mm'] - w_mm) < TOLERANCE and abs(seg['h_mm'] - h_mm) < TOLERANCE:
                seg['pages'].append(i)
                matched = True
                break
        if not matched:
            segments.append({'pages': [i], 'w_mm': w_mm, 'h_mm': h_mm, 'pos_set': False})
    return segments


def format_page_ranges(pages):
    if not pages: return ""
    pages = sorted(list(set(pages)))
    ranges = []
    start = end = pages[0]
    for p in pages[1:]:
        if p == end + 1:
            end = p
        else:
            ranges.append(f"{start + 1}-{end + 1}" if start != end else f"{start + 1}")
            start = end = p
    ranges.append(f"{start + 1}-{end + 1}" if start != end else f"{start + 1}")
    return ", ".join(ranges)


def parse_page_ranges(text):
    pages = set()
    for part in text.split(','):
        part = part.strip()
        if not part: continue
        if '-' in part:
            s, e = part.split('-')
            pages.update(range(int(s) - 1, int(e)))
        else:
            pages.add(int(part) - 1)
    return sorted(list(pages))


class UniversalSegmentDialog(QDialog):
    """通用的智能分段设置对话框(出图章和OCR通用，支持手动分离页面)"""

    def __init__(self, segments, mode_title="智能分段设置", parent=None):
        super().__init__(parent)
        self.setWindowTitle(mode_title)
        self.resize(800, 450)
        self.segments = segments
        self.parent_app = parent

        layout = QVBoxLayout(self)

        # 表格
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["页码范围", "尺寸(mm)", "状态", "操作"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.table)

        # 手动操作按钮区
        btn_layout = QHBoxLayout()
        btn_extract = QPushButton("✂️ 提取选中组的页面成新组 (分离封面)")
        btn_extract.setStyleSheet("background-color: #3498DB; color: white; padding: 6px; border-radius: 4px;")
        btn_extract.clicked.connect(self.extract_pages)

        btn_delete = QPushButton("❌ 删除选中组")
        btn_delete.setStyleSheet("background-color: #E74C3C; color: white; padding: 6px; border-radius: 4px;")
        btn_delete.clicked.connect(self.delete_group)

        btn_layout.addWidget(btn_extract);
        btn_layout.addWidget(btn_delete)
        layout.addLayout(btn_layout)

        # 确认返回
        btn_confirm = QPushButton("✅ 保存所有分段并返回")
        btn_confirm.setStyleSheet(
            "background-color: #2ECC71; color: white; font-weight: bold; padding: 10px; border-radius: 4px;")
        btn_confirm.clicked.connect(self.accept)
        layout.addWidget(btn_confirm)

        self.refresh_table()

    def refresh_table(self):
        self.table.setRowCount(len(self.segments))
        for i, seg in enumerate(self.segments):
            self.table.setItem(i, 0, QTableWidgetItem(format_page_ranges(seg['pages'])))
            self.table.setItem(i, 1, QTableWidgetItem(f"{int(seg['w_mm'])} x {int(seg['h_mm'])}"))
            status = "✅ 已设置" if seg.get('pos_set') else "❌ 未设置"
            self.table.setItem(i, 2, QTableWidgetItem(status))
            btn = QPushButton("设置位置")
            btn.clicked.connect(lambda checked, idx=i: self.set_position(idx))
            self.table.setCellWidget(i, 3, btn)

    def extract_pages(self):
        row = self.table.currentRow()
        if row < 0: return QMessageBox.warning(self, "提示", "请先在表格中点击选中一个组！")
        text, ok = QInputDialog.getText(self, "提取页面", "请输入要独立出来的页码 (例如封面是第一页输入: 1):")
        if ok and text:
            try:
                pages_to_extract = parse_page_ranges(text)
                current_pages = set(self.segments[row]['pages'])
                extract_set = set(pages_to_extract)

                if not extract_set.issubset(current_pages): return QMessageBox.warning(self, "错误",
                                                                                       "提取的页码不在当前选中的组内！")
                if extract_set == current_pages: return QMessageBox.warning(self, "错误",
                                                                            "不能提取全部页码，请直接对该组设置！")

                self.segments[row]['pages'] = sorted(list(current_pages - extract_set))
                new_group = {
                    'pages': sorted(list(extract_set)), 'w_mm': self.segments[row]['w_mm'],
                    'h_mm': self.segments[row]['h_mm'], 'pos_set': False
                }

                # 初始化新组的数据结构 (兼容 OCR和出图章)
                if 'pos_pct' in self.segments[row]:
                    if isinstance(self.segments[row]['pos_pct'], dict):
                        # OCR 模式
                        new_group['pos_pct'] = {'no': (0.8, 0.9, 150, 40), 'name': (0.8, 0.85, 150, 40)}
                    else:
                        # 图章模式
                        new_group['pos_pct'] = (0.5, 0.5)

                self.segments.append(new_group)
                self.refresh_table()
                QMessageBox.information(self, "成功", "提取成功，已在最下方生成新组！您可以为它单独设置位置。")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"输入格式错误: {str(e)}")

    def delete_group(self):
        row = self.table.currentRow()
        if row >= 0:
            del self.segments[row]
            self.refresh_table()

    def set_position(self, idx):
        self.hide()
        self.parent_app.enter_setting_mode_from_dialog(self.segments[idx], idx, self)