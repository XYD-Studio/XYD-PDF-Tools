import sys
import os
import site
import ctypes
import webbrowser
import multiprocessing
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QHBoxLayout,
                             QListWidget, QStackedWidget)
from PyQt5.QtGui import QIcon

from modules.mod_stamper import StamperWidget
from modules.mod_ocr import OCRExtractorWidget
from modules.mod_compressor import CompressorWidget
from modules.mod_toolkit import ToolkitWidget
from modules.mod_cropper import CropperWidget
from modules.mod_help import HelpWidget            # <--- 导入独立的帮助模块
from modules.mod_img2dxf import ImgToDxfWidget     # <--- 导入即将新增的 DXF 模块
from core.utils import get_base_path               # <--- 统一路径管理
from modules.mod_img_inserter import ImgInserterWidget

try:
    site_packages = site.getsitepackages()[0]
    qt_plugin_path = os.path.join(site_packages, 'PyQt5', 'Qt5', 'plugins')
    if os.path.exists(qt_plugin_path):
        os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = qt_plugin_path
except Exception:
    pass

try:
    myappid = 'mycompany.pdf_master_pro.v6.0'
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except Exception:
    pass

from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QHBoxLayout,
                             QListWidget, QStackedWidget, QTextBrowser)
from PyQt5.QtGui import QIcon

from modules.mod_stamper import StamperWidget
from modules.mod_ocr import OCRExtractorWidget
from modules.mod_compressor import CompressorWidget
from modules.mod_toolkit import ToolkitWidget
from modules.mod_cropper import CropperWidget
from core.utils import get_base_path





# ================= 全局现代扁平化 QSS 样式表 =================
GLOBAL_QSS = """
    /* 全局字体和颜色 */
    QWidget {
        font-family: "Microsoft YaHei", "Segoe UI", sans-serif;
        color: #2F3542;
        background-color: #F1F2F6;
    }

    /* 普通按钮基础样式 */
    QPushButton {
        background-color: #F1F2F6;
        color: #2F3542;
        border: 1px solid #CED6E0;
        border-radius: 4px;
        padding: 6px 12px;
        font-size: 13px;
    }
    QPushButton:hover { background-color: #DFE4EA; }
    QPushButton:pressed { background-color: #CED6E0; }

    /* 功能型按钮蓝色高亮 */
    QPushButton#ActionBtn {
        background-color: #3498DB;
        color: white;
        font-weight: bold;
        border: none;
    }
    QPushButton#ActionBtn:hover { background-color: #2980B9; }

    /* 导出/确认按钮绿色高亮 */
    QPushButton#ExportBtn {
        background-color: #2ECC71;
        color: white;
        font-weight: bold;
        border: none;
        padding: 10px;
    }
    QPushButton#ExportBtn:hover { background-color: #27AE60; }

    /* 警告/特殊按钮橙紫高亮 */
    QPushButton#SpecialBtn { background-color: #9B59B6; color: white; border: none; }
    QPushButton#SpecialBtn:hover { background-color: #8E44AD; }
    QPushButton#WarningBtn { background-color: #E67E22; color: white; font-weight: bold; padding: 10px; border: none; }
    QPushButton#WarningBtn:hover { background-color: #D35400; }

    /* 面板与分组框 */
    QGroupBox {
        background-color: #FFFFFF;
        border: 1px solid #DFE4EA;
        border-radius: 8px;
        margin-top: 15px;
        font-weight: bold;
        color: #2F3542;
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        left: 15px;
        padding: 0 5px;
        color: #34495E;
    }

    /* 输入框样式 */
    QLineEdit {
        border: 1px solid #CED6E0;
        border-radius: 4px;
        padding: 5px;
        background-color: #FFFFFF;
        selection-background-color: #3498DB;
    }
    QLineEdit:focus { border: 1px solid #3498DB; }

    /* 数据表格现代样式 */
    QTableWidget {
        background-color: #FFFFFF;
        border: 1px solid #DFE4EA;
        border-radius: 6px;
        gridline-color: #F1F2F6;
        selection-background-color: #D6EAF8;
        selection-color: #2F3542;
    }
    QHeaderView::section {
        background-color: #F8F9FA;
        color: #747D8C;
        padding: 6px;
        border: none;
        border-right: 1px solid #DFE4EA;
        border-bottom: 1px solid #DFE4EA;
        font-weight: bold;
    }

    /* 进度条 */
    QProgressBar {
        border: 1px solid #DFE4EA;
        border-radius: 4px;
        text-align: center;
        background-color: #FFFFFF;
        color: #2F3542;
    }
    QProgressBar::chunk {
        background-color: #2ECC71;
        border-radius: 3px;
    }

    /* QSplitter 分割线把手 */
    QSplitter::handle {
        background-color: #CED6E0;
        height: 2px;
    }

    /* QTextBrowser 样式 */
    QTextBrowser {
        background-color: #FFFFFF;
        border: none;
        padding: 20px;
        font-size: 14px;
        line-height: 1.6;
    }
"""


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF聚合工作站 V2.2")
        self.resize(1400, 900)

        icon_path = get_base_path('logo.ico')
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QHBoxLayout(main_widget)
        layout.setContentsMargins(0, 0, 0, 0)  # 去掉外边距，让侧边栏贴边
        layout.setSpacing(0)

        # --- 左侧深色导航菜单 ---
        self.menu_list = QListWidget()
        self.menu_list.setFixedWidth(220)
        self.menu_list.addItems([
            "🎯  批量图章工具",
            "📖  OCR智能提取目录",
            "✂️  超级裁剪拼接大师",
            "🗜️  PDF与图片压缩",
            "🛠️  PDF综合工具箱",
            "📐  线稿智能转DXF",
            "🖼️  批量一对一加图",
            "❓  操作指南与帮助"
        ])
        self.menu_list.currentRowChanged.connect(self.switch_tab)

        # 独立的极致深色侧边栏样式
        self.menu_list.setStyleSheet("""
            QListWidget { 
                background-color: #2C3E50; 
                color: #BDC3C7; 
                font-size: 15px; 
                border: none; 
            }
            QListWidget::item { 
                height: 55px; 
                padding-left: 15px; 
                border-bottom: 1px solid #34495E;
            }
            QListWidget::item:selected { 
                background-color: #34495E; 
                color: #FFFFFF; 
                border-left: 6px solid #3498DB; 
                font-weight: bold;
            }
            QListWidget::item:hover { 
                background-color: #3D566E; 
                color: #FFFFFF;
            }
        """)

        # --- 右侧堆叠区域 ---
        self.stack = QStackedWidget()
        self.stack.setStyleSheet("background-color: #F1F2F6;")

        self.page_stamp = StamperWidget()
        self.page_ocr = OCRExtractorWidget()
        self.page_crop = CropperWidget()
        self.page_comp = CompressorWidget()
        self.page_tool = ToolkitWidget()
        self.page_dxf = ImgToDxfWidget()   # <--- 实例化 DXF 模块
        self.page_inserter = ImgInserterWidget()
        self.page_help = HelpWidget()

        self.stack.addWidget(self.page_stamp)
        self.stack.addWidget(self.page_ocr)
        self.stack.addWidget(self.page_crop)
        self.stack.addWidget(self.page_comp)
        self.stack.addWidget(self.page_tool)
        self.stack.addWidget(self.page_dxf)  # <--- 加入 Stack
        self.stack.addWidget(self.page_inserter)
        self.stack.addWidget(self.page_help)

        layout.addWidget(self.menu_list)
        layout.addWidget(self.stack)

        self.menu_list.setCurrentRow(0)  # 改为0，默认选中第一个模块

    def switch_tab(self, index):
        self.stack.setCurrentIndex(index)

    def handle_link_click(self, url):
        """处理帮助文档中的链接点击事件"""
        webbrowser.open(url.toString())



if __name__ == "__main__":
    import multiprocessing

    multiprocessing.freeze_support()

    app = QApplication(sys.argv)

    # ================= 自定义动态闪屏类 =================
    from PyQt5.QtWidgets import QSplashScreen
    from PyQt5.QtGui import QPixmap, QPainter, QColor, QFont, QLinearGradient, QPen, QBrush
    from PyQt5.QtCore import Qt, QTimer


    class DynamicSplashScreen(QSplashScreen):
        """支持动态流光效果的闪屏"""

        def __init__(self):
            self.width = 1000
            self.height = 380
            self.progress = 0
            self.flow_offset = 0  # 流光偏移

            # 创建初始画布
            pixmap = QPixmap(self.width, self.height)
            pixmap.fill(Qt.transparent)
            super().__init__(pixmap, Qt.WindowStaysOnTopHint)

            # 启动动画定时器
            self.timer = QTimer()
            self.timer.timeout.connect(self._animate)
            self.timer.start(30)  # 33fps

            # 立即绘制
            self._draw()

        def _draw(self):
            """绘制闪屏"""
            pixmap = QPixmap(self.width, self.height)
            pixmap.fill(Qt.transparent)
            painter = QPainter(pixmap)
            painter.setRenderHint(QPainter.Antialiasing)
            painter.setRenderHint(QPainter.TextAntialiasing)

            # 渐变背景
            gradient = QLinearGradient(0, 0, self.width, self.height)
            gradient.setColorAt(0, QColor(0, 8, 35))
            gradient.setColorAt(1, QColor(0, 5, 28))
            painter.fillRect(0, 0, self.width, self.height, gradient)

            # 光晕效果
            painter.setBrush(QBrush(QColor(50, 100, 180, 30)))
            painter.setPen(Qt.NoPen)
            painter.drawEllipse(-50, -50, 150, 150)
            painter.drawEllipse(self.width - 100, self.height - 100, 150, 150)

            # 发光边框
            painter.setPen(QPen(QColor(100, 150, 255, 80), 2))
            painter.drawRoundedRect(2, 2, self.width - 4, self.height - 4, 12, 12)

            # 主标题（阴影）
            painter.setPen(QColor(30, 40, 60, 100))
            painter.setFont(QFont("Microsoft YaHei", 28, QFont.Bold))
            painter.drawText(2, -28, self.width, self.height, Qt.AlignCenter, "PDF 智能聚合工作站")

            # 主标题（渐变）
            title_grad = QLinearGradient(self.width // 2 - 150, 0, self.width // 2 + 150, 0)
            title_grad.setColorAt(0, QColor(200, 220, 255))
            title_grad.setColorAt(0.5, QColor(255, 255, 255))
            title_grad.setColorAt(1, QColor(150, 180, 255))
            painter.setPen(QPen(QBrush(title_grad), 1))
            painter.drawText(0, -30, self.width, self.height, Qt.AlignCenter, "PDF 智能聚合工作站")

            # 副标题
            painter.setPen(QColor(80, 110, 150))
            painter.setFont(QFont("Microsoft YaHei", 12))
            painter.drawText(0, 25, self.width, self.height, Qt.AlignCenter, "V 2.2 Professional Edition")
            painter.setPen(QColor(100, 120, 150))
            painter.setFont(QFont("Microsoft YaHei", 11))
            painter.drawText(0, 55, self.width, self.height, Qt.AlignCenter, "玄宇绘世设计工作室出品")

            # 装饰线
            center_x = self.width // 2
            painter.setPen(QPen(QColor(80, 120, 200, 100), 1))
            painter.drawLine(center_x - 120, self.height - 75, center_x + 120, self.height - 75)

            # ========== 动态进度条 ==========
            bar_width = 380
            bar_height = 3
            bar_x = (self.width - bar_width) // 2
            bar_y = self.height - 55

            # 背景
            painter.setBrush(QBrush(QColor(40, 50, 70, 100)))
            painter.setPen(Qt.NoPen)
            painter.drawRoundedRect(bar_x, bar_y, bar_width, bar_height, 2, 2)

            # 进度（如果 progress > 0）
            if self.progress > 0:
                progress_width = int(bar_width * self.progress / 100)
                if progress_width > 0:
                    # 进度渐变
                    prog_grad = QLinearGradient(bar_x, bar_y, bar_x + progress_width, bar_y)
                    prog_grad.setColorAt(0, QColor(80, 150, 255))
                    prog_grad.setColorAt(1, QColor(50, 100, 255))
                    painter.setBrush(QBrush(prog_grad))
                    painter.drawRoundedRect(bar_x, bar_y, progress_width, bar_height, 2, 2)

                    # 流光效果（在进度条上移动的光斑）
                    if progress_width > 20:
                        flow_x = bar_x + (self.flow_offset % progress_width)
                        painter.setBrush(QBrush(QColor(255, 255, 255, 120)))
                        painter.drawRoundedRect(flow_x - 10, bar_y - 1, 20, bar_height + 2, 3, 3)

            # 版本号
            painter.setPen(QColor(70, 90, 120, 150))
            painter.setFont(QFont("Microsoft YaHei", 8))
            painter.drawText(self.width - 95, self.height - 15, "Build 2026.04")

            painter.end()
            self.setPixmap(pixmap)

        def _animate(self):
            """动画更新"""
            self.flow_offset = (self.flow_offset + 3) % 200
            self._draw()

        def update_progress(self, value):
            """更新进度"""
            self.progress = min(100, max(0, value))
            self._draw()
            QApplication.processEvents()

        def set_message(self, msg):
            """设置消息"""
            self.showMessage(msg, Qt.AlignBottom | Qt.AlignCenter, QColor("#BDC3C7"))
            QApplication.processEvents()


    # ================= 使用动态闪屏 =================
    splash = DynamicSplashScreen()
    splash.show()
    app.processEvents()

    # 更新进度和消息
    splash.set_message("正在加载底层 PDF 渲染引擎与 OCR 模块...")
    splash.update_progress(30)
    app.processEvents()

    # 设置图标与全局样式
    icon_path = get_base_path('logo.ico')
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))
    app.setStyleSheet(GLOBAL_QSS)

    splash.set_message("正在初始化图形界面...")
    splash.update_progress(70)
    app.processEvents()

    # 初始化主窗口
    window = MainWindow()

    # 完成
    splash.update_progress(100)
    splash.set_message("启动完成！")
    app.processEvents()

    # 平滑切换
    splash.finish(window)
    window.show()

    sys.exit(app.exec_())


