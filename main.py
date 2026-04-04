import sys
import os
import site
import ctypes
import webbrowser
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
        self.setWindowTitle("PDF聚合工作站 V2.1")
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
        self.page_help = HelpWidget()

        # # 使用 QTextBrowser 替代 QLabel
        # self.page_help = QTextBrowser()
        # self.page_help.setOpenExternalLinks(False)  # 不自动打开外部链接
        # self.page_help.setOpenLinks(False)  # 不自动打开链接
        # self.page_help.anchorClicked.connect(self.handle_link_click)  # 链接点击信号
        # self.page_help.setHtml(self.generate_help_text())
        # self.page_help.setStyleSheet("padding: 20px; background-color: #FFFFFF;")

        self.stack.addWidget(self.page_stamp)
        self.stack.addWidget(self.page_ocr)
        self.stack.addWidget(self.page_crop)
        self.stack.addWidget(self.page_comp)
        self.stack.addWidget(self.page_tool)
        self.stack.addWidget(self.page_dxf)  # <--- 加入 Stack
        self.stack.addWidget(self.page_help)

        layout.addWidget(self.menu_list)
        layout.addWidget(self.stack)

        self.menu_list.setCurrentRow(0)  # 改为0，默认选中第一个模块

    def switch_tab(self, index):
        self.stack.setCurrentIndex(index)

    def handle_link_click(self, url):
        """处理帮助文档中的链接点击事件"""
        webbrowser.open(url.toString())

    # def generate_help_text(self):
    #     return """
    #     <html>
    #     <head>
    #     <style>
    #         h2 {{ color: #2C3E50; border-bottom: 2px solid #3498DB; padding-bottom: 5px; margin-top: 20px; }}
    #         h3 {{ color: #E67E22; margin-bottom: 5px; }}
    #         p {{ font-size: 14px; color: #34495E; line-height: 1.6; margin-top: 5px; }}
    #         ul {{ font-size: 14px; color: #34495E; line-height: 1.6; margin-top: 0; }}
    #         li {{ margin-bottom: 6px; }}
    #         .highlight {{ color: #E74C3C; font-weight: bold; }}
    #         .tip {{ background-color: #D6EAF8; padding: 8px; border-left: 4px solid #3498DB; border-radius: 4px; font-size: 13px; margin: 10px 0; }}
    #         .warning {{ background-color: #FCF3CF; padding: 8px; border-left: 4px solid #F1C40F; border-radius: 4px; font-size: 13px; margin: 10px 0; }}
    #         .legal {{ background-color: #F9EBEA; padding: 8px; border-left: 4px solid #E74C3C; border-radius: 4px; font-size: 13px; margin: 10px 0; }}
    #         .header {{ display: flex; align-items: center; gap: 15px; margin-bottom: 20px; }}
    #         .logo-link {{
    #             display: inline-block;
    #             transition: opacity 0.3s;
    #             cursor: pointer;
    #             text-decoration: none;
    #         }}
    #         .logo-link:hover {{ opacity: 0.8; }}
    #         .studio-link {{
    #             color: #2C3E50;
    #             text-decoration: none;
    #             border-bottom: 1px dashed #3498DB;
    #             cursor: pointer;
    #         }}
    #         .studio-link:hover {{
    #             color: #3498DB;
    #         }}
    #         .studio-name {{ color: #2C3E50; font-size: 24px; font-weight: bold; }}
    #         .legal-title {{
    #             color: #C0392B;
    #             font-size: 20px;
    #             font-weight: bold;
    #             margin-top: 40px;
    #             border-bottom: 2px solid #E74C3C;
    #             padding-bottom: 5px;
    #         }}
    #         .legal-section {{
    #             font-weight: bold;
    #             color: #C0392B;
    #             margin-top: 15px;
    #         }}
    #     </style>
    #     </head>
    #     <body>
    #         <div class="header">
    #             <span class="studio-name">📚 PDF 聚合工作站 V2.0 - 操作指南</span>
    #         </div>
    #         <p>欢迎使用本系统！本软件专为工程图纸、批量文档的高效处理而设计，请仔细阅读以下各模块的使用说明，掌握其中的高级特性。</p>
    #
    #         <h2>🎯 模块 1：批量出图章工具</h2>
    #         <p>专为不同尺寸图纸混合的 PDF 文件设计，支持多图章并行加盖与精准定位，具备强大的防篡改锁定功能。</p>
    #         <ul>
    #             <li><b>添加图章：</b>点击【增加印章/签名】导入图片（支持 PNG、JPG、JPEG）。系统会<b>自动计算并锁定原始宽高比</b>，您只需输入宽度，高度会自动换算。可设置图章名称、尺寸和旋转角度。</li>
    #             <li><b>大纲与分段：</b>导入 PDF 后，必须先点击【生成合并预览】（建议勾选“合并前用GS扁平化”以破除原文件的密码锁和可编辑批注），然后点击【① 智能按图纸尺寸分段组合】。系统会按图纸的物理尺寸（如 A3、A4）将页面分组，您只需为每种尺寸设置一次印章位置即可。</li>
    #             <li><b>终极自由微调：</b>点击【② 进入终极自由微调预览】，可在下方大图中对任意一页的图章进行独立调整。
    #                 <div class="tip"><b>💡 隐藏快捷键：</b>在微调视图中，选中图章<b>右键</b>可解锁比例或自定义旋转角度；按住键盘 <b>Alt 键</b> 并拖动图章，可直接快速复制一个新图章！</div>
    #             </li>
    #             <li><b>清除原批注：</b>点击【🧹 快速清除原PDF附带的图章批注】可一键删除文档中原有的所有表单控件和悬浮批注。注意：本软件盖的图章采用底层像素写入技术，已完全融入页面，无法被删除或修改。</li>
    #             <li><b>配置保存与复用：</b>排版完成后，可点击【💾 导出当前配置】保存为 JSON 文件。下次处理<b>同尺寸图纸</b>时，点击【📂 导入历史配置】，系统会智能识别页面尺寸，自动将印章匹配到之前的绝对方位。</li>
    #             <li><b>多重导出模式：</b>支持三种导出方式：
    #                 <ul>
    #                     <li><b>合并为单一文件：</b>所有盖章后的页面合并输出为一个 PDF。</li>
    #                     <li><b>批量按原文件拆分：</b>根据原始 PDF 文件的大纲书签，按原文件名拆分输出。</li>
    #                     <li><b>拆分为单页独立文件：</b>将每一页独立输出，可自定义前缀/后缀命名。</li>
    #                 </ul>
    #             </li>
    #             <li><b>画质与压缩控制：</b>可单独设置图章渲染 DPI（72/150/300/600），并支持 Ghostscript 二次全局压缩（/screen、/ebook、/printer 三级品质）。</li>
    #             <li><b>🛡️ 防篡改数字签名锁：</b>勾选【附加防篡改保护锁】后，系统会：
    #                 <ul>
    #                     <li>使用您提供的 .pfx 数字证书，在指定图章位置创建合法签名域。</li>
    #                     <li>对文档进行数字签名，锁定签名区域，防止他人使用任何 PDF 编辑器（包括 Adobe Acrobat）删除、替换或移动已加盖的图章。</li>
    #                     <li>签名后文档会显示“已签名且所有签名有效”，签名区域不可编辑，极大提升了文档的防伪性和法律效力。</li>
    #                 </ul>
    #                 <div class="tip"><b>💡 使用提示：</b>您可以在多图章列表中指定将签名锁定到哪一个具体图章位置（默认锁定到遇到的第一个图章处）。签名后的文档可以通过 Adobe Reader 验证签名有效性。</div>
    #             </li>
    #         </ul>
    #
    #         <h2>📖 模块 2：OCR 智能提取目录 (主打功能)</h2>
    #         <p>基于先进的 PaddleOCR 引擎，纯本地离线运行，安全且极速。</p>
    #         <ul>
    #             <li><b>动态字段管理：</b>不仅限于图纸名称和图号，您可以通过【➕ 添加目标字段】无限增加需要提取的信息（如版本号、设计人等）。</li>
    #             <li><b>设置提取框：</b>同样遵循“生成预览 -> 智能分段 -> 终极微调”的流程。您可以在屏幕上直观地拖拽彩色选框，<b>指哪打哪，绝对精准</b>。</li>
    #             <li><b>智能配置继承：</b>您可以将设好的提取框导出为 JSON。下次导入时，系统会<b>智能识别 PDF 的物理宽高</b>，自动将框对齐到之前的绝对方位，未匹配的新尺寸会提示您手动设置。</li>
    #             <li><b>数据处理与导出：</b>提取完毕后，可在右侧表格中进行修改、上下移、图号排序。支持导出为 Excel/JSON，也可<b>直接将表格第一二列作为书签写入 PDF</b>，或直接<b>按提取出来的名称将文档拆分为单页 PDF</b>！</li>
    #         </ul>
    #
    #         <h2>✂️ 模块 3：超级裁剪拼接大师</h2>
    #         <ul>
    #             <li>1. 支持混装拖入 JPG图片 和 PDF文件。</li>
    #             <li>2. 极度灵活的切线交互：上方点击加横线/竖线，左键按住红线拖拽移动，<b>右键点击红线删除</b>。</li>
    #             <li><b>点击任意被切割的格子，使其变暗即代表丢弃该区域！</b>导出时将完美过滤掉多余的白边或不要的印章。</li>
    #         </ul>
    #
    #         <h2>🗜️ 模块 4：PDF 与图片强力压缩</h2>
    #         <p>提供专业的媒体文件体积优化方案。</p>
    #         <ul>
    #             <li><b>PDF 强力压缩：</b>基于底层 Ghostscript 引擎。支持三种级别：Screen（极致压缩适合网络传输）、Ebook（平衡模式）、Printer（高清打印模式）。<span class="highlight">注意：必须在电脑上正确安装 Ghostscript 才能使用此功能。本软件内置便携版 GS，通常无需额外安装。</span></li>
    #             <li><b>图片批量转换压缩：</b>支持拖拽 JPG/PNG/WEBP 等图片，可选择按 DPI 缩放或按像素长边缩放，支持批量转换为 PDF 或其他图片格式。</li>
    #         </ul>
    #
    #         <h2>🛠️ 模块 5：PDF 综合工具箱</h2>
    #         <p>集成了日常高频的文档处理小工具。</p>
    #         <ul>
    #             <li><b>多图转 PDF / 多个 PDF 合并：</b>按文件列表顺序进行极速拼接。</li>
    #             <li><b>PDF 拆分为单页：</b>将文档逐页拆解。</li>
    #             <li><b>按书签拆分 PDF (高级)：</b><span class="highlight">超强功能！</span>选择此项后点击【⚙️ 配置拆分规则】，您可以选择“按书签”或“序号+书签”命名，更支持<b>按页码分段追加不同的前缀和后缀</b>（例如第1-9页加前缀A，10-20页加后缀B）。</li>
    #             <li><b>PDF 转图片型 PDF：</b>专治各种字体缺失、乱码、不可打印的“顽固 PDF”。它会将所有页面以设定 DPI 栅格化为图片，再重新打包为绝对兼容的只读 PDF。</li>
    #             <li><b>PDF 批量导出图片：</b>将 PDF 每一页导出为高清 JPG 或 PNG，支持开启透明通道（仅PNG）。</li>
    #         </ul>
    #
    #         <div class="legal-title">⚖️ 最终用户许可协议与免责声明</div>
    #         <p>欢迎使用"PDF聚合工作站"（以下简称"本软件"）。在安装、复制或使用本软件之前，请您务必仔细阅读并透彻理解本免责声明与许可协议。当您开始使用本软件，即表示您已阅读、理解并自愿接受本协议的所有条款。如果您不接受本协议的任何条款，请立即停止使用并销毁本软件的所有副本。</p>
    #
    #         <div class="legal-section">📋 第一条：合法使用原则</div>
    #         <ol style="color: #34495E; line-height: 1.6;">
    #             <li>本软件仅作为一款中立的 PDF 图像合并与排版工具，旨在提高用户处理电子文档的办公效率。</li>
    #             <li><span class="highlight">用户承诺：</span>在使用本软件添加任何图章、签名、标识或水印时，必须确保已获得相关图章或标识的【绝对且合法的拥有权或使用授权】。</li>
    #             <li><span class="highlight">严禁将本软件用于任何非法或违规用途</span>，包括但不限于：
    #                 <ul style="margin-top: 5px;">
    #                     <li>伪造、变造国家机关、企事业单位、人民团体、个人的公章、印章或签名；</li>
    #                     <li>制造虚假合同、伪造财务报表、资质证书等以进行诈骗或欺瞒；</li>
    #                     <li>侵犯他人的知识产权、名誉权、隐私权或其他合法权益；</li>
    #                     <li>其他任何违反国家法律法规及破坏社会公共利益的行为。</li>
    #                 </ul>
    #             </li>
    #         </ol>
    #
    #         <div class="legal-section">⚠️ 第二条：免责声明（核心条款）</div>
    #         <ol style="color: #34495E; line-height: 1.6;">
    #             <li>本软件按"现状"提供，开发者不对软件的绝对完美性、适用性或无技术缺陷做任何明示或暗示的保证。</li>
    #             <li><span class="highlight">开发者作为工具的提供方，不参与、不干预用户的任何使用过程。</span>用户通过本软件生成的任何带有图章的电子文档，其合法性、真实性和有效性均由【用户本人】全权负责。</li>
    #             <li><span class="highlight">因用户违规、违法使用本软件（如伪造印章等）所引发的一切法律纠纷、行政处罚、刑事责任及经济赔偿，均由【用户自行承担全部后果】</span>，本软件及其开发者、分发者概不承担任何直接、间接、连带的法律责任。</li>
    #             <li>开发者对用户因使用或无法使用本软件而造成的任何数据丢失、利润损失、业务中断或其他商业损害不承担任何赔偿责任。在使用本软件处理重要文档前，请务必自行做好原文件的备份工作。</li>
    #         </ol>
    #
    #         <div class="legal-section">©️ 第三条：知识产权</div>
    #         <p style="color: #34495E; line-height: 1.6;">本软件的架构、代码、界面设计及相关文档的知识产权归原开发者所有。未经明确授权，任何人不得对本软件进行反向工程、反编译、破解或用于二次商业倒卖。</p>
    #
    #         <div class="legal-section">🔚 第四条：协议终止</div>
    #         <p style="color: #34495E; line-height: 1.6;">如果用户违反了本协议的任何条款（尤其是第一条合法使用原则），本授权将自动立即终止。用户必须立即停止使用并销毁本软件的所有副本。开发者保留向违约方追究法律责任的权利。</p>
    #
    #         <div class="legal-section">📢 开源声明</div>
    #         <p style="color: #34495E; line-height: 1.6;">本软件为免费开源软件，任何人可在 GitHub 网站上进行下载使用，同时也欢迎同仁对该软件进行扩展升级。</p>
    #
    #         <div style="text-align: center; margin-top: 30px; padding-top: 20px; border-top: 1px solid #3498DB;">
    #             <a href="https://www.xy-d.top/" class="studio-link" style="text-decoration: none;">
    #                 <b>玄宇绘世设计工作室出品</b>
    #             </a>
    #             <br>
    #             <a href="https://www.xy-d.top/" class="logo-link">
    #                 <img src="file:///{0}" alt="玄宇绘世设计工作室" height="35" style="margin-top: 10px;">
    #             </a>
    #         </div>
    #     </body>
    #     </html>
    #     """.format(get_base_path('public/logo.png').replace('\\', '/'))

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
            painter.drawText(2, -28, self.width, self.height, Qt.AlignCenter, "PDF 聚合工作站")

            # 主标题（渐变）
            title_grad = QLinearGradient(self.width // 2 - 150, 0, self.width // 2 + 150, 0)
            title_grad.setColorAt(0, QColor(200, 220, 255))
            title_grad.setColorAt(0.5, QColor(255, 255, 255))
            title_grad.setColorAt(1, QColor(150, 180, 255))
            painter.setPen(QPen(QBrush(title_grad), 1))
            painter.drawText(0, -30, self.width, self.height, Qt.AlignCenter, "PDF 聚合工作站")

            # 副标题
            painter.setPen(QColor(80, 110, 150))
            painter.setFont(QFont("Microsoft YaHei", 12))
            painter.drawText(0, 25, self.width, self.height, Qt.AlignCenter, "V 2.1 Professional Edition")
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
            painter.drawText(self.width - 95, self.height - 15, "Build 2026.03")

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


