[Setup]
; ---------------- 基本信息设置 ----------------
AppName=PDF 智能聚合工作站
AppVersion=2.2 (Pro Max)
AppPublisher=玄宇绘世设计工作室
AppPublisherURL=https://www.xy-d.top
AppSupportURL=https://www.xy-d.top
AppUpdatesURL=https://github.com/XYD-Studio/XYD-PDF-Tools

; ---------------- 安装路径与名称设置 ----------------
; 默认安装到 C:\Program Files\PDF 智能聚合工作站
DefaultDirName={autopf}\PDF 智能聚合工作站
DefaultGroupName=PDF 智能聚合工作站

; 编译出的 Setup.exe 保存位置 (这里默认放在您的 D 盘根目录，您可随意更改)
OutputDir=D:\
OutputBaseFilename=PDF_Tools_Pro_V2.2_Setup

; ---------------- 核心功能设置 ----------------
; 【关键：强制展示刚才写的免责协议】
; 指向您刚才保存的 License.txt 文件的绝对路径
LicenseFile=F:\code_available\XYD_PDF_Tools_Pro_V2.1\License.txt

; 指定安装包的图标 (可选，指向您的 logo.ico)
SetupIconFile=F:\code_available\XYD_PDF_Tools_Pro_V2.1\logo.ico

; 开启中文支持弹窗
ShowLanguageDialog=yes
Compression=lzma2/ultra64
SolidCompression=yes
; 允许在安装完成后启动软件
PrivilegesRequired=admin

[Languages]
; 语言包设置 (如果您的 Inno 没有中文包，可以删掉这一段，默认英文安装不影响使用)
Name: "chinesesimplified"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"

[Tasks]
; 允许用户勾选创建桌面快捷方式
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; ---------------- 核心文件打包指令 ----------------
; 指向您 PyInstaller 打包出来的 dist 文件夹！
; 注意末尾的 \* 代表包含该文件夹下的所有内容
; Flags 里的 recursesubdirs 代表连同里面的 gs_portable, paddleocr 模型等子文件夹一起打包
Source: "F:\code_available\XYD_PDF_Tools_Pro_V2.1\dist\PDF智能聚合工作站\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; 生成开始菜单和桌面的快捷方式
Name: "{group}\PDF 智能聚合工作站"; Filename: "{app}\PDF智能聚合工作站.exe"
Name: "{commondesktop}\PDF 智能聚合工作站"; Filename: "{app}\PDF智能聚合工作站.exe"; Tasks: desktopicon

[Run]
; ---------------- 静默安装 C++ 运行库 (强烈建议保留) ----------------
; 只要把 VC_redist.x64.exe 提前放进您的 dist/PDF智能聚合工作站 文件夹里即可。
; 这一步会在安装即将结束时，毫无弹窗、毫无感知地在后台为用户的 Win7/Win10 老系统补齐缺失的 DLL 环境！
Filename: "{app}\VC_redist.x64.exe"; Parameters: "/install /quiet /norestart"; StatusMsg: "正在配置底层系统运行环境 (绝不弹窗，请耐心等待)..."; Flags: waituntilterminated skipifdoesntexist

; 安装完成后，弹出的最后界面上有一个“运行 PDF 智能聚合工作站”的勾选框
Filename: "{app}\PDF智能聚合工作站.exe"; Description: "立即启动 PDF 智能聚合工作站"; Flags: nowait postinstall skipifsilent