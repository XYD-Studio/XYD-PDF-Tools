#define MyAppName "PDF 聚合工作站 V4.0"
#define MyAppVersion "4.0.0"
#define MyAppExeName "XYD_PDF_Tools_Pro.exe"
#define ProjectRoot AddBackslash(SourcePath)

[Setup]
AppId={{D184509A-0E7F-47A7-89CA-A2499F6E4000}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher=玄宇绘世设计工作室
AppPublisherURL=https://www.xy-d.top
AppSupportURL=https://www.xy-d.top
AppUpdatesURL=https://github.com/XYD-Studio/XYD-PDF-Tools
DefaultDirName={autopf}\XYD_PDF_Tools_Pro_V4.0
DefaultGroupName=PDF 聚合工作站 V4.0
OutputDir={#ProjectRoot}Output\V4.0
OutputBaseFilename=XYD_PDF_Tools_Pro_V4.0_Setup
LicenseFile={#ProjectRoot}License.txt
SetupIconFile={#ProjectRoot}logo.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0
ShowLanguageDialog=yes
Compression=lzma2/ultra64
SolidCompression=yes
PrivilegesRequired=admin

[Languages]
Name: "chinesesimplified"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "{#ProjectRoot}Output\V4.0\Portable\XYD_PDF_Tools_Pro_V4.0\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\VC_redist.x64.exe"; Parameters: "/install /quiet /norestart"; StatusMsg: "正在配置系统运行环境..."; Flags: waituntilterminated skipifdoesntexist
Filename: "{app}\{#MyAppExeName}"; Description: "启动 {#MyAppName}"; Flags: nowait postinstall skipifsilent
