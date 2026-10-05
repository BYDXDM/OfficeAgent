; OfficeAgent Inno Setup installer script (design doc section 8.3, dual SKU).
; Build:  ISCC /DSKU=complete officeagent.iss   -> full offline (payload bundled)
;         ISCC /DSKU=lite officeagent.iss       -> lite online (boot fetches missing parts)
; Inno Setup 6.x; Win7 SP1+; PrivilegesRequired=lowest (per-user, no admin).
; NOTE: exe files are unsigned; sign with Authenticode before release (AV whitelist).
; File must be saved as UTF-8 with BOM (Inno requirement for non-ASCII).

#define AppName "OfficeAgent"
#define AppVersion "0.9.1"
#define AppVersionNum "0.9.1.0"
#define AppPublisher "OfficeAgent Project"

[Setup]
AppId={{8E1F4B7A-52C0-4A5E-9A2B-OFFICEAGENTM1}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={localappdata}\{#AppName}
PrivilegesRequired=lowest
DisableProgramGroupPage=yes
OutputDir=..\dist
OutputBaseFilename=OfficeAgent-{#AppVersion}-{#SKU}-setup
Compression=lzma2/max
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=6.1sp1
WizardStyle=modern
; 多语言：中文优先，按系统 UI 语言预选（中文系统直接中文向导），仍可手动切英文
ShowLanguageDialog=auto
SetupIconFile=..\app.ico
UninstallDisplayIcon={app}\host\OfficeAgent.exe
; Setup 程序 exe 的属性页元数据（右键 → 属性 → 详细信息）
VersionInfoCompany={#AppPublisher}
VersionInfoDescription={#AppName} Setup
VersionInfoProductName={#AppName}
VersionInfoProductVersion={#AppVersion}
VersionInfoVersion={#AppVersionNum}
VersionInfoCopyright=Copyright (C) 2026 {#AppPublisher}

[Languages]
; 中文放首位 = 默认语言。MessagesFile 相对本 .iss 所在目录（installer\）。
Name: "chinesesimplified"; MessagesFile: "lang\ChineseSimplified.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[CustomMessages]
; 任务/运行项描述需本地化：Inno 不会自动翻译 [Tasks]/[Run] 里的字符串，用 {cm:} 引用。
chinesesimplified.CreateDesktopIcon=创建桌面快捷方式(&D)
english.CreateDesktopIcon=Create a &desktop shortcut
chinesesimplified.LaunchApp=立即启动 OfficeAgent（首次运行自检）
english.LaunchApp=Launch OfficeAgent (first-run check)

[Files]
; host programs (both SKUs)
Source: "..\build\boot\OfficeAgentBoot.exe"; DestDir: "{app}\boot"; Flags: ignoreversion
Source: "..\build\boot\OfficeAgentBootAdmin.exe"; DestDir: "{app}\boot"; Flags: ignoreversion
Source: "..\build\host\OfficeAgent.exe"; DestDir: "{app}\host"; Flags: ignoreversion
Source: "..\build\host\OfficeAgent35.exe"; DestDir: "{app}\host"; Flags: ignoreversion
Source: "..\build\host\pdfium.dll"; DestDir: "{app}\host"; Flags: ignoreversion
Source: "..\store\components.ini"; DestDir: "{app}\store"; Flags: ignoreversion
Source: "..\store\wheels-requirements.txt"; DestDir: "{app}\store"; Flags: ignoreversion
Source: "..\skills\*"; DestDir: "{app}\skills"; Flags: ignoreversion recursesubdirs createallsubdirs
; 内置资源模板（make-ppt 的 PPT 模板 / excel-formulas.json 公式大全）。
; 0.7.x 漏打包 templates 导致全新安装的 create_presentation 找不到模板——两个 SKU 都必须带。
Source: "..\templates\*"; DestDir: "{app}\templates"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\tests\fixtures\recon\flow.csv"; DestDir: "{app}\tests"; Flags: ignoreversion
Source: "..\tests\fixtures\recon\ledger.csv"; DestDir: "{app}\tests"; Flags: ignoreversion
#if SKU == "complete"
; full offline SKU: payload bundled (~900MB incl. LibreOffice 7.6.7)
; 不设 skipifsourcedoesntexist —— complete SKU 的意义就是"离线齐套"，
; 缺件必须让 ISCC 直接编译失败，而不是静默产出缺 LO/wheels 的假完整包。
; 打包前先跑 tools\verify-payload.ps1 校验 SHA256。
Source: "..\payload\kb\*.msu"; DestDir: "{app}\payload\kb"; Flags: ignoreversion
Source: "..\payload\ndp48\*.exe"; DestDir: "{app}\payload\ndp48"; Flags: ignoreversion
Source: "..\payload\vcrt\*.exe"; DestDir: "{app}\payload\vcrt"; Flags: ignoreversion
Source: "..\payload\py38\*.zip"; DestDir: "{app}\payload\py38"; Flags: ignoreversion
Source: "..\payload\py38\get-pip.py"; DestDir: "{app}\payload\py38"; Flags: ignoreversion
Source: "..\payload\lo76\*.zip"; DestDir: "{app}\payload\lo76"; Flags: ignoreversion
Source: "..\payload\README.md"; DestDir: "{app}\payload"; Flags: ignoreversion
Source: "..\store\wheels\*.whl"; DestDir: "{app}\store\wheels"; Flags: ignoreversion
#endif
#if SKU == "lite"
; lite SKU: 不带 payload，由 boot 在线补全；缺件可容忍
Source: "..\payload\README.md"; DestDir: "{app}\payload"; Flags: ignoreversion skipifsourcedoesntexist
#endif

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\boot\OfficeAgentBoot.exe"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\boot\OfficeAgentBoot.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"

[Run]
Filename: "{app}\boot\OfficeAgentBoot.exe"; Description: "{cm:LaunchApp}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; runtime trees unpacked by the bootstrap + conversion outputs
Type: filesandordirs; Name: "{app}\runtime"
Type: filesandordirs; Name: "{app}\lo76"
Type: filesandordirs; Name: "{app}\output"
