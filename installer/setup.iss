; QA测试工具盒 Inno Setup script
; Build command: ISCC.exe installer\setup.iss
; This installer is per-user by default and does not require administrator rights.

#define MyAppName "QA测试工具盒"
#ifndef MyAppVersion
  #define MyAppVersion "0.0.0"
#endif
#define MyAppPublisher "测试助手"
#define MyAppExeName "XiaoLinAssistant.exe"
#define MyUpdaterExeName "XiaoLinUpdater.exe"
#define MySetupName "XiaoLinAssistant_Setup_v" + MyAppVersion

[Setup]
AppId={{A2F59A97-8D6B-4B74-9E8D-8E1C814E2410}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\QA测试工具盒
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=no
OutputDir=..\Output
OutputBaseFilename={#MySetupName}
SetupIconFile=..\assets\app.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma
DiskSpanning=no
SolidCompression=no
WizardStyle=modern
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
CloseApplications=yes
RestartApplications=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "快捷方式："; Flags: unchecked

[Dirs]
Name: "{app}\assets"
Name: "{app}\tools\ios"
Name: "{app}\App_Logs"
Name: "{app}\IOS_Logs"
Name: "{app}\Screenshots"
Name: "{app}\Recordings"
Name: "{app}\Update_Logs"

[Files]
Source: "..\dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\dist\{#MyUpdaterExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\README.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\INSTALL_AND_FIRST_USE.txt"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "..\assets\*"; DestDir: "{app}\assets"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\tools\ios\README.txt"; DestDir: "{app}\tools\ios"; Flags: ignoreversion skipifsourcedoesntexist

Source: "..\tools\ios\*"; DestDir: "{app}\tools\ios"; Flags: ignoreversion recursesubdirs createallsubdirs
[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\assets\app.ico"
Name: "{group}\卸载 {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\assets\app.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "启动 {#MyAppName}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}\__pycache__"
