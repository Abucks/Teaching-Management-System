; Inno Setup 脚本：生成 Windows 安装程序 Setup.exe
; 由 GitHub Actions 的 build-installers 工作流调用（需先安装 Inno Setup 6）
;
; 本地手动编译：
;   iscc /DAppVersion=1.4.0 /DSourceExe=dist\TeachingManager.exe packaging\windows\installer.iss

#ifndef AppVersion
  #define AppVersion "1.4.0"
#endif
#ifndef SourceExe
  #define SourceExe "..\..\dist\TeachingManager.exe"
#endif
#ifndef OutputDir
  #define OutputDir "..\..\dist\installers"
#endif
#ifndef OutputBase
  #define OutputBase "TeachingManager-windows-x64-setup"
#endif

[Setup]
AppId={{8F3C1B42-5D7A-4E19-9C3B-7A2E5F0D6C11}
AppName=教学管理系统
AppVersion={#AppVersion}
AppVerName=教学管理系统 {#AppVersion}
AppPublisher=Asell Bucks
AppPublisherURL=https://github.com/Abucks/Teaching-Management-System
AppSupportURL=https://github.com/Abucks/Teaching-Management-System/issues
DefaultDirName={autopf}\TeachingManager
DefaultGroupName=教学管理系统
DisableProgramGroupPage=yes
OutputDir={#OutputDir}
OutputBaseFilename={#OutputBase}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={app}\TeachingManager.exe
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
SetupLogging=yes
; 注：安装向导语言使用 Inno 默认语言（英文）；若需中文向导，可自行下载
; ChineseSimplified.isl 放入 Inno 的 Languages 目录后在此处添加 [Languages] 段

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "附加任务:"

[Files]
Source: "{#SourceExe}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\教学管理系统"; Filename: "{app}\TeachingManager.exe"
Name: "{group}\卸载 教学管理系统"; Filename: "{uninstallexe}"
Name: "{autodesktop}\教学管理系统"; Filename: "{app}\TeachingManager.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\TeachingManager.exe"; Description: "立即运行 教学管理系统"; Flags: nowait postinstall skipifsilent
