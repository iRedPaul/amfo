; Inno Setup Script für belegpilot
#define MyAppName "belegpilot"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "amsepa AG"
#define MyAppURL "https://belegpilot.io"
#define MyAppExeName "belegpilot.exe"
#define MyServiceExeName "belegpilot_service.exe"

[Setup]
AppId={{A7B3F4E2-9C8D-4F2A-B6E1-3E7F8A9E2C1B}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
OutputDir=installer
OutputBaseFilename=belegpilot_Setup_{#MyAppVersion}
SetupIconFile=gui/assets/icon.ico
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64compatible
ShowLanguageDialog=no
UninstallDisplayIcon={app}\{#MyAppExeName}
CloseApplications=yes
CloseApplicationsFilter=*.exe

[Languages]
Name: "german"; MessagesFile: "compiler:Languages\German.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
; Hauptprogramm
Source: "dist\belegpilot\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\belegpilot\_internal\*"; DestDir: "{app}\_internal"; Flags: ignoreversion recursesubdirs createallsubdirs

; Dienst-Programm
Source: "dist\belegpilot_service\{#MyServiceExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\belegpilot_service\_internal\*"; DestDir: "{app}\_internal"; Flags: ignoreversion recursesubdirs createallsubdirs

; Icons
Source: "gui/assets/icon.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "gui/assets/icon.png"; DestDir: "{app}"; Flags: ignoreversion

; Dependencies
Source: "dependencies\poppler\*"; DestDir: "{app}\dependencies\poppler"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "dependencies\Tesseract-OCR\*"; DestDir: "{app}\dependencies\Tesseract-OCR"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "dependencies\gs\*"; DestDir: "{app}\dependencies\gs"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; Startmenü-Eintrag f
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
; Optional: Desktop-Verknüpfung
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Registry]
; Registry-Einträge (falls benötigt)
Root: HKLM; Subkey: "Software\{#MyAppPublisher}\{#MyAppName}"; ValueType: string; ValueName: "InstallPath"; ValueData: "{app}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\{#MyAppPublisher}\{#MyAppName}"; ValueType: string; ValueName: "Version"; ValueData: "{#MyAppVersion}"

[Run]
; Installiert den Dienst mit automatischem Start
Filename: "{app}\{#MyServiceExeName}"; Parameters: "--startup auto install"; Flags: runhidden waituntilterminated
; Startet den Dienst direkt nach der Installation
Filename: "{app}\{#MyServiceExeName}"; Parameters: "start"; Flags: runhidden waituntilterminated

; Optional: Programm nach Installation starten
Filename: "{app}\{#MyAppExeName}"; Description: "{#MyAppName} starten"; Flags: nowait postinstall skipifsilent runascurrentuser

[UninstallRun]
; Stoppt den Dienst vor der Deinstallation
Filename: "{app}\{#MyServiceExeName}"; Parameters: "stop"; Flags: runhidden waituntilterminated
; Entfernt den Dienst aus dem System
Filename: "{app}\{#MyServiceExeName}"; Parameters: "remove"; Flags: runhidden waituntilterminated

[Messages]
; Deutsche Meldungen
SetupAppTitle=Installation von {#MyAppName}
SetupWindowTitle=Installation von {#MyAppName} {#MyAppVersion}

[CustomMessages]
CreateDesktopIcon=Desktop-Verknüpfung erstellen