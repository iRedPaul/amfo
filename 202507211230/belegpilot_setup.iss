; Inno Setup Script für belegpilot
#define MyAppName "belegpilot"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "amsepa AG"
#define MyAppURL "https://belegpilot.io"
#define MyAppExeName "belegpilot.exe"

[Setup]
; Eindeutige App-ID (generieren Sie eine neue GUID)
AppId={{A7B3F4E2-9C8D-4F2A-B6E1-3E6F8A9E2C1B}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
; Ausgabeverzeichnis und Dateiname
OutputDir=installer
OutputBaseFilename=belegpilot_Setup_{#MyAppVersion}
; Icon für das Setup
SetupIconFile=gui/assets/icon.ico
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
; Admin-Rechte erforderlich
PrivilegesRequired=admin
; 64-Bit Installation
ArchitecturesInstallIn64BitMode=x64compatible
; Keine Sprachauswahl - nur Deutsch
ShowLanguageDialog=no
; Icon auch für Deinstallation
UninstallDisplayIcon={app}\{#MyAppExeName}
; Anwendung beim Update schließen
CloseApplications=yes
CloseApplicationsFilter=*.exe

[Languages]
Name: "german"; MessagesFile: "compiler:Languages\German.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
; Hauptprogramm
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

; Icon (ico)
Source: "gui/assets/icon.ico"; DestDir: "{app}"; Flags: ignoreversion

; Icon (png)
Source: "gui/assets/icon.png"; DestDir: "{app}"; Flags: ignoreversion

; Dependencies - Poppler (PDF-Tools)
Source: "dependencies\poppler\*"; DestDir: "{app}\dependencies\poppler"; Flags: ignoreversion recursesubdirs createallsubdirs

; Dependencies - Tesseract OCR
Source: "dependencies\Tesseract-OCR\*"; DestDir: "{app}\dependencies\Tesseract-OCR"; Flags: ignoreversion recursesubdirs createallsubdirs

; Dependencies - Ghostscript
Source: "dependencies\gs\*"; DestDir: "{app}\dependencies\gs"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; Startmenü-Eintrag
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
; Desktop-Verknüpfung (optional)
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Registry]
; Registry-Einträge für die Anwendung
Root: HKLM; Subkey: "Software\{#MyAppPublisher}\{#MyAppName}"; ValueType: string; ValueName: "InstallPath"; ValueData: "{app}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\{#MyAppPublisher}\{#MyAppName}"; ValueType: string; ValueName: "Version"; ValueData: "{#MyAppVersion}"

; Registry-Eintrag damit die EXE immer als Admin ausgeführt wird
Root: HKLM; Subkey: "Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers"; ValueType: string; ValueName: "{app}\{#MyAppExeName}"; ValueData: "RUNASADMIN"; Flags: uninsdeletekeyifempty uninsdeletevalue

[Run]
; Hauptprogramm nach Installation starten
Filename: "{app}\{#MyAppExeName}"; Description: "belegpilot starten"; Flags: nowait postinstall skipifsilent runascurrentuser

[Messages]
; Deutsche Meldungen
BeveledLabel=
SetupAppTitle=Installation - {#MyAppName}
SetupWindowTitle=Installation - {#MyAppName} {#MyAppVersion}
WelcomeLabel1=Willkommen zur Installation von [name]
WelcomeLabel2=Dieses Programm installiert [name/ver] auf Ihrem Computer.%n%nEs wird empfohlen, alle anderen Anwendungen zu schließen, bevor Sie fortfahren.
FinishedLabel=Die Installation von [name] wurde erfolgreich abgeschlossen.
FinishedHeadingLabel=Installation abgeschlossen
UninstallStatusLabel=Entferne %1 vom Computer...

[CustomMessages]
CreateDesktopIcon=Desktop-Verknüpfung erstellen