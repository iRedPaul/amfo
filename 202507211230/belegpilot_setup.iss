; Inno Setup Script für belegpilot (angepasst für Dienst)
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
; Hauptprogramm (GUI Konfigurationstool)
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
; Dienst-Programm
Source: "dist\{#MyServiceExeName}"; DestDir: "{app}"; Flags: ignoreversion

; Icons
Source: "gui/assets/icon.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "gui/assets/icon.png"; DestDir: "{app}"; Flags: ignoreversion

; Dependencies
Source: "dependencies\poppler\*"; DestDir: "{app}\dependencies\poppler"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "dependencies\Tesseract-OCR\*"; DestDir: "{app}\dependencies\Tesseract-OCR"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "dependencies\gs\*"; DestDir: "{app}\dependencies\gs"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; Startmenü-Eintrag für das Konfigurationstool
Name: "{group}\{#MyAppName} Konfiguration"; Filename: "{app}\{#MyAppExeName}"
; Desktop-Verknüpfung für das Konfigurationstool (optional)
Name: "{autodesktop}\{#MyAppName} Konfiguration"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Registry]
; [cite_start]... (deine bisherigen Registry-Einträge bleiben unverändert) [cite: 21, 22, 23, 24]

[Run]
; Installiert den Dienst und setzt den Starttyp auf "Automatisch"
Filename: "{app}\{#MyServiceExeName}"; Parameters: "install --startup auto"; Flags: runhidden waituntilterminated
; Startet den Dienst direkt nach der Installation
Filename: "{app}\{#MyServiceExeName}"; Parameters: "start"; Flags: runhidden waituntilterminated

; Optional: GUI nach Installation starten
Filename: "{app}\{#MyAppExeName}"; Description: "{#MyAppName} Konfiguration starten"; Flags: nowait postinstall skipifsilent runascurrentuser

[UninstallRun]
; Stoppt den Dienst vor der Deinstallation (wichtig für sauberes Entfernen)
Filename: "{app}\{#MyServiceExeName}"; Parameters: "stop"; Flags: runhidden waituntilterminated
; Entfernt den Dienst aus dem System
Filename: "{app}\{#MyServiceExeName}"; Parameters: "remove"; Flags: runhidden waituntilterminated

[Messages]
; [cite_start]... (deine bisherigen Messages bleiben unverändert) [cite: 25, 26]

[CustomMessages]
CreateDesktopIcon=Desktop-Verknüpfung für Konfiguration erstellen
