; Inno Setup Script für belegpilot
#define MyAppName "belegpilot"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "amsepa AG"
#define MyAppURL "https://belegpilot.io"
#define MyAppExeName "belegpilot.exe"
#define MyServiceExeName "belegpilot_service.exe"

[Setup]
AppId={{A7B3F4E2-9C8D-4F2A-B6E1-4E7F8A9E2C1B}
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
; Startmenü-Eintrag
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"

; Optional: Desktop-Verknüpfung
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Registry]
; Registry-Einträge (falls benötigt)
Root: HKLM; Subkey: "Software\{#MyAppPublisher}\{#MyAppName}"; ValueType: string; ValueName: "InstallPath"; ValueData: "{app}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\{#MyAppPublisher}\{#MyAppName}"; ValueType: string; ValueName: "Version"; ValueData: "{#MyAppVersion}"

[Run]
; Installiert den Dienst OHNE automatischen Start (manuell)
Filename: "{app}\{#MyServiceExeName}"; Parameters: "--startup manual install"; Flags: runhidden waituntilterminated; StatusMsg: "Installiere belegpilot Dienst..."

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
FinishedLabel={#MyAppName} wurde erfolgreich installiert.%n%nWICHTIG: Der Windows-Dienst muss noch konfiguriert werden!%n%n1. Öffnen Sie services.msc%n2. Suchen Sie "belegpilot Hotfolder Service"%n3. Rechtklick → Eigenschaften → Anmelden%n4. Wählen Sie ein Benutzerkonto%n5. Setzen Sie Starttyp auf "Automatisch"%n6. Starten Sie den Dienst

[CustomMessages]
CreateDesktopIcon=Desktop-Verknüpfung erstellen