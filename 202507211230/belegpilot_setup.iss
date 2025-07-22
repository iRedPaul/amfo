; Inno Setup Script für belegpilot
#define MyAppName "belegpilot"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "amsepa AG"
#define MyAppURL "https://belegpilot.io"
#define MyAppExeName "belegpilot.exe"
#define MyServiceExeName "belegpilot_service.exe"

[Setup]
; Eindeutige App-ID (generieren Sie eine neue GUID)
AppId={{A7B3F4E2-9D8D-4F2A-B7E1-3E6F8A9E2C1B}
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
Name: "installservice"; Description: "Als Windows-Dienst installieren (für automatische Verarbeitung)"; GroupDescription: "Windows-Dienst:"; Flags: unchecked

[Files]
; Hauptprogramm
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

; Windows Service Executable
Source: "dist\{#MyServiceExeName}"; DestDir: "{app}"; Flags: ignoreversion

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
; Service-Verwaltung im Startmenü
Name: "{group}\Windows-Dienst verwalten"; Filename: "{app}\{#MyServiceExeName}"; Parameters: "status"; IconFilename: "{app}\icon.ico"
; Desktop-Verknüpfung (optional)
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Registry]
; Registry-Einträge für die Anwendung
Root: HKLM; Subkey: "Software\{#MyAppPublisher}\{#MyAppName}"; ValueType: string; ValueName: "InstallPath"; ValueData: "{app}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\{#MyAppPublisher}\{#MyAppName}"; ValueType: string; ValueName: "Version"; ValueData: "{#MyAppVersion}"

; Registry-Eintrag damit die EXE immer als Admin ausgeführt wird
Root: HKLM; Subkey: "Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers"; ValueType: string; ValueName: "{app}\{#MyAppExeName}"; ValueData: "RUNASADMIN"; Flags: uninsdeletekeyifempty uninsdeletevalue

[Run]
; Service installieren (wenn ausgewählt)
Filename: "{app}\{#MyServiceExeName}"; Parameters: "install"; StatusMsg: "Installiere Windows-Dienst..."; Flags: runhidden; Tasks: installservice
; Service starten (wenn installiert)
Filename: "{app}\{#MyServiceExeName}"; Parameters: "start"; StatusMsg: "Starte Windows-Dienst..."; Flags: runhidden; Tasks: installservice

; Hauptprogramm nach Installation starten
Filename: "{app}\{#MyAppExeName}"; Description: "belegpilot starten"; Flags: nowait postinstall skipifsilent runascurrentuser

[UninstallRun]
; Service stoppen und entfernen beim Deinstallieren
Filename: "{app}\{#MyServiceExeName}"; Parameters: "stop"; RunOnceId: "StopBelegpilotService"; Flags: runhidden
Filename: "{app}\{#MyServiceExeName}"; Parameters: "remove"; RunOnceId: "RemoveBelegpilotService"; Flags: runhidden

[Code]
var
  ServicePage: TInputOptionWizardPage;

procedure InitializeWizard;
begin
  // Erstelle eine benutzerdefinierte Seite für Service-Optionen
  ServicePage := CreateInputOptionPage(wpSelectTasks,
    'Windows-Dienst Konfiguration',
    'Möchten Sie belegpilot als Windows-Dienst installieren?',
    'Ein Windows-Dienst ermöglicht die automatische Verarbeitung von PDFs im Hintergrund, ' +
    'auch wenn kein Benutzer angemeldet ist. Der Dienst startet automatisch beim Systemstart.' + #13#10#13#10 +
    'Wählen Sie diese Option, wenn Sie eine dauerhafte Überwachung der Hotfolder wünschen.',
    True, False);
    
  ServicePage.Add('Windows-Dienst installieren und automatisch starten');
  ServicePage.Values[0] := True;
end;

function ShouldInstallService: Boolean;
begin
  Result := ServicePage.Values[0];
end;

// Prüfe ob der Service bereits installiert ist
function IsServiceInstalled: Boolean;
var
  ResultCode: Integer;
begin
  Result := Exec('sc.exe', 'query belegpilot', '', SW_HIDE, ewWaitUntilTerminated, ResultCode) and (ResultCode = 0);
end;

// Event vor der Installation
procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
begin
  if CurStep = ssInstall then
  begin
    // Stoppe den Service falls er läuft
    if IsServiceInstalled then
    begin
      Exec(ExpandConstant('{app}\{#MyServiceExeName}'), 'stop', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    end;
  end;
end;

// Event vor der Deinstallation
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  ResultCode: Integer;
begin
  if CurUninstallStep = usUninstall then
  begin
    // Stoppe und entferne den Service
    if IsServiceInstalled then
    begin
      Exec(ExpandConstant('{app}\{#MyServiceExeName}'), 'stop', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
      Sleep(2000); // Warte 2 Sekunden
      Exec(ExpandConstant('{app}\{#MyServiceExeName}'), 'remove', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    end;
  end;
end;

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
