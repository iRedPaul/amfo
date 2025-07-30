; Inno Setup Script für belegpilot
#define MyAppName "belegpilot"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "amsepa AG"
#define MyAppURL "https://belegpilot.io"
#define MyAppExeName "belegpilot.exe"
#define MyServiceExeName "belegpilot_service.exe"

[Setup]
AppId={{A7B3F4E2-9C8D-4F2A-B6E1-5E9F3A9E3C2B}
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
; Wichtig: Force kill bei Deinstallation
UninstallRestartComputer=no

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
; WICHTIG: Zuerst GUI beenden (falls läuft)
Filename: "{sys}\taskkill.exe"; Parameters: "/F /IM {#MyAppExeName}"; Flags: runhidden waituntilterminated; RunOnceId: "KillGUI"

; Dienst stoppen
Filename: "{app}\{#MyServiceExeName}"; Parameters: "stop"; Flags: runhidden waituntilterminated; RunOnceId: "StopService"

; Warte kurz
Filename: "{sys}\timeout.exe"; Parameters: "/T 2"; Flags: runhidden waituntilterminated; RunOnceId: "Wait1"

; Dienst entfernen
Filename: "{app}\{#MyServiceExeName}"; Parameters: "remove"; Flags: runhidden waituntilterminated; RunOnceId: "RemoveService"

; Force kill falls Service noch läuft
Filename: "{sys}\taskkill.exe"; Parameters: "/F /IM {#MyServiceExeName}"; Flags: runhidden waituntilterminated; RunOnceId: "KillService"

; Warte nochmal
Filename: "{sys}\timeout.exe"; Parameters: "/T 2"; Flags: runhidden waituntilterminated; RunOnceId: "Wait2"

[Dirs]
; Explizit zu löschende Verzeichnisse
Name: "{app}\_internal"; Flags: uninsalwaysuninstall
Name: "{app}\dependencies"; Flags: uninsalwaysuninstall
Name: "{app}\logs"; Flags: uninsalwaysuninstall
Name: "{app}"; Flags: uninsalwaysuninstall

[InstallDelete]
; Lösche alte Dateien vor Installation
Type: filesandordirs; Name: "{app}\_internal"
Type: filesandordirs; Name: "{app}\logs"

[UninstallDelete]
; Explizite Löschanweisungen bei Deinstallation
Type: filesandordirs; Name: "{app}\_internal"
Type: filesandordirs; Name: "{app}\dependencies"
Type: filesandordirs; Name: "{app}\logs"
Type: files; Name: "{app}\{#MyServiceExeName}"
Type: files; Name: "{app}\{#MyAppExeName}"
Type: files; Name: "{app}\*.log"
Type: files; Name: "{app}\*.tmp"
Type: filesandordirs; Name: "{app}"

[Messages]
; Deutsche Meldungen
SetupAppTitle=Installation von {#MyAppName}
SetupWindowTitle=Installation von {#MyAppName} {#MyAppVersion}
FinishedLabel={#MyAppName} wurde erfolgreich installiert.%n%nWICHTIG: Der Windows-Dienst muss noch konfiguriert werden!%n%n1. Öffnen Sie services.msc%n2. Suchen Sie "belegpilot Service"%n3. Rechtklick → Eigenschaften → Anmelden%n4. Wählen Sie ein Benutzerkonto%n5. Setzen Sie Starttyp auf "Automatisch"%n6. Starten Sie den Dienst

[CustomMessages]
CreateDesktopIcon=Desktop-Verknüpfung erstellen

[Code]
{ Pascal Script Code für erweiterte Deinstallation }

procedure StopServiceWithSC();
var
  ResultCode: Integer;
begin
  { Versuche Service mit sc.exe zu stoppen }
  Exec('sc.exe', 'stop BelegpilotService', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Sleep(2000);
  
  { Versuche Service zu löschen }
  Exec('sc.exe', 'delete BelegpilotService', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Sleep(1000);
end;

procedure KillRunningProcesses();
var
  ResultCode: Integer;
begin
  { Kill alle laufenden Prozesse }
  Exec('taskkill.exe', '/F /IM belegpilot.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Exec('taskkill.exe', '/F /IM belegpilot_service.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Sleep(2000);
end;

function InitializeUninstall(): Boolean;
begin
  { Vor der Deinstallation: Services und Prozesse beenden }
  KillRunningProcesses();
  StopServiceWithSC();
  Result := True;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  ResultCode: Integer;
  AppDir: String;
begin
  if CurUninstallStep = usPostUninstall then
  begin
    { Nach der Deinstallation: Versuche verbleibende Dateien zu löschen }
    AppDir := ExpandConstant('{app}');
    
    { Nochmal kill falls wieder gestartet }
    KillRunningProcesses();
    
    { Lösche _internal Ordner mit rmdir }
    Exec('cmd.exe', '/c rmdir /s /q "' + AppDir + '\_internal"', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    
    { Lösche Service-EXE }
    DeleteFile(AppDir + '\belegpilot_service.exe');
    
    { Lösche App-Verzeichnis }
    Exec('cmd.exe', '/c rmdir /s /q "' + AppDir + '"', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  end;
end;