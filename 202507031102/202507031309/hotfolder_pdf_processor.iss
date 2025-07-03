; Hotfolder PDF Processor - Inno Setup Script

#define MyAppName "Hotfolder PDF Processor"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Ihre Firma GmbH"
#define MyAppURL "https://www.ihrefirma.de"
#define MyAppExeName "Hotfolder PDF Processor.exe"
#define MyServiceExeName "HotfolderPDFService.exe"

[Setup]
AppId={{B5A7F2D8-3C4D-5E6F-7A8B-9C0D1E2F3A4B}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}/support
AppUpdatesURL={#MyAppURL}/updates
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
DisableDirPage=no
DisableReadyPage=no
DisableFinishedPage=no
OutputBaseFilename=HotfolderPDFProcessor_Setup_{#MyAppVersion}
Compression=lzma2/ultra64
SolidCompression=yes
CompressionThreads=auto
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64
MinVersion=10.0
WizardStyle=modern
PrivilegesRequired=admin
SetupIconFile=icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallDisplayName={#MyAppName}
CloseApplications=yes
RestartApplications=no
ShowLanguageDialog=no
VersionInfoVersion={#MyAppVersion}
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription={#MyAppName} Setup
VersionInfoCopyright=© 2024 {#MyAppPublisher}
WizardImageFile=compiler:WizModernImage-IS.bmp
WizardSmallImageFile=compiler:WizModernSmallImage-IS.bmp

[Languages]
Name: "german"; MessagesFile: "compiler:Languages\German.isl"

[Messages]
german.BeveledLabel=Installation

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; OnlyBelowVersion: 0,6.1
Name: "service"; Description: "Als Windows-Dienst installieren (empfohlen für Server)"; GroupDescription: "Service-Optionen:"; Flags: unchecked

[Files]
; Hauptprogramm und Service
Source: "dist\Hotfolder PDF Processor\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Dirs]
Name: "{app}\logs"
Name: "{app}\errors"
Name: "{app}\temp"
Name: "{commonappdata}\{#MyAppName}"
Name: "{commonappdata}\{#MyAppName}\logs"
Name: "{commonappdata}\{#MyAppName}\errors"
Name: "{commonappdata}\{#MyAppName}\hotfolders"
Name: "{userappdata}\{#MyAppName}"

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: quicklaunchicon
Name: "{autoprograms}\{#MyAppName}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autoprograms}\{#MyAppName}\Logs"; Filename: "{commonappdata}\{#MyAppName}\logs"

[Registry]
Root: HKLM; Subkey: "SOFTWARE\{#MyAppPublisher}"; Flags: uninsdeletekeyifempty
Root: HKLM; Subkey: "SOFTWARE\{#MyAppPublisher}\{#MyAppName}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\{#MyAppPublisher}\{#MyAppName}"; ValueType: string; ValueName: "InstallPath"; ValueData: "{app}"
Root: HKLM; Subkey: "SOFTWARE\{#MyAppPublisher}\{#MyAppName}"; ValueType: string; ValueName: "Version"; ValueData: "{#MyAppVersion}"
Root: HKLM; Subkey: "SOFTWARE\{#MyAppPublisher}\{#MyAppName}"; ValueType: string; ValueName: "Publisher"; ValueData: "{#MyAppPublisher}"

[Run]
; Service installieren (wenn ausgewählt)
Filename: "{sys}\sc.exe"; Parameters: "create HotfolderPDFProcessor binPath= ""{app}\{#MyServiceExeName}"" DisplayName= ""{#MyAppName} Service"" start= auto"; Flags: runhidden; Tasks: service
Filename: "{sys}\sc.exe"; Parameters: "description HotfolderPDFProcessor ""Überwacht Hotfolder und verarbeitet PDF-Dateien automatisch"""; Flags: runhidden; Tasks: service
Filename: "{sys}\sc.exe"; Parameters: "start HotfolderPDFProcessor"; Flags: runhidden; Tasks: service

; Programm starten
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent unchecked

[UninstallRun]
; Service stoppen und entfernen
Filename: "{sys}\sc.exe"; Parameters: "stop HotfolderPDFProcessor"; Flags: runhidden; RunOnceId: "StopService"
Filename: "{sys}\sc.exe"; Parameters: "delete HotfolderPDFProcessor"; Flags: runhidden; RunOnceId: "DeleteService"

; Prozesse beenden
Filename: "{sys}\taskkill.exe"; Parameters: "/F /IM ""{#MyAppExeName}"""; Flags: runhidden; RunOnceId: "KillApp"
Filename: "{sys}\taskkill.exe"; Parameters: "/F /IM ""{#MyServiceExeName}"""; Flags: runhidden; RunOnceId: "KillService"

[UninstallDelete]
Type: filesandordirs; Name: "{app}\logs"
Type: filesandordirs; Name: "{app}\errors"
Type: filesandordirs; Name: "{app}\temp"
Type: filesandordirs; Name: "{app}\__pycache__"
Type: dirifempty; Name: "{app}"
Type: dirifempty; Name: "{commonappdata}\{#MyAppName}"

[Code]
var
  DependencyPage: TOutputMsgWizardPage;

function InitializeSetup: Boolean;
begin
  Result := True;
  
  // Prüfe ob bereits installiert
  if RegKeyExists(HKEY_LOCAL_MACHINE, 'SOFTWARE\{#MyAppPublisher}\{#MyAppName}') then
  begin
    if MsgBox('Eine Version von {#MyAppName} ist bereits installiert.' + #13#10 + 
              'Möchten Sie die Installation fortsetzen?' + #13#10 + #13#10 +
              'Die vorhandene Installation wird aktualisiert.', 
              mbConfirmation, MB_YESNO) = IDNO then
    begin
      Result := False;
    end;
  end;
end;

procedure InitializeWizard;
begin
  // Erstelle Informationsseite für externe Abhängigkeiten
  DependencyPage := CreateOutputMsgPage(wpSelectDir,
    'Externe Abhängigkeiten',
    'Optionale Programme für erweiterte Funktionen',
    'Die folgenden Programme werden für spezielle Funktionen benötigt:' + #13#10 + #13#10 +
    
    'Tesseract OCR (Texterkennung):' + #13#10 +
    '  • Wird für OCR-Funktionen benötigt' + #13#10 +
    '  • Download: https://github.com/UB-Mannheim/tesseract/wiki' + #13#10 + #13#10 +
    
    'Ghostscript (PDF-Komprimierung):' + #13#10 +
    '  • Wird für erweiterte PDF-Komprimierung benötigt' + #13#10 +
    '  • Download: https://www.ghostscript.com/download/gsdnld.html' + #13#10 + #13#10 +
    
    'Diese Programme müssen separat installiert werden.' + #13#10 +
    'Das Hauptprogramm funktioniert auch ohne diese Abhängigkeiten.'
  );
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
  TesseractPath, GhostscriptPath: String;
begin
  if CurStep = ssPostInstall then
  begin
    // Suche nach Tesseract
    TesseractPath := '';
    if FileExists(ExpandConstant('{pf}\Tesseract-OCR\tesseract.exe')) then
      TesseractPath := ExpandConstant('{pf}\Tesseract-OCR\tesseract.exe')
    else if FileExists(ExpandConstant('{pf32}\Tesseract-OCR\tesseract.exe')) then
      TesseractPath := ExpandConstant('{pf32}\Tesseract-OCR\tesseract.exe');
    
    // Suche nach Ghostscript
    GhostscriptPath := '';
    if FileExists(ExpandConstant('{pf}\gs\gs10.04.0\bin\gswin64c.exe')) then
      GhostscriptPath := ExpandConstant('{pf}\gs\gs10.04.0\bin\gswin64c.exe')
    else if FileExists(ExpandConstant('{pf32}\gs\gs10.04.0\bin\gswin32c.exe')) then
      GhostscriptPath := ExpandConstant('{pf32}\gs\gs10.04.0\bin\gswin32c.exe');
    
    // Registriere gefundene Pfade
    if TesseractPath <> '' then
      RegWriteStringValue(HKEY_LOCAL_MACHINE, 'SOFTWARE\{#MyAppPublisher}\{#MyAppName}', 
                          'TesseractPath', TesseractPath);
    
    if GhostscriptPath <> '' then
      RegWriteStringValue(HKEY_LOCAL_MACHINE, 'SOFTWARE\{#MyAppPublisher}\{#MyAppName}', 
                          'GhostscriptPath', GhostscriptPath);
  end;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  ResultCode: Integer;
begin
  // Beende laufende Instanzen
  Exec('taskkill.exe', '/F /IM "{#MyAppExeName}"', '', SW_HIDE, 
       ewWaitUntilTerminated, ResultCode);
  Exec('taskkill.exe', '/F /IM "{#MyServiceExeName}"', '', SW_HIDE, 
       ewWaitUntilTerminated, ResultCode);
  
  // Service stoppen falls vorhanden
  Exec(ExpandConstant('{sys}\sc.exe'), 'stop HotfolderPDFProcessor', '', 
       SW_HIDE, ewWaitUntilTerminated, ResultCode);
  
  Result := '';
end;

function UninstallShouldRemoveData: Boolean;
begin
  Result := MsgBox('Möchten Sie alle Programmeinstellungen und Benutzerdaten entfernen?' + #13#10 + #13#10 +
                   'Dies umfasst:' + #13#10 +
                   '  • Hotfolder-Konfigurationen' + #13#10 +
                   '  • Einstellungen' + #13#10 +
                   '  • Log-Dateien' + #13#10 + #13#10 +
                   'Wählen Sie "Nein", um die Daten für eine spätere Neuinstallation zu behalten.',
                   mbConfirmation, MB_YESNO) = IDYES;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
  begin
    if UninstallShouldRemoveData then
    begin
      // Entferne Benutzerdaten
      DelTree(ExpandConstant('{userappdata}\{#MyAppName}'), True, True, True);
      DelTree(ExpandConstant('{commonappdata}\{#MyAppName}'), True, True, True);
    end;
  end;
end;