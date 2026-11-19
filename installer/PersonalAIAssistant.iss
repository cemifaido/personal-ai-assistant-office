; Personal AI Assistant - installatore per utente (Inno Setup 6)
; Si compila con build.ps1, che prepara manifest, chiave pubblica e add-in Outlook.
;
; Architettura (verificata con lo spike su utente Windows standard):
;  - tutto per l'utente corrente, SENZA privilegi (PrivilegesRequired=lowest);
;  - Word/Excel: catalogo attendibile \\localhost\PersonalAIAssistant$ (condivisione in sola lettura
;    creata da condivisione.ps1 con UNA elevazione, solo se non esiste gia');
;  - Outlook classico: add-in VSTO in HKCU + elenco di inclusione VSTO (nessun avviso di fiducia).

#ifndef AppVer
  #define AppVer "0.1.0"
#endif
#ifndef BuildDir
  #define BuildDir "..\build"
#endif
#define AddinId "ecb39046-2d14-45ba-b5ea-94524ea141db"
#define ShareUnc "\\localhost\PersonalAIAssistant$"
#define CatalogId "{{0B7E4C2A-91D3-4F6E-8A25-6C1D9E3F7A40}"
#define InclusionId "{{5D8A1F36-2C47-4B9E-A0E3-8F61B2D74C95}"
#define OutlookKey "Software\Microsoft\Office\Outlook\Addins\PersonalAIAssistant.Outlook"
#ifndef SiteUrl
  #error SiteUrl non definito: compilare con build.ps1
#endif
#ifndef RepoUrl
  #error RepoUrl non definito: compilare con build.ps1
#endif
#define InstallPage SiteUrl + "/installa.html"

[Setup]
AppId={{E3B6F0A2-7C41-4D8E-9B15-2A6F8C3D1E74}
AppName=Personal AI Assistant
AppVersion={#AppVer}
AppVerName=Personal AI Assistant {#AppVer} (beta)
AppPublisher=Personal AI Assistant
AppPublisherURL={#RepoUrl}
AppSupportURL={#RepoUrl}/issues
VersionInfoVersion={#AppVer}.0
PrivilegesRequired=lowest
DefaultDirName={localappdata}\Programs\PersonalAIAssistant
DisableDirPage=yes
DisableProgramGroupPage=yes
OutputDir=..\release
OutputBaseFilename=PersonalAIAssistant-{#AppVer}-Setup
SetupIconFile=icona.ico
UninstallDisplayIcon={app}\icona.ico
UninstallDisplayName=Personal AI Assistant
InfoBeforeFile=info-prima.txt
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
CloseApplications=no

[Languages]
Name: "it"; MessagesFile: "compiler:Languages\Italian.isl"

[Messages]
it.FinishedLabel=Installazione completata.%n%nWORD ED EXCEL - solo la prima volta:%nHome > Componenti aggiuntivi (Add-ins) > Altri componenti aggiuntivi > scheda CARTELLA CONDIVISA > Personal AI Assistant > Aggiungi.%nDa quel momento trovi la scheda "Personal AI".%n%nOUTLOOK: gruppo "Personal AI" nella scheda Home.%n%nAl primo utilizzo inserisci la tua chiave API (Impostazioni / icona dell'ingranaggio).

[Tasks]
Name: "wordexcel"; Description: "Word ed Excel (pannello Personal AI; la prima volta chiede una conferma da amministratore)"
Name: "outlook"; Description: "Outlook classico per Windows (gruppo Personal AI nella scheda Home)"; Check: OutlookSupported

[Files]
Source: "icona.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildDir}\manifest.xml"; DestDir: "{app}\WordExcel"; Tasks: wordexcel; Flags: ignoreversion
Source: "condivisione.ps1"; DestDir: "{app}\WordExcel"; Flags: ignoreversion
Source: "{#BuildDir}\outlook\PersonalAIAssistant.Outlook.dll"; DestDir: "{app}\Outlook"; Tasks: outlook; Flags: ignoreversion
Source: "{#BuildDir}\outlook\PersonalAIAssistant.Outlook.dll.manifest"; DestDir: "{app}\Outlook"; Tasks: outlook; Flags: ignoreversion
Source: "{#BuildDir}\outlook\PersonalAIAssistant.Outlook.vsto"; DestDir: "{app}\Outlook"; Tasks: outlook; Flags: ignoreversion
Source: "{#BuildDir}\outlook\Microsoft.Office.Tools.Common.v4.0.Utilities.dll"; DestDir: "{app}\Outlook"; Tasks: outlook; Flags: ignoreversion
Source: "{#BuildDir}\outlook\Microsoft.Office.Tools.Outlook.v4.0.Utilities.dll"; DestDir: "{app}\Outlook"; Tasks: outlook; Flags: ignoreversion

[InstallDelete]
; Solo la cache di QUESTO add-in (versioni precedenti); il resto della cache di Office non si tocca.
Type: filesandordirs; Name: "{localappdata}\Microsoft\Office\16.0\Wef\AddinInfo\1\filesystem\Word\1\{#AddinId}_*"
Type: filesandordirs; Name: "{localappdata}\Microsoft\Office\16.0\Wef\AddinInfo\1\filesystem\Excel\1\{#AddinId}_*"

[UninstallDelete]
Type: filesandordirs; Name: "{localappdata}\Microsoft\Office\16.0\Wef\AddinInfo\1\filesystem\Word\1\{#AddinId}_*"
Type: filesandordirs; Name: "{localappdata}\Microsoft\Office\16.0\Wef\AddinInfo\1\filesystem\Excel\1\{#AddinId}_*"

[Registry]
; Pulizia delle prove precedenti: metodo "sviluppatore" (causa l'errore all'avvio) e add-in Outlook di prova
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\WEF\Developer"; ValueName: "{#AddinId}"; ValueType: none; Flags: deletevalue
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\WEF\Developer\{#AddinId}"; ValueType: none; Flags: deletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\PersonalAIAssistant.Spike"; ValueType: none; Flags: deletekey

; Word/Excel: catalogo attendibile mostrato nel menu
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{#CatalogId}"; ValueType: string; ValueName: "Id"; ValueData: "{#CatalogId}"; Tasks: wordexcel; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{#CatalogId}"; ValueType: string; ValueName: "Url"; ValueData: "{#ShareUnc}"; Tasks: wordexcel
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{#CatalogId}"; ValueType: dword; ValueName: "Flags"; ValueData: "1"; Tasks: wordexcel

; Outlook: add-in VSTO per l'utente corrente
Root: HKCU; Subkey: "{#OutlookKey}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Personal AI Assistant"; Tasks: outlook; Flags: uninsdeletekey
Root: HKCU; Subkey: "{#OutlookKey}"; ValueType: string; ValueName: "Description"; ValueData: "Riassunti e bozze di risposta con AI (Gemini, OpenAI, Claude)"; Tasks: outlook
Root: HKCU; Subkey: "{#OutlookKey}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Tasks: outlook
Root: HKCU; Subkey: "{#OutlookKey}"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:VstoUrl}|vstolocal"; Tasks: outlook
; Elenco di inclusione VSTO: fiducia limitata a QUESTO manifest firmato con QUESTA chiave pubblica
Root: HKCU; Subkey: "Software\Microsoft\VSTO\Security\Inclusion\{#InclusionId}"; ValueType: string; ValueName: "Url"; ValueData: "{code:VstoUrl}"; Tasks: outlook; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\VSTO\Security\Inclusion\{#InclusionId}"; ValueType: string; ValueName: "PublicKey"; ValueData: "{#FileRead(FileOpen(BuildDir + '\publickey.txt'))}"; Tasks: outlook

[Run]
Filename: "{#InstallPage}"; Description: "Apri la pagina con le istruzioni illustrate"; Flags: postinstall shellexec skipifsilent nowait

[Code]
function VstoUrl(Param: String): String;
var
  P: String;
begin
  P := ExpandConstant('{app}\Outlook\PersonalAIAssistant.Outlook.vsto');
  StringChangeEx(P, '\', '/', True);
  Result := 'file:///' + P;
end;

{ ---------- prerequisiti ---------- }

function OfficeInstalled(): Boolean;
begin
  Result := RegKeyExists(HKLM, 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration') or
            RegKeyExists(HKLM, 'SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot') or
            RegKeyExists(HKLM, 'SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Common\InstallRoot');
end;

function NetFx48Installed(): Boolean;
var
  Release: Cardinal;
begin
  Result := RegQueryDWordValue(HKLM, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', Release) and (Release >= 528040);
end;

function VstoRuntimeInstalled(): Boolean;
begin
  Result := RegKeyExists(HKLM, 'SOFTWARE\Microsoft\VSTO Runtime Setup\v4R') or
            RegKeyExists(HKLM, 'SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup\v4R') or
            RegKeyExists(HKLM, 'SOFTWARE\Microsoft\VSTO Runtime Setup\v4') or
            RegKeyExists(HKLM, 'SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup\v4');
end;

function OutlookClassicInstalled(): Boolean;
begin
  Result := RegKeyExists(HKLM, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE') or
            RegKeyExists(HKLM, 'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE');
end;

function OutlookSupported(): Boolean;
begin
  Result := OutlookClassicInstalled() and NetFx48Installed() and VstoRuntimeInstalled();
end;

function InitializeSetup(): Boolean;
var
  Msg: String;
begin
  Result := True;
  if not OfficeInstalled() then
    Result := MsgBox('Microsoft Office desktop non sembra installato su questo PC.' + #13#10 +
                     'Personal AI Assistant funziona con Word, Excel e Outlook per Windows. Continuare comunque?',
                     mbConfirmation, MB_YESNO) = IDYES;
  if Result and OutlookClassicInstalled() and not OutlookSupported() then
  begin
    Msg := 'La parte per Outlook non verra'' installata perche'' manca:';
    if not NetFx48Installed() then Msg := Msg + #13#10 + '- .NET Framework 4.8';
    if not VstoRuntimeInstalled() then Msg := Msg + #13#10 + '- Visual Studio 2010 Tools for Office Runtime';
    MsgBox(Msg + #13#10#13#10 + 'Word ed Excel funzionano comunque.', mbInformation, MB_OK);
  end;
end;

{ ---------- Office aperto? ---------- }

function OfficeRunning(): String;
begin
  Result := '';
  if FindWindowByClassName('OpusApp') <> 0 then Result := Result + ' Word';
  if FindWindowByClassName('XLMAIN') <> 0 then Result := Result + ' Excel';
  if FindWindowByClassName('rctrl_renwnd32') <> 0 then Result := Result + ' Outlook';
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  Running: String;
begin
  Result := '';
  repeat
    Running := OfficeRunning();
    if Running = '' then Exit;
  until MsgBox('Chiudi questi programmi prima di continuare:' + Running + #13#10#13#10 +
               'Premi OK quando li hai chiusi, Annulla per uscire.', mbInformation, MB_OKCANCEL) = IDCANCEL;
  Result := 'Chiudi Word, Excel e Outlook e avvia di nuovo l''installazione.';
end;

{ ---------- condivisione (unica parte con amministratore) ---------- }

function RunShareScript(Action: String; Uninstalling: Boolean; var RC: Integer): Boolean;
var
  Params: String;
begin
  Params := '-NoProfile -ExecutionPolicy Bypass -File "' + ExpandConstant('{app}\WordExcel\condivisione.ps1') + '" -Azione ' + Action;
  if Action = 'Crea' then
    Params := Params + ' -ManifestSource "' + ExpandConstant('{app}\WordExcel\manifest.xml') + '"';
  Result := ShellExec('runas', ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe'), Params, '', SW_HIDE, ewWaitUntilTerminated, RC);
end;

function ShareHasCurrentManifest(): Boolean;
var
  Shared, Local: AnsiString;
begin
  Result := LoadStringFromFile('{#ShareUnc}\manifest.xml', Shared) and
            LoadStringFromFile(ExpandConstant('{app}\WordExcel\manifest.xml'), Local) and (Shared = Local);
end;

procedure EnsureShare();
var
  RC: Integer;
begin
  if ShareHasCurrentManifest() then
  begin
    Log('Condivisione gia'' presente e aggiornata: nessuna elevazione.');
    Exit;
  end;
  if not RunShareScript('Crea', False, RC) then
  begin
    MsgBox('Word ed Excel richiedono una conferma da amministratore, che non e'' stata concessa.' + #13#10 +
           'Il resto e'' installato. Per Word ed Excel avvia di nuovo l''installazione quando puoi.', mbInformation, MB_OK);
    Exit;
  end;
  if RC = 2 then
    MsgBox('Il servizio di Windows "Server" non e'' attivo: Word ed Excel non possono usare il catalogo.' + #13#10 +
           'Chiedi all''amministratore del PC di attivarlo, poi ripeti l''installazione.', mbError, MB_OK)
  else if RC <> 0 then
    MsgBox('Creazione della cartella condivisa non riuscita (codice ' + IntToStr(RC) + ').' + #13#10 +
           'Dettagli: C:\ProgramData\PersonalAIAssistant\condivisione.log', mbError, MB_OK);
end;

{ ---------- cache di Office: solo le voci di QUESTO add-in ---------- }

procedure DeleteMatching(Dir, Pattern: String);
var
  F: TFindRec;
begin
  if FindFirst(Dir + Pattern, F) then
  try
    repeat
      if (F.Name <> '.') and (F.Name <> '..') then
      begin
        if (F.Attributes and FILE_ATTRIBUTE_DIRECTORY) <> 0 then
          DelTree(Dir + F.Name, True, True, True)
        else
          DeleteFile(Dir + F.Name);
        Log('Cache rimossa: ' + Dir + F.Name);
      end;
    until not FindNext(F);
  finally
    FindClose(F);
  end;
end;

{ Office conserva copie del manifest (anche di versioni precedenti, che considera piu' recenti)
  in Wef\<archivio>\<provider>\Manifests e la configurazione della scheda in AppCommands.
  Si rimuovono solo i file con l'Id di questo add-in: gli altri add-in non vengono toccati. }
procedure DeleteAddinCache();
var
  Wef: String;
  A, B: TFindRec;
begin
  Wef := ExpandConstant('{localappdata}\Microsoft\Office\16.0\Wef\');
  if FindFirst(Wef + '*', A) then
  try
    repeat
      if ((A.Attributes and FILE_ATTRIBUTE_DIRECTORY) <> 0) and (A.Name <> '.') and (A.Name <> '..') then
        if FindFirst(Wef + A.Name + '\*', B) then
        try
          repeat
            if ((B.Attributes and FILE_ATTRIBUTE_DIRECTORY) <> 0) and (B.Name <> '.') and (B.Name <> '..') then
              DeleteMatching(Wef + A.Name + '\' + B.Name + '\Manifests\', '{#AddinId}_*');
          until not FindNext(B);
        finally
          FindClose(B);
        end;
    until not FindNext(A);
  finally
    FindClose(A);
  end;
  DeleteMatching(Wef + 'AppCommands\11.0\TrustedCatalog\', '{#AddinId}*');
  DeleteMatching(Wef + 'AddinInfo\1\filesystem\Word\1\', '{#AddinId}_*');
  DeleteMatching(Wef + 'AddinInfo\1\filesystem\Excel\1\', '{#AddinId}_*');
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssInstall then
    DeleteAddinCache();
  if (CurStep = ssPostInstall) and WizardIsTaskSelected('wordexcel') then
    EnsureShare();
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  RC: Integer;
begin
  if (CurUninstallStep = usUninstall) and DirExists('{#ShareUnc}') then
    if MsgBox('Rimuovere anche la cartella condivisa del PC (\\localhost\PersonalAIAssistant$)?' + #13#10#13#10 +
              'Rispondi No se altri utenti di questo PC usano Personal AI Assistant.' + #13#10 +
              'Serve una conferma da amministratore.', mbConfirmation, MB_YESNO) = IDYES then
      RunShareScript('Rimuovi', True, RC);
  if CurUninstallStep = usPostUninstall then
    DeleteAddinCache();
end;
