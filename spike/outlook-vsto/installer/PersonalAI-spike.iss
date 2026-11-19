; Personal AI Assistant - INSTALLATORE DI PROVA (spike, non per la distribuzione)
; 0.0.2: Word/Excel tramite catalogo in cartella condivisa (opzione A).
;   - l'installatore gira SENZA privilegi (impostazioni nel profilo dell'utente che installa);
;   - solo la creazione della condivisione \\localhost\PersonalAIAssistant$ chiede l'elevazione,
;     e solo se non esiste gia';
;   - opzione sperimentale: pre-inserisce l'add-in nella cache di Office (boot.json, non documentato).

#define AppVer "0.0.2"
#define PubKeyFile "..\installer-publickey.txt"
#define VstoDir "..\PaiOutlookSpike\bin\Release"
#define WordExcelManifest "..\..\..\manifest\manifest-word-excel.xml"
#define WordExcelAddinId "ecb39046-2d14-45ba-b5ea-94524ea141db"
#define WordExcelManifestVersion "2.0.0.0"
#define ShareUnc "\\localhost\PersonalAIAssistant$"
#define CatalogId "{{6F1D2B9E-3C4A-4E8B-9A57-2D0C8E1B4A63}"
#define InclusionId "{{3B2E7C41-8F0A-4C5D-9E61-7A2B4C9D0E11}"

[Setup]
AppId={{9C1F4A2E-5B7D-4E3A-A8C6-2D9E0F1B3C57}
AppName=Personal AI Assistant (prova tecnica)
AppVersion={#AppVer}
AppPublisher=Personal AI Assistant
PrivilegesRequired=lowest
DefaultDirName={localappdata}\PersonalAIAssistant-spike
DisableDirPage=yes
DisableProgramGroupPage=yes
OutputDir=..\out
OutputBaseFilename=PersonalAI-spike-{#AppVer}-Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
UninstallDisplayName=Personal AI Assistant (prova tecnica)

[Languages]
Name: "it"; MessagesFile: "compiler:Languages\Italian.isl"

[Tasks]
Name: "wordexcel"; Description: "Word ed Excel (catalogo in cartella condivisa; puo' chiedere una conferma da amministratore)"
Name: "wordexcel\subito"; Description: "Mostra subito la scheda Personal AI (sperimentale)"
Name: "outlook"; Description: "Outlook classico (add-in di prova)"
Name: "outlook\fiducia"; Description: "Considera attendibile l'add-in Outlook (evita la richiesta di conferma)"

[Files]
Source: "{#WordExcelManifest}"; DestDir: "{app}\WordExcel"; DestName: "manifest.xml"; Tasks: wordexcel; Flags: ignoreversion
Source: "crea-condivisione.ps1"; DestDir: "{app}\WordExcel"; Tasks: wordexcel; Flags: ignoreversion
Source: "{#VstoDir}\PaiOutlookSpike.dll"; DestDir: "{app}\Outlook"; Tasks: outlook; Flags: ignoreversion
Source: "{#VstoDir}\PaiOutlookSpike.dll.manifest"; DestDir: "{app}\Outlook"; Tasks: outlook; Flags: ignoreversion
Source: "{#VstoDir}\PaiOutlookSpike.vsto"; DestDir: "{app}\Outlook"; Tasks: outlook; Flags: ignoreversion
Source: "{#VstoDir}\Microsoft.Office.Tools.Common.v4.0.Utilities.dll"; DestDir: "{app}\Outlook"; Tasks: outlook; Flags: ignoreversion
Source: "{#VstoDir}\Microsoft.Office.Tools.Outlook.v4.0.Utilities.dll"; DestDir: "{app}\Outlook"; Tasks: outlook; Flags: ignoreversion

[Registry]
; Pulizia della prova 0.0.1: il metodo "sviluppatore" causa l'errore all'avvio di Word
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\WEF\Developer"; ValueName: "{#WordExcelAddinId}"; ValueType: none; Flags: deletevalue

; Word/Excel: catalogo attendibile (cartella condivisa locale), mostrato nel menu
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{#CatalogId}"; ValueType: string; ValueName: "Id"; ValueData: "{#CatalogId}"; Tasks: wordexcel; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{#CatalogId}"; ValueType: string; ValueName: "Url"; ValueData: "{#ShareUnc}"; Tasks: wordexcel
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{#CatalogId}"; ValueType: dword; ValueName: "Flags"; ValueData: "1"; Tasks: wordexcel

; Outlook: add-in VSTO per l'utente corrente
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\PersonalAIAssistant.Spike"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Personal AI Assistant (prova)"; Tasks: outlook; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\PersonalAIAssistant.Spike"; ValueType: string; ValueName: "Description"; ValueData: "Prova tecnica di installazione"; Tasks: outlook
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\PersonalAIAssistant.Spike"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Tasks: outlook
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\PersonalAIAssistant.Spike"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:VstoUrl}|vstolocal"; Tasks: outlook
Root: HKCU; Subkey: "Software\Microsoft\VSTO\Security\Inclusion\{#InclusionId}"; ValueType: string; ValueName: "Url"; ValueData: "{code:VstoUrl}"; Tasks: outlook\fiducia; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\VSTO\Security\Inclusion\{#InclusionId}"; ValueType: string; ValueName: "PublicKey"; ValueData: "{#FileRead(FileOpen(PubKeyFile))}"; Tasks: outlook\fiducia

[UninstallDelete]
Type: filesandordirs; Name: "{localappdata}\Microsoft\Office\16.0\Wef\AddinInfo\1\filesystem\Word\1\{#WordExcelAddinId}_{#WordExcelManifestVersion}"
Type: filesandordirs; Name: "{localappdata}\Microsoft\Office\16.0\Wef\AddinInfo\1\filesystem\Excel\1\{#WordExcelAddinId}_{#WordExcelManifestVersion}"

[Code]
function VstoUrl(Param: String): String;
var
  P: String;
begin
  P := ExpandConstant('{app}\Outlook\PaiOutlookSpike.vsto');
  StringChangeEx(P, '\', '/', True);
  Result := 'file:///' + P;
end;

{ Crea la condivisione solo se non e' gia' raggiungibile; chiede l'elevazione solo in quel caso. }
function EnsureShare(): Boolean;
var
  RC: Integer;
  Params: String;
begin
  Result := FileExists('{#ShareUnc}\manifest.xml');
  if Result then
  begin
    Log('Condivisione gia'' presente: nessuna elevazione necessaria.');
    Exit;
  end;
  Params := '-NoProfile -ExecutionPolicy Bypass -File "' + ExpandConstant('{app}\WordExcel\crea-condivisione.ps1') +
            '" -ManifestSource "' + ExpandConstant('{app}\WordExcel\manifest.xml') + '"';
  if not ShellExec('runas', 'powershell.exe', Params, '', SW_HIDE, ewWaitUntilTerminated, RC) then
  begin
    MsgBox('Per Word ed Excel serve una conferma da amministratore, che non e'' stata concessa.' + #13#10 +
           'Outlook e'' comunque installato. Puoi rilanciare l''installazione in seguito.', mbInformation, MB_OK);
    Exit;
  end;
  if RC = 2 then
    MsgBox('Il servizio Windows "Server" non e'' attivo: Word ed Excel non possono usare il catalogo.', mbError, MB_OK)
  else if RC <> 0 then
    MsgBox('Creazione della cartella condivisa non riuscita (codice ' + IntToStr(RC) + ').' + #13#10 +
           'Dettagli in C:\ProgramData\PersonalAIAssistant\installazione-condivisione.log', mbError, MB_OK);
  Result := FileExists('{#ShareUnc}\manifest.xml');
end;

{ Sperimentale e non documentato: segnala a Office che l'add-in e' gia' stato inserito. }
procedure PreInsertAddin(Host: String);
var
  Base, Dir: String;
begin
  Base := ExpandConstant('{localappdata}\Microsoft\Office\16.0\Wef\AddinInfo\1\filesystem\') + Host;
  Dir := Base + '\1\{#WordExcelAddinId}_{#WordExcelManifestVersion}';
  ForceDirectories(Dir);
  if not FileExists(Base + '\meta.json') then
    SaveStringToFile(Base + '\meta.json', '{"minorversion":1,"lastupdate":0}', False);
  SaveStringToFile(Dir + '\boot.json',
    '{"solid":"{#WordExcelAddinId}","storeid":"\\\\localhost\\PersonalAIAssistant$","appversion":"{#WordExcelManifestVersion}","equivalentaddins":[]}', False);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if (CurStep = ssPostInstall) and WizardIsTaskSelected('wordexcel') then
  begin
    if EnsureShare() and WizardIsTaskSelected('wordexcel\subito') then
    begin
      PreInsertAddin('Word');
      PreInsertAddin('Excel');
    end;
  end;
end;
