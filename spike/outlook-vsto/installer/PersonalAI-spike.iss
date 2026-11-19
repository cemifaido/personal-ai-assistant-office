; Personal AI Assistant - INSTALLATORE DI PROVA (spike, non per la distribuzione)
; Verifica: installazione per utente senza admin, sideload Word/Excel solo da registro,
; caricamento add-in VSTO Outlook, avvisi di fiducia, disinstallazione pulita.

#define AppVer "0.0.1"
#define PubKeyFile "..\installer-publickey.txt"
#define VstoDir "..\PaiOutlookSpike\bin\Release"
#define WordExcelManifest "..\..\..\manifest\manifest-word-excel.xml"
#define WordExcelAddinId "ecb39046-2d14-45ba-b5ea-94524ea141db"
; GUID fisso dell'elenco di inclusione VSTO (per poterlo rimuovere in disinstallazione)
#define InclusionId "{{3B2E7C41-8F0A-4C5D-9E61-7A2B4C9D0E11}"

[Setup]
AppId={{9C1F4A2E-5B7D-4E3A-A8C6-2D9E0F1B3C57}
AppName=Personal AI Assistant (prova tecnica)
AppVersion={#AppVer}
AppPublisher=Personal AI Assistant
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=
DefaultDirName={localappdata}\PersonalAIAssistant-spike
DisableDirPage=yes
DisableProgramGroupPage=yes
OutputDir=..\out
OutputBaseFilename=PersonalAI-spike-{#AppVer}-Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
CloseApplicationsFilter=*.exe
UninstallDisplayName=Personal AI Assistant (prova tecnica)

[Languages]
Name: "it"; MessagesFile: "compiler:Languages\Italian.isl"

[Tasks]
Name: "wordexcel"; Description: "Word ed Excel (registrazione solo da registro, senza catalogo)"
Name: "outlook"; Description: "Outlook classico (add-in di prova)"
Name: "outlook\fiducia"; Description: "Considera attendibile l'add-in Outlook (elenco di inclusione VSTO, evita la richiesta di conferma)"; Flags: unchecked

[Files]
Source: "{#WordExcelManifest}"; DestDir: "{app}\WordExcel"; DestName: "manifest.xml"; Tasks: wordexcel; Flags: ignoreversion
Source: "{#VstoDir}\PaiOutlookSpike.dll"; DestDir: "{app}\Outlook"; Tasks: outlook; Flags: ignoreversion
Source: "{#VstoDir}\PaiOutlookSpike.dll.manifest"; DestDir: "{app}\Outlook"; Tasks: outlook; Flags: ignoreversion
Source: "{#VstoDir}\PaiOutlookSpike.vsto"; DestDir: "{app}\Outlook"; Tasks: outlook; Flags: ignoreversion
Source: "{#VstoDir}\Microsoft.Office.Tools.Common.v4.0.Utilities.dll"; DestDir: "{app}\Outlook"; Tasks: outlook; Flags: ignoreversion
Source: "{#VstoDir}\Microsoft.Office.Tools.Outlook.v4.0.Utilities.dll"; DestDir: "{app}\Outlook"; Tasks: outlook; Flags: ignoreversion

[Registry]
; Word/Excel: SOLO il valore WEF\Developer (nome = Id add-in, dato = percorso manifest)
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\WEF\Developer"; ValueType: string; ValueName: "{#WordExcelAddinId}"; ValueData: "{app}\WordExcel\manifest.xml"; Tasks: wordexcel; Flags: uninsdeletevalue

; Outlook: registrazione add-in VSTO per l'utente corrente
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\PersonalAIAssistant.Spike"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Personal AI Assistant (prova)"; Tasks: outlook; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\PersonalAIAssistant.Spike"; ValueType: string; ValueName: "Description"; ValueData: "Prova tecnica di installazione"; Tasks: outlook
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\PersonalAIAssistant.Spike"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Tasks: outlook
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\PersonalAIAssistant.Spike"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:VstoUrl}|vstolocal"; Tasks: outlook

; Facoltativo: elenco di inclusione VSTO (fiducia per QUESTO manifest e QUESTA chiave pubblica)
Root: HKCU; Subkey: "Software\Microsoft\VSTO\Security\Inclusion\{#InclusionId}"; ValueType: string; ValueName: "Url"; ValueData: "{code:VstoUrl}"; Tasks: outlook\fiducia; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\VSTO\Security\Inclusion\{#InclusionId}"; ValueType: string; ValueName: "PublicKey"; ValueData: "{#FileRead(FileOpen(PubKeyFile))}"; Tasks: outlook\fiducia

[Code]
function VstoUrl(Param: String): String;
var
  P: String;
begin
  P := ExpandConstant('{app}\Outlook\PaiOutlookSpike.vsto');
  StringChangeEx(P, '\', '/', True);
  Result := 'file:///' + P;
end;
