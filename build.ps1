<#
  Personal AI Assistant - build della release Windows.

    powershell -ExecutionPolicy Bypass -File .\build.ps1 [-Versione 0.1.0] [-BaseUrl https://...]

  Passi: compila l'add-in Outlook (MSBuild + VSTO, manifest firmati con il certificato
  "CN=Personal AI Assistant" dell'utente corrente), genera il manifest Word/Excel dal modello,
  estrae la chiave pubblica per l'elenco di inclusione VSTO, compila il Setup con Inno Setup 6
  e scrive l'impronta SHA-256 in release\SHA256SUMS.txt.
#>
param(
  [string]$Versione = '0.1.0',
  # Unico punto con gli indirizzi del progetto (sito GitHub Pages e repository).
  [string]$BaseUrl = 'https://pai-office-automation.github.io/personal-ai-assistant-office',
  [string]$RepoUrl = 'https://github.com/pai-office-automation/personal-ai-assistant-office'
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$build = Join-Path $root 'build'
$release = Join-Path $root 'release'
$proj = Join-Path $root 'outlook\PersonalAIAssistant.Outlook\PersonalAIAssistant.Outlook.csproj'

function Step($m) { Write-Host "-> $m" -ForegroundColor Cyan }

# Strumenti
$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
$vs = & $vswhere -latest -requires Microsoft.VisualStudio.Component.Sharepoint.Tools -property installationPath
if (-not $vs) { throw 'Visual Studio con "Sviluppo di Microsoft 365" (VSTO) non trovato.' }
$msbuild = Join-Path $vs 'MSBuild\Current\Bin\MSBuild.exe'
$vsMajor = (& $vswhere -latest -requires Microsoft.VisualStudio.Component.Sharepoint.Tools -property installationVersion).Split('.')[0]
$iscc = "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe"
if (-not (Test-Path $iscc)) { throw 'Inno Setup 6 non trovato.' }

# Certificato di firma dei manifest VSTO (deve coincidere con il thumbprint nel .csproj)
$cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -eq 'CN=Personal AI Assistant' -and $_.HasPrivateKey } | Select-Object -First 1
if (-not $cert) { throw 'Certificato "CN=Personal AI Assistant" non trovato in CurrentUser\My: importare il .pfx di firma.' }
$thumbInProj = ([xml](Get-Content $proj)).Project.PropertyGroup.ManifestCertificateThumbprint | Where-Object { $_ } | Select-Object -First 1
if ($thumbInProj -ne $cert.Thumbprint) { throw "Thumbprint nel progetto ($thumbInProj) diverso dal certificato ($($cert.Thumbprint))." }

Remove-Item $build -Recurse -Force -ErrorAction SilentlyContinue
New-Item "$build\outlook" -ItemType Directory -Force | Out-Null
New-Item $release -ItemType Directory -Force | Out-Null

Step "Compilazione add-in Outlook"
& $msbuild $proj /t:Rebuild /p:Configuration=Release "/p:VisualStudioVersion=$vsMajor.0" /nologo /v:minimal
if ($LASTEXITCODE -ne 0) { throw 'Compilazione add-in Outlook non riuscita.' }

# La compilazione VSTO registra l'add-in per chi compila, puntando alla cartella bin del repository
# e con lo stesso nome usato dal Setup. Si annulla: Outlook deve usare la copia installata, se c'e'.
$addinKey = 'HKCU:\Software\Microsoft\Office\Outlook\Addins\PersonalAIAssistant.Outlook'
$installed = Join-Path $env:LOCALAPPDATA 'Programs\PersonalAIAssistant\Outlook\PersonalAIAssistant.Outlook.vsto'
$m = (Get-ItemProperty $addinKey -ErrorAction SilentlyContinue).Manifest
if ($m -and $m -like "*$($root.Replace('\','/'))*") {
  if (Test-Path $installed) { Set-ItemProperty $addinKey -Name Manifest -Value ('file:///' + $installed.Replace('\','/') + '|vstolocal'); Write-Host '   registrazione Outlook riportata alla copia installata' }
  else { Remove-Item $addinKey -Recurse -Force; Write-Host '   registrazione di sviluppo di Outlook rimossa' }
}
Get-ChildItem 'HKCU:\Software\Microsoft\VSTO\Security\Inclusion' -ErrorAction SilentlyContinue | Where-Object {
  (Get-ItemProperty $_.PSPath).Url -like "*$($root.Replace('\','/'))*" } | Remove-Item -Recurse -Force

$bin = Join-Path (Split-Path $proj) 'bin\Release'
Copy-Item "$bin\PersonalAIAssistant.Outlook.dll", "$bin\PersonalAIAssistant.Outlook.dll.manifest", "$bin\PersonalAIAssistant.Outlook.vsto",
          "$bin\Microsoft.Office.Tools.Common.v4.0.Utilities.dll", "$bin\Microsoft.Office.Tools.Outlook.v4.0.Utilities.dll" "$build\outlook\"

Step "Manifest Word/Excel per $BaseUrl"
$manifest = (Get-Content (Join-Path $root 'manifest-word-excel.template.xml') -Raw).Replace('__BASE_URL__', $BaseUrl.TrimEnd('/'))
if ($manifest -match '__BASE_URL__') { throw 'Segnaposto non sostituito nel manifest.' }
[xml]$check = $manifest
$manVer = $check.OfficeApp.Version
if (-not $manVer.StartsWith($Versione)) { throw "Versione manifest ($manVer) diversa da $Versione." }
[IO.File]::WriteAllText("$build\manifest.xml", $manifest, (New-Object Text.UTF8Encoding $false))

Step "Comandi elevati per la condivisione (script e manifest incorporati, nessun file da leggere)"
# Lo script eseguito come amministratore non deve leggere file da cartelle scrivibili dall'utente:
# codice e manifest (compresso) viaggiano dentro -EncodedCommand, generato qui e incorporato nel Setup.
$shareScript = Get-Content (Join-Path $root 'installer\condivisione.ps1') -Raw
$shareScript = [regex]::Replace($shareScript, '(?s)<#.*?#>', '')                 # commento di testa
$shareScript = ($shareScript -split "`n" | Where-Object { $_ -notmatch '^\s*#' -and $_.Trim() }) -join "`n"
$raw = [IO.File]::ReadAllBytes("$build\manifest.xml")
$msOut = New-Object IO.MemoryStream
$gzw = New-Object IO.Compression.GZipStream($msOut, [IO.Compression.CompressionMode]::Compress)
$gzw.Write($raw, 0, $raw.Length); $gzw.Close()
$manifestGz = [Convert]::ToBase64String($msOut.ToArray())
foreach ($a in @(@('crea', "-Azione Crea -ManifestGzB64 '$manifestGz'"), @('rimuovi', '-Azione Rimuovi'),
                 @('verifica', "-Azione Verifica -ManifestGzB64 '$manifestGz'"))) {
  $cmd = "& {`n$shareScript`n} $($a[1])"
  $enc = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($cmd))
  if ($enc.Length -gt 30000) { throw "Comando elevato troppo lungo ($($enc.Length) caratteri)." }
  [IO.File]::WriteAllText("$build\enc-$($a[0]).txt", $enc, [Text.Encoding]::ASCII)
  Write-Host "   $($a[0]): $($enc.Length) caratteri"
}

Step "Chiave pubblica per l'elenco di inclusione VSTO"
$pub = $cert.PublicKey.Key.ToXmlString($false)
[IO.File]::WriteAllText("$build\publickey.txt", $pub, [Text.Encoding]::ASCII)

Step "Installatore"
& $iscc /Q "/DAppVer=$Versione" "/DBuildDir=$build" "/DSiteUrl=$($BaseUrl.TrimEnd('/'))" "/DRepoUrl=$($RepoUrl.TrimEnd('/'))" (Join-Path $root 'installer\PersonalAIAssistant.iss')
if ($LASTEXITCODE -ne 0) { throw 'Compilazione installatore non riuscita.' }

$setup = Join-Path $release "PersonalAIAssistant-$Versione-Setup.exe"
$hash = (Get-FileHash $setup -Algorithm SHA256).Hash
"$hash  $(Split-Path $setup -Leaf)" | Set-Content (Join-Path $release 'SHA256SUMS.txt') -Encoding ascii
Write-Host "`nOK: $setup" -ForegroundColor Green
Write-Host "SHA-256: $hash"
