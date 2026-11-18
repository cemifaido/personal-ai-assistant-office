<#
  Personal AI Assistant - installazione / pulizia (solo utente corrente, niente admin)

  Esempi:
    # 1) Rimuove la vecchia versione con server locale (manifest Outlook, catalogo, certificato localhost)
    .\installa.ps1 -PulisciVecchiaVersione

    # 2) Registra l'add-in Word/Excel che punta ai file pubblicati (es. GitHub Pages)
    .\installa.ps1 -BaseUrl https://TUO-UTENTE.github.io/personal-ai-assistant

    # 3) Copia le chiavi dal vecchio file .env nelle impostazioni della macro Outlook
    .\installa.ps1 -ImportaChiaviDaEnv

    # Disinstallazione completa (add-in + impostazioni della macro)
    .\installa.ps1 -Disinstalla

  Le opzioni si possono combinare. Chiudere Word, Excel e Outlook prima di eseguire.
#>
param(
  [string]$BaseUrl,
  [switch]$PulisciVecchiaVersione,
  [switch]$ImportaChiaviDaEnv,
  [switch]$Disinstalla
)

$ErrorActionPreference = 'Stop'
$root      = $PSScriptRoot
$template  = Join-Path $root 'manifest-word-excel.template.xml'
$manifest  = Join-Path $root 'manifest-word-excel.xml'
$addinId   = 'ecb39046-2d14-45ba-b5ea-94524ea141db'   # Id del manifest Word/Excel
$oldOutlookId = '4306873d-fb66-4fe3-87ac-9de2efc7586d' # vecchio add-in Outlook web (non usabile su POP)
$wefDev    = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
$wefCat    = 'HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs'
$wefCache  = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\Wef'
$vbaKey    = 'HKCU:\Software\VB and VBA Program Settings\PersonalAIAssistant\Config'

if (-not ($BaseUrl -or $PulisciVecchiaVersione -or $ImportaChiaviDaEnv -or $Disinstalla)) {
  Get-Help $MyInvocation.MyCommand.Path -Detailed
  exit 0
}

function Step($msg) { Write-Host "-> $msg" -ForegroundColor Cyan }
function Ok($msg)   { Write-Host "   OK  $msg" -ForegroundColor Green }
function Skip($msg) { Write-Host "   --  $msg" -ForegroundColor DarkGray }

$running = Get-Process WINWORD, EXCEL, OUTLOOK -ErrorAction SilentlyContinue
if ($running) {
  Write-Warning ("Chiudi prima: " + (($running | Select-Object -ExpandProperty ProcessName -Unique) -join ', '))
  exit 1
}

function Clear-WefCache {
  if (Test-Path $wefCache) { Remove-Item $wefCache -Recurse -Force; Ok "cache add-in di Office svuotata" }
  else { Skip "cache add-in gia' vuota" }
}

# ---------------------------------------------------------------
if ($PulisciVecchiaVersione -or $Disinstalla) {
  Step "Rimozione della vecchia versione (server locale)"

  if (Test-Path "$wefDev\$oldOutlookId") { Remove-Item "$wefDev\$oldOutlookId" -Recurse -Force; Ok "manifest Outlook web rimosso" }
  else { Skip "manifest Outlook web non registrato" }

  if ((Get-ItemProperty $wefDev -ErrorAction SilentlyContinue).OutlookSideloadManifestPath) {
    Remove-ItemProperty $wefDev -Name OutlookSideloadManifestPath -Force; Ok "valore OutlookSideloadManifestPath rimosso"
  } else { Skip "OutlookSideloadManifestPath assente" }

  if (Test-Path $wefCat) {
    $folderName = Split-Path $root -Leaf
    Get-ChildItem $wefCat | ForEach-Object {
      $url = (Get-ItemProperty $_.PSPath).Url
      if ($url -and $url -like "*$folderName*") { Remove-Item $_.PSPath -Recurse -Force; Ok "catalogo attendibile rimosso: $url" }
    }
  }

  $crt = @((Join-Path $root 'certs\localhost.crt'), (Join-Path $root '_obsoleto\certs\localhost.crt')) |
         Where-Object { Test-Path $_ } | Select-Object -First 1
  if ($crt) {
    $thumb = (New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $crt).Thumbprint
    foreach ($store in 'Cert:\CurrentUser\Root', 'Cert:\CurrentUser\My') {
      $c = Get-ChildItem $store | Where-Object Thumbprint -eq $thumb
      if ($c) {
        # Per lo store Root Windows mostra una finestra di conferma: rispondere Si'.
        $c | Remove-Item -Force; Ok "certificato localhost rimosso da $store"
      } else { Skip "certificato localhost non presente in $store" }
    }
  } else { Skip "file del certificato del progetto non trovato: nessun certificato toccato" }
}

# ---------------------------------------------------------------
if ($Disinstalla) {
  Step "Disinstallazione add-in Word/Excel e impostazioni macro"
  if (Test-Path "$wefDev\$addinId") { Remove-Item "$wefDev\$addinId" -Recurse -Force; Ok "add-in Word/Excel rimosso" }
  else { Skip "add-in Word/Excel non registrato" }
  if (Test-Path $vbaKey) { Remove-Item $vbaKey -Recurse -Force; Ok "impostazioni e chiavi della macro rimosse" }
  else { Skip "nessuna impostazione della macro" }
  Clear-WefCache
  Write-Host "`nRicorda di eliminare il modulo VBA da Outlook (Alt+F11) se non serve piu'."
}

# ---------------------------------------------------------------
if ($BaseUrl) {
  Step "Registrazione add-in Word/Excel su $BaseUrl"
  $BaseUrl = $BaseUrl.TrimEnd('/')
  if ($BaseUrl -notmatch '^https://') { throw "BaseUrl deve iniziare con https:// (Office non carica pagine http)." }

  try {
    $r = Invoke-WebRequest "$BaseUrl/index.html" -UseBasicParsing -TimeoutSec 20
    if ($r.Content -notmatch 'ai\.js') { Write-Warning "index.html raggiungibile ma non sembra la versione nuova (manca ai.js)." }
    else { Ok "index.html raggiungibile" }
  } catch { Write-Warning "Impossibile scaricare $BaseUrl/index.html : $($_.Exception.Message). Controlla la pubblicazione." }

  (Get-Content $template -Raw).Replace('__BASE_URL__', $BaseUrl) | Set-Content $manifest -Encoding UTF8
  Ok "creato manifest-word-excel.xml"

  New-Item "$wefDev\$addinId" -Force | Out-Null
  Set-ItemProperty "$wefDev\$addinId" -Name ManifestPath -Value $manifest
  Ok "manifest registrato per l'utente corrente"
  Clear-WefCache
  Write-Host "`nApri Word o Excel: compare la scheda 'Personal AI' (se no: Inserisci > Componenti aggiuntivi > Componenti aggiuntivi per sviluppatori)."
}

# ---------------------------------------------------------------
if ($ImportaChiaviDaEnv) {
  Step "Importazione chiavi da .env nelle impostazioni della macro Outlook"
  $envFile = @((Join-Path $root '.env'), (Join-Path $root '_obsoleto\.env')) | Where-Object { Test-Path $_ } | Select-Object -First 1
  if (-not $envFile) { throw "File .env non trovato." }
  $map = @{ GEMINI_API_KEY = 'gemini'; OPENAI_API_KEY = 'openai'; ANTHROPIC_API_KEY = 'claude' }
  New-Item $vbaKey -Force | Out-Null
  foreach ($line in Get-Content $envFile) {
    if ($line -match '^\s*([A-Z_]+)\s*=\s*(.+?)\s*$' -and $map.ContainsKey($Matches[1]) -and $Matches[2].Length -gt 5) {
      Set-ItemProperty $vbaKey -Name ("key_" + $map[$Matches[1]]) -Value $Matches[2]
      Ok ("chiave " + $map[$Matches[1]] + " importata")
    }
  }
  if (-not (Get-ItemProperty $vbaKey -ErrorAction SilentlyContinue).provider) {
    Set-ItemProperty $vbaKey -Name provider -Value 'gemini'
  }
  Write-Host "`nPer il pannello Word/Excel le chiavi vanno inserite una volta dall'icona impostazioni (le due app non condividono lo stesso archivio)."
}
