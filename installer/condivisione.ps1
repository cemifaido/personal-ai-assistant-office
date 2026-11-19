<#
  Personal AI Assistant - parte "per computer" (richiede amministratore).
  Office desktop accetta i cataloghi di componenti aggiuntivi solo da percorsi di rete: questo
  script crea C:\ProgramData\PersonalAIAssistant\Catalogo (scrivibile solo dagli amministratori)
  e la condivide in SOLA LETTURA come \\localhost\PersonalAIAssistant$ (condivisione nascosta).

  -Azione Crea    -ManifestSource <file>   crea o aggiorna cartella, manifest e condivisione
  -Azione Rimuovi                          rimuove condivisione e cartella

  Codici di uscita: 0 ok, 2 servizio "Server" non attivo, 3 errore (vedi log).
#>
param(
  [ValidateSet('Crea', 'Rimuovi')][string]$Azione = 'Crea',
  [string]$ManifestSource
)

$ErrorActionPreference = 'Stop'
$shareName = 'PersonalAIAssistant$'
$root = Join-Path $env:ProgramData 'PersonalAIAssistant'
$dir = Join-Path $root 'Catalogo'
$log = Join-Path $root 'condivisione.log'

function Write-Log($msg) {
  New-Item $root -ItemType Directory -Force | Out-Null
  Add-Content -Path $log -Value ("{0} [{1}] {2}" -f (Get-Date -Format s), $env:USERNAME, $msg)
}

try {
  if ($Azione -eq 'Rimuovi') {
    if (Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue) { Remove-SmbShare -Name $shareName -Force; Write-Log 'condivisione rimossa' }
    if (Test-Path $dir) { Remove-Item $dir -Recurse -Force; Write-Log 'cartella catalogo rimossa' }
    exit 0
  }

  if (-not $ManifestSource -or -not (Test-Path -LiteralPath $ManifestSource)) { Write-Log "manifest non trovato: $ManifestSource"; exit 3 }

  $srv = Get-Service LanmanServer -ErrorAction SilentlyContinue
  if (-not $srv -or $srv.Status -ne 'Running') { Write-Log 'servizio Server (LanmanServer) non attivo'; exit 2 }

  New-Item $dir -ItemType Directory -Force | Out-Null
  Copy-Item -LiteralPath $ManifestSource -Destination (Join-Path $dir 'manifest.xml') -Force

  # NTFS: amministratori e SYSTEM controllo completo, utenti solo lettura (nessuno puo' sostituire il manifest).
  $acl = New-Object System.Security.AccessControl.DirectorySecurity
  $acl.SetAccessRuleProtection($true, $false)
  $inh = [System.Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit'
  $none = [System.Security.AccessControl.PropagationFlags]::None
  foreach ($sid in 'S-1-5-32-544', 'S-1-5-18') {
    $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule((New-Object System.Security.Principal.SecurityIdentifier $sid), 'FullControl', $inh, $none, 'Allow')))
  }
  $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule((New-Object System.Security.Principal.SecurityIdentifier 'S-1-5-32-545'), 'ReadAndExecute', $inh, $none, 'Allow')))
  Set-Acl -LiteralPath $dir -AclObject $acl

  # Condivisione in sola lettura per "Everyone" (nome localizzato ricavato dal SID S-1-1-0).
  $everyone = (New-Object System.Security.Principal.SecurityIdentifier 'S-1-1-0').Translate([System.Security.Principal.NTAccount]).Value
  $existing = Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue
  if ($existing -and $existing.Path -ne $dir) { Remove-SmbShare -Name $shareName -Force; $existing = $null }
  if (-not $existing) {
    New-SmbShare -Name $shareName -Path $dir -ReadAccess $everyone -Description 'Personal AI Assistant: catalogo componenti aggiuntivi Office (sola lettura)' | Out-Null
    Write-Log "condivisione creata: \\localhost\$shareName -> $dir"
  } else {
    Write-Log 'condivisione gia presente, manifest aggiornato'
  }
  exit 0
}
catch {
  Write-Log ('errore: ' + $_.Exception.Message)
  exit 3
}
