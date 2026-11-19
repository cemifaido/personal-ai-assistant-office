<#
  Personal AI Assistant - parte "per computer" (richiede amministratore, eseguita una sola volta).
  Crea una cartella in sola lettura con il manifest di Word/Excel e la condivide come
  \\localhost\PersonalAIAssistant$ (condivisione nascosta, sola lettura per tutti),
  perche' Office desktop accetta i cataloghi di componenti aggiuntivi solo da percorsi di rete.
  Codici di uscita: 0 ok, 2 servizio Server non attivo, 3 errore generico.
#>
param([Parameter(Mandatory = $true)][string]$ManifestSource)

$ErrorActionPreference = 'Stop'
$shareName = 'PersonalAIAssistant$'
$dir = Join-Path $env:ProgramData 'PersonalAIAssistant\Catalogo'
$log = Join-Path $env:ProgramData 'PersonalAIAssistant\installazione-condivisione.log'

function Write-Log($msg) { Add-Content -Path $log -Value ("{0} {1}" -f (Get-Date -Format s), $msg) }

try {
  New-Item $dir -ItemType Directory -Force | Out-Null
  Write-Log "avvio, utente elevato: $env:USERNAME"

  # Il servizio "Server" (LanmanServer) serve per qualunque condivisione, anche locale.
  $srv = Get-Service LanmanServer -ErrorAction SilentlyContinue
  if (-not $srv -or $srv.Status -ne 'Running') { Write-Log 'servizio Server non attivo'; exit 2 }

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
    New-SmbShare -Name $shareName -Path $dir -ReadAccess $everyone -Description 'Personal AI Assistant: catalogo componenti aggiuntivi (sola lettura)' | Out-Null
    Write-Log "condivisione creata: \\localhost\$shareName -> $dir"
  } else {
    Write-Log "condivisione gia' presente, manifest aggiornato"
  }
  exit 0
}
catch {
  Write-Log ("errore: " + $_.Exception.Message)
  exit 3
}
