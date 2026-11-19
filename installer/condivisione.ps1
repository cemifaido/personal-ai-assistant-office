<#
  Personal AI Assistant - parte "per computer" (richiede amministratore).
  Office desktop accetta i cataloghi di componenti aggiuntivi solo da percorsi di rete: questo
  script crea C:\ProgramData\PersonalAIAssistant\Catalogo (scrivibile solo da amministratori e
  SYSTEM) e la condivide in SOLA LETTURA come \\localhost\PersonalAIAssistant$ (condivisione nascosta).

  -Azione Crea    -ManifestSource <file>   crea o aggiorna cartella, manifest e condivisione
                  (oppure -ManifestGzB64 <manifest gzip+base64>, come fa l'installatore)
  -Azione Rimuovi                          rimuove condivisione e SOLO i propri file; le cartelle solo se vuote
  -Azione Verifica -ManifestGzB64 <...>    SENZA privilegi: 0 se condivisione, cartella, permessi e manifest
                                           sono conformi; 1 se non conformi o non verificabili

  Regole di sicurezza (lo script gira elevato):
  - non segue collegamenti (junction/symlink) e rifiuta cartelle create da utenti non amministratori;
  - imposta proprietario e permessi PRIMA di copiare il manifest; il manifest viene ricreato, non sovrascritto;
  - una condivisione con lo stesso nome viene toccata solo se punta alla nostra cartella e porta la nostra descrizione.

  Codici di uscita: 0 ok, 2 servizio "Server" non attivo, 3 errore (vedi log),
                    4 cartella esistente non sicura, 5 condivisione con lo stesso nome non nostra,
                    6 rimozione completata ma restano file non creati da Personal AI Assistant.
#>
param(
  [ValidateSet('Crea', 'Rimuovi', 'Verifica')][string]$Azione = 'Crea',
  [string]$ManifestSource,
  [string]$ManifestGzB64   # manifest compresso e incorporato (usato dall'installatore: nessun file da leggere)
)

$ErrorActionPreference = 'Stop'
$shareName = 'PersonalAIAssistant$'
$shareDescription = 'Personal AI Assistant: catalogo componenti aggiuntivi Office (sola lettura)'
$root = Join-Path $env:ProgramData 'PersonalAIAssistant'
$dir = Join-Path $root 'Catalogo'
$sidAdmins = New-Object System.Security.Principal.SecurityIdentifier 'S-1-5-32-544'
$sidSystem = New-Object System.Security.Principal.SecurityIdentifier 'S-1-5-18'
$sidUsers = New-Object System.Security.Principal.SecurityIdentifier 'S-1-5-32-545'
$sidEveryone = New-Object System.Security.Principal.SecurityIdentifier 'S-1-1-0'
$trustedOwners = @($sidAdmins.Value, $sidSystem.Value, 'S-1-5-80-956008885-3418522649-1831038044-1853292631-2271478464') # + TrustedInstaller
$script:logReady = $false

function Write-Log($msg) {
  if (-not $script:logReady) { return }   # si scrive il log solo dopo aver messo in sicurezza la cartella
  Add-Content -Path (Join-Path $root 'condivisione.log') -Value ("{0} [{1}] {2}" -f (Get-Date -Format s), $env:USERNAME, $msg)
}

# Una cartella esistente e' accettabile solo se non e' un collegamento e appartiene ad amministratori/SYSTEM.
function Test-SafeFolder($path) {
  if (-not (Test-Path -LiteralPath $path)) { return $true }
  $item = Get-Item -LiteralPath $path -Force
  if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { return $false }
  if (-not $item.PSIsContainer) { return $false }
  $owner = (Get-Acl -LiteralPath $path).GetOwner([System.Security.Principal.SecurityIdentifier]).Value
  return $trustedOwners -contains $owner
}

# Nessun permesso di scrittura per chi non e' amministratore/SYSTEM/TrustedInstaller/CREATOR OWNER.
function Test-NoForeignWrite($path) {
  $write = [System.Security.AccessControl.FileSystemRights]'WriteData, AppendData, WriteExtendedAttributes, WriteAttributes, Delete, DeleteSubdirectoriesAndFiles, ChangePermissions, TakeOwnership'
  $allowed = $trustedOwners + @('S-1-3-0')
  foreach ($ace in (Get-Acl -LiteralPath $path).GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier])) {
    if ($ace.AccessControlType -ne 'Allow') { continue }
    if (($ace.FileSystemRights -band $write) -and ($allowed -notcontains $ace.IdentityReference.Value)) { return $false }
  }
  return $true
}

function Get-ManifestBytes {
  if ($ManifestGzB64) {
    $gz = New-Object IO.Compression.GZipStream((New-Object IO.MemoryStream(,[Convert]::FromBase64String($ManifestGzB64))), [IO.Compression.CompressionMode]::Decompress)
    $ms = New-Object IO.MemoryStream; $gz.CopyTo($ms); return ,$ms.ToArray()
  }
  if ($ManifestSource -and (Test-Path -LiteralPath $ManifestSource -PathType Leaf)) { return ,[IO.File]::ReadAllBytes($ManifestSource) }
  return $null
}

function Set-SecureAcl($path) {
  $acl = New-Object System.Security.AccessControl.DirectorySecurity
  $acl.SetOwner($sidAdmins)
  $acl.SetAccessRuleProtection($true, $false)   # nessuna ereditarieta' dal padre
  $inh = [System.Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit'
  $none = [System.Security.AccessControl.PropagationFlags]::None
  foreach ($sid in $sidAdmins, $sidSystem) {
    $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($sid, 'FullControl', $inh, $none, 'Allow')))
  }
  $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($sidUsers, 'ReadAndExecute', $inh, $none, 'Allow')))
  Set-Acl -LiteralPath $path -AclObject $acl
}

# Descrizioni usate dalle versioni di prova (stessa cartella): riconosciute come nostre e aggiornate.
$legacyDescriptions = @('Personal AI Assistant: catalogo componenti aggiuntivi (sola lettura)')

function Get-OurShare {
  $s = Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue
  if (-not $s) { return $null }
  if ($s.Path -ne $dir) { return 'altra' }
  if ($legacyDescriptions -contains $s.Description) {
    Set-SmbShare -Name $shareName -Description $shareDescription -Force
    return Get-SmbShare -Name $shareName
  }
  if ($s.Description -ne $shareDescription) { return 'altra' }
  return $s
}

if ($Azione -eq 'Verifica') {
  try {
    $expected = Get-ManifestBytes
    if (-not $expected) { exit 1 }
    $share = Get-CimInstance -ClassName Win32_Share -Filter "Name='PersonalAIAssistant`$'" -ErrorAction Stop
    if (-not $share -or $share.Path -ne $dir) { exit 1 }
    foreach ($p in $root, $dir) { if (-not (Test-Path -LiteralPath $p) -or -not (Test-SafeFolder $p) -or -not (Test-NoForeignWrite $p)) { exit 1 } }
    $file = Join-Path $dir 'manifest.xml'
    if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { exit 1 }
    if ((Get-Item -LiteralPath $file -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) { exit 1 }
    if (-not (Test-NoForeignWrite $file)) { exit 1 }
    $actual = [IO.File]::ReadAllBytes($file)
    if ([Convert]::ToBase64String($actual) -ne [Convert]::ToBase64String($expected)) { exit 1 }
    exit 0
  } catch { exit 1 }
}

try {
  foreach ($p in $root, $dir) {
    if (-not (Test-SafeFolder $p)) { exit 4 }
  }

  if ($Azione -eq 'Rimuovi') {
    $s = Get-OurShare
    if ($s -eq 'altra') { exit 5 }
    if ($s) { Remove-SmbShare -Name $shareName -Force }
    # Solo i file creati da Personal AI Assistant; le cartelle si tolgono solo se restano vuote.
    foreach ($f in (Join-Path $dir 'manifest.xml'), (Join-Path $root 'condivisione.log'), (Join-Path $root 'installazione-condivisione.log')) {
      if (Test-Path -LiteralPath $f -PathType Leaf) { Remove-Item -LiteralPath $f -Force }
    }
    $foreign = $false
    foreach ($d in $dir, $root) {
      if (Test-Path -LiteralPath $d) {
        if (@(Get-ChildItem -LiteralPath $d -Force).Count -eq 0) { Remove-Item -LiteralPath $d -Force } else { $foreign = $true }
      }
    }
    if ($foreign) { exit 6 }
    exit 0
  }

  # Il manifest arriva incorporato nel comando (installatore) oppure da file (uso manuale).
  $manifestBytes = Get-ManifestBytes
  if (-not $manifestBytes) { exit 3 }
  try { [void][xml][Text.Encoding]::UTF8.GetString($manifestBytes) } catch { exit 3 }
  $existing = Get-OurShare
  if ($existing -eq 'altra') { exit 5 }

  $srv = Get-Service LanmanServer -ErrorAction SilentlyContinue
  if (-not $srv -or $srv.Status -ne 'Running') { exit 2 }

  # 1) cartelle con proprietario e permessi sicuri, PRIMA di qualunque scrittura del contenuto
  if (-not (Test-Path -LiteralPath $root)) { New-Item -Path $root -ItemType Directory | Out-Null }
  Set-SecureAcl $root
  $script:logReady = $true
  if (-not (Test-Path -LiteralPath $dir)) { New-Item -Path $dir -ItemType Directory | Out-Null }
  Set-SecureAcl $dir

  # 2) manifest ricreato (eredita i permessi della cartella; nessuna ACL esplicita preesistente)
  $target = Join-Path $dir 'manifest.xml'
  if (Test-Path -LiteralPath $target) { Remove-Item -LiteralPath $target -Force }
  [IO.File]::WriteAllBytes($target, $manifestBytes)

  # 3) condivisione: solo lettura per Everyone, nient'altro
  $everyone = $sidEveryone.Translate([System.Security.Principal.NTAccount]).Value
  if ($existing) {
    $access = @(Get-SmbShareAccess -Name $shareName)
    $ok = $access.Count -eq 1 -and $access[0].AccessControlType -eq 'Allow' -and $access[0].AccessRight -eq 'Read' -and
          (New-Object System.Security.Principal.NTAccount $access[0].AccountName).Translate([System.Security.Principal.SecurityIdentifier]).Value -eq $sidEveryone.Value
    if (-not $ok) {
      Remove-SmbShare -Name $shareName -Force
      $existing = $null
      Write-Log 'permessi della condivisione non conformi: condivisione ricreata'
    }
  }
  if (-not $existing) {
    New-SmbShare -Name $shareName -Path $dir -ReadAccess $everyone -Description $shareDescription | Out-Null
    Write-Log "condivisione creata: \\localhost\$shareName -> $dir"
  } else {
    Write-Log 'condivisione presente e conforme, manifest aggiornato'
  }
  exit 0
}
catch {
  Write-Log ('errore: ' + $_.Exception.Message)
  exit 3
}
