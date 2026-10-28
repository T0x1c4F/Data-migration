param(
 [Parameter(Mandatory=$true)][string]$Package,
 [Parameter(Mandatory=$true)][string]$AppFolder,
 [string]$SourceFolder='',
 [ValidateRange(0,600)][int]$WaitSeconds=120
)
$ErrorActionPreference='Stop'
$appRoot=(Resolve-Path -LiteralPath $AppFolder).Path
$packagePath=(Resolve-Path -LiteralPath $Package).Path
$target=Join-Path $appRoot 'Contract Desk.exe'
if ($packagePath -eq $target) { throw 'Choose the new package, not the installed executable.' }
$expectedHash=(Get-FileHash -LiteralPath $packagePath).Hash
$oldNames=@('Contract Desk updated.exe','Contract Desk website update.exe','Contract Desk rounding update.exe')
$launchers=@($target)+@($oldNames | ForEach-Object { Join-Path $appRoot $_ })
function Test-AppClosed {
 foreach ($path in $launchers) {
  if (!(Test-Path -LiteralPath $path)) { continue }
  try {
   $stream=[System.IO.File]::Open($path,[System.IO.FileMode]::Open,[System.IO.FileAccess]::Read,[System.IO.FileShare]::None)
   $stream.Dispose()
  } catch [System.IO.IOException] { return $false }
 }
 return $true
}
$deadline=[DateTime]::UtcNow.AddSeconds($WaitSeconds)
$announced=$false
while (!(Test-AppClosed)) {
 if (!$announced) { Write-Output 'Close Contract Desk to finish updating the same app. Your work and history stay in this folder.'; $announced=$true }
 if ([DateTime]::UtcNow -ge $deadline) { throw 'The app is still open. Close it, then retry this update. The installed app was not replaced.' }
 Start-Sleep -Seconds 1
}
$backup=Join-Path $appRoot ('backups\before-update-'+(Get-Date -Format 'yyyyMMdd-HHmmss-fff'))
New-Item -ItemType Directory -Path $backup | Out-Null
$pending=Join-Path $backup 'new-package.pending'
Copy-Item -LiteralPath $packagePath -Destination $pending
if ((Get-FileHash -LiteralPath $pending).Hash -ne $expectedHash) { throw 'The new package did not copy correctly.' }
if (Test-Path -LiteralPath $target) {
 # Atomic replacement never deletes the installed executable before the new one is ready.
 [System.IO.File]::Replace($pending,$target,(Join-Path $backup 'Contract Desk.exe'))
} else { [System.IO.File]::Move($pending,$target) }
if ((Get-FileHash -LiteralPath $target).Hash -ne $expectedHash) { throw 'Installed package verification failed.' }
foreach ($name in $oldNames) {
 $old=Join-Path $appRoot $name
 if (Test-Path -LiteralPath $old) { Move-Item -LiteralPath $old -Destination (Join-Path $backup $name) }
}
if ($SourceFolder) {
 $sourceRoot=(Resolve-Path -LiteralPath $SourceFolder).Path
 $installedSource=Join-Path $appRoot 'source\ContractDesk'
 New-Item -ItemType Directory -Path $installedSource -Force | Out-Null
 $sourceBackup=Join-Path $backup 'source'
 New-Item -ItemType Directory -Path $sourceBackup | Out-Null
 foreach ($file in (Get-ChildItem -LiteralPath $sourceRoot -File | Where-Object { $_.Extension -in @('.py','.txt','.ps1') })) {
  $destination=Join-Path $installedSource $file.Name
  if (Test-Path -LiteralPath $destination) { Copy-Item -LiteralPath $destination -Destination $sourceBackup }
  Copy-Item -LiteralPath $file.FullName -Destination $destination -Force
 }
 $readme=Join-Path $sourceRoot 'READ ME.txt'
 if (Test-Path -LiteralPath $readme) { Copy-Item -LiteralPath $readme -Destination (Join-Path $appRoot 'READ ME.txt') -Force }
}
Write-Output "Updated: $target"
Write-Output "Backup: $backup"
Write-Output 'Always open Contract Desk.exe. No alternate launcher is created.'
