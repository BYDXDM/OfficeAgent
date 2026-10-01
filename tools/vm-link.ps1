# vm-link.ps1 -- Map drive X: to <repo>\build\vm (ASCII path) so VBoxManage never
# sees non-ASCII characters. Repo root derived from script location.
# Also power off the VM before storage changes.
$repo = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$vmDir = Join-Path (Join-Path $repo "build\vm") ""
$vbm = "C:\Program Files\Oracle\VirtualBox\VBoxManage.exe"
& $vbm controlvm Win7SP1x64 poweroff 2>$null | Out-Null
Start-Sleep -Seconds 5
$subst = "X:"
$existing = (subst)
if ("$existing" -match "^X:") { subst X: /D 2>$null | Out-Null }
subst $subst $vmDir
if ($LASTEXITCODE -ne 0) { Write-Host "subst failed"; exit 2 }
Write-Host ("X: mapped to " + $vmDir)
& $vbm storageattach Win7SP1x64 --storagectl SATA --port 0 --device 0 --type hdd --medium "X:\win7.raw"
if ($LASTEXITCODE -ne 0) { Write-Host "storageattach failed"; exit 2 }
Write-Host "raw disk attached"
& $vbm startvm Win7SP1x64 --type headless | Out-Null
Write-Host "VM started headless"
exit 0
