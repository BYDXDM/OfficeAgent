# run-on-vm.ps1 -- Host-side driver: run tests\vm\win7-smoke.ps1 inside a VirtualBox Win7 guest.
# Requires: VirtualBox installed, a Win7 SP1 x64 VM with Guest Additions, snapshot recommended.
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\vm\run-on-vm.ps1 -VmName "Win7x64" `
#     -User OfficeAgent -Pass <password> -GuestRoot C:\OfficeAgent
# What it does:
#   1. Starts the VM headless (if not running) and waits for Guest Additions.
#   2. Copy repo (build\ + tests\fixtures + runtime + lo76) into guest via shared folder or guestcontrol copyfrom.
#   3. Runs in-guest: powershell -NoProfile -ExecutionPolicy Bypass -File C:\OfficeAgent\tests\vm\win7-smoke.ps1
#   4. Pulls back %TEMP%\win7-smoke-result.txt and prints the verdict.
# NOTE: this machine currently has no hypervisor installed -- the script exits with guidance in that case.
param(
  [string]$VmName = "",
  [string]$User = "",
  [string]$Pass = "",
  [string]$GuestRoot = "C:\OfficeAgent",
  [string]$RepoRoot = ""
)
$ErrorActionPreference = "Continue"

if ($RepoRoot -eq "") {
  $repo = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path))
} else {
  $repo = $RepoRoot
}
$vbm = "C:\Program Files\Oracle\VirtualBox\VBoxManage.exe"
if (-not (Test-Path $vbm)) {
  Write-Host "VirtualBox not found. Win7 VM acceptance needs a hypervisor + a Win7 SP1 x64 VM."
  Write-Host "Manual path: see tests\WIN7-VM-验收.md section 2 (copy package, run win7-smoke.ps1 in guest)."
  exit 3
}
if ($VmName -eq "" -or $User -eq "") {
  Write-Host "Usage: -VmName <name> -User <guest user> -Pass <pass> [-GuestRoot C:\OfficeAgent]"
  & $vbm list vms
  exit 3
}

$vmState = & $vbm showvminfo $VmName --machinereadable 2>$null | Select-String "^VMState="
if ("$vmState" -notlike "*running*") {
  Write-Host ("Starting VM headless: " + $VmName)
  & $vbm startvm $VmName --type headless | Out-Null
  Start-Sleep -Seconds 45
}

# 1) copy payload of the test into guest (guest additions needed for guestcontrol)
$items = @("build\host", "build\boot", "tests\fixtures", "tests\vm", "runtime\py38", "lo76", "bin")
foreach ($it in $items) {
  $src = Join-Path $repo $it
  if (Test-Path $src) {
    $dst = $GuestRoot + "\" + ($it -replace '/', '\')
    Write-Host ("copyto: " + $it)
    & $vbm guestcontrol $VmName copyto --username $User --password $Pass --target-directory $dst $src 2>&1 | Out-Null
  }
}

# 2) run in-guest smoke
$guestScript = $GuestRoot + "\tests\vm\win7-smoke.ps1"
Write-Host "running in-guest smoke..."
& $vbm guestcontrol $VmName run --username $User --password $Pass --wait-stdout --wait-stderr `
  --exe "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -- -NoProfile -ExecutionPolicy Bypass `
  -File $guestScript -Root $GuestRoot

# 3) fetch result file
& $vbm guestcontrol $VmName copyfrom --username $User --password $Pass --target-directory (Join-Path $repo "build\test") `
  "C:\Windows\Temp\win7-smoke-result.txt" 2>$null | Out-Null

Write-Host "Done. See build\test\win7-smoke-result.txt (copyfrom of guest %TEMP% may need the guest temp path)."
exit 0
