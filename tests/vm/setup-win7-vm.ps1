# setup-win7-vm.ps1 -- One-command Win7 SP1 x64 VM: create -> unattended install -> snapshot.
# Requires: VirtualBox 7.x installed, a Win7 SP1 x64 ISO. Keep ASCII.
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File tests\vm\setup-win7-vm.ps1 -IsoPath "D:\iso\cn_windows_7_ultimate_with_sp1_x64_dvd_u_677408.iso"
# After install (~20-40 min headless) it takes snapshot "clean-install" and prints how to run the smoke.
param(
  [string]$IsoPath = "",
  [string]$VmName = "Win7SP1x64",
  [string]$User = "OfficeAgent",
  [string]$Pass = "OA-smoke-2026",
  [int]$MemMB = 4096,
  [int]$Cpus = 2,
  [int]$DiskGB = 60
)
$ErrorActionPreference = "Continue"
$vbm = "C:\Program Files\Oracle\VirtualBox\VBoxManage.exe"
if (-not (Test-Path $vbm)) { Write-Host "VBoxManage not found"; exit 3 }
if ($IsoPath -eq "" -or -not (Test-Path $IsoPath)) { Write-Host "ISO not found: $IsoPath"; exit 3 }
Write-Host ("ISO=" + $IsoPath)

$gaIso = "C:\Program Files\Oracle\VirtualBox\VBoxGuestAdditions.iso"

# 1. create + register (remove leftovers from a previous failed run)
& $vbm unregistervm $VmName --delete 2>$null | Out-Null
& $vbm createvm --name $VmName --ostype "Windows7_64" --register | Out-Null
if ($LASTEXITCODE -ne 0) { Write-Host "createvm failed"; exit 2 }
& $vbm modifyvm $VmName --memory $MemMB --cpus $Cpus --vram 32 `
  --nic1 nat --boot1 dvd --boot2 disk --boot3 none --boot4 none `
  --clipboard-mode bidirectional --drag-and-drop bidirectional | Out-Null

# 2. disks: SATA (system) + IDE (install ISO + guest additions)
$disk = Join-Path $env:USERPROFILE ("VirtualBox VMs\" + $VmName + "\" + $VmName + ".vdi")
& $vbm createmedium disk --filename $disk --size $DiskGB --variant Standard | Out-Null
& $vbm storagectl $VmName --name "SATA" --add sata --controller IntelAHCI | Out-Null
& $vbm storageattach $VmName --storagectl "SATA" --port 0 --device 0 --type hdd --medium $disk | Out-Null
& $vbm storagectl $VmName --name "IDE" --add ide --controller PIIX4 | Out-Null
& $vbm storageattach $VmName --storagectl "IDE" --port 0 --device 0 --type dvddrive --medium $IsoPath | Out-Null
if (Test-Path $gaIso) {
  & $vbm storageattach $VmName --storagectl "IDE" --port 1 --device 0 --type dvddrive --medium $gaIso | Out-Null
}

# 3. unattended install (preconfigures user/password/locale/timezone + runs guest additions setup)
Write-Host "starting unattended install (headless, 20-40 min)..."
& $vbm unattended install $VmName --iso $IsoPath --user $User --password $Pass `
  --full-user-name "OfficeAgent Test" --locale zh_CN --country CN --time-zone "Asia/Shanghai" `
  --install-additions
if ($LASTEXITCODE -ne 0) { Write-Host "unattended config failed"; exit 2 }

# 4. start headless and wait for install to finish (poll guest logon until reachable)
& $vbm startvm $VmName --type headless | Out-Null
Write-Host "polling for guest availability (this takes 20-40 min)..."
$deadline = (Get-Date).AddMinutes(60)
$ready = $false
while ((Get-Date) -lt $deadline -and -not $ready) {
  Start-Sleep -Seconds 60
  $probe = & $vbm guestcontrol $VmName run --username $User --password $Pass --wait-stdout `
    --exe "C:\Windows\System32\cmd.exe" -- /c echo READY 2>$null
  if ("$probe" -like "*READY*") { $ready = $true; break }
  $st = & $vbm showvminfo $VmName --machinereadable 2>$null | Select-String "^VMState="
  Write-Host ("  " + (Get-Date -Format HH:mm) + " " + "$st".Trim())
}
if (-not $ready) { Write-Host "TIMEOUT waiting for guest (check GUI or VBoxManage showvminfo)"; exit 2 }
Write-Host "guest is up."

# 5. snapshot for instant rollback before every acceptance run
& $vbm snapshot $VmName take clean-install | Out-Null
Write-Host "snapshot taken: clean-install"
Write-Host ("NEXT: tests\vm\run-on-vm.ps1 -VmName `"$VmName`" -User $User -Pass $Pass -GuestRoot C:\OfficeAgent")
exit 0
