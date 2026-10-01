# vm-attach.ps1 -- Attach the partitioned raw disk (and fall back to a VMDK wrapper)
# to the VM, using properly-encoded PowerShell paths (never bash-passed args).
# Keep ASCII; repo root derived from script location.
$repo = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$raw = Join-Path (Join-Path $repo "build\vm") "win7.raw"
$vmdk = Join-Path (Join-Path $repo "build\vm") "win7.vmdk"
$vbm = "C:\Program Files\Oracle\VirtualBox\VBoxManage.exe"
$vm = "Win7SP1x64"

& $vbm controlvm $vm poweroff 2>$null | Out-Null
Start-Sleep -Seconds 5

if (-not (Test-Path $vmdk)) {
  $sectors = [UInt64]64424509440 / [UInt64]512
  $cyl = [UInt64]($sectors / (16 * 63))
  $desc = "# Disk DescriptorFile" + [Environment]::NewLine +
          "version=1" + [Environment]::NewLine +
          "CID=fffffffe" + [Environment]::NewLine +
          "parentCID=ffffffff" + [Environment]::NewLine +
          'createType="fullDevice"' + [Environment]::NewLine +
          [Environment]::NewLine +
          "# Extent description" + [Environment]::NewLine +
          "RW $sectors FLAT `"win7.raw`" 0" + [Environment]::NewLine +
          [Environment]::NewLine +
          "# The Disk Data Base" + [Environment]::NewLine +
          'ddb.virtualHWVersion = "4"' + [Environment]::NewLine +
          "ddb.geometry.cylinders = `"$cyl`"" + [Environment]::NewLine +
          'ddb.geometry.heads = "16"' + [Environment]::NewLine +
          'ddb.geometry.sectors = "63"' + [Environment]::NewLine +
          'ddb.adapterType = "lsilogic"' + [Environment]::NewLine
  [IO.File]::WriteAllText($vmdk, $desc, [Text.Encoding]::ASCII)
}

$err1 = & $vbm storageattach $vm --storagectl SATA --port 0 --device 0 --type hdd --medium $raw 2>&1
if ($LASTEXITCODE -eq 0) { Write-Host "attached RAW"; }
else {
  Write-Host ("raw attach failed, trying vmdk wrapper: " + ($err1 | Select-Object -First 1))
  $err2 = & $vbm storageattach $vm --storagectl SATA --port 0 --device 0 --type hdd --medium $vmdk 2>&1
  if ($LASTEXITCODE -eq 0) { Write-Host "attached VMDK wrapper" }
  else { Write-Host ("vmdk attach failed too: " + ($err2 | Select-Object -First 1)); exit 2 }
}
& $vbm startvm $vm --type headless | Out-Null
Write-Host "VM started headless"
exit 0
