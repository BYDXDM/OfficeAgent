# vm-copy-kit3.ps1 -- Create missing dir levels individually, then re-copy failed files.
$repo = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$vbm = "C:\Program Files\Oracle\VirtualBox\VBoxManage.exe"
$vm = "Win7SP1x64"
$u = "OfficeAgent"
$p = "OA-smoke-2026"

$dirs = @(
  "C:\OfficeAgent\build",
  "C:\OfficeAgent\build\host",
  "C:\OfficeAgent\build\boot",
  "C:\OfficeAgent\tests\fixtures\recon",
  "C:\OfficeAgent\tests\vm"
)
foreach ($d in $dirs) {
  & $vbm guestcontrol $vm mkdir --username $u --password $p $d 2>$null | Out-Null
}
Start-Sleep -Seconds 2

$files = @(
  @("build\host\OfficeAgent.exe",  "C:\OfficeAgent\build\host\OfficeAgent.exe"),
  @("build\host\pdfium.dll",       "C:\OfficeAgent\build\host\pdfium.dll"),
  @("build\boot\OfficeAgentBoot.exe", "C:\OfficeAgent\build\boot\OfficeAgentBoot.exe"),
  @("tests\fixtures\recon\flow.csv",   "C:\OfficeAgent\tests\fixtures\recon\flow.csv"),
  @("tests\fixtures\recon\ledger.csv", "C:\OfficeAgent\tests\fixtures\recon\ledger.csv"),
  @("tests\vm\win7-smoke.ps1",     "C:\OfficeAgent\tests\vm\win7-smoke.ps1")
)
foreach ($f in $files) {
  $src = Join-Path $repo $f[0]
  $ok = $false
  for ($i = 0; $i -lt 3 -and -not $ok; $i++) {
    & $vbm guestcontrol $vm copyto --username $u --password $p --target-directory ((Split-Path -Parent $f[1]) + "\") $src 2>$null | Out-Null
    Start-Sleep -Seconds 1
    $chk = & $vbm guestcontrol $vm stat --username $u --password $p $f[1] 2>$null
    if ("$chk" -notlike "*VBOX_E*" -and "$chk" -notlike "*error*") { $ok = $true }
  }
  Write-Host ((Get-Date -Format HH:mm:ss) + "  " + $(if ($ok) {"OK  "} else {"FAIL"}) + "  " + $f[0])
}
Write-Host "COPY3 DONE"
