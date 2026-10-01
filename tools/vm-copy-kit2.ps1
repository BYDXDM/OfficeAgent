# vm-copy-kit2.ps1 -- Robust kit copy: mkdir each dir, copy each file with retries, verify.
$repo = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$vbm = "C:\Program Files\Oracle\VirtualBox\VBoxManage.exe"
$vm = "Win7SP1x64"
$u = "OfficeAgent"
$p = "OA-smoke-2026"

# 1. dirs
$dirs = @(
  "C:\OfficeAgent\build\host",
  "C:\OfficeAgent\build\boot",
  "C:\OfficeAgent\bin",
  "C:\OfficeAgent\tests\fixtures\recon",
  "C:\OfficeAgent\tests\vm"
)
foreach ($d in $dirs) {
  & $vbm guestcontrol $vm mkdir --username $u --password $p --parents $d 2>$null | Out-Null
}
Start-Sleep -Seconds 2

# 2. files (source relative to repo, dest full path) with retry
$files = @(
  @("build\host\OfficeAgent.exe",  "C:\OfficeAgent\build\host\OfficeAgent.exe"),
  @("build\host\pdfium.dll",       "C:\OfficeAgent\build\host\pdfium.dll"),
  @("build\boot\OfficeAgentBoot.exe", "C:\OfficeAgent\build\boot\OfficeAgentBoot.exe"),
  @("bin\pdfium.dll",              "C:\OfficeAgent\bin\pdfium.dll"),
  @("tests\fixtures\test-smoke.pdf", "C:\OfficeAgent\tests\fixtures\test-smoke.pdf"),
  @("tests\fixtures\dirty.xlsx",   "C:\OfficeAgent\tests\fixtures\dirty.xlsx"),
  @("tests\fixtures\gb18030.csv",  "C:\OfficeAgent\tests\fixtures\gb18030.csv"),
  @("tests\fixtures\recon\flow.csv",   "C:\OfficeAgent\tests\fixtures\recon\flow.csv"),
  @("tests\fixtures\recon\ledger.csv", "C:\OfficeAgent\tests\fixtures\recon\ledger.csv"),
  @("tests\vm\win7-smoke.ps1",     "C:\OfficeAgent\tests\vm\win7-smoke.ps1")
)
foreach ($f in $files) {
  $src = Join-Path $repo $f[0]
  $ok = $false
  for ($i = 0; $i -lt 3 -and -not $ok; $i++) {
    & $vbm guestcontrol $vm copyto --username $u --password $p --target-directory (Split-Path -Parent $f[1]) $src 2>$null | Out-Null
    Start-Sleep -Seconds 1
    $chk = & $vbm guestcontrol $vm stat --username $u --password $p $f[1] 2>$null
    if ("$chk" -notlike "*error*" -and "$chk" -notlike "") { $ok = $true }
    else { $ok = $false }
  }
  Write-Host ((Get-Date -Format HH:mm:ss) + "  " + $(if ($ok) {"OK  "} else {"FAIL"}) + "  " + $f[0])
}
Write-Host "COPY PHASE DONE"
