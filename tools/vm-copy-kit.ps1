# vm-copy-kit.ps1 -- Copy the test kit into the guest via guestcontrol. Keep ASCII.
$repo = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$vbm = "C:\Program Files\Oracle\VirtualBox\VBoxManage.exe"
$vm = "Win7SP1x64"
$u = "OfficeAgent"
$p = "OA-smoke-2026"

# guest: create target dirs
& $vbm guestcontrol $vm mkdir --username $u --password $p --parents `
  "C:\OfficeAgent\build" "C:\OfficeAgent\tests" "C:\OfficeAgent\bin" 2>$null | Out-Null

$items = @(
  @("build\host\OfficeAgent.exe", "C:\OfficeAgent\build\host\OfficeAgent.exe"),
  @("build\host\pdfium.dll",      "C:\OfficeAgent\build\host\pdfium.dll"),
  @("build\boot\OfficeAgentBoot.exe", "C:\OfficeAgent\build\boot\OfficeAgentBoot.exe"),
  @("bin\pdfium.dll",             "C:\OfficeAgent\bin\pdfium.dll"),
  @("tests\fixtures",             "C:\OfficeAgent\tests\fixtures"),
  @("tests\vm\win7-smoke.ps1",    "C:\OfficeAgent\tests\vm\win7-smoke.ps1")
)
foreach ($it in $items) {
  $src = Join-Path $repo $it[0]
  $dst = $it[1]
  Write-Host ("copy: " + $it[0])
  & $vbm guestcontrol $vm copyto --username $u --password $p --target-directory $dst $src 2>&1 | Select-Object -First 1
}
Write-Host "kit copied"
