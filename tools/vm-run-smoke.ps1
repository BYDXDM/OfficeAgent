# vm-run-smoke.ps1 -- Run win7-smoke.ps1 inside the guest, capture output to host file.
$repo = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$vbm = "C:\Program Files\Oracle\VirtualBox\VBoxManage.exe"
$vm = "Win7SP1x64"
$u = "OfficeAgent"
$p = "OA-smoke-2026"
$outDir = Join-Path $repo "build\test"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$out = Join-Path $outDir "guest-smoke-result.txt"
& $vbm guestcontrol $vm run --username $u --password $p --wait-stdout --wait-stderr `
  --exe "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -- `
  -NoProfile -ExecutionPolicy Bypass -File "C:\OfficeAgent\tests\vm\win7-smoke.ps1" -Root "C:\OfficeAgent" `
  2>&1 | Out-File $out -Encoding UTF8
Write-Host ("smoke done, output=" + $out)
