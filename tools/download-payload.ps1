# Rebuild offline payload from locked official URLs (all sha256 recorded in store\components.ini).
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File tools\download-payload.ps1 [-Proxy http://127.0.0.1:7897]
# NOTE: keep pure ASCII.
param([string]$Proxy = "")
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)

$items = @(
  @{ url = "https://www.python.org/ftp/python/3.8.10/python-3.8.10-embed-win32.zip"; out = "payload\py38\python-3.8.10-embed-win32.zip" },
  @{ url = "https://www.python.org/ftp/python/3.8.10/python-3.8.10-embed-amd64.zip"; out = "payload\py38\python-3.8.10-embed-amd64.zip" },
  @{ url = "https://bootstrap.pypa.io/get-pip.py"; out = "payload\py38\get-pip.py" },
  @{ url = "https://catalog.s.download.windowsupdate.com/msdownload/update/software/ftpk/2013/02/windows6.1-kb2670838-x64_9f667ff60e80b64cbed2774681302baeaf0fc6a6.msu"; out = "payload\kb\windows6.1-kb2670838-x64.msu" },
  @{ url = "https://go.microsoft.com/fwlink/?linkid=2088631"; out = "payload\ndp48\ndp48-x86x64allos-enu.exe" },
  @{ url = "https://download.visualstudio.microsoft.com/download/pr/bd1c8d9d-ba95-4eee-bc6e-df1fcc876373/0C09F2611660441084CE0DF425C51C11E147E6447963C3690F97E0B25C55ED64/VC_redist.x86.exe"; out = "payload\vcrt\VC_redist.x86.exe" },
  @{ url = "https://download.visualstudio.microsoft.com/download/pr/bd1c8d9d-ba95-4eee-bc6e-df1fcc876373/CC0FF0EB1DC3F5188AE6300FAEF32BF5BEEBA4BDD6E8E445A9184072096B713B/VC_redist.x64.exe"; out = "payload\vcrt\VC_redist.x64.exe" }
)

foreach ($it in $items) {
  $dst = Join-Path $root $it.out
  New-Item -ItemType Directory -Force -Path (Split-Path -Parent $dst) | Out-Null
  if (Test-Path $dst) { Write-Host ("SKIP (exists): " + $it.out); continue }
  Write-Host ("GET: " + $it.url)
  if ($Proxy -ne "") {
    Invoke-WebRequest -Uri $it.url -OutFile $dst -Proxy $Proxy -UseBasicParsing
  } else {
    Invoke-WebRequest -Uri $it.url -OutFile $dst -UseBasicParsing
  }
}
Write-Host "DONE. Verify hashes against store\components.ini."
