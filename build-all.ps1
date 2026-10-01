# Build both boot (net35) and host (net4x). Keep pure ASCII.
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
& (Join-Path $root "boot\build.ps1")
& (Join-Path $root "host\build.ps1")
Write-Host "ALL DONE"
