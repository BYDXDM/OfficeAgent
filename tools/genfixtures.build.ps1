# Build GenFixtures fixture generator (net4x console exe). Keep pure ASCII.
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

$csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (-not (Test-Path $csc)) { $csc = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe" }
if (-not (Test-Path $csc)) { throw "csc.exe for .NET 4.x not found" }

$fw = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319"
if (-not (Test-Path (Join-Path $fw "System.IO.Compression.FileSystem.dll"))) { $fw = "C:\Windows\Microsoft.NET\Framework\v4.0.30319" }

$outDir = Join-Path $root "..\build\test"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

& $csc /nologo /target:exe /platform:anycpu /optimize+ /codepage:65001 `
  ("/r:" + (Join-Path $fw "System.dll")) `
  ("/r:" + (Join-Path $fw "System.Core.dll")) `
  ("/r:" + (Join-Path $fw "System.IO.Compression.dll")) `
  ("/r:" + (Join-Path $fw "System.IO.Compression.FileSystem.dll")) `
  ("/out:" + (Join-Path $outDir "GenFixtures.exe")) `
  (Join-Path $root "GenFixtures.cs")
if ($LASTEXITCODE -ne 0) { throw "compile failed, exit code $LASTEXITCODE" }
Write-Host ("OK: " + (Join-Path $outDir "GenFixtures.exe"))
