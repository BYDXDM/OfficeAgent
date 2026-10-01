# OfficeAgent bootstrapper build: .NET 3.5 targeted WinForms exe (Win7 RTM zero-prereq).
# AnyCPU: on x64 the boot process runs native 64-bit, so System32\wusa.exe resolves
# without WOW64 redirection (x86 boot would look for wusa in SysWOW64 and fail);
# on 32-bit OS it JITs to 32-bit as before. Boot has no native dll deps.
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File boot\build.ps1
# NOTE: keep this file pure ASCII (PowerShell 5.1 parses BOM-less UTF-8 as GBK).
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

$csc = "C:\Windows\Microsoft.NET\Framework64\v3.5\csc.exe"
if (-not (Test-Path $csc)) { $csc = "C:\Windows\Microsoft.NET\Framework\v3.5\csc.exe" }
if (-not (Test-Path $csc)) { throw "csc.exe for .NET 3.5 not found (enable Windows feature: .NET Framework 3.5)" }

$fw = "C:\Windows\Microsoft.NET\Framework64\v2.0.50727"
if (-not (Test-Path (Join-Path $fw "System.Windows.Forms.dll"))) { $fw = "C:\Windows\Microsoft.NET\Framework\v2.0.50727" }

$refs = @(
  ("/r:" + (Join-Path $fw "System.dll")),
  ("/r:" + (Join-Path $fw "System.Windows.Forms.dll")),
  ("/r:" + (Join-Path $fw "System.Drawing.dll")),
  ("/r:" + (Join-Path $fw "System.Management.dll"))
)

$outDir = Join-Path $root "..\build\boot"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$sources = @()
$sources += (Join-Path $root "src\Program.cs")
$sources += (Join-Path $root "src\Report.cs")
$sources += (Join-Path $root "src\BootForm.cs")
$sources += (Join-Path $root "src\BootLaunch.cs")
$sources += (Join-Path $root "src\Installer.cs")
$sources += (Join-Path $root "src\Unzip.cs")
$sources += (Join-Path $root "..\core\src\Detect.cs")
$sources += (Join-Path $root "..\core\src\ProcRunner.cs")

$cscargs = @()
$cscargs += "/nologo"
$cscargs += "/target:winexe"
$cscargs += "/platform:anycpu"
$cscargs += "/optimize+"
$cscargs += "/codepage:65001"
$cscargs += "/win32manifest:" + (Join-Path $root "app.manifest")
$cscargs += "/win32icon:" + (Join-Path $root "..\app.ico")
$cscargs += "/out:" + (Join-Path $outDir "OfficeAgentBoot.exe")
$cscargs += $refs
$cscargs += $sources

& $csc $cscargs
if ($LASTEXITCODE -ne 0) { throw "compile failed, exit code $LASTEXITCODE" }
Write-Host ("OK: " + (Join-Path $outDir "OfficeAgentBoot.exe"))

# Second target: elevated variant. Same sources, requireAdministrator manifest.
# The asInvoker UI launches this exe via plain Process.Start (ShellExecute path);
# Windows then shows the UAC consent itself -- no Verb/UseShellExecute needed.
$cscargsAdmin = @()
$cscargsAdmin += "/nologo"
$cscargsAdmin += "/target:winexe"
$cscargsAdmin += "/platform:anycpu"
$cscargsAdmin += "/optimize+"
$cscargsAdmin += "/codepage:65001"
$cscargsAdmin += "/win32manifest:" + (Join-Path $root "app.admin.manifest")
$cscargsAdmin += "/win32icon:" + (Join-Path $root "..\app.ico")
$cscargsAdmin += "/out:" + (Join-Path $outDir "OfficeAgentBootAdmin.exe")
$cscargsAdmin += $refs
$cscargsAdmin += $sources

& $csc $cscargsAdmin
if ($LASTEXITCODE -ne 0) { throw "admin compile failed, exit code $LASTEXITCODE" }
Write-Host ("OK: " + (Join-Path $outDir "OfficeAgentBootAdmin.exe"))
