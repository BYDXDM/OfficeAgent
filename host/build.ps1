# OfficeAgent host build: .NET Framework 4.x WinForms exe (runs on .NET 4.8, Win7 SP1+).
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File host\build.ps1
# NOTE: keep this file pure ASCII (PowerShell 5.1 parses BOM-less UTF-8 as GBK).
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

$csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (-not (Test-Path $csc)) { $csc = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe" }
if (-not (Test-Path $csc)) { throw "csc.exe for .NET 4.x not found" }

$fw = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319"
if (-not (Test-Path (Join-Path $fw "System.Windows.Forms.dll"))) { $fw = "C:\Windows\Microsoft.NET\Framework\v4.0.30319" }

$refs = @(
  ("/r:" + (Join-Path $fw "System.dll")),
  ("/r:" + (Join-Path $fw "System.Core.dll")),
  ("/r:" + (Join-Path $fw "System.Data.dll")),
  ("/r:" + (Join-Path $fw "System.Drawing.dll")),
  ("/r:" + (Join-Path $fw "System.Management.dll")),
  ("/r:" + (Join-Path $fw "System.Windows.Forms.dll")),
  ("/r:" + (Join-Path $fw "System.Xml.dll"))
)

$outDir = Join-Path $root "..\build\host"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$sources = @()
$sources += (Join-Path $root "src\Program.cs")
$sources += (Join-Path $root "src\MainForm.cs")
$sources += (Join-Path $root "src\ChatPanel.cs")
$sources += (Join-Path $root "src\CaretHelper.cs")
$sources += (Join-Path $root "src\WorkspaceStore.cs")
$sources += (Join-Path $root "src\OpenFolder.cs")
$sources += (Join-Path $root "src\SetupDialog.cs")
$sources += (Join-Path $root "src\AppConfig.cs")
$sources += (Join-Path $root "src\MemoryStore.cs")
$sources += (Join-Path $root "src\HostGuard.cs")
$sources += (Join-Path $root "src\LlmClient.cs")
  $sources += (Join-Path $root "src\MiniXlsx.cs")
  $sources += (Join-Path $root "src\MiniCsv.cs")
  $sources += (Join-Path $root "src\MiniXlsxWrite.cs")
  $sources += (Join-Path $root "src\ActionPlan.cs")
  $sources += (Join-Path $root "src\IntentRouter.cs")
  $sources += (Join-Path $root "src\ActionExecutor.cs")
  $sources += (Join-Path $root "src\ReconEngine.cs")
  $sources += (Join-Path $root "src\ReconTemplate.cs")
  $sources += (Join-Path $root "src\MaskEngine.cs")
  $sources += (Join-Path $root "src\MiniJson.cs")
  $sources += (Join-Path $root "src\MergeTemplate.cs")
  $sources += (Join-Path $root "src\MergeEngine.cs")
  $sources += (Join-Path $root "src\InvoiceEngine.cs")
  $sources += (Join-Path $root "src\AuditLog.cs")
  $sources += (Join-Path $root "src\PlainTips.cs")
  $sources += (Join-Path $root "src\ColumnSuggest.cs")
  $sources += (Join-Path $root "src\SkillSystem.cs")
  $sources += (Join-Path $root "src\SkillRunner.cs")
  $sources += (Join-Path $root "src\SkillToolBridge.cs")
  $sources += (Join-Path $root "src\ModelText.cs")
  $sources += (Join-Path $root "src\HopBudget.cs")
  $sources += (Join-Path $root "src\ArtifactRegistry.cs")
  $sources += (Join-Path $root "src\AgentTools.cs")
  $sources += (Join-Path $root "src\AgentLoop.cs")
  $sources += (Join-Path $root "src\SessionStore.cs")
  $sources += (Join-Path $root "src\PptWriter.cs")
  $sources += (Join-Path $root "src\RepairLauncher.cs")
  $sources += (Join-Path $root "src\MainForm.Recon.cs")
$sources += (Join-Path $root "src\ConvertEngine.cs")
$sources += (Join-Path $root "src\ToolRunner.cs")
$sources += (Join-Path $root "src\ToolJobGuard.cs")
$sources += (Join-Path $root "src\Pdfium.cs")
$sources += (Join-Path $root "src\PdfTextReader.cs")
  $sources += (Join-Path $root "..\core\src\Detect.cs")
  $sources += (Join-Path $root "..\core\src\ProcRunner.cs")
  $sources += (Join-Path $root "..\core\src\JobObject.cs")
  $sources += (Join-Path $root "..\core\src\JobProbeRunner.cs")
  $sources += (Join-Path $root "..\core\src\EnvSafe.cs")
  $sources += (Join-Path $root "..\core\src\BootGate.cs")
  $sources += (Join-Path $root "..\core\src\MiniZip.cs")

$cscargs = @()
$cscargs += "/nologo"
$cscargs += "/target:winexe"
$cscargs += "/platform:x86"
$cscargs += "/optimize+"
$cscargs += "/codepage:65001"
$cscargs += "/win32manifest:" + (Join-Path $root "app.manifest")
$cscargs += "/win32icon:" + (Join-Path $root "..\app.ico")
$cscargs += "/out:" + (Join-Path $outDir "OfficeAgent.exe")
$cscargs += $refs
$cscargs += $sources

& $csc $cscargs
if ($LASTEXITCODE -ne 0) { throw "compile failed, exit code $LASTEXITCODE" }
Write-Host ("OK: " + (Join-Path $outDir "OfficeAgent.exe"))

# Fallback shell target: same sources on .NET 3.5 (Win7 RTM built-in runtime).
# Design doc 8.4: when .NET 4.8 cannot be installed, the 3.5 shell keeps
# recon/merge/invoice/convert working. Sources are kept C# 3.0-clean for this.
$fw35 = "C:\Windows\Microsoft.NET\Framework64\v2.0.50727"
if (-not (Test-Path (Join-Path $fw35 "System.Windows.Forms.dll"))) { $fw35 = "C:\Windows\Microsoft.NET\Framework\v2.0.50727" }
$csc35 = "C:\Windows\Microsoft.NET\Framework64\v3.5\csc.exe"
if (-not (Test-Path $csc35)) { $csc35 = "C:\Windows\Microsoft.NET\Framework\v3.5\csc.exe" }

$refs35 = @(
  ("/r:" + (Join-Path $fw35 "System.dll")),
  ("/r:" + (Join-Path $fw35 "System.Data.dll")),
  ("/r:" + (Join-Path $fw35 "System.Drawing.dll")),
  ("/r:" + (Join-Path $fw35 "System.Management.dll")),
  ("/r:" + (Join-Path $fw35 "System.Windows.Forms.dll")),
  ("/r:" + (Join-Path $fw35 "System.Xml.dll"))
)
$cscargs35 = @()
$cscargs35 += "/nologo"
$cscargs35 += "/target:winexe"
$cscargs35 += "/platform:x86"
$cscargs35 += "/optimize+"
$cscargs35 += "/codepage:65001"
$cscargs35 += "/warn:1"
$cscargs35 += "/win32manifest:" + (Join-Path $root "app.manifest")
$cscargs35 += "/win32icon:" + (Join-Path $root "..\app.ico")
$cscargs35 += "/out:" + (Join-Path $outDir "OfficeAgent35.exe")
$cscargs35 += $refs35
$cscargs35 += $sources

& $csc35 $cscargs35
if ($LASTEXITCODE -ne 0) { throw "net35 fallback compile failed, exit code $LASTEXITCODE" }
Write-Host ("OK: " + (Join-Path $outDir "OfficeAgent35.exe"))
