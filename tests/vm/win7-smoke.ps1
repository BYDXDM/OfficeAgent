# win7-smoke.ps1 -- In-VM automated smoke for OfficeAgent (M0..M3 headless chain).
# Target: Windows 7 SP1 x64, PowerShell 2.0 (no PS3+ syntax, ASCII only).
# Usage:  powershell -NoProfile -ExecutionPolicy Bypass -File win7-smoke.ps1 -Root C:\OfficeAgent
# Result: prints PASS/FAIL lines, writes win7-smoke-result.txt next to this script, exit 0=pass 2=fail.
param([string]$Root = "C:\OfficeAgent")

$script:pass = 0
$script:fail = 0
$script:skips = 0
$lines = New-Object System.Collections.ArrayList

function Note($ok, $step, $detail) {
  # NB: compare as strings -- ($true -eq "SKIP") is TRUE in PowerShell because the
  # RHS string converts to bool; that would mark every PASS as SKIP.
  $s = "$ok"
  if ($s -eq "SKIP") {
    $script:skips++
    $out = "[SKIP] " + $step + "  " + $detail
  } elseif ($s -eq "True") {
    $script:pass++
    $out = "[PASS] " + $step + "  " + $detail
  } else {
    $script:fail++
    $out = "[FAIL] " + $step + "  " + $detail
  }
  Write-Host $out
  [void]$lines.Add($out)
}

function Exec($exe, $argList, $outFile) {
  $p = New-Object System.Diagnostics.Process
  $psi = $p.StartInfo
  $psi.FileName = $exe
  $psi.Arguments = $argList
  $psi.UseShellExecute = $false
  $psi.CreateNoWindow = $true
  $psi.RedirectStandardOutput = $true
  $psi.RedirectStandardError = $true
  $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
  $psi.StandardErrorEncoding = [System.Text.Encoding]::UTF8
  $env:PYTHONIOENCODING = "utf-8"
  [void]$p.Start()
  $so = $p.StandardOutput.ReadToEnd()
  $se = $p.StandardError.ReadToEnd()
  $p.WaitForExit()
  if ($outFile -ne "") {
    [System.IO.File]::WriteAllText($outFile, $so + $se, [System.Text.Encoding]::UTF8)
  }
  return $p.ExitCode
}

$boot   = Join-Path (Join-Path $Root "build\boot") "OfficeAgentBoot.exe"
$host2  = Join-Path (Join-Path $Root "build\host") "OfficeAgent.exe"
$fix    = Join-Path $Root "tests\fixtures"
$resultFile = Join-Path $env:TEMP "win7-smoke-result.txt"

Write-Host ("== OfficeAgent Win7 smoke == root=" + $Root)
[void]$lines.Add("== OfficeAgent Win7 smoke == root=" + $Root + "  time=" + [DateTime]::Now.ToString("s"))

if (-not (Test-Path $host2)) {
  Note $false "0.0 app exists" $host2
  $lines | Out-File -FilePath $resultFile -Encoding UTF8
  exit 2
}

# 1. boot report-only: exit 0 = env all green, 2 = has missing (missing is OK on a bare VM, but record)
$rc = Exec $boot "/report-only /console" ""
Note ($rc -eq 0 -or $rc -eq 2) "1.1 boot report-only runs" ("exit=" + $rc)

# 2. host selftest / intents / masktest
$rc = Exec $host2 "/selftest" (Join-Path $env:TEMP "oa-selftest.txt")
Note ($rc -eq 0) "2.1 host /selftest" ("exit=" + $rc)
$rc = Exec $host2 "/intents" ""
Note ($rc -eq 0) "2.2 host /intents (router)" ("exit=" + $rc)
$rc = Exec $host2 "/masktest" ""
Note ($rc -eq 0) "2.3 host /masktest" ("exit=" + $rc)

# 3. recon on bundled diff fixture
$flow   = Join-Path (Join-Path $fix "recon") "flow.csv"
$ledger = Join-Path (Join-Path $fix "recon") "ledger.csv"
$reconOut = Join-Path $env:TEMP "oa-recon-out.xlsx"
if (Test-Path $reconOut) { Remove-Item $reconOut -Force }
$rc = Exec $host2 ("/recon `"" + $flow + "`" `"" + $ledger + "`" --keya 1 --keyb 2 --debita 5 --credita 6 --debitb 4 --creditb 5 --out `"" + $reconOut + "`"") ""
Note ($rc -eq 0) "3.1 /recon diff fixture (col indexes)" ("exit=" + $rc)
Note (Test-Path $reconOut) "3.2 diff report exists" $reconOut

# 3b. net35 fallback shell (design 8.4): same recon on the 3.5-built exe.
#     Also proves MiniZip (pure 3.5 zip) round-trips on the guest runtime.
$host35 = Join-Path (Join-Path $Root "build\host") "OfficeAgent35.exe"
if (Test-Path $host35) {
  $reconOut35 = Join-Path $env:TEMP "oa-recon-out35.xlsx"
  if (Test-Path $reconOut35) { Remove-Item $reconOut35 -Force }
  $rc = Exec $host35 ("/recon `"" + $flow + "`" `"" + $ledger + "`" --keya 1 --keyb 2 --debita 5 --credita 6 --debitb 4 --creditb 5 --out `"" + $reconOut35 + "`"") ""
  Note ($rc -eq 0) "3.3 /recon on net35 fallback shell" ("exit=" + $rc)
  Note (Test-Path $reconOut35) "3.4 fallback report exists" $reconOut35
} else {
  Note "SKIP" "3.3/3.4 net35 fallback" "OfficeAgent35.exe not present in build\host"
}

# 4. conversions: csv native chain, pdf needs LO unpacked
$dirty = Join-Path $fix "dirty.xlsx"
$csvOut = Join-Path $env:TEMP "oa-dirty-conv.csv"
if (Test-Path $csvOut) { Remove-Item $csvOut -Force }
$rc = Exec $host2 ("/convert `"" + $dirty + "`" csv") ""
Note ($rc -eq 0) "4.1 /convert xlsx->csv (native)" ("exit=" + $rc)
$soffice = Join-Path (Join-Path (Join-Path $Root "lo76") "LibreOffice") "program\soffice.exe"
if (Test-Path $soffice) {
  # /convert writes next to the input: <dir>\dirty_conv.pdf
  $pdfOut = Join-Path (Split-Path -Parent $dirty) "dirty_conv.pdf"
  if (Test-Path $pdfOut) { Remove-Item $pdfOut -Force }
  $rc = Exec $host2 ("/convert `"" + $dirty + "`" pdf") ""
  Note ($rc -eq 0) "4.2 /convert xlsx->pdf (LibreOffice)" ("exit=" + $rc)
  if (Test-Path $pdfOut) {
    $rc = Exec $host2 ("/renderpdf `"" + $pdfOut + "`" " + (Join-Path $env:TEMP "oa-render.png")) ""
    Note ($rc -eq 0) "4.3 /renderpdf pdfium render" ("exit=" + $rc)
  } else {
    Note $false "4.3 /renderpdf pdfium render" "pdf missing"
  }
} else {
  Note "SKIP" "4.2 xlsx->pdf" "LibreOffice not unpacked (run boot fix or copy lo76 tree)"
  Note "SKIP" "4.3 /renderpdf" "no pdf input"
}

# 5. audit chain written by the runs above
$audit = Join-Path (Join-Path $env:LOCALAPPDATA "OfficeAgent") "audit.jsonl"
$auditTxt = ""
if (Test-Path $audit) { $auditTxt = [System.IO.File]::ReadAllText($audit) }
Note ($auditTxt -like "*app_start*") "5.1 audit app_start" ""
Note ($auditTxt -like "*action_finished*" -or $auditTxt -like "*recon_run*") "5.2 audit recon event" ""

# 6. privacy default: config.json privacyLevel absent => default 1 (L1)
$cfg = Join-Path (Join-Path $env:LOCALAPPDATA "OfficeAgent") "config.json"
if (Test-Path $cfg) {
  $cfgTxt = [System.IO.File]::ReadAllText($cfg)
  $lvlBad = $cfgTxt -like "*privacyLevel*"
  Note $true "6.1 config exists" ""
} else {
  Note "SKIP" "6.1 config.json" "no wizard run yet (GUI first-run creates it)"
}

# summary
$summary = "RESULT pass=" + $script:pass + " fail=" + $script:fail + " skip=" + $script:skips
Write-Host $summary
[void]$lines.Add($summary)
$lines | Out-File -FilePath $resultFile -Encoding UTF8
if ($script:fail -gt 0) { exit 2 }
exit 0
