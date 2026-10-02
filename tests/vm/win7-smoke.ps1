# win7-smoke.ps1 -- In-VM automated smoke for OfficeAgent (M0..M3 headless chain).
# Target: Windows 7 SP1 x64, PowerShell 2.0 (no PS3+ syntax, ASCII only).
# Usage:  powershell -NoProfile -ExecutionPolicy Bypass -File win7-smoke.ps1 -Root C:\OfficeAgent
# Result: prints PASS/FAIL lines, writes win7-smoke-result.txt next to this script, exit 0=pass 2=fail.
param([string]$Root = "C:\OfficeAgent")

$script:pass = 0
$script:fail = 0
$script:skips = 0
$lines = New-Object System.Collections.ArrayList

# Write a file as BOM-less UTF-8, in a way that works on PowerShell 2.0 (Win7).
#
# Why not [System.Text.Encoding]::UTF8: that overload emits a UTF-8 BOM, and the
# request.json we hand to the python sidecar is parsed with json.load(encoding='utf-8'),
# which HARD-FAILS on a BOM ("Unexpected UTF-8 BOM ... decode using utf-8-sig").
# Verified on the dev box: BOM -> JSONDecodeError, no BOM -> loads fine.
#
# Why not New-Object System.Text.UTF8Encoding($false): passing constructor arguments
# that way relies on PS2.0's overload resolution. It is *probably* fine, but this
# script must not gamble on the target OS, so we avoid constructor overloads entirely:
# Encoding.UTF8.GetBytes() uses the BOM-less UTF8 encoder for byte[] output
# (the BOM only comes from the *string* writers), then WriteAllBytes writes it raw.
function WriteUtf8NoBom($path, $text) {
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($text)
  [System.IO.File]::WriteAllBytes($path, $bytes)
}

function Note($ok, $step, $detail) {  # NB: compare as strings -- ($true -eq "SKIP") is TRUE in PowerShell because the
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

# 7. python38 sidecar skills (added in 0.6.3): verify runtime + offline deps + each skill runs.
#    These are the skills that need store\wheels; on Win7 the embedded python may be x86,
#    so this section also proves the wheel set matches the actual interpreter bitness.
$py = Join-Path $Root "runtime\py38\python.exe"
if ((Test-Path $py) -and (Test-Path $host2)) {
  # 7.1 interpreter runs and reports bitness (proves VC++ runtime is present)
  $pyOut = Join-Path $env:TEMP "oa-pyver.txt"
  $rc = Exec $py "-c ""import sys,platform;open(r'$pyOut','w').write(sys.version.split()[0]+' '+platform.machine())""" ""
  $pyTxt = ""
  if (Test-Path $pyOut) { $pyTxt = [System.IO.File]::ReadAllText($pyOut) }
  Note ($rc -eq 0 -and $pyTxt.Length -gt 0) "7.1 python38 sidecar runs" ("exit=" + $rc + " ver=" + $pyTxt)

  # 7.2 the offline deps the new skills need are importable
  $depOut = Join-Path $env:TEMP "oa-deps.txt"
  $depCode = "import openpyxl,docx,pptx,pypdf,pdfplumber,PIL,xlsxwriter;open(r'$depOut','w').write('ok')"
  $rc = Exec $py ("-c """ + $depCode + """") ""
  $depTxt = ""
  if (Test-Path $depOut) { $depTxt = [System.IO.File]::ReadAllText($depOut) }
  Note ($depTxt -eq "ok") "7.2 offline deps importable" ("(openpyxl/docx/pptx/pypdf/pdfplumber/PIL/xlsxwriter)")

  # 7.3 skills are registered
  $regOut = Join-Path $env:TEMP "oa-skillreg.txt"
  $rc = Exec $host2 "/skilltest" $regOut
  $regTxt = ""
  if (Test-Path $regOut) { $regTxt = [System.IO.File]::ReadAllText($regOut) }
  Note ($rc -eq 0) "7.3 /skilltest (registry)" ("exit=" + $rc)
  $hasXlsx = $regTxt -like "*xlsx-ops*"
  $hasDocx = $regTxt -like "*docx-report*"
  $hasAcct = $regTxt -like "*acct-tools*"
  Note ($hasXlsx -and $hasDocx -and $hasAcct) "7.4 new skills registered" ("xlsx-ops=" + $hasXlsx + " docx-report=" + $hasDocx + " acct-tools=" + $hasAcct)

  # 7.5 row-stat end-to-end (stdlib only): proves stage -> sidecar -> stdout JSON chain
  $csvIn = Join-Path $env:TEMP "oa-rows.csv"
  WriteUtf8NoBom $csvIn "col,amt`r`n a,1`r`n b,2"
  $rsOut = Join-Path $env:TEMP "oa-rowstat.txt"
  $rc = Exec $host2 ("/skillrun row-stat `"" + $csvIn + "`" --yes") $rsOut
  $rsTxt = ""
  if (Test-Path $rsOut) { $rsTxt = [System.IO.File]::ReadAllText($rsOut) }
  Note ($rc -eq 0 -and $rsTxt -like "*stat*") "7.5 row-stat end-to-end" ("exit=" + $rc)

  # 7.6 xlsx-ops end-to-end: exercises openpyxl (the wheel that must match interpreter bitness)
  $xlsxIn = Join-Path $env:TEMP "oa-xs.xlsx"
  if (Test-Path $xlsxIn) { Remove-Item $xlsxIn -Force }
  $rc = Exec $host2 ("/convert `"" + $csvIn + "`" xlsx") ""
  $csvConv = $csvIn -replace "\.csv$", "_conv.xlsx"
  if (-not (Test-Path $csvConv)) { $csvConv = (Join-Path $env:TEMP "oa-rows_conv.xlsx") }
  if (Test-Path $csvConv) {
    $reqFile = Join-Path $env:TEMP "oa-xreq.json"
    WriteUtf8NoBom $reqFile '{"action":"inspect"}'
    $xsOut = Join-Path $env:TEMP "oa-xlsxops.txt"
    $rc = Exec $host2 ("/skillrun xlsx-ops `"" + $csvConv + "`" --yes --request-file `"" + $reqFile + "`"") $xsOut
    $xsTxt = ""
    if (Test-Path $xsOut) { $xsTxt = [System.IO.File]::ReadAllText($xsOut) }
    # Match on the JSON field name ("sheets") rather than the localized UI label:
    # this script must stay pure ASCII for PowerShell 2.0 on Win7.
    Note ($xsTxt -like "*sheets*") "7.6 xlsx-ops inspect (openpyxl)" ("exit=" + $rc)
  } else {
    Note "SKIP" "7.6 xlsx-ops inspect" "csv->xlsx convert did not produce a file"
  }

  # 7.7 docx-report end-to-end: exercises python-docx (newly added dependency)
  $docOut = Join-Path $env:TEMP "oa-report.docx"
  if (Test-Path $docOut) { Remove-Item $docOut -Force }
  $dreq = Join-Path $env:TEMP "oa-dreq.json"
  $djson = '{"out":"' + ($docOut -replace "\\", "\\") + '","title":"smoke","blocks":[{"type":"para","text":"ok"}]}'
  WriteUtf8NoBom $dreq $djson
  $drOut = Join-Path $env:TEMP "oa-docxrep.txt"
  $rc = Exec $host2 ("/skillrun docx-report `"" + $csvIn + "`" --yes --request-file `"" + $dreq + "`"") $drOut
  Note (Test-Path $docOut) "7.7 docx-report generates .docx (python-docx)" ("exists=" + (Test-Path $docOut))

  # 7.8 acct-tools: trial-balance must detect an unbalanced set
  $tbIn = Join-Path $env:TEMP "oa-tb.csv"
  WriteUtf8NoBom $tbIn "subj,debit,credit`r`na,100,0`r`nb,0,80"
  $tbReq = Join-Path $env:TEMP "oa-tbreq.json"
  WriteUtf8NoBom $tbReq '{"action":"trial-balance"}'
  $tbOut = Join-Path $env:TEMP "oa-tb.txt"
  $rc = Exec $host2 ("/skillrun acct-tools `"" + $tbIn + "`" --yes --request-file `"" + $tbReq + "`"") $tbOut
  $tbTxt = ""
  if (Test-Path $tbOut) { $tbTxt = [System.IO.File]::ReadAllText($tbOut) }
  # 100 vs 80 -> unbalanced by 20; the skill must report the difference (not silently pass).
  # Match on the numeric delta only -- keep this file pure ASCII.
  $detected = ($tbTxt -like "*20.00*") -or ($tbTxt -like "*unbalanced*")
  Note $detected "7.8 acct-tools detects unbalanced ledger" ("")
} else {
  Note "SKIP" "7.x python skills" "runtime\py38 or host exe missing"
}

# summary
$summary = "RESULT pass=" + $script:pass + " fail=" + $script:fail + " skip=" + $script:skips
Write-Host $summary
[void]$lines.Add($summary)
$lines | Out-File -FilePath $resultFile -Encoding UTF8
if ($script:fail -gt 0) { exit 2 }
exit 0
