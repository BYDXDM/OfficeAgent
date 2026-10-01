# verify-payload.ps1 -- Check every store\components.ini entry: file exists + SHA256 matches.
# Also verifies bin\pdfium.dll presence and runtime\py38. Exit 0=all good, 2=problems. Keep ASCII.
$ErrorActionPreference = "Continue"
$repo = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$ini = Join-Path $repo "store\components.ini"
$sha = New-Object System.Security.Cryptography.SHA256Managed

$pass = 0; $fail = 0; $skip = 0
$cur = @{}
function Flush($c) {
  if (-not $c.ContainsKey("file")) { return }
  $f = Join-Path $repo ($c["file"] -replace '/', '\')
  if (-not (Test-Path $f)) {
    Write-Host ("[MISS] " + $c["file"])
    $script:fail++
    return
  }
  if (-not $c.ContainsKey("sha256")) { Write-Host ("[INFO] " + $c["file"] + " (no hash)"); $script:skip++; return }
  $h = (Get-FileHash $f -Algorithm SHA256).Hash.ToLowerInvariant()
  if ($h -eq $c["sha256"].ToLowerInvariant()) { Write-Host ("[OK]   " + $c["file"]); $script:pass++ }
  else { Write-Host ("[HASH] " + $c["file"] + "  expected=" + $c["sha256"].Substring(0,12) + " actual=" + $h.Substring(0,12)); $script:fail++ }
}

Get-Content $ini | ForEach-Object {
  $line = $_.Trim()
  if ($line -like "component *") {
    Flush $cur
    $cur = @{}
  } elseif ($line -match "^(\w+)\s*=\s*(.*)$") {
    $cur[$Matches[1].ToLowerInvariant()] = $Matches[2].Trim()
  } elseif ($line -eq "end") {
    Flush $cur
    $cur = @{}
  }
}
Flush $cur

# key runtime bits
foreach ($p in @("bin\pdfium.dll", "runtime\py38\python.exe", "lo76\LibreOffice\program\soffice.exe")) {
  $f = Join-Path $repo ($p -replace '/', '\')
  if (Test-Path $f) { Write-Host ("[OK]   " + $p); $pass++ }
  else { Write-Host ("[MISS] " + $p); $fail++ }
}
Write-Host ("RESULT pass=" + $pass + " fail=" + $fail + " info=" + $skip)
if ($fail -gt 0) { exit 2 } else { exit 0 }
