# download-kb-x86.ps1 -- Fetch the x86 (32-bit Win7 SP1) MSU variants for the four
# offline patches from the Microsoft Update Catalog, verify and place into payload\kb\.
# Filename filter: must start with windows6.1- and contain -x86_ (excludes 6.0/Server and x64).
# After download, print SHA256 lines to append into store\components.ini.
# Keep pure ASCII. Requires network to www.catalog.update.microsoft.com and
# *.download.windowsupdate.com (both reachable from dev sandbox).
$ErrorActionPreference = "Stop"
$repo = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$kbDir = Join-Path $repo "payload\kb"

$targets = @(
  @{ kb = "KB4490628"; patterns = @("windows6.1-kb4490628-x86_") },
  @{ kb = "KB4474419"; patterns = @("windows6.1-kb4474419-v3-x86_", "windows6.1-kb4474419-x86_") },
  @{ kb = "KB3140245"; patterns = @("windows6.1-kb3140245-x86_") },
  @{ kb = "KB2670838"; patterns = @("windows6.1-kb2670838-x86_") }
)

function Get-CatalogUrls($kb) {
  $search = Invoke-WebRequest -UseBasicParsing ("https://www.catalog.update.microsoft.com/Search.aspx?q=" + $kb)
  $ms = [regex]::Matches($search.Content, 'goToDetails\("([0-9a-fA-F-]{36})"\)')
  $guids = @()
  foreach ($m in $ms) { if ($guids -notcontains $m.Groups[1].Value) { $guids += $m.Groups[1].Value } }
  $urls = @()
  foreach ($g in $guids) {
    $body = 'updateIDs=%5B%7B%22size%22%3A0%2C%22updateID%22%3A%22' + $g + '%22%2C%22uidInfo%22%3A%22' + $g + '%22%7D%5D'
    try {
      $resp = Invoke-WebRequest -UseBasicParsing -Method Post `
        -Uri "https://www.catalog.update.microsoft.com/DownloadDialog.aspx" `
        -ContentType "application/x-www-form-urlencoded" -Body $body
      $um = [regex]::Matches($resp.Content, "files\[\d+\]\.url\s*=\s*'([^']+)'")
      foreach ($u in $um) { $urls += $u.Groups[1].Value.Replace("\u0026", "&") }
    } catch { }
  }
  return $urls
}

foreach ($t in $targets) {
  $dest = $null
  foreach ($pat in $t.patterns) {
    if (Test-Path (Join-Path $kbDir ($pat + "msu"))) { $dest = Join-Path $kbDir ($pat + "msu"); break }
  }
  if ($dest -ne $null) { Write-Host ("SKIP exists: " + (Split-Path -Leaf $dest)); continue }

  Write-Host ("searching " + $t.kb + " ...")
  $urls = Get-CatalogUrls $t.kb
  $picked = $null
  foreach ($pat in $t.patterns) {
    foreach ($u in $urls) {
      $name = ([Uri]$u).AbsolutePath.Split("/")[-1]
      if ($name -like ($pat + "*")) { $picked = $u; break }
    }
    if ($picked -ne $null) { break }
  }
  if ($picked -eq $null) {
    Write-Host ("FAIL no match for " + $t.kb + " (" + $urls.Count + " urls seen)")
    continue
  }
  $name = ([Uri]$picked).AbsolutePath.Split("/")[-1]
  $dest = Join-Path $kbDir $name
  Write-Host ("downloading " + $name)
  Invoke-WebRequest -UseBasicParsing -Uri $picked -OutFile $dest
  $sha = (Get-FileHash $dest -Algorithm SHA256).Hash.ToLowerInvariant()
  Write-Host ("OK " + $name + "  bytes=" + (Get-Item $dest).Length + "  sha256=" + $sha)
}
Write-Host "ALL DONE"
