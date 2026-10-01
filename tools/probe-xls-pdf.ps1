# probe-xls-pdf.ps1 -- Bisect the Excel COM ExportAsFixedFormat hang with variants.
# Each variant runs in an isolated child powershell with a 25s timeout. Keep ASCII.
param([string]$Variant)
$repo = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$xls = Join-Path (Join-Path (Join-Path $repo "tests\fixtures") "office") "test.xls"
$pdf = Join-Path $env:TEMP "oa-probe.pdf"
if (Test-Path $pdf) { Remove-Item $pdf -Force }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Open($xls)
  if ($Variant -eq "D") { $excel.PrintCommunication = $false }
  if ($Variant -eq "A") { $wb.ExportAsFixedFormat(0, $pdf) }
  elseif ($Variant -eq "B") { $wb.ExportAsFixedFormat(0, $pdf, 0, $false, $true) }
  elseif ($Variant -eq "C") { $wb.Worksheets.Item(1).ExportAsFixedFormat(0, $pdf, 0, $false, $true) }
  elseif ($Variant -eq "D") { $wb.ExportAsFixedFormat(0, $pdf, 0, $false, $true); $excel.PrintCommunication = $true }
  elseif ($Variant -eq "F") { $wb.SaveAs($pdf, 57) }
  elseif ($Variant -eq "I") { $wb.ExportAsFixedFormat(0, $pdf, 0, $false, $true, 1, 1, $false) }
  if (Test-Path $pdf) { $msg = "OK size=" + (Get-Item $pdf).Length } else { $msg = "NOFILE" }
} catch { $msg = "EXC " + $_.Exception.Message } finally {
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
$log = Join-Path (Join-Path $repo "build\test") ("probe-" + $Variant + ".txt")
[System.IO.File]::WriteAllText($log, $msg, [System.Text.Encoding]::UTF8)
