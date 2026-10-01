# gen-office-fixtures.ps1 -- Create doc/docx/ppt/pptx/xls fixtures via Office COM (headless).
# Output: tests\fixtures\office\{test.doc,test.docx,test.ppt,test.pptx,test.xls}
# Keep ASCII. Each file gets Chinese text so CP936/encoding handling is exercised.
$ErrorActionPreference = "Stop"
$repo = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$out = Join-Path (Join-Path $repo "tests\fixtures") "office"
if (-not (Test-Path $out)) { New-Item -ItemType Directory -Force -Path $out | Out-Null }

# ---------- Word: docx + doc ----------
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  $doc = $word.Documents.Add()
  $sel = $word.Selection
  $sel.TypeText("OfficeAgent conversion fixture - test document")
  $sel.TypeParagraph()
  $sel.TypeText("Account reconciliation: travel expenses 12345.67 yuan; office supplies 880.00 yuan. (encoding check)")
  $doc.SaveAs([ref](Join-Path $out "test.docx"), [ref]12)   # wdFormatXMLDocument
  $doc.SaveAs([ref](Join-Path $out "test.doc"),  [ref]0)    # wdFormatDocument97
  $doc.Close($false)
  Write-Host "word ok"
} finally { $word.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null }

# ---------- PowerPoint: pptx + ppt ----------
$pp = New-Object -ComObject PowerPoint.Application
try {
  $pres = $pp.Presentations.Add(0)   # msoFalse: WithWindow=false (headless)
  $slide = $pres.Slides.Add(1, 2)    # ppLayoutText (title + body placeholders)
  $slide.Shapes.Item(1).TextFrame.TextRange.Text = "OfficeAgent PPT fixture"
  $slide.Shapes.Item(2).TextFrame.TextRange.Text = "Total amount 45000.00 yuan - including tax (encoding check)"
  $pres.SaveAs((Join-Path $out "test.pptx"), 24)   # ppSaveAsOpenXMLPresentation
  $pres.SaveAs((Join-Path $out "test.ppt"),  1)    # ppSaveAsPresentation
  $pres.Close()
  Write-Host "powerpoint ok"
} finally { $pp.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null }

# ---------- Excel: xls (xlsx already exists via GenFixtures) ----------
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $ws.Cells.Item(1,1) = "Account"
  $ws.Cells.Item(1,2) = "Amount"
  $ws.Cells.Item(2,1) = "Cash on hand"
  $ws.Cells.Item(2,2) = 1234.56
  $ws.Cells.Item(3,1) = "Bank deposit"
  $ws.Cells.Item(3,2) = 9876.00
  $wb.SaveAs((Join-Path $out "test.xls"), 56)   # xlExcel8 (97-2003)
  $wb.Close($false)
  Write-Host "excel ok"
} finally { $excel.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }

[System.GC]::Collect()
Get-ChildItem $out | ForEach-Object { Write-Host ("  " + $_.Name + "  " + $_.Length) }
