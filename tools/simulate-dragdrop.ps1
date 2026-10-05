# drag-drop simulation: focus OA, start drag source (waits for handshake flag),
# inject mouse (hold left button -> move -> release), screenshot.
# ASCII-only on purpose: PS 5.1 reads .ps1 without BOM as ANSI.
Add-Type -AssemblyName System.Windows.Forms,System.Drawing
Add-Type '
using System;using System.Runtime.InteropServices;
public class M9 { [DllImport("user32.dll")] public static extern bool SetCursorPos(int x,int y);
  [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint dx, uint dy, uint d, UIntPtr e);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out R r);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  public struct R { public int L,T,Rt,B; } }'

$p = Get-Process OfficeAgent | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
if (-not $p) { Write-Host "NO OA PROCESS"; exit 1 }
$r = New-Object 'M9+R'
[M9]::GetWindowRect($p.MainWindowHandle, [ref]$r) | Out-Null

# 1) click title bar (left side, away from buttons) to take foreground
[M9]::SetCursorPos($r.L + 250, $r.T + 12) | Out-Null
Start-Sleep -Milliseconds 120
[M9]::mouse_event(0x0002,0,0,0,[UIntPtr]::Zero)
Start-Sleep -Milliseconds 50
[M9]::mouse_event(0x0004,0,0,0,[UIntPtr]::Zero)
Start-Sleep -Milliseconds 400

# 2) write drag-source script (STA; waits for handshake flag before DoDragDrop;
#    logs the returned drop effect)
$srcPath = "$env:TEMP\oa_dragsrc2.ps1"
$src = @'
Add-Type -AssemblyName System.Windows.Forms,System.Drawing
$f = New-Object System.Windows.Forms.Form
$f.FormBorderStyle = "None"; $f.StartPosition = "Manual"
$f.Location = New-Object System.Drawing.Point(432, 200)
$f.Size = New-Object System.Drawing.Size(80, 40)
$f.TopMost = $true; $f.BackColor = [System.Drawing.Color]::FromArgb(62,99,221)
$lbl = New-Object System.Windows.Forms.Label; $lbl.Text = "DRAG-SRC"; $lbl.ForeColor = "White"; $lbl.Dock = "Fill"
$f.Controls.Add($lbl)
$f.Add_Shown({
  $f.BeginInvoke([Action]{
    $go = "$env:TEMP\oa_go.flag"
    $waited = 0
    while (-not (Test-Path $go) -and $waited -lt 6000) { Start-Sleep -Milliseconds 100; $waited += 100 }
    Remove-Item $go -ErrorAction SilentlyContinue
    $p1 = $env:TEMP + "\1.frp"
    $data = New-Object System.Windows.Forms.DataObject("FileDrop", (,[string]$p1))
    $eff = [System.Windows.Forms.Control]::DoDragDrop($lbl, $data, [System.Windows.Forms.DragDropEffects]::Copy)
    Set-Content "$env:TEMP\oa_drag_result.txt" ("effect=" + $eff) -Encoding UTF8
    $f.Close()
  }) | Out-Null
})
[void]$f.ShowDialog()
'@
Set-Content $srcPath $src -Encoding UTF8
Start-Process powershell -ArgumentList '-NoProfile','-STA','-ExecutionPolicy','Bypass','-File',$srcPath -WindowStyle Hidden
Start-Sleep -Milliseconds 1200

# 3) injector: hold LEFT button, signal source, move to OA chat area, release
$srcX = 472; $srcY = 220
$tgtX = [int](($r.L + $r.Rt) / 2); $tgtY = [int](($r.T + $r.B) / 2 + 60)
[M9]::SetCursorPos($srcX, $srcY) | Out-Null
Start-Sleep -Milliseconds 250
[M9]::mouse_event(0x0002,0,0,0,[UIntPtr]::Zero)          # LEFT DOWN
Set-Content "$env:TEMP\oa_go.flag" "go"                   # signal source
Start-Sleep -Milliseconds 400                             # DoDragDrop starts while held
for ($s = 1; $s -le 20; $s++) {
  $x = $srcX + [int](($tgtX - $srcX) * $s / 20)
  $y = $srcY + [int](($tgtY - $srcY) * $s / 20)
  [M9]::SetCursorPos($x, $y) | Out-Null
  Start-Sleep -Milliseconds 55
}
Start-Sleep -Milliseconds 500
[M9]::mouse_event(0x0004,0,0,0,[UIntPtr]::Zero)          # LEFT UP (drop)
Start-Sleep -Milliseconds 600

# 4) screenshot OA
$bmp = New-Object System.Drawing.Bitmap(($r.Rt-$r.L), ($r.B-$r.T))
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($r.L, $r.T, 0, 0, $bmp.Size)
$bmp.Save("$env:TEMP\oa_afterdrop4.png")
Write-Host 'done'
