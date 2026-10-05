# 诊断：OfficeAgent 窗口位置/是否被遮挡/能否抬到 z 序顶层（不抢焦点）
Add-Type -AssemblyName System.Windows.Forms,System.Drawing
Add-Type '
using System;using System.Runtime.InteropServices;
public class W3 {
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out R r);
  [DllImport("user32.dll")] public static extern IntPtr WindowFromPoint(POINT p);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
  [DllImport("user32.dll")] public static extern int GetWindowTextLength(IntPtr h);
  public struct R { public int L,T,Rt,B; }
  public struct POINT { public int X,Y; }
}'
$p = Get-Process OfficeAgent | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
$h = $p.MainWindowHandle
$r = New-Object 'W3+R'
[W3]::GetWindowRect($h, [ref]$r) | Out-Null
$wa = [W3]::GetForegroundWindow()
$fpid = 0
[W3]::GetWindowThreadProcessId($wa, [ref]$fpid) | Out-Null
$fname = try { (Get-Process -Id $fpid).ProcessName } catch { "?" }
Write-Host ("OA hwnd=" + $h + " rect=" + $r.L + "," + $r.T + "," + $r.Rt + "," + $r.B + " iconic=" + [W3]::IsIconic($h))
Write-Host ("foreground pid=" + $fpid + " name=" + $fname)
$sc = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
Write-Host ("screen=" + $sc.Width + "x" + $sc.Height)
# 窗口中心点落在谁身上
$pt = New-Object 'W3+POINT'
$pt.X = [int](($r.L + $r.Rt) / 2); $pt.Y = [int](($r.T + $r.B) / 2)
$under = [W3]::WindowFromPoint($pt)
$upid = 0
[W3]::GetWindowThreadProcessId($under, [ref]$upid) | Out-Null
$uname = try { (Get-Process -Id $upid).ProcessName } catch { "?" }
Write-Host ("point " + $pt.X + "," + $pt.Y + " -> hwnd=" + $under + " pid=" + $upid + " name=" + $uname)
# 尝试抬到顶层：HWND_TOP=0, flags = NOSIZE|NOMOVE|NOACTIVATE
$ok = [W3]::SetWindowPos($h, [IntPtr]::Zero, 0, 0, 0, 0, 0x0001 -bor 0x0002 -bor 0x0010)
Write-Host ("raise noactivate ok=" + $ok)
Start-Sleep -Milliseconds 500
$under2 = [W3]::WindowFromPoint($pt)
$upid2 = 0
[W3]::GetWindowThreadProcessId($under2, [ref]$upid2) | Out-Null
$uname2 = try { (Get-Process -Id $upid2).ProcessName } catch { "?" }
Write-Host ("after raise: point -> " + $uname2 + " (pid " + $upid2 + ")")
$fg2 = [W3]::GetForegroundWindow()
$fpid2 = 0
[W3]::GetWindowThreadProcessId($fg2, [ref]$fpid2) | Out-Null
$fname2 = try { (Get-Process -Id $fpid2).ProcessName } catch { "?" }
Write-Host ("foreground after=" + $fname2)
# 截图 OfficeAgent 区域与全屏
$bmp = New-Object System.Drawing.Bitmap(($r.Rt - $r.L), ($r.B - $r.T))
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($r.L, $r.T, 0, 0, $bmp.Size)
$bmp.Save("$env:TEMP\oa_probe_oa.png"); $g.Dispose(); $bmp.Dispose()
$full = New-Object System.Drawing.Bitmap($sc.Width, $sc.Height)
$g2 = [System.Drawing.Graphics]::FromImage($full)
$g2.CopyFromScreen(0, 0, 0, 0, $full.Size)
$full.Save("$env:TEMP\oa_probe_full.png"); $g2.Dispose(); $full.Dispose()
Write-Host "done"
