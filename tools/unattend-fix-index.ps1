# unattend-fix-index.ps1 -- Patch the VBox unattended floppy's autounattend.xml:
# /IMAGE/INDEX 1 (HomeBasic) -> 4 (Ultimate) on the cn Ultimate x64 SP1 DVD wim.
# Single-byte in-place patch inside the raw floppy image (same length, FAT intact).
$repo = Split-Path -Parent $MyInvocation.MyCommand.Path
$img = "C:\Users\Administrator\VirtualBox VMs\Win7SP1x64\Unattended-03251e83-119e-4177-938a-0116a79aea57-aux-floppy.img"
$bytes = [IO.File]::ReadAllBytes($img)
$text = [Text.Encoding]::ASCII.GetString($bytes)
$old = "<Key>/IMAGE/INDEX</Key><Value>1</Value>"
# The XML on the floppy contains newline+indent between </Key> and <Value>; search both parts:
$k = $text.IndexOf("/IMAGE/INDEX</Key>")
if ($k -lt 0) { Write-Host "marker /IMAGE/INDEX not found"; exit 2 }
$v = $text.IndexOf("<Value>1</Value>", $k)
if ($v -lt 0 -or ($v - $k) -gt 80) { Write-Host ("Value node not adjacent (k=" + $k + " v=" + $v + ")"); exit 2 }
$bytes[$v + 7] = 0x34    # '1' -> '4'  (offset of the digit inside <Value>1</Value> is +7)
$fs = [IO.File]::Open($img, 'Open', 'ReadWrite')
$fs.Seek($v, 'Begin') | Out-Null
$fs.Write($bytes, $v, 8)
$fs.Close()
Write-Host ("patched: byte at " + ($v + 7) + " now '4'; context=" + $text.Substring($k, 60))
