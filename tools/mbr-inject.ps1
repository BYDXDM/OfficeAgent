# mbr-inject.ps1 -- Write a single-partition MBR into the raw disk image.
# Partition: bootable, type 07 (NTFS), LBA start 2048, covering the rest of the disk.
# NOTE: keep ASCII (PS 5.1 parses BOM-less UTF-8 as GBK); repo root derived from script location.
$repo = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$raw = Join-Path (Join-Path $repo "build\vm") "win7.raw"
$mbr = New-Object byte[] 512
$mbr[446] = 0x80                                    # bootable
$mbr[450] = 0x07                                    # NTFS/exFAT
$start = [UInt64]2048
$sectors = [UInt64]64424509440 / [UInt64]512 - $start
$b1 = [BitConverter]::GetBytes([UInt32]$start)
$b2 = [BitConverter]::GetBytes([UInt32]$sectors)
[Array]::Copy($b1, 0, $mbr, 454, 4)                 # LBA start (LE)
[Array]::Copy($b2, 0, $mbr, 458, 4)                 # sector count (LE)
$mbr[510] = 0x55
$mbr[511] = 0xAA
$fs = [IO.File]::Open($raw, 'Open', 'ReadWrite')
$fs.Seek(0, 'Begin') | Out-Null
$fs.Write($mbr, 0, 512)
$fs.Close()
Write-Host ("MBR injected: start=" + $start + " sectors=" + $sectors + " raw=" + $raw)
