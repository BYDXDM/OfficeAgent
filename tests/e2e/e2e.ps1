# OfficeAgent E2E v3 —— 纯坐标模拟真人点击/输入 + UIA(PID 过滤)文本断言 + 截图 + 磁盘产物断言
# 向导为固定 560x452 居中对话框；主窗口布局坐标与源码一致。
# 用法：powershell -NoProfile -STA -ExecutionPolicy Bypass -File e2e.ps1
$ErrorActionPreference = "Continue"
$repo = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path))

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class W32 {
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr h, out R r);
  [DllImport("user32.dll")] public static extern bool ClientToScreen(IntPtr h, ref P p);
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
  [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint dx, uint dy, uint d, IntPtr e);
  [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, System.Text.StringBuilder sb, int max);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, System.Text.StringBuilder sb, int max);
  public struct R { public int L, T, Rt, B; }
  public struct P { public int X, Y; }
  public static IntPtr FindDialog(uint targetPid, string titleLike) {
    IntPtr found = IntPtr.Zero;
    EnumWindows(delegate(IntPtr h, IntPtr l) {
      uint pid; GetWindowThreadProcessId(h, out pid);
      if (pid == targetPid && IsWindowVisible(h)) {
        var cn = new System.Text.StringBuilder(256); GetClassName(h, cn, 256);
        if (cn.ToString() == "#32770") {
          var sb = new System.Text.StringBuilder(256); GetWindowText(h, sb, 256);
          if (sb.ToString().Contains(titleLike)) { found = h; return false; }
        }
      }
      return true;
    }, IntPtr.Zero);
    return found;
  }
}
"@

$auto = [System.Windows.Automation.AutomationElement]
$exe        = Join-Path $repo "build\host\OfficeAgent.exe"
$mockLog    = Join-Path $repo "tests\e2e\mock-server.log"
$shotDir    = Join-Path $repo "build\test\e2e"
$cfgFile    = Join-Path $env:LOCALAPPDATA "OfficeAgent\config.json"
$sampleCsv  = Join-Path $repo "tests\e2e\sample.csv"
$pdfFixture = Join-Path $repo "tests\fixtures\test-smoke.pdf"
$outXlsx    = Join-Path $repo "tests\e2e\sample_conv.xlsx"

$script:pass = 0; $script:fail = 0
function Note($ok, $step, $detail) {
  if ($ok) { $script:pass++; Write-Host ("[PASS] " + $step + "  " + $detail) }
  else     { $script:fail++; Write-Host ("[FAIL] " + $step + "  " + $detail) }
}
function Shot($proc, $name) {
  try {
    [void][W32]::SetForegroundWindow($proc.MainWindowHandle)
    Start-Sleep -Milliseconds 500
    $r = New-Object W32+R
    [void][W32]::GetClientRect($proc.MainWindowHandle, [ref]$r)
    $po = New-Object W32+P; $po.X = 0; $po.Y = 0
    [void][W32]::ClientToScreen($proc.MainWindowHandle, [ref]$po)
    $bmp = New-Object System.Drawing.Bitmap($r.Rt, $r.B)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($po.X, $po.Y, 0, 0, $bmp.Size)
    $g.Dispose()
    $bmp.Save((Join-Path $shotDir $name), [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    Write-Host ("[SHOT] " + $name)
  } catch { Write-Host ("[SHOT-ERR] " + $_.Exception.Message) }
}
function Click-XY($hwnd, $cx, $cy) {
  $po = New-Object W32+P; $po.X = [int]$cx; $po.Y = [int]$cy
  [void][W32]::ClientToScreen($hwnd, [ref]$po)
  [void][W32]::SetCursorPos($po.X, $po.Y)
  Start-Sleep -Milliseconds 150
  [void][W32]::mouse_event(2, 0, 0, 0, [IntPtr]::Zero)
  [void][W32]::mouse_event(4, 0, 0, 0, [IntPtr]::Zero)
}
# 前台确保 + 点击（模拟器焦点被抢时，点击会落在别的窗口上 → 一律先拉前台）
function Click-Fg($proc, $cx, $cy) {
  [void][W32]::SetForegroundWindow($proc.MainWindowHandle)
  Start-Sleep -Milliseconds 200
  Click-XY $proc.MainWindowHandle $cx $cy
}
function Type-Clip($text) {
  Set-Clipboard -Value $text
  Start-Sleep -Milliseconds 250
  [System.Windows.Forms.SendKeys]::SendWait("^v")
  Start-Sleep -Milliseconds 250
}
function Paste-Enter($text) {
  Type-Clip $text
  [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
}
function Wait-Until($script, $sec) {
  $deadline = (Get-Date).AddSeconds($sec)
  while ((Get-Date) -lt $deadline) { if (& $script) { return $true } Start-Sleep -Milliseconds 400 }
  return $false
}
function Count-LogPost($log, $pattern) {
  if (-not (Test-Path $log)) { return 0 }
  return @(Select-String -Path $log -Pattern $pattern -SimpleMatch).Count
}
function Count-Audit($audit, $pattern) {
  if (-not (Test-Path $audit)) { return 0 }
  return @(Select-String -Path $audit -Pattern $pattern -SimpleMatch).Count
}
function Find-NameLike($proc, $pattern, $sec) {
  $deadline = (Get-Date).AddSeconds($sec)
  while ((Get-Date) -lt $deadline) {
    $cond = New-Object System.Windows.Automation.PropertyCondition($auto::ProcessIdProperty, $proc.Id)
    $all = $auto::RootElement.FindAll([System.Windows.Automation.TreeScope]::Descendants, $cond)
    foreach ($e in $all) {
      if ($e.Current.IsOffscreen) { continue }
      if ($e.Current.Name -like $pattern) { return $e }
    }
    Start-Sleep -Milliseconds 400
  }
  return $null
}

function Wait-AppDialog($proc, $titleLike, $sec) {
  $deadline = (Get-Date).AddSeconds($sec)
  while ((Get-Date) -lt $deadline) {
    $h = [W32]::FindDialog([uint32]$proc.Id, $titleLike)
    if ($h -ne [IntPtr]::Zero) { return $h }
    Start-Sleep -Milliseconds 300
  }
  return [IntPtr]::Zero
}

function Wait-AppDialogGone($proc, $titleLike, $sec) {
  $deadline = (Get-Date).AddSeconds($sec)
  while ((Get-Date) -lt $deadline) {
    $h = [W32]::FindDialog([uint32]$proc.Id, $titleLike)
    if ($h -eq [IntPtr]::Zero) { return $true }
    Start-Sleep -Milliseconds 300
  }
  return $false
}

# ---------- 准备 ----------
New-Item -ItemType Directory -Force -Path $shotDir | Out-Null
Get-Process OfficeAgent -ErrorAction SilentlyContinue | Stop-Process -Force
try {
  Get-NetTCPConnection -LocalPort 18080 -State Listen -ErrorAction SilentlyContinue | ForEach-Object { Stop-Process -Id $_.OwningProcess -Force }
} catch {}
Start-Sleep -Milliseconds 500
if (Test-Path $cfgFile) { Remove-Item $cfgFile -Force }
if (Test-Path $mockLog) { Remove-Item $mockLog -Force }
if (Test-Path $outXlsx) { Remove-Item $outXlsx -Force }
Remove-Item (Join-Path $env:LOCALAPPDATA "OfficeAgent\memory.json") -Force -ErrorAction SilentlyContinue
Remove-Item (Join-Path $env:LOCALAPPDATA "OfficeAgent\history.jsonl") -Force -ErrorAction SilentlyContinue
Set-Content -LiteralPath $sampleCsv -Value "科目,借方,贷方,日期`r`n库存现金,100.50,0,2026-09-19`r`n银行存款,0,100.50,2026-09-19" -Encoding UTF8
Start-Process powershell -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", (Join-Path $repo "tests\e2e\mock-llm.ps1") -WindowStyle Hidden
Start-Sleep -Seconds 2

# ---------- 0. CLI 冒烟：M2 引擎无头验收 ----------
$flowCsv   = Join-Path $repo "tests\fixtures\recon\flow.csv"
$ledgerCsv = Join-Path $repo "tests\fixtures\recon\ledger.csv"
$reconOut  = Join-Path $repo "tests\fixtures\recon\核对_flow_vs_ledger.xlsx"
if (Test-Path $reconOut) { Remove-Item $reconOut -Force }
$rc = (Start-Process -FilePath $exe -ArgumentList "/recon", $flowCsv, $ledgerCsv, "--keya", "账号", "--keyb", "对方账号", "--debita", "收入", "--credita", "支出", "--debitb", "借方", "--creditb", "贷方", "--tol", "0.01", "--all" -PassThru -Wait -WindowStyle Hidden).ExitCode
Note ($rc -eq 0) "0.1 CLI /recon 核对引擎（差异集 16/2/3/2）" ("exit=" + $rc)
Note (Test-Path $reconOut) "0.2 差异表产物存在" ""
$rc = (Start-Process -FilePath $exe -ArgumentList "/masktest" -PassThru -Wait -WindowStyle Hidden).ExitCode
Note ($rc -eq 0) "0.3 CLI /masktest 脱敏引擎" ("exit=" + $rc)
# 光标定位回归：修复前 EM_POSFROMCHAR 失败被当坐标 → 光标塌到第一个字符之前
$rc = (Start-Process -FilePath $exe -ArgumentList "/carettest" -PassThru -Wait -WindowStyle Hidden).ExitCode
Note ($rc -eq 0) "0.3b CLI /carettest 光标定位（行尾/多行/非负）" ("exit=" + $rc)
$sub1 = Join-Path $repo "build\test\e2e-sub1.csv"; $sub2 = Join-Path $repo "build\test\e2e-sub2.csv"
$mergeOut = Join-Path $repo "build\test\e2e-merge.xlsx"
Set-Content -LiteralPath $sub1 -Value "科目,金额,备注`r`n差旅费,1234.50,北京`r`n办公费,880.00,打印纸" -Encoding UTF8
Set-Content -LiteralPath $sub2 -Value "备注,科目,金额`r`n年会,会务费,2000.00`r`n上海,差旅费,765.40" -Encoding UTF8
$rc = (Start-Process -FilePath $exe -ArgumentList "/merge", $sub1, $sub2, "--targets", "科目,金额", "--srcs", "科目,金额", "--amounts", "金额", "--out", $mergeOut -PassThru -Wait -WindowStyle Hidden).ExitCode
Note ($rc -eq 0 -and (Test-Path $mergeOut)) "0.4 CLI /merge 多簿汇总（列序不同按表头名归集）" ("exit=" + $rc)
$invOut = Join-Path $repo "build\test\e2e-invoice.xlsx"
$rc = (Start-Process -FilePath $exe -ArgumentList "/invoices", (Join-Path $repo "tests\fixtures\invoice-sample.pdf"), "--out", $invOut -PassThru -Wait -WindowStyle Hidden).ExitCode
Note ($rc -eq 0 -and (Test-Path $invOut)) "0.5 CLI /invoices 发票提取" ("exit=" + $rc)
# xlsx 网格读取回归：曾因 StreamRows 每节点重复提交 + LoadGrid 复用缓冲，
# 导致表头丢失、各行都显示最后一行（预览页表现为只有一列、数据重复）
$gridOut = Join-Path $repo "build\test\e2e-grid.txt"
$gridCsv = Join-Path $repo "build\test\e2e-grid.csv"
$gridXlsx = Join-Path $repo "build\test\e2e-grid_conv.xlsx"   # /convert 产物带 _conv 后缀
Set-Content -LiteralPath $gridCsv -Value "科目,金额`r`n库存现金,100.50`r`n银行存款,9876.00" -Encoding UTF8
$null = (Start-Process -FilePath $exe -ArgumentList "/convert", $gridCsv, "xlsx" -PassThru -Wait -WindowStyle Hidden).ExitCode
$rc = (Start-Process -FilePath $exe -ArgumentList "/gridtest", $gridXlsx -PassThru -Wait -WindowStyle Hidden -RedirectStandardOutput $gridOut).ExitCode
$gridTxt = if (Test-Path $gridOut) { Get-Content $gridOut -Raw -Encoding UTF8 } else { "" }   # exe 输出是 UTF-8，PS5.1 默认按 ANSI 读会变乱码
Note (($rc -eq 0) -and ($gridTxt -like '*表头行=科目|金额*') -and ($gridTxt -like '*总行数=3*')) "0.5b xlsx 网格读取（表头保留、行不重复）" ("exit=" + $rc)
$srOut = Join-Path $repo "build\test\e2e-skillrun.txt"
# 确认门禁：不带 --yes 必须先拒绝（exit=2，零执行）；带 --yes 才真正执行
$rcGate = (Start-Process -FilePath $exe -ArgumentList "/skillrun", "row-stat", $flowCsv -PassThru -Wait -WindowStyle Hidden -RedirectStandardOutput $srOut).ExitCode
$gateTxt = if (Test-Path $srOut) { Get-Content $srOut -Raw } else { "" }
Note (($rcGate -eq 2) -and ($gateTxt -notlike '*"ok": true*')) "0.6a CLI /skillrun 无 --yes 被拒（确认门禁）" ("exit=" + $rcGate)

$rc = (Start-Process -FilePath $exe -ArgumentList "/skillrun", "row-stat", $flowCsv, "--yes" -PassThru -Wait -WindowStyle Hidden -RedirectStandardOutput $srOut).ExitCode
$srTxt = if (Test-Path $srOut) { Get-Content $srOut -Raw } else { "" }
Note (($srTxt -like '*"ok": true*') -or ($rc -eq 0)) "0.6b CLI /skillrun --yes 执行 python 技能（row-stat）" ("exit=" + $rc)

# ---------- 1. 启动 ----------
$p = Start-Process -FilePath $exe -PassThru
Start-Sleep -Seconds 6
$win = New-Object W32+R
[void][W32]::GetClientRect($p.MainWindowHandle, [ref]$win)
$org = New-Object W32+P; $org.X = 0; $org.Y = 0
[void][W32]::ClientToScreen($p.MainWindowHandle, [ref]$org)
$cw = $win.Rt; $ch = $win.B
$contentTop = 52; $contentH = $ch - $contentTop
Note ($cw -gt 800) "1.1 启动主窗口" ("client=" + $cw + "x" + $ch)
# 向导居中于主窗口客户区
$dx = [int]($cw / 2) - 280
$dy = [int]($ch / 2) - 226
Shot $p "01-wizard-fresh.png"

# ---------- 2. 向导 ----------
Click-Fg $p ($dx + 40) ($dy + 336)                            # 勾选"允许内网端点"
Start-Sleep -Milliseconds 300
Click-Fg $p ($dx + 280) ($dy + 153)                           # URL 输入框
Type-Clip "http://127.0.0.1:18080/v1"
Click-Fg $p ($dx + 280) ($dy + 211)                           # Key 输入框
Type-Clip "e2e-fake-key"
$okFetch = $null
for ($i = 0; $i -lt 3 -and ($okFetch -eq $null); $i++) {
  Click-Fg $p ($dx + 40) ($dy + 336)                          # 允许内网/本机端点（首击易被窗口激活吞掉，循环内重勾）
  Start-Sleep -Milliseconds 200
  Click-Fg $p ($dx + 88) ($dy + 252)                          # 获取模型列表
  $okFetch = Find-NameLike $p "✓ 获取到*" 8
}
Note ($okFetch -ne $null) "2.1 自动获取模型列表（mock 返回 2 个模型）" ""
$okTest = $null
for ($i = 0; $i -lt 3 -and ($okTest -eq $null); $i++) {
  Click-Fg $p ($dx + 88) ($dy + 290)                          # 测试连接
  $okTest = Find-NameLike $p "✓ 模型可用*" 20
}
Note ($okTest -ne $null) "2.2 测试连接（mock 真实应答）" ""
Shot $p "02-wizard-filled.png"
Click-Fg $p ($dx + 447) ($dy + 456)                           # 保存并开始使用（按钮 y=440+16；工作区行加入后从 412 下移）
Start-Sleep -Seconds 2
Note ((Find-NameLike $p "三步开始*" 2) -eq $null) "2.3 保存并关闭向导" ""

# ---------- 3. 会话发消息 ----------
[void][W32]::SetForegroundWindow($p.MainWindowHandle)
Click-Fg $p 114 129                # 会话（侧边栏会话列表首行）
Start-Sleep -Milliseconds 700
$memFile = Join-Path $env:LOCALAPPDATA "OfficeAgent\memory.json"
$chatInputY = $contentTop + $contentH - 118 + 34
# 3.0 记忆模块命令（离线可用）："记住 我是一名会计"——粘贴带重试（模拟器偶发剪贴板竞态）
$okMem = $false
for ($i = 0; $i -lt 3 -and -not $okMem; $i++) {
  Click-Fg $p 500 $chatInputY
  Paste-Enter "记住 我是一名会计"
  $okMem = Wait-Until { Test-Path $memFile } 8
}
Note $okMem "3.0 记忆模块：「记住」命令落盘 memory.json" ""
# 3.1 正常对话（模型应答）——以 mock 日志新增 POST 为准（聊天气泡是自绘项，UIA 不可读）
$basePosts = Count-LogPost $mockLog "REQ POST /v1/chat/completions"
$okChat = $false
for ($i = 0; $i -lt 3 -and -not $okChat; $i++) {
  Click-Fg $p 500 $chatInputY
  Paste-Enter "你好，E2E 测试"
  $okChat = Wait-Until { (Count-LogPost $mockLog "REQ POST /v1/chat/completions") -ge ($basePosts + 1) } 15
}
Note ($okChat) "3.1 对话请求到达 mock（model=mock-a）" ""
Start-Sleep -Seconds 2
$logTxt = if (Test-Path $mockLog) { Get-Content $mockLog -Raw } else { "" }
Note ($logTxt -like '*"tools"*' -or $logTxt -like '*function*') "3.2 Agent 工具声明随请求发送（function calling）" ""
Note ($logTxt -like "*会计*") "3.3 长期记忆注入模型提示词" ""
Shot $p "03-chat-reply.png"
# 3.4 上下文回归（用户实机抓到的 bug：模型答"这是我们的第一次对话"）——
# 发 HISTTEST，mock 数请求里的 user 消息：>=2 = 模型真收到历史；回 HIST-OK 落进 history.jsonl 断言
$histFile = Join-Path $env:LOCALAPPDATA "OfficeAgent\history.jsonl"
$histBase = if (Test-Path $histFile) { @(Get-Content $histFile).Count } else { 0 }
$okHist = $false
for ($i = 0; $i -lt 3 -and -not $okHist; $i++) {
  Click-Fg $p 500 $chatInputY
  Paste-Enter "HISTTEST 我上一句说了什么"
  $okHist = Wait-Until {
    (Test-Path $histFile) -and (@(Get-Content $histFile).Count -ge ($histBase + 2)) -and
    ((Get-Content $histFile -Raw) -like '*HIST-OK*')
  } 20
}
Note ($okHist) "3.4 多轮上下文真实到达模型（HIST-OK）" ""

# ---------- 4. 任务台：全键盘流（中性点击 → TAB → 方向键选 XLSX → TAB+空格开对话框） ----------
Click-Fg $p 114 316                                 # 导航:任务台
Start-Sleep -Milliseconds 700
$dlg4 = [IntPtr]::Zero
for ($i = 0; $i -lt 3 -and ($dlg4 -eq [IntPtr]::Zero -or $dlg4 -eq $null); $i++) {
  Click-Fg $p 620 260                                         # 中性区域点击（聚焦 ListView，不触发下拉）
  Start-Sleep -Milliseconds 400
  [System.Windows.Forms.SendKeys]::SendWait("{TAB}")          # → 目标组合框
  Start-Sleep -Milliseconds 300
  [System.Windows.Forms.SendKeys]::SendWait("{DOWN}{DOWN}")   # 闭态直接改选 → XLSX（不开下拉）
  Start-Sleep -Milliseconds 300
  [System.Windows.Forms.SendKeys]::SendWait("{TAB}")          # → 添加文件
  Start-Sleep -Milliseconds 300
  [System.Windows.Forms.SendKeys]::SendWait(" ")              # 空格按下 → 文件对话框
  $dlg4 = Wait-AppDialog $p "要转换的文件" 8
}
Note ($dlg4 -ne [IntPtr]::Zero) "4.1a 文件对话框就绪" ""
if ($dlg4 -eq [IntPtr]::Zero) { Shot $p "04a-dialog-missing.png" }
if ($dlg4 -ne [IntPtr]::Zero) { [void][W32]::SetForegroundWindow($dlg4); Start-Sleep -Milliseconds 300 }
Paste-Enter $sampleCsv
[void](Wait-AppDialogGone $p "要转换的文件" 10)
Start-Sleep -Seconds 1
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")            # 焦点回添加文件 → TAB → 开始转换
Start-Sleep -Milliseconds 300
[System.Windows.Forms.SendKeys]::SendWait(" ")
$okQueue = Find-NameLike $p "队列完成：成功 1*" 35
Note ($okQueue -ne $null) "4.1 转换队列完成（成功 1）" ""
Note (Test-Path $outXlsx) "4.2 转换产物存在" "sample_conv.xlsx"
Shot $p "04-task-done.png"

# ---------- 4b. 表格核对页：填参数 → 核对 → 产物断言 ----------
Click-Fg $p 114 356                # 导航:表格核对
Start-Sleep -Milliseconds 800
Click-Fg $p 412 (52 + 51)                    # 文件A 输入框
Paste-Enter $flowCsv
Start-Sleep -Milliseconds 400
Click-Fg $p 812 (52 + 107)                   # 文件B 输入框（B 块整体低一行）
Type-Clip $ledgerCsv
Start-Sleep -Milliseconds 400
Click-Fg $p 327 (52 + 183); Type-Clip "账号"          # A 键列
Start-Sleep -Milliseconds 300
Click-Fg $p 527 (52 + 183); Type-Clip "对方账号"      # B 键列
Start-Sleep -Milliseconds 300
Click-Fg $p 307 (52 + 257); Type-Clip "收入"          # A 借方列
Start-Sleep -Milliseconds 300
Click-Fg $p 437 (52 + 257); Type-Clip "支出"          # A 贷方列
Start-Sleep -Milliseconds 300
Click-Fg $p 577 (52 + 257); Type-Clip "借方"          # B 借方列
Start-Sleep -Milliseconds 300
Click-Fg $p 707 (52 + 257); Type-Clip "贷方"          # B 贷方列
Start-Sleep -Milliseconds 300
if (Test-Path $reconOut) { Remove-Item $reconOut -Force }
Click-Fg $p 312 (52 + 459)                   # 开始核对
$okRecon = Find-NameLike $p "完成（*" 40
Note ($okRecon -ne $null) "4b.1 表格核对页运行完成（Verifier 全过）" ""
Note (Test-Path $reconOut) "4b.2 UI 核对差异表产物存在" ""
Shot $p "04b-recon.png"

# ---------- 4c. M3 自然语言任务：计划展示 → 取消 → 再计划 → 确认执行 ----------
# 聊天气泡为自绘 ListBox 项（UIA 不可读），断言一律走 audit.jsonl / 磁盘产物
[void][W32]::SetForegroundWindow($p.MainWindowHandle)
Click-Fg $p 114 129                # 会话（侧边栏会话列表首行）
Start-Sleep -Milliseconds 800
$auditFile = Join-Path $env:LOCALAPPDATA "OfficeAgent\audit.jsonl"
$chatInputY = $contentTop + $contentH - 118 + 34
$reconDir   = Join-Path $repo "tests\fixtures\recon"
$beforeCount = @(Get-ChildItem -Path $reconDir -Filter "核对_*.xlsx" -ErrorAction SilentlyContinue).Count
# 基线必须与断言用同一个 pattern：此前基线用泛化的 "plan_created"（含 /skillrun 写的
# "skill ...; cli" 行），而断言用 Recon 专属 pattern，两个计数基数不同 → 断言恒不成立。
$planPattern = 'plan_created","detail":"Recon; executable=True'
$planBase   = Count-Audit $auditFile $planPattern
$confBase   = Count-Audit $auditFile "confirmation"
$fnBase     = Count-Audit $auditFile "action_finished"
# 4c.1 计划创建（audit 行内同时含 plan_created 与 executable=True → 映射模板已复用）
$okPlan = $false
for ($i = 0; $i -lt 3 -and -not $okPlan; $i++) {
  Click-Fg $p 500 $chatInputY
  Paste-Enter ("核对 " + $flowCsv + " 和 " + $ledgerCsv)
  $okPlan = Wait-Until { (Count-Audit $auditFile $planPattern) -ge ($planBase + 1) } 15
}
Note ($okPlan) "4c.1 计划创建且复用 __last 模板（executable=True）" ""
$planNow = Count-Audit $auditFile "plan_created"
# 4c.3 取消 → 零产物
$okCancel = $false
for ($i = 0; $i -lt 3 -and (Count-Audit $auditFile "confirmation") -lt ($confBase + 1); $i++) {
  Click-Fg $p 500 $chatInputY
  Paste-Enter "取消"
  $okCancel = Wait-Until { (Count-Audit $auditFile "confirmation") -ge ($confBase + 1) } 10
}
Note ($okCancel) "4c.3 取消确认（audit: confirmation=cancelled）" ""
$afterCancel = @(Get-ChildItem -Path $reconDir -Filter "核对_*.xlsx" -ErrorAction SilentlyContinue).Count
Note ($afterCancel -eq $beforeCount) "4c.4 取消后零产物" ""
# 4c.5 再计划 → 确认 → ReconEngine 真执行
Click-Fg $p 500 $chatInputY
Paste-Enter ("核对 " + $flowCsv + " 和 " + $ledgerCsv)
[void](Wait-Until { (Count-Audit $auditFile "plan_created") -ge ($planNow + 1) } 20)
Click-Fg $p 500 $chatInputY
Paste-Enter "确认"
$okRun = Wait-Until { (Count-Audit $auditFile "action_finished") -ge ($fnBase + 1) } 60
Note ($okRun) "4c.5 确认后真执行（ReconEngine）" ""
Start-Sleep -Seconds 1
$afterRun = @(Get-ChildItem -Path $reconDir -Filter "核对_*.xlsx" -ErrorAction SilentlyContinue).Count
Note ($afterRun -gt $beforeCount) "4c.6 确认执行产物落盘" ("count=" + $beforeCount + "→" + $afterRun)
$auditTail = if (Test-Path $auditFile) { (Get-Content $auditFile -Tail 60) -join "`n" } else { "" }
Note ($auditTail -like "*plan_created*" -and $auditTail -like "*confirmation*") "4c.7 审计含 plan_created/confirmation" ""
Note ($auditTail -like "*action_started*" -and $auditTail -like "*action_finished*") "4c.8 审计含 action_started/finished" ""
Shot $p "04c-chat-action.png"

# ---------- 5. 预览 PDF ----------
Click-Fg $p 114 436                                 # 导航:预览
Start-Sleep -Milliseconds 700
$dlg5 = [IntPtr]::Zero
for ($i = 0; $i -lt 3 -and ($dlg5 -eq [IntPtr]::Zero -or $dlg5 -eq $null); $i++) {
  Click-Fg $p (228 + 79) ($contentTop + 21)                   # 打开预览文件
  $dlg5 = Wait-AppDialog $p "预览文件" 8
}
Note ($dlg5 -ne [IntPtr]::Zero) "5.0 预览对话框就绪" ""
if ($dlg5 -ne [IntPtr]::Zero) { [void][W32]::SetForegroundWindow($dlg5); Start-Sleep -Milliseconds 300 }
Paste-Enter $pdfFixture
[void](Wait-AppDialogGone $p "预览文件" 10)
Start-Sleep -Seconds 3
$okPage = Find-NameLike $p "第 1 / 1 页" 10
Note ($okPage -ne $null) "5.1 PDF 渲染（页码标签出现）" ""
Shot $p "05-preview-pdf.png"

# ---------- 6. 汇总·发票页（载入断言） + 系统状态 ----------
Click-Fg $p 114 396                                 # 导航:汇总·发票
Start-Sleep -Milliseconds 700
$okMerge = Find-NameLike $p "*报表汇总*" 6
Note ($okMerge -ne $null) "6.0 汇总·发票页载入" ""
Shot $p "06-extract.png"
Click-Fg $p 114 516                                 # 导航:系统状态
Start-Sleep -Milliseconds 700
Click-Fg $p (228 + 68) ($contentTop + $contentH - 52 + 35)    # 重新检测
$okEnv = Find-NameLike $p "*环境齐备*" 30
Note ($okEnv -ne $null) "6.1 系统状态检测列表" ""
Shot $p "06-status.png"

# ---------- 7. 退出 ----------
[void][W32]::PostMessage($p.MainWindowHandle, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)
Start-Sleep -Seconds 3
$alive = Get-Process -Id $p.Id -ErrorAction SilentlyContinue
Note ($null -eq $alive) "7.1 正常退出（无残留进程）" ""

if (Test-Path $cfgFile) { Remove-Item $cfgFile -Force }
Write-Host ("===== E2E RESULT: PASS=" + $script:pass + " FAIL=" + $script:fail + " =====")
if ($script:fail -gt 0) { exit 2 } else { exit 0 }
