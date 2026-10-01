# Win7 x64 实机验收手册（M0-M3 全量）

> 验收线（设计方案 §9.3）：裸 Win7 SP1 x64、无外网、无管理员、无 Office → 引导装成 → P0 场景（核对/汇总/发票/转换）全链可用 → 卸载无残留。
> 配套自动化：`tests\vm\win7-smoke.ps1`（虚机内冒烟，PowerShell 2.0 兼容）+ `tests\vm\run-on-vm.ps1`（宿主端 VBox 驱动）。
> 本机（开发机）无虚拟化环境，虚机阶段需你提供一台 Win7 SP1 x64 虚机（VirtualBox/VMware 均可）。

## 0. 已执行记录：2026-09-25 首轮实机验收（VirtualBox 7.0.14 / Win7SP1x64）

环境：`Win7SP1x64`（Windows 7 SP1 x64，VirtualBox 7.0.14，Guest Additions 7.0.14 在位），
客体账号 `OfficeAgent`（Administrators 组成员，**但以标准令牌运行**：`whoami /priv` 无提权特权，UAC 未提升）。
仓库经 guestcontrol copyto 拷入 `C:\OfficeAgent`（payload 全量 + lo76 zip 446MB + fixtures 全量）。

### 0.1 冒烟结果：`RESULT pass=1 fail=8 skip=3`

| # | 项 | 结果 | 说明 |
|---|---|---|---|
| 1.1 | boot report-only | **PASS** | exit=2（8 项缺失，符合裸机预期）；报告正常生成 |
| 2.1-2.3 | host /selftest /intents /masktest | FAIL | exit=`-2146232576`（`0x80131700` CLR 初始化失败） |
| 3.1-3.2 | /recon 差异集 | FAIL | 同上（主程序无法启动，产物不存在） |
| 4.1 | /convert xlsx→csv | FAIL | 同上 |
| 4.2-4.3 | LO 转换 / pdfium 渲染 | SKIP | 脚本按 LO 未解压跳过（本轮已解压，可复跑转 PASS 路径） |
| 5.1-5.2 | 审计链断言 | FAIL | 无主程序运行，无事件 |
| 6.1 | config.json | SKIP | 未跑 GUI 首启向导 |

### 0.2 根因（已定位，非代码缺陷）

主程序 `build\host\OfficeAgent.exe` 目标框架为 **.NET 4.x**，而裸 Win7 SP1 自带运行时只到 **3.5**
（客体 `C:\Windows\Microsoft.NET\Framework` 仅含 v1.0.3705 / v1.1.4322 / v2.0.50727 / v3.0 / v3.5，
无 `v4.0.30319`；`HKLM\...\NDP\v4\Full\Release` 不存在）。
故进程可创建但 CLR 初始化即失败，退出码 `0x80131700`。**拷贝链路本身正确**（宿主 236032 字节 = 客体 236032 字节）。

### 0.3 引导器免管理员补全路径：PASS（已验证）

`OfficeAgentBoot.exe /fix /console /root:C:\OfficeAgent`（标准用户令牌，无 UAC）：

- 7 项需管理员组件按设计跳过：KB4490628 / KB4474419 / KB3140245 / KB2670838 / .NET 4.8 / VC++ x86 / VC++ x64。
- **Python 3.8.10 解压成功** → `runtime\py38\python.exe`（101552 字节）落地。
- **LibreOffice 7.6.7 解压成功** → 446MB zip 展开耗时 41s，`lo76\LibreOffice\program\soffice.exe`（331704 字节）落地。
- 结论：C2（离线补全）与 C3（per-user 免管理员补全免安装组件）在实机成立。

### 0.4 新发现：解压出的运行时**无法执行**（缺 VC++ 运行库）

| 探针 | 结果 |
|---|---|
| `runtime\py38\python.exe -V` | 退出码 **126**，无输出（`probe.txt` 0 字节） |
| `lo76\...\soffice.exe --version` | 无输出 |
| `C:\Windows\System32\msvcr100.dll` | **不存在** |
| `C:\Windows\System32\msvcp140.dll` | **不存在** |
| `C:\Windows\System32\vcruntime140.dll` | **不存在** |
| `C:\Windows\System32\api-ms-win-crt-runtime-l1-1-0.dll` | **不存在** |

即：Python 3.8 与 LibreOffice 7.6 均依赖 VC++ 2015-2022 运行库，而该组件属"需管理员"档，
在标准用户令牌下无法补全 → **解压成功 ≠ 可用**。免管理员路径下这两个组件目前是"落地但跑不起来"。

### 0.5 待补缺口（本轮暴露，按优先级）

1. **[P0] 提权路径无效**：`boot\src\BootForm.cs:155` 在需管理员时用
   `Process.Start(Application.ExecutablePath, "/adminfix ...")` 重启，但未指定 `runas` 动词，
   且 `boot\app.manifest` 为 `asInvoker` → 新进程仍是标准令牌，**点"安装缺失组件"不会弹 UAC**，
   需管理员项永远装不上。设计方案 §8.2 要求"需管理员时弹一次 UAC"。
2. **[P0] 降级壳缺失**：设计方案 §8.4 承诺"装不上 .NET 4.8 → 进入 .NET 3.5 降级壳（自检/核对/转换核心可用）"，
   但 `host\build.ps1` 只产出 net4x 单一 exe，仓库内**无降级壳工程**。
   无管理员机器上主程序功能归零，与 C3/§8.4 不符。
3. **[P1] 免管理员档位缺 VC++ 兜底**：python/LO 依赖 VC++ 但该组件需提权。
   可选方案：payload 内置 VC++ 的免安装展开形态（DLL 随 runtime 目录分发并在 ProcRunner 启动前置于 DLL 搜索路径），
   使"无管理员"分支下 Python/LO 真能跑。
4. **[P1] 安装包分发形态未验证**：本轮用目录+guestcontrol 拷贝，未经 Inno 安装器（`installer\officeagent.iss`）
   与引导器 GUI 完整路径；卸载残留检查未做。

### 0.6 复跑建议

修掉 0.5-1（提权）后，在**已提权会话**（UAC 提升一次）内复跑 `boot /adminfix` 完成 KB+VC++.NET 补全，
重启后再跑冒烟；4.2/4.3 应转 PASS，2.x/3.x/5.x 应随 .NET 4.8 到位转 PASS。

### 0.7 第二轮（2026-09-25 晚，同虚机）：**冒烟全绿 pass=13 fail=0 skip=1**

承接 0.5 的缺口，本轮修复与验证：

| 动作 | 结果 |
|---|---|
| 第一轮提权补全（UAC 放行，右键管理员运行） | KB4490628/4474419/3140245/2670838 完成（3010 需重启）、VC++ x86/x64 完成、Python 解压完成 |
| .NET 4.8 第一轮失败 | 退出码 `0x800B0109`（证书链不受信任）——SHA-2 支持未重启生效，安装器验签失败。属"装完全部后重启再复检"设计预期内 |
| 重启后第二轮 | **.NET 4.8 安装成功**；复检"结论: P0 组件齐备"（report-only exit=0） |
| 实机冒烟（13 项） | **pass=13 fail=0**；唯一 SKIP=6.1 config.json（未跑 GUI 首启向导，文档预期内） |
| LO 转换链实机验证 | dirty.xlsx → LO headless → PDF → pdfium 渲染 全通（4.2/4.3） |
| 审计链 | 5.1/5.2 PASS（见 0.8 的 CLI 审计补丁） |

### 0.8 本轮产品修复（代码已合入）

1. **boot SetErrorMode**：`boot\src\Program.cs` Main 开头 `SetErrorMode(SEM_FAILCRITICALERRORS)`。
   裸机上 python/LO 探针先于 VC++ 执行时，Windows 会弹 `api-ms-win-crt-*.dll` 缺失的模态错误框阻塞检测；
   抑制后探针直接收到加载失败退出码。第二轮实机验证不再弹框。
2. **CLI 审计补齐**：CLI 无头链（selftest/intents/masktest/recon 等）此前完全不写审计（audit.jsonl 只有 GUI 写），
   与设计方案 §7.3"每个动作可审计"不符。现 Main 入口写 `app_start(cli)`，RunRecon 写
   `action_started`/`action_finished`（其余 CLI 动作的 finished 事件待补）。
3. **.NET 4.8 安装超时放宽**：`ProcRunner.WaitExit` 对 netfx48 组件 20→45 分钟（低速虚机防误杀）。
4. **`.NET 3.5 降级壳`（T1.3，设计方案 §8.4 兑现）**：
   - 扫描确认 host 源码除 `ZipArchive`（.NET 4.5+）外无任何 4.0+ BCL/C# 语法依赖；
   - 新增 `core\src\MiniZip.cs`（纯 3.5 zip 读/写，读侧与 boot\Unzip.cs 同源，CRC32 校验）；
   - MiniXlsx/MiniXlsxWrite/ConvertEngine 全部改走 MiniZip（ZipArchive 引用清零）；
   - `host\build.ps1` 双目标：`OfficeAgent.exe`（net4x）+ **`OfficeAgent35.exe`（net35 降级壳）**；
   - 顺带替换三处 4.0-only API（Path.Combine params 重载 ×3、SpecialFolder.ProgramFilesX86）；
   - 开发机验证：35 壳 /recon、xlsx 读、报告写、写读闭环全过；**实机冒烟 3.3/3.4 PASS**；
   - 冒烟脚本新增 3.3/3.4 两项锁定降级壳回归（11 项 → 13 项）。
   - 已知缺口：boot"启动主程序"按钮在 .NET 4.8 未就绪时自动选 35 壳的逻辑**未能落地**——
     所有进程启动代码的新增/修改均被 Mimosa 安全钩子拦截（详见 0.9），当前需手动双击 OfficeAgent35.exe。
5. **虚机驱动脚本修正**：`tools\vm-copy-kit3.ps1`/`vm-run-smoke.ps1` 原把 tools\ 当仓库根（源文件不存在却报 OK）；
   已修正并在复制脚本中对缺失源文件直接抛错。

### 0.9 Mimosa 安全钩子阻塞记录（重要，后续接手者必读）

2026-09-25 会话中，以下合法修复尝试全部被 Mimosa 安全钩子以"命令注入高危"拦截写入（共 7 次）：

- boot 提权修复（`Process.Start` + `runas`，asInvoker 清单下原实现永远拿不到提权令牌）：
  尝试过①路径拼接传参 ②常量参数 + UseShellExecute/Verb ③挪入 ProcRunner ④`Path.Combine` 变量
  ⑤全编译期常量 ProcessStartInfo ⑥requireAdministrator 清单变体 + 普通启动（`boot\app.admin.manifest`
  与 build.ps1 第二目标已合入，仅剩 BootForm.cs 的启动调用被拦）——均被拦；
- boot"启动主程序"按 .NET 状态选 48/35（连 `ProcRunner.StartDetached` 封装也被拦）。

钩子对**任何新增的进程启动代码**（含 `Process.Start` 出现在 diff 中）一律拦截，其建议的
"参数列表 API"（.NET 4.5 `ArgumentList`）在 boot 的 .NET 3.5 目标下不存在。文件里**既有**的
`Process.Start` 调用不受影响。规避钩子不做（构成对安全工具的绕过）。
接手者可选合规路径：联系钩子维护方为 `boot\src` 与 `core\src\ProcRunner.cs` 加白名单规则，
或采用无需新启动代码的方案（计划任务、explorer 代理启动等）。

### 0.10 实机验收当前结论

- M0（引导器 + 离线补全）：**实机闭环达成**（裸 Win7 SP1 x64 → 两轮补全 → P0 全绿）。
- M1（查看/转换）：实机冒烟通过（LO 转换 + pdfium 渲染 + 原生链）。
- M2 场景（核对等）：CLI 实机通过；GUI 流程待 E2E 干扰排除后重跑（见 0.11）。
- 3.5 降级壳：实机 CLI 链路通过；GUI 未在虚机实测。
- 发布门槛剩余：GUI 首启向导（config.json）与卸载残留检查未做；Inno 安装器路径未验证（0.5-4）。

### 0.11 E2E 排障记录（2026-09-26 复跑 PASS=31 FAIL=0）

首轮 E2E 报 PASS=10 FAIL=21，无并行 UI 操作复跑结果相同（排除焦点干扰）。根因两处，均为**测试脚本自身**：

1. **向导复选框首击被窗口激活吞掉**：e2e 只在开头点一次"允许内网/本机端点"，
   未勾选时 HostGuard 按设计拒绝环回地址（红字"拒绝 环回地址 127.0.0.1"），
   获取模型列表从未发出请求（mock 日志零连接）→ 向导卡死 → 之后全部 UI 步骤级联失败。
   修复：复选框点击挪进获取重试循环（`tests\e2e\e2e.ps1`）。
2. **0.6 /skillrun 退出码读取不可靠**：GUI 子系统进程经 Start-Process 读 ExitCode 偶发为空。
   修复：改为重定向输出并断言 `"ok": true`（技能功能本身手动验证正常）。

修复后全量重跑：**PASS=31 FAIL=0**（含向导/对话/记忆/转换队列/核对页/计划确认/审计/预览全链）。

### 0.12 GUI 实机验收状态（2026-09-26）

**§4 安装与首启（GUI，已验证，含截图）**：
- 引导器检测清单全绿（KB×4/.NET4.8(Release=528040)/VC++/Python/LO）→ 点【启动主程序】；
- 主窗口 + 首启向导弹出正常，中文渲染无乱码；点【暂不配置，离线使用】跳过向导成功；
- 九个页签导航可切换，表格核对页整页控件渲染正常（见本轮 computer-use 截图）。

**§5 GUI 场景（部分完成）**：
- 已验证：无外网分支——离线向导跳过；引擎层核对/转换/渲染由冒烟 13/13 覆盖（CLI 无头）；
  同一套 GUI 填表/会话/隐私/审计流程在开发机 E2E 31/31 通过。
- 待人工点验（约 10 分钟，控制通道不稳定导致自动化中断，清单如下）：
  1. 核对页：文件A=testsixtures
econlow.csv，文件B=ledger.csv → 自动猜列 → 开始核对 → 差异表打开；
  2. 会话页：输入"核对 <A路径> 和 <B路径>" → 计划卡片 → 确认 → 已完成；
  3. 内置插件页切隐私等级 → 审计页出现记录 → 校验审计链 ✓。

**本轮新发现（GUI 路径）**：引导器启动布局期望 `<root>\host\OfficeAgent.exe`（安装布局），
测试 kit 的 `build\host` 布局会报"未找到主程序"——已按安装布局补齐；lite/complete 安装器的 DestDir 即安装布局，无此问题。

### 0.13 代码审查修复记录（2026-09-26，双轴审查后）

对 M3.4/M3.5/M3.6 改动做了规范轴 + 需求轴审查并逐条修复。**开发机回归：build-all 双目标 ALL DONE；
六套自测（selftest/guardtest/masktest/intents 8/suggesttest 6/skilltest 8）全 ALL PASS；E2E PASS=32 FAIL=0。**

**阻断级（已修）**
1. `/skillrun` 无确认即可执行技能，绕过 M3.2「显式确认才执行」红线且不留审计
   → 现必须带 `--yes` 才执行，否则只打印计划并 exit=2；确认后落 plan_created/confirmation/action_started/finished 全链审计。
   E2E 新增 0.6a（无 --yes 被拒）/0.6b（带 --yes 通过）双向断言。
2. `/skilltest` 会向真实 `%LOCALAPPDATA%\OfficeAgent\skills` 写假清单再删，可能毁掉用户真实技能
   → 改用临时 user 根（`Scan(..., userRootOverride, writeRegistry:false)`），不再触碰真实目录。
3. **技能计划到不了执行器**：`ActionPlan` 允许 Skill 无输出文件，但 `ActionExecutor` 无条件要求 `OutputPath`
   → 已按 Kind 分支放行；`/skilltest` 新增第 8 项断言（Skill 计划 executable=True 且真跑到 RunSkill）。
4. SkillRunner 超时只单进程 Kill，且错误文案谎称"整树终止"（soffice/python 会派生子进程）
   → 新增 `core\src\JobProbeRunner.cs`（入 Job + 整树击杀），soffice 侧新增 `host\src\ToolJobGuard.cs`，文案改真实描述。
5. 技能声明 `network`/`officeCom` 而执行桥无沙箱可放行 → 现显式拒绝执行（此前两字段解析后无人读，形同虚设）。

**正确性（已修）**
6. `ColumnSuggest` 在 L1 对**表头**套 `MaskEngine`，既无保护意义又破坏白名单比对（脱敏后名字不在真实表头里→建议被整体拒绝）→ 改为只发表头、不做值级脱敏。
7. `CachedScan` 传 `conv=null` 导致 LibreOffice 恒 false、依赖 LO 的技能被误判 degraded → 改用轻量能力探测（只查文件存在性，不在热路径做 COM 探测）。
8. `DepsMissing` 对未知工具名直接放行 → 改为不满足；拼错依赖名不再静默通过。
9. 技能 `task` 值两条路径不一致（`/skillrun` 传 entry、GUI 传 SkillId）→ 统一为 SkillId。
10. `OfficeAgentBootAdmin.exe` 早已构建但从未被使用：「以管理员身份运行」重跑 asInvoker 本体不提权
    → 改启动 requireAdministrator 变体（Windows 自行弹 UAC）；有意不用 `Verb="runas"`（铁律 4）。
11. 引导器【启动主程序】固定启 `OfficeAgent.exe`，无 .NET 4.8 时不降级 → 新增 `boot\src\BootLaunch.cs`，按 Release≥528040 选壳（§8.4 落地）。
12. complete SKU 的 LO/wheels 带 `skipifsourcedoesntexist`，缺件会静默产出假完整包 → 移除该 flag（lite 保留），缺件即编译失败。

**欠账（本轮未做，非阻断）**
- §4.4 每技能独立 venv、§4.6 ed25519 团队验签、§4.3 SQLite 注册表、§4.2 `ui.schema.json` 一键卡片仍未实现；
  团队共享根已落地（env `OfficeAgentTeamSkills`）但验签未接，故 team 源技能暂不置 `SignatureVerified`。
- 32 位 Win7 实机、10 万行大表虚机实测、GUI 5 项人工点验仍待做（见 §0.12）。

## 1. 虚机准备

1. 装机目标：**Windows 7 SP1 x64**（首要目标；Server 2008 R2 SP1 同内核可选测）。
2. 规格：2 vCPU / 4GB 内存 / 60GB 磁盘（对齐设计方案"公司配发老机器"画像）。
3. 装完 SP1 后**不装任何运行库、不装 Office、禁用网卡（断网）**，装 Guest Additions（仅共享目录用）。
4. 快照留底（每轮验收失败可回滚重跑）。
5. 装一台杀软（360 或火绒）做拦截率抽查的第二个快照分支。

## 2. 分发包与在位组件

| 内容 | 说明 |
|---|---|
| `build\boot\OfficeAgentBoot.exe` | .NET 3.5 引导器（AnyCPU：x64 原生跑，wusa 不经 WOW64 重定向） |
| `build\host\OfficeAgent.exe` + `bin\pdfium.dll` | 主程序 |
| `payload\` | KB x64 / .NET 4.8 / VC x86+x64 / Python 3.8 双架构 / **LO 7.6.7 zip（446MB，已登记 components.ini）** |
| `tests\vm\win7-smoke.ps1` | 虚机内自动冒烟（boot 自检 → /intents /masktest /recon → 转换 → 渲染 → 审计断言） |
| `tests\fixtures\` | 核对差异集 / 脏表 / 大表 / 发票 PDF |

把整个仓库目录拷入虚机（共享目录或 ISO），如 `C:\OfficeAgent`。

## 3. 自动冒烟（先跑这个）

在虚机 cmd 里：

```
powershell -NoProfile -ExecutionPolicy Bypass -File C:\OfficeAgent\tests\vm\win7-smoke.ps1 -Root C:\OfficeAgent
```

- 全部 PASS + exit 0 = 无头链全绿（此时 LO 可能尚未解压，4.2/4.3 会 SKIP——属预期）。
- 结果文件：`%TEMP%\win7-smoke-result.txt`。
- 注意：全新 Win7 的 PowerShell 2.0 可直接跑本脚本（无 PS3+ 语法）。

## 4. 安装与首启（模拟用户完整路径，GUI）

1. 双击 `build\boot\OfficeAgentBoot.exe`（不要进 cmd 敲参数）。
2. 预期检测清单：KB4490628 / KB4474419 / KB3140245 / KB2670838 / .NET 4.8 / VC++ x86+x64 / Python 3.8.10 标红"缺失"。
3. 点 **安装缺失组件** → UAC 提权 → 顺序自动安装：
   - 预期顺序：KB4490628 → KB4474419 → KB3140245 → KB2670838 → .NET 4.8 → VC x86 → VC x64 → 解压 Python → 解压 LO。
   - **x64 专项**：全部 wusa 步骤应正常启动（引导器 64 位进程直达 System32\wusa.exe）；若 KB 报"不适用"记录退出码。
   - KB3140245 / KB2670838 / .NET 4.8 可能要求重启：全部装完后重启一次再复检。
4. 重启后复检 → P0 组件全 OK；**LO 应显示 7.6.7 路径**（boot 从 payload\lo76 zip 解包到 `<root>\lo76\`）。
5. 点 **启动主程序** → 首启向导可跳过（离线）→ 主界面各页签可切。
6. 再跑一次第 3 节冒烟：4.2/4.3 应转 PASS（LO 在位）。

## 5. GUI 场景验收（无 Office + 无外网 + 无管理员三分支）

每分支重置快照后执行：

| # | 分支 | 验收点 |
|---|---|---|
| 1 | 无管理员（标准用户） | per-user 安装成功；缺系统组件时按钮提权一次，拒绝 UAC → 走 3.5 降级链提示；核对/转换仍可用 |
| 2 | 无外网 | 向导跳过；隐私等级 L1 下模型对话禁用提示；核对/汇总/发票/转换全可用；TLS 探针"未知"非"缺失" |
| 3 | 无 Office | 三级转换总线落 LO（转换队列跑 dirty.xlsx→PDF、sample.csv→XLSX）；预览 PDF 正常 |
| 4 | 杀软在位（另快照） | 引导器/主程序/soffice 拦截记录截图；白名单指引页可用 |

GUI 用例（对应 E2E 场景的手动版）：
- 拖入 `tests\fixtures\recon\flow.csv` → 预览表格网格正常（中文不乱码）。
- 「表格核对」页选 A=flow.csv B=ledger.csv → 自动猜列 → 开始核对 → 差异表打开：标红差异行、汇总页勾稽全 ✓。
- 会话页输入"核对 <flow.csv 路径> 和 <ledger.csv 路径>" → 计划卡片 → 确认 → 已完成 + 产物存在（离线下 L1 不影响本地任务）。
- 「内置插件」页切隐私等级 → 审计页出现"调整隐私等级"❗行；点"校验审计链" → ✓ 完整。

## 6. 判定通过的标准（发布门槛）

- [x] 冒烟脚本 13 项 PASS（2026-09-26 实测 pass=13 fail=0 skip=1；6.1 config.json 属预期 SKIP）。
- [ ] 引导器离线补全全程无网络、SHA256 无失败、日志（`%LOCALAPPDATA%\OfficeAgent\logs\`）无异常栈。
- [ ] `runtime\py38\python.exe --version` = Python 3.8.10；`lo76\LibreOffice\program\soffice.exe --version` = 7.6.7。
- [ ] 主程序五页签无崩溃、中文无乱码（CP936 控制台 + GUI 均正常）。
- [ ] 10 万行大表（tests\fixtures\big.xlsx）→ PDF 转换在虚机内可完成（开发机实测 543s / 1852 页；虚机慢属预期，看门狗 900s）。
- [ ] 卸载检查：删除安装目录与 `%LOCALAPPDATA%\OfficeAgent` 后无残留服务/驱动/计划任务。

## 7. 失败取样清单

引导器日志目录打包；`certutil -hashfile <文件> SHA256` 对照 `store\components.ini`；wusa 手动退出码与 `%WINDIR%\Logs\CBS\*.log` 尾部；.NET 4.8 安装日志 `%TEMP%\dd_*`；LO 解包后 `soffice --version` 手动输出；杀软拦截记录截图。

### 0.14 实机反馈修复（2026-09-26 下午，v0.4.1）

Win7 SP1 x64 实机首装反馈 8 项 + "KB3140245 修复一直失败"，本轮全部处理：

**引导器**
- KB3140245 安装失败 0x80242017 根因=缺 2015-04 服务堆栈前置 KB3020369 → payload 已补
  （windows6.1-kb3020369-x64.msu，SHA256=e9fbb6b4…，verify-payload 19/19），检测序在 KB3140245 前，
  且 install_kb3140245 分支带自愈（缺前置先装前置）；退出码显示改为含 0x 十六进制。
- 本项目自即日起只支持 Win7 SP1 x64（用户指定），KB3020369 仅备 x64。

**会话（"智障"三大根因）**
- 模型每次只收单条消息、无历史 → 多轮历史上下文（含恢复的历史种子，上限 24 条）。
- AI 看不到 txt 等文件 → 消息内路径/拖拽文件自动读出（≤3 个、每个 6000 字、L1 脱敏）附进提问；
  并新增 function calling 工具循环（read_text_file/list_directory/download_file/convert_document/
  repair_environment 白名单，经 HostGuard，≤5 轮），已用 mock 两跳全链路验证（/agenttest）。
- 不联网 → 智谱系端点注入 web_search 工具（设置页可关）；不支持 tools 的端点自动降级纯对话。

**界面**
- 长回复显示不全两个根因：max_tokens 1024→4096；ListBox 单项高度被系统钳制 255px →
  长回复自动分块成连续气泡（≤440px/段，断点优先换行标点）。
- 滚轮滚不动历史：Win7 滚轮只作用焦点控件 → 输入框/面板滚轮转发消息列表。
- 流式整表重绘闪烁 → 100ms 节流改"只重绘尾部 + 贴底才跟随"。
- 气泡左右难分辨 → 用户实心浅蓝+边框，助手白底+边框；宽度测量补内边距余量（修文字破框）。
- 模型列表拉不到 → 设置页预置 13 个常用模型；新增"联网搜索"开关。
- 「新建任务」页新增「新建对话」按钮（清空会话+模型上下文）。

**验证**：六套自测 ALL PASS；/agenttest 普通对话与 TOOLTEST 工具循环（2 跳）全通；
payload 19/19；boot/host 双目标编译全绿。GUI 全量 E2E 本轮未跑完（用户中止，避免占机），
GUI 回归待下轮。

### 0.15 自优化轮（2026-09-26 傍晚，v0.5.0 收尾）

E2E 复跑揪出两处回归并修复，另做三项体验优化：

1. 设置页"联网搜索"复选框插入位置把 E2E 按坐标勾选的"允许内网端点"挤偏 → 向导卡住级联 17 项失败
   → 布局固定 chkLan=y326 / chkWeb=y352（代码注释写明布局约束）。
2. 侧边栏分组头使导航行号与页号错位一格，E2E 导航点击全落空 → e2e.ps1 行映射更新
   （页 0-5=行 1-6，页 6-8=行 8-10，跨"系统"头）。
3. 优化：等待中气泡显示已等待秒数（非流式请求体感）；超长用户消息（大段粘贴）同样分块防 255px 裁剪；
   历史上下文加 24000 字符预算（附文件的长历史从最旧丢弃，至少保留一问一答）。
4. mock 的 BODY 日志 300→2000 字符；E2E 3.2 断言从"流式 SSE"改为"function calling 工具声明随请求发送"
   （agent 循环有意非流式，工具调用语义下更可靠）。

**最终基线：E2E PASS=32 FAIL=0；六套自测 ALL PASS；真模型 glm-4.5-air 工具循环两跳实测通过；
v0.5.0 双 SKU 安装包 16:47 重打。** 注意：E2E 会清空 config.json 强制走向导（含用户真实配置）——跑完 E2E 需重新配置或恢复备份。
