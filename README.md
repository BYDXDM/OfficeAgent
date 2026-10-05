# OfficeAgent —— Win7 优先的办公 AI Agent

面向会计/打工人高频办公任务的 AI Agent（不是聊天框：会自己调工具把活干完，并把产物文件直接交给你）。
目标环境 **Windows 7 SP1 x64**（含 Server 2008 R2 SP1；无外网可离线补全）。

> 当前版本 **0.8.2**。设计方案见 `..\OfficeAgent-设计方案.md`；上手说明见 [新手手册.md](新手手册.md)；给下一个开发 agent 的接手计划见 [下一步优化提示词.md](下一步优化提示词.md)。
>
> 近期变化：
> - **0.8.2**：安全兜底——C 盘文件写/删、删除表格强制弹窗确认（可"始终允许此目录"，盘根除外）；新增受守卫的 `delete_file` 工具；工作区支持选磁盘；安装向导中文化。
> - **0.8.1**：Excel 弹窗两根因修复（孤儿 styles 部件=「部分内容有问题」、独占读=「文件正在使用」）；模板打开即落在可用表；流水首页汇总速览；预览页「打开源文件所在地」。
> - **0.8.0**：**带公式工作簿生成**（工资表/增值税台账/流水账三模板 + 自由定制，填数即自动算、改明细汇总联动）+ 内置 Excel 公式大全查询 + 表达不清必须反问的铁律。
> - **0.6.3**：agent 多轮效率（先定位→一次做对、重复调用拦截、轮次 10→12）；最终答复流式；工具失败给「为什么+怎么办」；产物同名自动改名不覆盖。


## 它能做什么

**对话即任务**：说“核对这两个表”“把这些表汇总成底稿”“建一个 3 页 AI 发展 PPT”，agent 会自己列目录、读文件、调工具，完成后把生成的 xlsx/pptx 以**文件卡片**交给你（点击即可预览）。

| 能力 | 说明 |
| ---------------------------------- | ------------------------------------------ |
| 会话式 agent | 工具循环（读文本/PDF、列目录、下载、转换、建表、**建带公式工作簿**、**删文件（受确认守卫）**、建 PPT、环境修复）；多轮上下文；可选联网搜索 |
| 带公式表格 | 三模板（工资表套账 / 增值税台账 / 流水账）+ 按口述自由定制；预置空白公式行，填数即算；改明细汇总自动联动；内置 95 条会计向公式离线查询 |
| 两表核对 | 归一化键匹配 + FIFO 配对 + 金额容差；差异表带双侧单元格引用；Verifier 勾稽自证 |
| 多簿汇总 | 按表头名归集（列序无关）→ 合并底稿（SUM 公式）+ 逐文件勾稽校验页 |
| 发票提取 | pdfium 文本层直提（号码/日期/税号/金额/税额/价税合计）；单票勾稽；扫描件进未识别清单 |
| 文档转换 | xlsx/csv/pdf/doc/ppt 互转；frp 打印模板解析→xlsx（查看/输出）：原生直写 → LibreOffice → Office COM（Excel→PDF 有挂死看门狗） |
| 隐私分级 | L0 全本地 / **L1 脱敏出网（默认）** / L2 全量；手机号/身份证/银行卡/税号/邮箱规则脱敏 |
| 审计哈希链 | SHA256 链式 JSONL，删改可检出；文件读写/LLM 调用/动作全记录 |
| 工作区（项目） | 侧边栏按项目分组，每个项目挂自己的文件夹；产物/@ 引用/下载默认落到当前项目；工作区可选磁盘 |
| 安全兜底 | C 盘文件写入/删除、删除表格 → 强制弹窗确认（可"始终允许此目录"，盘根不提供）；临时目录与应用自管路径豁免 |
| 会话管理 | 归档 / 恢复 / 删除；输入 `@` 引用工作区文件 |
| 离线补全 | 引导器自检 KB/.NET/VC++/LibreOffice/wheels 并离线补全，**组件齐全时直接启动不弹窗** |

## 目录

```
core/      共享层：Detect.cs（环境检测）+ ProcRunner.cs（进程启动唯一出口）+ JobObject/MiniZip
boot/      .NET 3.5 引导器：检测 → 离线补全 → 复检 → 启动主程序（/report-only /fix /adminfix /ui /console）
host/      .NET 4.8 主程序（同源码可编 3.5 降级壳 OfficeAgent35.exe）
             文档引擎：MiniXlsx / MiniCsv / MiniXlsxWrite / ConvertEngine / Pdfium / PdfTextReader / PptWriter
             agent：AgentLoop（工具循环+收尾跳）/ AgentTools（工具白名单）/ LlmClient
             场景：ReconEngine / MergeEngine / InvoiceEngine
             骨架：MaskEngine / AuditLog / HostGuard / SessionStore / WorkspaceStore / CaretHelper
             UI：MainForm(+partial Recon)、ChatPanel、SetupDialog
installer/ officeagent.iss（Inno 双 SKU：/DSKU=complete 全离线 或 lite 在线补全）
bin/       pdfium.dll（Chromium 109.0.5406，Win7 可用）
store/     components.ini（组件↔官方源↔SHA256 权威清单）+ wheels 缓存
payload/   离线组件仓（KB/.NET4.8/VCRT/Python/LO，大文件不入库，由 tools/download-payload.ps1 重建）
tools/     构建与验证脚本（payload 下载/校验、fixture 生成、VM 辅助）
tests/     e2e\e2e.ps1（35 用例）+ fixtures + WIN7-VM-验收.md
```

## 构建与自测

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File build-all.ps1        # boot(net35) + host(net4x, 双目标)

# 无头自测（全部返回 0 为通过）
build\host\OfficeAgent.exe /guardtest          # 出网守卫
build\host\OfficeAgent.exe /masktest           # 脱敏引擎
build\host\OfficeAgent.exe /intents            # 意图路由
build\host\OfficeAgent.exe /suggesttest        # 列映射建议
build\host\OfficeAgent.exe /skilltest          # 技能体系
build\host\OfficeAgent.exe /carettest          # 输入框光标定位回归
build\host\OfficeAgent.exe /gridtest <xlsx>    # 表格读取（表头/行序）回归
build\host\OfficeAgent.exe /bridgetest         # 技能→工具投影（规划层一期）
build\host\OfficeAgent.exe /plantest          # 编排：预算/产物/换路/三期
build\host\OfficeAgent.exe /perftest          # 低配优化：预览有界性 + 流式解压一致性
build\host\OfficeAgent.exe /toolidtest        # tool_call_id 解析/归一化（并行工具调用不串号）
build\host\OfficeAgent.exe /formulatest       # 带公式工作簿：公式落位/页序/页内速览/括号配平
build\host\OfficeAgent.exe /urltest           # 服务地址诊断：裸域名补 /v1、多账号密钥槽（#tag）
build\host\OfficeAgent.exe /safetytest        # 安全兜底：C盘/删表确认、delete_file、豁免目录
build\host\OfficeAgent.exe /audittest         # 审计哈希链校验
build\host\OfficeAgent.exe /agenttest "读一下 D:\a.txt"     # 真模型 agent 循环

# 场景 CLI
build\host\OfficeAgent.exe /recon <A> <B> --keya 账号 --keyb 对方账号 --debita 收入 --credita 支出 --debitb 借方 --creditb 贷方 --tol 0.01
build\host\OfficeAgent.exe /merge <files...> --targets "科目,金额" --srcs "科目,金额" --amounts "金额"
build\host\OfficeAgent.exe /invoices <pdf...> --out 发票清单.xlsx
build\host\OfficeAgent.exe /convert <file> <pdf|csv|xlsx>
build\host\OfficeAgent.exe /pdftext <in.pdf>

# E2E（CLI 冒烟 + UI 全流程；会覆盖本机 config，注意备份）
powershell -NoProfile -STA -ExecutionPolicy Bypass -File tests\e2e\e2e.ps1

# 带公式工作簿实算复验：/formulatest 生成 → LibreOffice 无头重算 → python tools\verify-recalc.py
# （LO 转换命令与期望值见 交接-0.8.1.md §3.3；LO 在仓库外 D:\ai\工作\OfficeAgent-deps-lo76\）

# 出安装包
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" /DSKU=lite installer\officeagent.iss
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" /DSKU=complete installer\officeagent.iss
```

## 工程红线（评审必查）

1. 子进程只经 `ProcRunner` / 白名单形态启动：编译期常量路径 + 字面量参数，禁 shell 拼接。
2. 源码与构建脚本同守 **C# 3.0** 语法（无 LINQ/HashSet）；`build*.ps1` 纯 ASCII，含中文的 `.ps1` 必须带 UTF-8 BOM。
3. 安装器先 SHA256 校验 → 复制到固定暂存路径 → 再执行；清单/网络数据不参与命令行构造。
4. API Key 仅经 **DPAPI** 落盘，支持环境变量覆盖；源码/测试不含任何凭据字面量。
5. 输出一律 UTF-8；子进程 stdio 显式 UTF-8。

## Win7 兼容覆盖

| 目标环境 | 状态 | 说明 |
| -------------------------- | ------- | ---------------------------------------- |
| Win7 SP1 x64 | ✅ 主目标 | 引导器 AnyCPU（x64 原生跑，wusa 不被 WOW64 重定向） |
| Server 2008 R2 SP1 | ✅ 同内核 | 检测项显式识别（SM_SERVERR2） |
| Win7 SP1 x86 | ⚠️ 降级可用 | KB x86 变体已就位，引导器按位宽自动选 |
| Win7 RTM（无 SP1） | ⚠️ 阻塞 | .NET 4.8 硬性要求 SP1，引导页明示不假装可装 |
| 无管理员 | ✅ | per-user 安装；系统组件需提权一次，拒绝则走 3.5 降级壳 |
| 无外网 | ✅ | payload 离线补全，SHA256 锁定 |


## x64 专项修复（细节留档）

- **wusa 启动**：x86 引导进程在 x64 上 `System32\wusa.exe` 会被 WOW64 重定向到不存在的 `SysWOW64\wusa.exe` → 引导器改 AnyCPU（x64 原生跑），32 位系统仍 32 位运行。
- **VC++ x64 误检**：`GetSystemWow64Directory` 返回的是 SysWOW64（非原生 System32），旧逻辑把 x86 运行库当成 x64 → 新增 `EnvDetect.DirOfBitness(bool)`，32/64 位进程下都取语义正确的目录（32 位进程经 Sysnative 探测 x64 库）。
- **LibreOffice 只查 x86 目录**：32 位进程的 `ProgramFiles` 指向 `(x86)` → 探测候选补 `ProgramW6432` 与 `ProgramFiles(x86)`，x64 版 LO 安装可被发现（Detect + ConvertEngine.FindSoffice 双处）。
- **.NET 4.8 注册表缺失回退**：精简镜像 Release 值缺失时，用 `Framework64\v4.0.30319\clr.dll` 文件版本 ≥ 4.0.30319.42000 回退判定（该目录不经 WOW64 重定向）。
