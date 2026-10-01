# OfficeAgent —— Win7 优先的办公 AI Agent

面向会计/打工人高频办公任务的 AI Agent。设计方案见 `..\OfficeAgent-设计方案.md`（当前阶段对应 §9.3 里程碑表 M1/M1.5）。


## 当前状态：M3 第二阶段完成（LLM 列映射建议 + 技能体系骨架 + Win7 x64 兼容专项）

| 能力                                                                                                | 状态                                 |
| ------------------------------------------------------------------------------------------------- | ---------------------------------- |
| M0：.NET 3.5 引导器 + 18 项检测 + 离线 payload 补全（SHA256 锁定）                                               | ✅ 已验证                              |
| M1：MiniXlsx/MiniCsv 读写、三级转换总线、pdfium(Chromium109) 渲染、批量队列                                         | ✅ 已验证                              |
| M1.5 UI：侧边栏导航 + 气泡会话页 + 首启向导（自动拉模型/测连接）                                                           | ✅ 截图验证                             |
| M1.6 内置插件（记忆/转换偏好/会话历史，默认全开）                                                                      | ✅ E2E 实测                           |
| **M2.1 两表核对（设计方案 §6.1 旗舰场景）**                                                                     | ✅ fixture 验收                       |
| · 归一化键匹配（去空白/全半角/大小写）+ 同键 FIFO 配对 + 金额容差（decimal）                                                 | ✅                                  |
| · 差异分类：金额不等 / 仅A有 / 仅B有；差异表 Excel（红/橙标记 + 双侧单元格引用 + 表头冻结）                                         | ✅                                  |
| · Verifier 勾稽自证：两侧合计差 = 差异合计 + 容差内匹配残差；行数守恒；差异行抽样重算                                               | ✅                                  |
| · 模板记忆：映射配置存 recon-templates.json，「上月对账再来一次」直接载入；UI 自动猜列                                          | ✅                                  |
| · 金额解析失败行剔除并列错误清单（绝不静默当 0）；.xls 经转换总线预转                                                           | ✅                                  |
| **M2.2 报表汇总（§6.2）**                                                                               | ✅ CLI/单测验证                         |
| · 多簿按表头名归集（列序可不同）→ 合并底稿（SUM 公式）+ 校验页（逐文件小计勾稽）                                                     | ✅                                  |
| · 模板存取 merge-templates.json                                                                       | ✅                                  |
| **M2.3 发票提取（§6.4）**                                                                               | ✅ fixture 验收                       |
| · pdfium 文本层直提（FPDFText 字符基线聚类成行）：号码/日期/购销方/税号/金额/税额/价税合计                                         | ✅                                  |
| · 行内伪空格归一化 + 标签空格容忍正则；单票勾稽 金额+税额=价税合计；扫描件进未识别清单（OCR 后续）                                           | ✅                                  |
| **M2.4 隐私分级与脱敏（§7.2，C5 硬约束）**                                                                     | ✅ /masktest 全过                     |
| · L0 全本地（禁对话）/ L1 脱敏出网（默认，手机号/身份证/银行卡/税号/邮箱）/ L2 全量；设置在「内置插件」页                                    | ✅                                  |
| **M2.5 审计哈希链（§7.3）**                                                                              | ✅                                  |
| · SHA256 链式 JSONL（audit.jsonl，5MB 轮转）：文件读写/LLM 调用(含脱敏命中)/核对/汇总/提取/隐私变更全记录                         | ✅                                  |
| · 审计页实装（最近 500 条 + 一键校验链 + 打开目录）；校验发现删改即报告断点                                                      | ✅                                  |
| **M2.6 运行隔离 Job Object（§4.5）**                                                                    | ✅                                  |
| · 工具进程（soffice 等）入 Job：1.5GB 内存上限 + KILL_ON_JOB_CLOSE + 超时整树终止；安装器不进 Job                          | ✅                                  |
| MiniXlsx 加固：合并单元格解析与回填、公式无缓存值标记"="、共享字符串 O(n)、zip 解压体积防护、原始行号                                     | ✅ 脏表 fixture                       |
| 性能：10 万行 xlsx → CSV 流式转换 565ms（开发机）；大表 fixture big.xlsx 入回归                                       | ✅                                  |
| E2E 真人操作模拟（CLI 冒烟 0.1-0.5 + 向导→记忆→对话→转换→核对→计划确认→预览→检测→退出）                                         | ✅ 全绿                               |
| **M3.1 规则优先意图路由（IntentRouter，纯规则不调模型）**                                                           | ✅ /intents 8 用例                    |
| · 「核对/对账/勾稽」「汇总/归集」「发票/提取」→ ActionPlan（Kind/输入/模板/输出/缺口/状态机）                                      | ✅                                  |
| · 消息内文件路径正则抽取（中文路径可用，句读尾巴剥离）+ 拖入文件上下文 + 核对页选文件联动                                                  | ✅                                  |
| · 普通问句/确认词不误触；缺文件/缺映射 → 计划标"缺参数"，绝不执行                                                             | ✅                                  |
| **M3.2 确认执行闭环（ChatPanel + ActionExecutor）**                                                       | ✅ E2E 4c.1-4c.8                    |
| · 计划卡片（输入/模板/容差/输出/覆盖提示）→「确认/执行/好/开始」才执行；「取消/算了」零产物                                               | ✅                                  |
| · 确认后后台线程调 ReconEngine/MergeEngine/InvoiceEngine（LLM 不参与执行链），结果回填气泡                               | ✅                                  |
| · 审计全链：plan_created / confirmation / action_started / action_finished                             | ✅ E2E 断言                           |
| **M3.3 界面通俗化（PlainTips）**                                                                         | ✅                                  |
| · 核对页检查结果改为大白话列表（"总数对得上/一行都没漏/抽查复核…"），每行带 ❗ 悬停显示通俗解释                                              | ✅                                  |
| · 审计页事件名汉化（"使用 AI 模型/读取文件/写出结果…"）+ ❗ 悬停注释；页头说明改为用户视角                                              | ✅                                  |
| · 只改界面显示；Excel 报告与审计日志内容保持专业口径不变                                                                  | ✅                                  |
| **M3.4 LLM 列映射建议通道（ColumnSuggest）**                                                               | ✅ /suggesttest 6/6                 |
| · 缺列映射时只发【表头名单】给模型（L1 先脱敏、记审计）；建议列名超白名单 → 整体拒绝                                                    | ✅                                  |
| · 建议仅"待确认"回填计划，用户仍须显式「确认」才执行；L0/未配置/失败一律回退手工页                                                     | ✅                                  |
| **M3.5 技能体系骨架（SkillSystem，设计方案 §4）**                                                              | ✅ /skilltest 8/8                   |
| · skill.json 清单校验（id/version semver/host 兼容/runtime 白名单）+ 坏清单拒绝                                   | ✅                                  |
| · 三级扫描根（内置 → 用户 → 团队共享 `OfficeAgentTeamSkills`）+ 同 id 高版本覆盖 + 注册表落盘 + 禁用跨扫描 + skill.json 变更监视失效缓存 | ✅                                  |
| · 能力表（LO/COM/OCR）+ 依赖匹配（"?"可选依赖降级不阻塞；未知工具名不再静默放行）；内置示例技能 bank-recon                                | ✅                                  |
| **M3.6 python 技能执行桥（SkillRunner）**                                                                | ✅ 端到端                              |
| · 经 JobProbeRunner 启动（ProcRunner 同形态 + Job 隔离）：exe 来自安装目录、参数字面量 -m oa_skill_main；清单/数据只进文件              | ✅                                  |
| · 技能声明 network/officeCom 时拒绝执行（执行桥无沙箱可放行，宁显式失败）；进程入 Job 1.5GB + 超时整树终止                        | ✅                                  |
| · embeddable .\_pth 隔离忽略 PYTHONPATH → 暂存目录幂等追加进 python38.\_pth（PatchPth 同族）                       | ✅                                  |
| · 协议：stdout 末行 JSON；示例技能 row-stat（stdlib CSV 统计+GB18030 探测）端到端入回归                                 | ✅                                  |
| · **python 技能接入自然语言闭环**：场景词→Skill 计划→确认→ActionExecutor→SkillRunner→对话返回                           | ✅ /intents 8/8 + 回归项               |
| · **/skillrun 内部入口需 --yes 显式确认**（不带只打印计划、零执行）；确认后全链落审计 plan_created/confirmation/action_*        | ✅ 实测                              |
| **M2.7 Win7 x64 兼容专项（详见文末兼容矩阵）**                                                                  | ✅ 双分支自检                            |
| **Win7 实机验收（VirtualBox Win7SP1x64）**：裸机两轮离线补全 → P0 组件齐备 → 冒烟 13 项 pass=13 fail=0（详见 tests\WIN7-VM-验收.md §0） | ⚠️ 冒烟全绿；GUI 5 项待人工点验 |
| **LibreOffice 7.6.7 落位**：USTC 镜像官方便携 paf → 免安装展开 → payload zip 446MB（components.ini 登记）           | ✅ 无 Office 转换链实测                   |
| · 无 Office 转换实测：10 万行 xlsx → LO headless → 1852 页 PDF → pdfium 渲染全通                               | ✅                                  |
| · 修复：LO 超时被杀后的半成品文件不再误判为成功；PDF 目标超时上限 900s（大表导出）                                                  | ✅                                  |
| **离线 payload 补齐（2026-09-25）**                                                                     | ✅                                  |
| · store\wheels：py3.8 双架构技能依赖 21 wheel + pip 24.3.1 自身（NJU 镜像 pip download）                        | ✅                                  |
| · KB x86 变体四个（Update Catalog 脚本化抓取，tools\download-kb-x86.ps1）→ boot 按位宽选变体 → **32 位 Win7 离线补全齐套** | ✅                                  |
| · 意图路由接入技能场景词（SkillSystem 注册表缓存 → builtin 技能 entry 映射引擎任务）                                        | ✅                                  |
| **M2.8 .NET 3.5 降级壳（设计方案 §8.4 兑现）**                                                                  | ✅ 实机冒烟通过（3.3/3.4）                |
| · MiniZip（纯 3.5 zip 读/写）替换 ZipArchive，host 全源码可 csc 3.5 编译；build 双目标 OfficeAgent.exe + OfficeAgent35.exe | ✅                                  |
| · boot SetErrorMode：裸机 python/LO 探针不再被加载错误框阻塞；CLI 无头链补写审计（§7.3）                        | ✅                                  |
| · 引导器启动按 .NET 环境选壳：有 4.8 启 OfficeAgent.exe，否则启 OfficeAgent35.exe（此前固定启 4.x，与 §8.4 承诺不符）        | ✅ 源码修复                              |
| · 「以管理员身份运行」改启动 requireAdministrator 变体 OfficeAgentBootAdmin.exe（Windows 自行弹 UAC）；此前重跑 asInvoker 本体不提权 | ✅ 源码修复                              |
| LibreOffice 7.6.7.2 便携包                                                                           | ✅ 已就位（446MB，USTC 官方便携展开，components.ini 登记） |
| COM→PDF 无打印机会话阻塞                                                                                  | STA 看门狗兜底，实机待验收                    |
| Win7 虚机验收矩阵（§9.2）—— GUI 首启向导与卸载残留检查                                                  | ⏸ 剩余项（补全/转换/核对链已实机全绿）    |

## 目录

```
core/      共享层：Detect.cs（环境检测）+ ProcRunner.cs（进程启动唯一出口）+ JobObject.cs（工具进程隔离）
boot/      .NET 3.5 引导器：检测 → 离线补全 → 复检 → 启动主程序（/report-only /fix /adminfix /console /root:）
host/      .NET 4.8 主程序：
             文档引擎：MiniXlsx(读)/MiniCsv/MiniXlsxWrite(报表写出)/ConvertEngine/ToolRunner/Pdfium(渲染+文本层)
             M2 场景：ReconEngine+ReconTemplate(两表核对) / MergeEngine+MergeTemplate(汇总) / InvoiceEngine(发票)
             骨架件：MaskEngine(脱敏) / AuditLog(哈希链审计) / MiniJson(扁平 JSON) / HostGuard(出网守卫)
             UI：MainForm(+partial Recon/Extract 页)、ChatPanel、SetupDialog
installer/ officeagent.iss（Inno 双 SKU：/DSKU=complete 或 lite）
新手手册.md 会计/新手向用户的第一份使用说明（大白话）
bin/       pdfium.dll（Chromium 109.0.5406 构建）
store/     components.ini（组件↔官方源↔SHA256 权威清单）+ wheels-requirements.txt（M2 依赖钉死）
payload/   离线组件仓（KB/.NET4.8/VCRT/Python，见 payload\README.md）
tools/     download-payload.ps1（URL+哈希锁定的一键重建）
tests/     WIN7-VM-验收.md + fixtures\test-smoke.pdf（pdfium 渲染冒烟固定输入）
```

## 构建与冒烟

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File build-all.ps1        # boot(net35) + host(net4x)
build\boot\OfficeAgentBoot.exe /report-only /console                     # exit 0=全绿, 2=有缺失
build\host\OfficeAgent.exe /convert <file> <pdf|csv|xlsx>                # 无头转换（三级总线）
build\host\OfficeAgent35.exe /recon ...                                   # 3.5 降级壳（同源码 csc 3.5 编译，无 .NET 4.8 机器用）
build\host\OfficeAgent.exe /renderpdf <in.pdf> <out.png>                 # 无头渲染（pdfium）
build\host\OfficeAgent.exe /recon <A.csv|A.xlsx> <B...> --keya 账号 --keyb 对方账号 `
    --debita 收入 --credita 支出 --debitb 借方 --creditb 贷方 --tol 0.01 [--all] [--save 模板名] [--tpl 模板名]
build\host\OfficeAgent.exe /merge <files...> --targets "科目,金额" --srcs "科目,金额" --amounts "金额" [--out 底稿.xlsx]
build\host\OfficeAgent.exe /invoices <pdf或文件夹...> --out 发票清单.xlsx
build\host\OfficeAgent.exe /masktest                                     # 脱敏引擎自测
build\host\OfficeAgent.exe /intents                                      # 意图路由自测（8 用例）
build\host\OfficeAgent.exe /suggesttest                                  # 列映射建议解析自测（6 用例）
build\host\OfficeAgent.exe /skilltest                                    # 技能体系自测（8 用例，含 python 技能端到端 + Skill 计划执行回归）
build\host\OfficeAgent.exe /skillrun row-stat <csv...>                   # 仅显示计划，不执行（缺 --yes）
build\host\OfficeAgent.exe /skillrun row-stat <csv...> --yes             # 显式确认后执行（全链落审计）
build\host\OfficeAgent.exe /pdftext <in.pdf>                             # PDF 文本层调试
runtime\py38\python.exe tools\gen_test_pdf.py out.pdf                    # 生成渲染测试输入
powershell -NoProfile -ExecutionPolicy Bypass -File tools\genfixtures.build.ps1   # 构建 fixture 生成器
build\test\GenFixtures.exe                                               # 重建 tests\fixtures（核对差异集/脏表/大表/发票PDF）
powershell -NoProfile -STA -ExecutionPolicy Bypass -File tests\e2e\e2e.ps1        # E2E 全量（CLI 冒烟 + UI 全流程）
```

## 工程红线（评审必查，与设计方案附录 A 一致）

1. 一切子进程只经 `ProcRunner`：编译期常量路径 + switch 白名单内联字面量参数，禁 cmd/shell 拼接。
2. 源码与构建脚本同守 C# 3.0 语法（无 LINQ/HashSet/System.Core），build 脚本纯 ASCII。
3. 安装器先 SHA256 校验、复制到固定暂存路径再执行；清单/网络数据不参与命令行构造。
4. 输出一律 UTF-8（文件带 BOM 由后续引擎决定）；子进程 stdio 显式 UTF-8。

## Win7 x64 兼容覆盖（M2.7）

| 目标环境                       | 支持状态    | 说明                                                     |
| -------------------------- | ------- | ------------------------------------------------------ |
| Win7 SP1 x64（各补丁级）         | ✅ 主目标   | 引导器 AnyCPU：x64 上原生 64 位进程，wusa/System32 探测不经 WOW64 重定向 |
| Windows Server 2008 R2 SP1 | ✅ 同内核支持 | 检测项显式识别（SM_SERVERR2），KB 组件与 Win7 SP1 通用                |
| Win7 SP1 x86               | ⚠️ 降级可用 | 引导器/宿主可装可跑；KB x86 变体四个已就位（2026-09-25），引导器按位宽自动选变体 |
| Win7 RTM（无 SP1）            | ⚠️ 阻塞引导 | .NET 4.8 微软硬性要求 SP1；引导页明示，不假装可装                        |
| 无管理员                       | ✅       | per-user 安装；.NET/VC 等系统组件需要时提权一次，拒绝则走 3.5 降级链          |
| 无外网                        | ✅       | payload 离线补全，SHA256 锁定                                 |

x64 专项修复（2026-09-25）：

- **wusa 启动**：x86 引导进程在 x64 上 `System32\wusa.exe` 会被 WOW64 重定向到不存在的 `SysWOW64\wusa.exe` → 引导器改 AnyCPU（x64 原生跑），32 位系统仍 32 位运行。
- **VC++ x64 误检**：`GetSystemWow64Directory` 返回的是 SysWOW64（非原生 System32），旧逻辑把 x86 运行库当成 x64 → 新增 `EnvDetect.DirOfBitness(bool)`，32/64 位进程下都取语义正确的目录（32 位进程经 Sysnative 探测 x64 库）。
- **LibreOffice 只查 x86 目录**：32 位进程的 `ProgramFiles` 指向 `(x86)` → 探测候选补 `ProgramW6432` 与 `ProgramFiles(x86)`，x64 版 LO 安装可被发现（Detect + ConvertEngine.FindSoffice 双处）。
- **.NET 4.8 注册表缺失回退**：精简镜像 Release 值缺失时，用 `Framework64\v4.0.30319\clr.dll` 文件版本 ≥ 4.0.30319.42000 回退判定（该目录不经 WOW64 重定向）。
