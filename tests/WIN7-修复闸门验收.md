# Win7 修复闸门闭环验收（拷到真机执行）

> 目的：验证"缺失 → 只提醒一次 → 之后直接开 agent"这条链路在 **Win7 真机**上成立。
> 为什么必须上真机：开发机（Win11）检测结果是"组件齐全"，`ShouldPrompt` 永远返回 false，
> **弹窗分支根本跑不到**；而且本机无 .NET 3.5，编不出真正的 net35 引导器。
>
> 对应提交：`25d21de`（闸门）、`c4cf2da`（防抖动）、`c60ee3a`（EnvSafe）。
> 闸门状态文件：`%LOCALAPPDATA%\OfficeAgent\boot-gate.json`

---

## 0. 前置：先确认闸门代码真的在包里

在 Win7 上打开 cmd：

```bat
cd /d C:\OfficeAgent
findstr /C:"ShouldPrompt" boot\src\Program.cs
```
- 有输出 → 提交是对的，继续。
- 没输出 → **你拷过去的是旧包**，闸门不在里面，后面的步骤全无意义。

> 如果拷的是**编译好的 exe** 而不是源码，改用日志判断：跑一次看
> `%LOCALAPPDATA%\OfficeAgent\logs\bootstrap-*.txt` 里有没有
> "无需修复提示（组件齐全，或缺失项此前已提醒过）" 这句话。有就是新包。

---

## 1. 场景 A：组件齐全 → 应直接开 agent，不弹窗

```bat
del "%LOCALAPPDATA%\OfficeAgent\boot-gate.json" 2>nul
C:\OfficeAgent\boot\OfficeAgentBoot.exe
```

**期望**
- 不出现"环境引导"窗口
- 直接出现 agent 主窗口
- 日志末尾是：`无需修复提示（组件齐全，或缺失项此前已提醒过），直接启动主程序`
- **`boot-gate.json` 不应被创建**（无缺失项就不写盘）

---

## 2. 场景 B：有缺失 → 第一次弹窗，第二次直开（★核心）

先人为制造一个缺失项（把 Python 挪走）：

```bat
move C:\OfficeAgent\runtime C:\OfficeAgent\runtime_off
del "%LOCALAPPDATA%\OfficeAgent\boot-gate.json" 2>nul
```

**第 1 次**（从未提醒过）
```bat
C:\OfficeAgent\boot\OfficeAgentBoot.exe
```
**期望**
- **弹出** "OfficeAgent 环境引导" 窗口
- 日志：`发现新的缺失组件，显示修复窗口一次：Python 3.8.10 (embeddable)`
- 窗口里那一行是红色"缺失"
- 生成 `boot-gate.json`，内容形如：
  ```json
  { "promptedIds": ["py38"], "promptedAt": "..." }
  ```
- 点窗口的 **"启动主程序"**，agent 能起来（或直接关掉窗口）

**第 2 次**（同一个缺失项，已提醒过）← 这就是用户要的行为
```bat
C:\OfficeAgent\boot\OfficeAgentBoot.exe
```
**期望**
- **不再弹窗**
- 直接出现 agent 主窗口
- 日志：`无需修复提示（组件齐全，或缺失项此前已提醒过），直接启动主程序`
- `boot-gate.json` 内容**不变**（还是 `["py38"]`）

**第 3 次**（再确认一次，排除偶发）
```bat
C:\OfficeAgent\boot\OfficeAgentBoot.exe
```
**期望**：同第 2 次。

---

## 3. 场景 C：修好之后不该再弹（★我改过的语义）

```bat
move C:\OfficeAgent\runtime_off C:\OfficeAgent\runtime
C:\OfficeAgent\boot\OfficeAgentBoot.exe
```
**期望**：不弹窗，直接开 agent。
（缺失项消失了；即使它又坏回来，因为 `py38` 已在 `promptedIds` 里，也不会再弹——
这是"老问题不必反复提醒"的设计。）

---

## 4. 场景 D：出现**新**缺失项 → 应再提醒一次

再制造一个**不同**的缺失项（例如挪走 VC++ 相关或把 payload 改名）：

```bat
rename C:\OfficeAgent\payload payload_off
C:\OfficeAgent\boot\OfficeAgentBoot.exe
```
**期望**
- **重新弹窗一次**（因为出现了从未提示过的新缺失项）
- 日志：`发现新的缺失组件，显示修复窗口一次：...`
- `boot-gate.json` 的 `promptedIds` **变长**（新项被并入，旧项保留）

验完复原：
```bat
rename C:\OfficeAgent\payload_off payload
```

---

## 5. 场景 E：`/ui` 强制显示（手工修复入口）

```bat
C:\OfficeAgent\boot\OfficeAgentBoot.exe /ui
```
**期望**：无论闸门怎么判，**都弹出**自检窗口。
（这正是 `RepairLauncher` 现在传的参数——agent 里让模型调 `repair_environment`
应该弹 UAC 然后**出现窗口**；此前没传 `/ui` 会导致"弹了 UAC 却什么都没有"。）

---

## 6. 场景 F：防抖动（可选，但很有价值）

**F1 磁盘迟滞**：把某个盘塞到只剩 1.9GB 再删到 2.2GB 反复跑引导器。
**期望**：`disk` 项不会在 Ok/缺失之间来回跳导致反复弹窗。

**F2 WMI 抖动**：临时停掉 WMI 服务再跑检测。
```bat
net stop Winmgmt /y
C:\OfficeAgent\boot\OfficeAgentBoot.exe /report-only /console
net start Winmgmt
```
**期望**：KB 相关项显示为 **"未知"（橙色）**，而**不是"缺失"（红色）**。
即：查不出来 ≠ 没装。这样 WMI 超时不会污染闸门。

---

## 7. 结果记录模板

| 场景 | 期望 | 实际 | 通过 |
|---|---|---|---|
| A 组件齐全 | 不弹窗、直开 agent、无状态文件 | | |
| B1 首次缺失 | 弹窗一次、写入 `["py38"]` | | |
| B2 缺失仍在·第 2 次 | **不弹窗、直开 agent** | | |
| B3 第 3 次 | 同上 | | |
| C 修好后 | 不弹窗 | | |
| D 新缺失项 | 再弹一次、集合变长 | | |
| E `/ui` | 强制弹窗 | | |
| F1 磁盘迟滞 | 不来回跳 | | |
| F2 WMI 停掉 | KB 显示"未知"而非"缺失" | | |

---

## 8. 出问题怎么定位

| 现象 | 先看什么 |
|---|---|
| 每次都弹窗 | `boot-gate.json` 是否被写得**每次都不一样**？对比两次运行后的 `promptedIds`。若内容在变，说明有探测项在抖动 → 看 `/report-only` 输出里哪个项在 Ok/缺失/未知之间变 |
| 从不弹窗（但确实缺东西） | `boot-gate.json` 里是否已有该项（正常）；或删除该文件再试 |
| 弹了 UAC 但无窗口 | 确认 `RepairLauncher` 传了 `/ui`（场景 E 单独测）；老包没有这个修复 |
| 状态文件写不进去 | `%LOCALAPPDATA%\OfficeAgent` 是否可写；写失败会 fail-open（每次都提醒），不会静默吞掉 |
| agent 起不来 | 看 `bootstrap-*.txt` 最后一行；若为"未找到主程序"，确认 `<root>\host\OfficeAgent.exe` 存在（boot 按**安装后布局**找，不是 `build\host\`） |

日志目录：`%LOCALAPPDATA%\OfficeAgent\logs\bootstrap-*.txt`
