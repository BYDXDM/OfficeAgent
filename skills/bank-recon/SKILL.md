# bank-recon —— 银行流水核对（内置技能示例）

## 用途
任意两个 xlsx/xls/csv（银行流水 vs 账面记录）按键列勾稽，输出差异表（金额不等/仅A有/仅B有）
与 Verifier 勾稽自证。

## 使用
- 会话：「核对 <文件A> 和 <文件B>」（IntentRouter 场景词：表格核对/对账/勾稽/流水核对）
- 表格核对页（可自动猜列、模板复用）
- CLI：`OfficeAgent.exe /recon A B --keya 账号 --keyb 对方账号 --debita 收入 --credita 支出 --debitb 借方 --creditb 贷方`

## 边界
- `runtime: builtin`：执行体内置于主程序（ReconEngine），不经 sidecar。
- 金额全部 decimal；解析失败的金额行剔除并列错误清单，绝不静默当 0。
- 原始文件只读；差异表写出在 A 侧文件目录。
- `libreoffice?` 为可选依赖：仅当输入是 .xls 需要预转 xlsx 时用到，缺失时给出降级提示。
