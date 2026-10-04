# installer/lang —— Inno Setup 语言文件

## ChineseSimplified.isl

- **来源**：https://github.com/kira-96/Inno-Setup-Chinese-Simplified-Translation
- **分支/标签**：`6.1.0+v2`
- **许可**：该仓库为社区翻译（基于网络资源整理），非 jrsoftware 官方发行
- **编码**：UTF-8 **带 BOM**（Inno 要求；无 BOM 会导致中文乱码）

## ★ 为什么固定用 6.1.0+v2，而不是最新版

本机安装的是 **Inno Setup 6.3.3**（`ISCC.exe`，路径 `C:\Program Files (x86)\Inno Setup 6\`）。

Inno 的语言文件是**版本敏感**的：

| 候选版本 | 与 6.3.3 的结果 |
|---|---|
| `6.1.0+v2` | ✅ **零警告**，276 条 message 与 6.3.3 的 `Default.isl` **完全一致**（0 缺失 / 0 多余） |
| `6.4.0+` | ⚠️ 5 条未知 message 警告（`ExtractionLabel`、`ButtonStopExtraction`、`StopExtraction`、`ErrorExtractionAborted`、`ErrorExtractionFailed` —— 这些是 6.4.0 新增的） |
| `6.5.0+`（main） | ⚠️ 更多未知 message，风险更高 |

> 升级 Inno Setup 后，需同步换成对应版本的语言文件，并重跑一次编译确认零警告。

## 校验方法

```bash
# 1. 确认 message 覆盖与当前 Default.isl 一致（缺失会静默回落英文，多余会导致编译警告）
#    对照 installer/lang/ChineseSimplified.isl 与 C:\Program Files (x86)\Inno Setup 6\Default.isl 的 message 名集合

# 2. 编译并检查有无 Warning
ISCC /DSKU=lite installer/officeagent.iss
# 期望：Successful compile，且无 "Message name ... is not recognized" 警告
```
