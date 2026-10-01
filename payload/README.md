# payload —— 离线组件仓

引导器首启自检后，缺失组件从这里补全（全程无需联网）。组件与 SHA256 的权威对照见 `store\components.ini`。

## 当前已就位（x64 Win7 SP1 全套，约 700MB）

| 目录 | 文件 | 用途 |
|---|---|---|
| `kb\` | kb4490628-x64.msu / kb4474419-v3-x64.msu / kb3140245-x64.msu / windows6.1-kb2670838-x64.msu | SP1 后置补丁：SSU、SHA-2 签名、TLS1.2、.NET 4.8 硬前置 |
| `ndp48\` | ndp48-x86x64allos-enu.exe | .NET Framework 4.8 离线安装器（主程序运行时） |
| `vcrt\` | VC_redist.x86.exe / VC_redist.x64.exe | VC++ 2015-2022 固定版运行库 |
| `py38\` | python-3.8.10-embed-win32.zip / -amd64.zip / get-pip.py | 技能与文档引擎 sidecar（免安装解压即用） |
| `lo76\` | **LibreOffice_7.6.7_Win_x86-64.zip（446MB，2026-09-25 就位）** | 万能转换引擎（Win7 末代分支 7.6.x，MultilingualAll 含 zh-CN）；boot 解包到 `<root>\lo76\`。来源：USTC 镜像的 TDF 官方便携 paf.exe 免安装展开后重打包，URL+SHA256 见 `store\components.ini` |
| `kb\`（x86 变体） | **四个 KB 的 x86 msu（2026-09-25 就位）** | 32 位 Win7 SP1 补丁齐套：boot 按系统位宽自动选择变体（Installer.cs）。来源：Microsoft Update Catalog（脚本 `tools\download-kb-x86.ps1`），SHA256 见 `store\components.ini` |
| `store\wheels\` | **21 个 wheel（2026-09-25 就位，26MB）** | py3.8 双架构（win32+amd64）技能依赖全量：openpyxl/pdfplumber/python-pptx/lxml/Pillow/pypdf 等 + pip 自身；来源 NJU pypi 镜像（pip download 双平台跑批） |

## 待补位（M3+）

无 —— pdfium、wheels（21 个）、x86 KB（四个）均已落实（2026-09-25），待补位清单清空。

## 32 位 Win7 注意

x86 KB 变体四个已就位（2026-09-25，Microsoft Update Catalog 抓取，脚本 `tools\download-kb-x86.ps1`），SHA256 见 `store\components.ini`；引导器已按系统位宽自动选择变体（`boot\src\Installer.cs` 的 `EnvDetect.Is64OS()` 分支）。

## 重建 payload

断网机器可跳过本文件直接使用随包 payload；有网机器可用 `tools\download-payload.ps1` 一键重建（URL 与哈希全部锁定）。
