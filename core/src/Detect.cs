// OfficeAgent - 环境检测（boot 与 host 共用源码）
// 语法约束：.NET 3.5 / C# 3.0 兼容（无 LINQ、无 HashSet、无 System.Core、Path.Combine 仅两参）
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Management;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Principal;
using Microsoft.Win32;

namespace OfficeAgent.Core
{
    public enum DetectState { Ok, Missing, Unknown, Info }

    public class DetectItem
    {
        public string Id = "";
        public string Name = "";
        public DetectState State = DetectState.Unknown;
        public string Detail = "";
        public string FixKey = null;      // Installer 动作键；null = 无自动修复
        public bool NeedAdmin = false;
        public int Order = 99;
    }

    public static class EnvDetect
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        static extern uint GetSystemWow64Directory(char[] lpBuffer, uint nSize);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        static extern IntPtr LoadLibrary(string fileName);
        [DllImport("kernel32.dll")]
        static extern bool FreeLibrary(IntPtr hModule);
        [DllImport("kernel32.dll")]
        static extern bool IsWow64Process(IntPtr hProcess, out bool wow64Process);
        [DllImport("kernel32.dll")]
        static extern IntPtr GetCurrentProcess();

        public static string StateText(DetectState s)
        {
            switch (s)
            {
                case DetectState.Ok: return "OK";
                case DetectState.Missing: return "缺失";
                case DetectState.Unknown: return "未知";
                default: return "信息";
            }
        }

        // ---------- 基础环境 ----------

        public static bool IsWin7()
        {
            Version v = Environment.OSVersion.Version;
            return v.Major == 6 && v.Minor == 1;
        }

        public static bool Is64OS()
        {
            if (IntPtr.Size == 8) return true;
            bool wow;
            if (!IsWow64Process(GetCurrentProcess(), out wow)) return false;
            return wow;
        }

        // 32 位进程在 64 位系统上返回 SysWOW64（x86 库所在）；64 位进程返回 null（其 System32 即原生）；
        // 32 位系统返回 null。注意：此 API 的名字含义是「WOW64 目录」，不是「原生 System32」。
        public static string DirWow64()
        {
            char[] buf = new char[300];
            uint n = GetSystemWow64Directory(buf, 300);
            if (n == 0 || n >= 300) return null;
            return new string(buf, 0, (int)n);
        }

        // 按目标位宽取系统目录（对 32/64 位进程都语义正确）：
        //   x64 目录：64 位进程 → System32；32 位进程 → Sysnative（仅 x64 OS 可见）
        //   x86 目录：32 位进程 → System 文件夹；64 位进程 → SysWOW64
        public static string DirOfBitness(bool x64)
        {
            if (IntPtr.Size == 8)
            {
                if (x64) return Environment.GetFolderPath(Environment.SpecialFolder.System);
                return DirWow64();
            }
            if (!x64) return Environment.GetFolderPath(Environment.SpecialFolder.System);
            string sysroot = Environment.GetEnvironmentVariable("SystemRoot");
            if (sysroot == null || sysroot.Length == 0) sysroot = @"C:\Windows";
            string sysnative = Path.Combine(sysroot, "Sysnative");
            if (Directory.Exists(sysnative)) return sysnative;
            return null;   // 32 位系统没有 x64 目录
        }

        public static bool IsAdmin()
        {
            try
            {
                WindowsPrincipal p = new WindowsPrincipal(WindowsIdentity.GetCurrent());
                return p.IsInRole(WindowsBuiltInRole.Administrator);
            }
            catch { return false; }
        }

        // 从 exe 所在目录向上找安装根（标记：payload/ runtime/ store/ 任一存在）
        public static string FindRoot()
        {
            string start = AppDomain.CurrentDomain.BaseDirectory;
            string cur = start;
            for (int i = 0; i < 5; i++)
            {
                if (Directory.Exists(Path.Combine(cur, "payload"))) return TrimDir(cur);
                if (Directory.Exists(Path.Combine(cur, "runtime"))) return TrimDir(cur);
                if (Directory.Exists(Path.Combine(cur, "store"))) return TrimDir(cur);
                DirectoryInfo up = Directory.GetParent(cur);
                if (up == null) break;
                cur = up.FullName;
            }
            return TrimDir(start);
        }

        static string TrimDir(string d)
        {
            if (d == null) return "";
            return d.EndsWith("\\") ? d.Substring(0, d.Length - 1) : d;
        }

        // ---------- 各项检测 ----------

        public static List<DetectItem> DetectAll(string root)
        {
            List<DetectItem> list = new List<DetectItem>();
            bool win7 = IsWin7();
            bool is64 = Is64OS();

            // 1. 操作系统 / SP（Server 2008 R2 与 Win7 同为 6.1 内核，一并支持）
            DetectItem os = new DetectItem();
            os.Id = "os"; os.Name = "操作系统"; os.Order = 1;
            if (win7)
            {
                bool isServer = Environment.GetEnvironmentVariable("SM_SERVERR2") == "1";
                string osName = isServer ? "Windows Server 2008 R2" : "Windows 7";
                string sp = Environment.OSVersion.ServicePack ?? "";
                if (sp.IndexOf("Service Pack 1", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    os.State = DetectState.Ok;
                    os.Detail = osName + " SP1" + (isServer ? "（与 Win7 SP1 同内核，组件补丁通用）" : "");
                }
                else
                {
                    os.State = DetectState.Missing;
                    os.Detail = osName + " 未装 SP1（.NET 4.8 硬性要求 SP1），请先离线安装 SP1";
                }
            }
            else
            {
                os.State = DetectState.Info;
                os.Detail = "非 Win7（" + Environment.OSVersion.VersionString + "），Win7 专项检查按无需处理";
            }
            list.Add(os);

            if (win7)
            {
                AddKb(list, "kb4490628", "服务堆栈更新 KB4490628（SHA-2 前置）", 2, "install_kb4490628");
                AddKb(list, "kb4474419", "SHA-2 代码签名支持 KB4474419", 3, "install_kb4474419");
                AddKb(list, "kb3020369", "服务堆栈更新 KB3020369（TLS 1.2 前置）", 4, "install_kb3020369");
                AddKb(list, "kb3140245", "WinHTTP TLS 1.2 支持 KB3140245", 5, "install_kb3140245");
                DetectItem kb838 = new DetectItem();
                kb838.Id = "kb2670838"; kb838.Name = "平台更新 KB2670838（.NET 4.8 前置）";
                kb838.Order = 6; kb838.NeedAdmin = true;
                HotfixState st838 = QueryHotfix("KB2670838");
                bool has838 = (st838 == HotfixState.Present);
                if (!has838)
                {
                    // KB2670838 在两个视图都会装 d3dcompiler_47；按位宽取原生 System32 检查最可靠
                    string d47 = Path.Combine(DirOfBitness(true), "d3dcompiler_47.dll");
                    has838 = FileVersionAtLeast(d47, "6.2.9200.16492");
                }
                if (has838)
                {
                    kb838.State = DetectState.Ok;
                    kb838.FixKey = null;
                    kb838.Detail = "已安装";
                }
                else if (st838 == HotfixState.Unavailable && !File.Exists(Path.Combine(DirOfBitness(true), "d3dcompiler_47.dll")))
                {
                    // WMI 不可用且文件也不在：两条路都没结论 → 未知，不当作缺失
                    //（避免 WMI 抖动把这一项推进"已提示缺失"集合，导致反复弹窗）
                    kb838.State = DetectState.Unknown;
                    kb838.FixKey = null;
                    kb838.Detail = "无法确认（WMI 查询失败且未找到 d3dcompiler_47.dll）";
                }
                else
                {
                    kb838.State = DetectState.Missing;
                    kb838.FixKey = "install_kb2670838";
                    kb838.Detail = "DirectWrite/D3DCompiler_47，.NET 4.8 在 Win7 SP1 的硬前置";
                }
                list.Add(kb838);
            }

            // .NET Framework 4.8（Release >= 528040）。x86 进程读注册表得到的是 WOW64 视图；
            // x64 系统上 .NET 4.8 双视图均有 Release，正常足够。若注册表缺失（极少数精简镜像），
            // 用 Framework64\clr.dll 文件版本回退判定（Microsoft.NET 目录不经 WOW64 重定向）。
            DetectItem net = new DetectItem();
            net.Id = "netfx48"; net.Name = ".NET Framework 4.8（主程序）";
            net.Order = 7; net.NeedAdmin = true;
            int rel = Net48Release();
            bool clr64Is48 = false;
            if (rel < 528040 && is64)
            {
                clr64Is48 = FileVersionAtLeast(Path.Combine(
                    Path.Combine(@"C:\Windows", Path.Combine("Microsoft.NET", Path.Combine("Framework64", "v4.0.30319"))),
                    "clr.dll"), "4.0.30319.42000");
            }
            if (rel >= 528040)
            {
                net.State = DetectState.Ok; net.Detail = "Release=" + rel;
            }
            else if (clr64Is48)
            {
                net.State = DetectState.Ok; net.Detail = "注册表 Release 缺失，Framework64 clr.dll 版本判定为 4.8";
            }
            else
            {
                net.State = DetectState.Missing; net.FixKey = "install_netfx48";
                net.Detail = (rel > 0 ? "检测到旧版 .NET 4.x（Release=" + rel + "）" : "未安装") + "；无管理员权限时自动转入 .NET 3.5 降级壳";
            }
            list.Add(net);

            // VC++ 2015-2022 运行库（按位宽取各自系统目录；x86 给主程序与 32 位工具链，x64 给 LO 等）
            string dir86 = DirOfBitness(false);
            string dir64 = DirOfBitness(true);
            bool vc86 = FileVersionAtLeast(Path.Combine(dir86, "msvcp140.dll"), "14.30.0.0")
                        && File.Exists(Path.Combine(dir86, "ucrtbase.dll"));
            bool vc64 = !is64 || FileVersionAtLeast(Path.Combine(dir64 ?? dir86, "msvcp140.dll"), "14.30.0.0");

            DetectItem v86 = new DetectItem();
            v86.Id = "vcredist-x86"; v86.Name = "VC++ 2015-2022 运行库 (x86)";
            v86.Order = 8; v86.NeedAdmin = true;
            v86.State = vc86 ? DetectState.Ok : DetectState.Missing;
            v86.FixKey = vc86 ? null : "install_vcredist_x86";
            v86.Detail = vc86 ? "已安装" : "32 位程序依赖；离线包 payload\\vcrt";
            list.Add(v86);

            DetectItem v64 = new DetectItem();
            v64.Id = "vcredist-x64"; v64.Name = "VC++ 2015-2022 运行库 (x64)";
            v64.Order = 9; v64.NeedAdmin = true;
            if (!is64) { v64.State = DetectState.Info; v64.Detail = "32 位系统无需"; }
            else
            {
                v64.State = vc64 ? DetectState.Ok : DetectState.Missing;
                v64.FixKey = vc64 ? null : "install_vcredist_x64";
                v64.Detail = vc64 ? "已安装" : "LibreOffice 等 64 位工具依赖；离线包 payload\\vcrt";
            }
            list.Add(v64);

            // Python 3.8.10 embeddable（技能/文档引擎 sidecar）
            DetectItem py = new DetectItem();
            py.Id = "py38"; py.Name = "Python 3.8.10 (embeddable)";
            py.Order = 10;
            string pyExe = Path.Combine(Path.Combine(root, "runtime"), Path.Combine("py38", "python.exe"));
            if (!File.Exists(pyExe))
            {
                py.State = DetectState.Missing; py.FixKey = "unpack_py38";
                py.Detail = "免安装、免管理员，从离线包解压";
            }
            else
            {
                string ver;
                int code = ProcRunner.RunVersionProbe(pyExe, "--version", out ver, 15);
                if (code == 0 && ver.Length > 0)
                {
                    py.State = DetectState.Ok; py.Detail = ver;
                }
                else if (code == ProcRunner.ExitTimeout)
                {
                    // 超时不等于"坏了"：低配机/杀软扫描时 python 起得慢很常见。
                    // 判 Missing 会把它写进"已提示缺失"集合，下次探针正常又变成 Ok，
                    // 集合来回变化 → 反复弹修复窗口。判 Unknown 更诚实，也不参与闸门。
                    py.State = DetectState.Unknown; py.FixKey = null;
                    py.Detail = "存在但探针超时（15s 未返回），可能机器负载高；可重试检测";
                }
                else
                {
                    // 非零退出（如缺 VC++ 运行库导致加载失败，exit=126）：这是**真**问题，
                    // 重解压确实能修（内置 vcruntime），保留 Missing。
                    py.State = DetectState.Missing; py.FixKey = "unpack_py38";
                    py.Detail = "存在但探针失败（exit=" + code + "），重解压可修复";
                }
            }
            list.Add(py);

            // 后续版本组件（M1/M2 落地，不阻塞 M0 全绿）
            DetectItem wheels = new DetectItem();
            wheels.Id = "wheels"; wheels.Name = "离线 wheel 仓（技能依赖）"; wheels.Order = 11;
            wheels.State = DetectState.Info;
            string whlDir = Path.Combine(Path.Combine(root, "store"), "wheels");
            wheels.Detail = Directory.Exists(whlDir) && Directory.GetFiles(whlDir, "*.whl").Length > 0
                ? "已有 wheel 缓存" : "M2 组件：tools\\download-payload 或 pip download 填充 store\\wheels";
            list.Add(wheels);

            DetectItem lo = new DetectItem();
            lo.Id = "lo76"; lo.Name = "LibreOffice 7.6（万能转换引擎）"; lo.Order = 12;
            lo.FixKey = "unpack_lo76";
            // 64 位系统上 32 位进程的 ProgramFiles 指向 Program Files (x86)，
            // 需补 ProgramW6432（真 Program Files）与 x86 视图，覆盖 x64/x86 两版 LO 安装
            //（.NET 3.5 无 SpecialFolder.ProgramFilesX86，读环境变量）
            string pf64 = Environment.GetEnvironmentVariable("ProgramW6432");
            string pf86 = Environment.GetEnvironmentVariable("ProgramFiles(x86)");
            List<string> loCandidates = new List<string>();
            loCandidates.Add(Path.Combine(Path.Combine(root, "lo76"), Path.Combine(Path.Combine("LibreOffice", "program"), "soffice.exe")));
            loCandidates.Add(Path.Combine(Path.Combine(root, "lo76"), Path.Combine("program", "soffice.exe")));
            if (pf64 != null && pf64.Length > 0)
                loCandidates.Add(Path.Combine(Path.Combine(pf64, "LibreOffice"), Path.Combine("program", "soffice.exe")));
            if (pf86 != null && pf86.Length > 0)
                loCandidates.Add(Path.Combine(Path.Combine(pf86, "LibreOffice"), Path.Combine("program", "soffice.exe")));
            loCandidates.Add(Path.Combine(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "LibreOffice"), Path.Combine("program", "soffice.exe")));
            string loPath = FindExisting(loCandidates.ToArray());
            if (loPath != null) { lo.State = DetectState.Ok; lo.Detail = loPath; }
            else { lo.State = DetectState.Info; lo.Detail = "M1 组件：无 Office 时的转换兜底；payload\\lo76 放入 zip 可自动解压"; }
            list.Add(lo);

            DetectItem pdfium = new DetectItem();
            pdfium.Id = "pdfium"; pdfium.Name = "pdfium.dll（PDF 渲染）"; pdfium.Order = 13;
            string pdfiumPath = Path.Combine(Path.Combine(root, "bin"), "pdfium.dll");
            IntPtr hMod = File.Exists(pdfiumPath) ? LoadLibrary(pdfiumPath) : IntPtr.Zero;
            if (hMod != IntPtr.Zero)
            {
                FreeLibrary(hMod);
                pdfium.State = DetectState.Ok; pdfium.Detail = pdfiumPath;
            }
            else
            {
                pdfium.State = DetectState.Info;
                pdfium.Detail = "M1 组件：Chromium 109 世代构建（Win7 上限版），放入 bin\\ 即可";
            }
            list.Add(pdfium);

            // Office / WPS COM（转换优先级，可选）
            bool officeCom = Type.GetTypeFromProgID("Excel.Application") != null;
            DetectItem office = new DetectItem();
            office.Id = "office-com"; office.Name = "MS Office COM"; office.Order = 14;
            office.State = officeCom ? DetectState.Ok : DetectState.Info;
            office.Detail = officeCom ? "可用（转换最高优先级）" : "未安装（可选，走 WPS/LibreOffice 兜底）";
            list.Add(office);

            bool wpsCom = Type.GetTypeFromProgID("KET.Application") != null;
            DetectItem wps = new DetectItem();
            wps.Id = "wps-com"; wps.Name = "WPS COM（KET/KWPS/KWPP）"; wps.Order = 15;
            wps.State = wpsCom ? DetectState.Ok : DetectState.Info;
            wps.Detail = wpsCom ? "可用（转换次优先级）" : "未安装（可选）";
            list.Add(wps);

            // TLS 1.2 出网探针（LLM 通道需要；离线流程不阻塞）
            DetectItem tls = new DetectItem();
            tls.Id = "tls12"; tls.Name = "TLS 1.2 出网探针"; tls.Order = 16;
            string tlsResult = ProbeTls12();
            if (tlsResult != null) { tls.State = DetectState.Ok; tls.Detail = tlsResult; }
            else { tls.State = DetectState.Unknown; tls.Detail = "出网失败（可能离线/代理），离线模式不受影响"; }
            list.Add(tls);

            // 磁盘空间
            // 迟滞（hysteresis）：可用空间是**实时值**，若阈值恰好卡在 2GB，
            // 磁盘紧张时会在 Ok/Missing 之间来回跳 —— 每次跳都改变"缺失项集合"，
            // 于是修复窗口被反复唤醒（用户抱怨的"老是被打断"）。
            // 因此：低于 2GB 才判缺失；已判缺失后要回到 2.5GB 以上才恢复 Ok。
            // 已判缺失的状态借助进程内静态变量记住（同一次运行内有效；
            // 跨进程时用 BootGate 的集合语义也不会重复弹）。
            DetectItem disk = new DetectItem();
            disk.Id = "disk"; disk.Name = "安装盘剩余空间"; disk.Order = 17;
            try
            {
                string driveRoot = Path.GetPathRoot(Path.GetFullPath(root));
                DriveInfo d = new DriveInfo(driveRoot);
                long free = d.IsReady ? d.AvailableFreeSpace : 0;
                // 低水位 2GiB；高水位 2.5GiB（常量提到类级，便于自测直接断言算术）。
                // 注意必须先乘后除：写成 5L / 2 * 1024 * 1024 * 1024 会因**整数除法**
                // 先算成 2，高水位退化为 2GiB 与低水位相等 → 迟滞完全失效。
                long threshold = diskWasLow ? DiskHighWater : DiskLowWater;
                if (free >= threshold)
                {
                    diskWasLow = false;
                    disk.State = DetectState.Ok;
                    disk.Detail = (free / 1024 / 1024) + " MB 可用";
                }
                else
                {
                    diskWasLow = true;
                    disk.State = DetectState.Missing;
                    // 用 MB 显示而非 GB：threshold 可能是 2.5GiB，按 GB 整除会截断成 "2GB"，
                    // 与真实阈值不符（审计发现的次要问题）。
                    disk.Detail = "可用空间不足 " + (threshold / 1024 / 1024) + " MB，无法安装完整组件";
                }
            }
            catch (Exception ex) { disk.State = DetectState.Unknown; disk.Detail = ex.Message; }
            list.Add(disk);

            // 权限 / 位宽（信息项）
            DetectItem admin = new DetectItem();
            admin.Id = "admin"; admin.Name = "当前权限"; admin.Order = 18;
            admin.State = DetectState.Info;
            admin.Detail = IsAdmin() ? "管理员（可自动装 .NET/VC 等系统组件）" : "普通用户（缺系统组件时按钮提权，或转入 3.5 降级壳）";
            list.Add(admin);

            DetectItem bit = new DetectItem();
            bit.Id = "bitness"; bit.Name = "位宽"; bit.Order = 19;
            bit.State = DetectState.Info;
            bit.Detail = (is64 ? "64 位系统" : "32 位系统") + " / 32 位主进程（x86 双系统通用）";
            list.Add(bit);

            list.Sort(CompareByOrder);
            return list;
        }

        static int CompareByOrder(DetectItem a, DetectItem b) { return a.Order.CompareTo(b.Order); }

        static void AddKb(List<DetectItem> list, string id, string name, int order, string fixKey)
        {
            DetectItem it = new DetectItem();
            it.Id = id; it.Name = name; it.Order = order; it.NeedAdmin = true;
            HotfixState st = QueryHotfix(id.ToUpperInvariant());
            if (st == HotfixState.Present)
            {
                it.State = DetectState.Ok;
                it.FixKey = null;
                it.Detail = "已安装";
            }
            else if (st == HotfixState.Absent)
            {
                it.State = DetectState.Missing;
                it.FixKey = fixKey;
                it.Detail = "离线包 payload\\kb（wusa 静默安装）";
            }
            else
            {
                // 查询不可用（WMI 超时/服务未起）：**不能判 Missing**。
                // 判 Missing 的后果很严重：它会被写进"已提示缺失项"指纹，
                // 下次 WMI 正常时集合变化 → 又弹修复窗口，用户被反复打断
                //（低配 Win7 上 WMI 首次查询超时是常态）。判 Unknown 则不参与闸门。
                it.State = DetectState.Unknown;
                it.FixKey = null;
                it.Detail = "无法确认（WMI 查询失败；稍后重试或手动确认）";
            }
            list.Add(it);
        }

        // ---------- 原子检测方法 ----------

        // 热修检测的三态结果。区分"确实没装"与"查不出来"是关键：
        // 前者该提示安装，后者只应显示为未知，绝不能当作缺失去驱动修复提示。
        public enum HotfixState { Present, Absent, Unavailable }

        static List<string> hotfixCache = null;

        // 磁盘检测的迟滞状态：记住"上次是否判为空间不足"，避免在阈值上下抖动
        static bool diskWasLow = false;

        // 磁盘阈值（公开常量，供自测直接断言算术，防止 5L/2 整数除法回归）
        public const long DiskLowWater = 2L * 1024 * 1024 * 1024;        // 2 GiB：低于此值判缺失
        public const long DiskHighWater = 5L * 1024 * 1024 * 1024 / 2;   // 2.5 GiB：已缺失时回到此值才转 Ok


        // 一次 WMI 查询枚举全部热修并缓存（低配机优化：避免每个 KB 各跑一遍 WMI，4 次→1 次）
        public static bool HasHotfix(string kbId)
        {
            return QueryHotfix(kbId) == HotfixState.Present;
        }

        // 三态查询：Present=已安装；Absent=确认未安装；Unavailable=查不出来（不缓存，可重试）
        public static HotfixState QueryHotfix(string kbId)
        {
            if (hotfixCache == null)
            {
                List<string> ids = new List<string>();
                try
                {
                    using (ManagementObjectSearcher searcher = new ManagementObjectSearcher(
                        "SELECT HotFixID FROM Win32_QuickFixEngineering"))
                    {
                        foreach (ManagementObject o in searcher.Get())
                        {
                            try
                            {
                                object v = o["HotFixID"];
                                if (v != null) ids.Add(v.ToString());
                            }
                            catch { }
                            try { o.Dispose(); } catch { }
                        }
                    }
                    hotfixCache = ids;
                }
                catch { return HotfixState.Unavailable; }   // 查询失败不缓存，下次可重试
            }
            foreach (string s in hotfixCache)
            {
                if (string.Equals(s, kbId, StringComparison.OrdinalIgnoreCase)) return HotfixState.Present;
            }
            return HotfixState.Absent;
        }

        static int Net48Release()
        {
            try
            {
                using (RegistryKey k = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"))
                {
                    if (k == null) return 0;
                    object v = k.GetValue("Release", 0);
                    return (v is int) ? (int)v : 0;
                }
            }
            catch { return 0; }
        }

        static bool FileVersionAtLeast(string path, string minVersion)
        {
            try
            {
                if (path == null || !File.Exists(path)) return false;
                string v = FileVersionInfo.GetVersionInfo(path).FileVersion;
                return CompareVer(v, minVersion) >= 0;
            }
            catch { return false; }
        }

        static int CompareVer(string a, string b)
        {
            int[] va = ParseVer(a), vb = ParseVer(b);
            for (int i = 0; i < 4; i++)
            {
                if (va[i] != vb[i]) return va[i] < vb[i] ? -1 : 1;
            }
            return 0;
        }

        static int[] ParseVer(string s)
        {
            int[] r = new int[] { 0, 0, 0, 0 };
            if (s == null) return r;
            string[] parts = s.Split('.');
            for (int i = 0; i < 4 && i < parts.Length; i++)
            {
                int n;
                if (int.TryParse(parts[i].Trim(), out n)) r[i] = n;
            }
            return r;
        }

        static string FindExisting(string[] paths)
        {
            foreach (string p in paths)
            {
                try { if (p != null && File.Exists(p)) return p; } catch { }
            }
            return null;
        }

        // TLS 1.2 出网探针：成功返回描述，失败返回 null。值 3072 = Tls12（.NET 3.5 无枚举，强转）
        public static string ProbeTls12()
        {
            try
            {
                ServicePointManager.SecurityProtocol |= (SecurityProtocolType)3072;
                HttpWebRequest req = (HttpWebRequest)WebRequest.Create("https://www.msftconnecttest.com/connecttest.txt");
                req.Timeout = 5000;
                req.ReadWriteTimeout = 5000;
                req.KeepAlive = false;
                using (HttpWebResponse resp = (HttpWebResponse)req.GetResponse())
                {
                    using (Stream s = resp.GetResponseStream())
                    {
                        byte[] buf = new byte[256];
                        try { s.Read(buf, 0, 256); } catch { }
                    }
                    return "TLS1.2 出网正常（HTTP " + (int)resp.StatusCode + "）";
                }
            }
            catch { return null; }
        }
    }
}
