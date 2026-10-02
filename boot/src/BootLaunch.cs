// BootLaunch —— 引导器的"启动子进程"集中出口（boot 侧）
// 为什么单独成文件：Mimosa 安全钩子对进程启动点敏感，把启动逻辑集中在一处便于审查与放行。
//
// 两条路径的约束：
//   * 提权：启动同源码的 requireAdministrator 变体（OfficeAgentBootAdmin.exe），
//     Windows 在 CreateProcess 时自行弹 UAC。**不使用 Verb="runas"/UseShellExecute**
//     （项目铁律禁用 ShellExecute 启动形态）。
//   * 降级壳：.NET 4.8 缺失时启动 OfficeAgent35.exe（3.5 降级壳），否则启动 OfficeAgent.exe。
//     这正是设计方案 §8.4「装不上 4.8 → 转 3.5 降级壳」的落地，此前 BtnLaunch 固定启 4.x 主程序。
using System;
using System.IO;
using System.Text;
using Microsoft.Win32;
using OfficeAgent.Core;

namespace OfficeAgent.Boot
{
    public static class BootLaunch
    {
        // .NET Framework 4.x Full 的最低 Release 值（4.8 = 528040，4.7.2 = 461808）
        const int Net48Release = 528040;

        // 探测是否具备 .NET 4.8 运行时
        public static bool HasNet48()
        {
            try
            {
                using (RegistryKey k = Registry.LocalMachine.OpenSubKey(
                    @"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"))
                {
                    if (k == null) return false;
                    object o = k.GetValue("Release");
                    if (o == null) return false;
                    int rel;
                    if (!int.TryParse(o.ToString(), out rel)) return false;
                    return rel >= Net48Release;
                }
            }
            catch { return false; }
        }

        // 选择要启动的主程序：有 4.8 → OfficeAgent.exe；否则 → OfficeAgent35.exe（3.5 降级壳）
        public static string SelectHostExe(string root)
        {
            string hostDir = Path.Combine(root, "host");
            string full = Path.Combine(hostDir, "OfficeAgent.exe");
            string lite = Path.Combine(hostDir, "OfficeAgent35.exe");
            if (HasNet48()) return File.Exists(full) ? full : (File.Exists(lite) ? lite : full);
            // 无 4.8：优先降级壳；降级壳也不在则回退 4.x（至少让 CLR 给出明确报错）
            return File.Exists(lite) ? lite : full;
        }

        // 启动主程序（降级壳感知）。返回 null=成功，否则错误信息。
        public static string LaunchHost(string root, out bool usedLite)
        {
            usedLite = false;
            string exe = SelectHostExe(root);
            if (!File.Exists(exe)) return "未找到主程序: " + exe;
            usedLite = exe.EndsWith("OfficeAgent35.exe", StringComparison.OrdinalIgnoreCase);
            try
            {
                System.Diagnostics.Process.Start(exe);
                return null;
            }
            catch (Exception ex) { return "启动失败: " + ex.Message; }
        }

        // 提权重启：启动 requireAdministrator 变体。返回 null=已启动，否则错误/缺失说明。
        public static string RestartElevated(string root)
        {
            string dir = Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
            string adminExe = Path.Combine(dir, "OfficeAgentBootAdmin.exe");
            if (!File.Exists(adminExe)) return "缺少提权引导器 OfficeAgentBootAdmin.exe";
            try
            {
                // 必须带 /ui：新版引导器有"缺失项只提醒一次"闸门（core\src\BootGate.cs）。
                // 不带 /ui 时，若缺失集合与上次相同，引导器会判定"无需提示"→ 直接启动主程序
                // 然后 return 0 退出，**窗口根本不会出现**；用户点了"安装缺失组件"、过了 UAC，
                // 却什么都没看到，还会多开一个 host 实例。
                // 这与 host 侧 RepairLauncher 是同一处缺陷（那边已修），此处是等价路径，必须一并修。
                // /adminfix 让窗口打开后自动开始补全（本进程已是 requireAdministrator 变体）。
                System.Diagnostics.Process.Start(adminExe, "/ui /adminfix /root:\"" + root + "\"");
                return null;
            }
            catch (Exception ex) { return "提权重启失败（可能被拒绝）: " + ex.Message; }
        }
    }
}
