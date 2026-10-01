// RepairLauncher —— 宿主侧"启动环境修复器"的独立出口（boot 侧 BootLaunch 同形态）
// 独立成文件的原因：Mimosa 安全钩子对启动点敏感，全项目只保留这一处提权启动。
// 参数为本地已知路径的固定模板，不使用 ShellExecute/Verb（铁律 4）。
using System;
using System.IO;

namespace OfficeAgent.Host
{
    public static class RepairLauncher
    {
        // 启动 requireAdministrator 变体引导器（UAC 由系统弹出）。返回 null=已启动，否则错误说明。
        public static string Start(string root)
        {
            string adminExe = Path.Combine(Path.Combine(root, "boot"), "OfficeAgentBootAdmin.exe");
            if (!File.Exists(adminExe)) return "未找到修复器: " + adminExe;
            try
            {
                string args = "/root:\"" + root + "\"";
                System.Diagnostics.Process.Start(adminExe, args);
                return null;
            }
            catch (Exception ex)
            {
                return "启动修复器失败（可能 UAC 被拒绝）: " + ex.Message;
            }
        }
    }
}
