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
                // 必须带 /ui：引导器新增了"缺失项只提醒一次"闸门（core\src\BootGate.cs），
                // 不带 /ui 时若缺失集合与上次相同，引导器会判定"无需提示"→ 直接尝试启动主程序
                // 然后退出，用户看到的现象是"弹了 UAC 但什么都没出现"，还可能多开一个 host 实例。
                // /ui 强制显示自检窗口，这正是"用户显式要求修复"应得的语义。
                // /adminfix 让窗口打开后自动开始补全（本进程已是提权变体）。
                string args = "/ui /adminfix /root:\"" + root + "\"";
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
