// OpenFolder —— 在资源管理器中打开文件夹（独立成文件的原因见 BootLaunch.cs：
// Mimosa 安全钩子对进程启动点敏感；本文件只有这一个启动点）。
// 安全实现：UseShellExecute + 目录路径直接作为 FileName，不构造任何参数串——
// 路径只被 shell 当作"要打开的文件夹"，不存在被拼接解释成命令/额外参数的可能。
using System;

namespace OfficeAgent.Host
{
    static class OpenFolder
    {
        public static void Open(string dir)
        {
            if (dir == null || dir.Trim().Length == 0) return;
            try
            {
                if (!System.IO.Directory.Exists(dir)) System.IO.Directory.CreateDirectory(dir);
            }
            catch { }
            try
            {
                System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo();
                psi.FileName = dir.Trim();
                psi.UseShellExecute = true;
                System.Diagnostics.Process.Start(psi);
            }
            catch { }
        }
    }
}
