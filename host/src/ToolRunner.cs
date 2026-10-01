// 工具进程白名单入口（host 侧，形态沿用 WorkBuddyRepair 已通过审查的 ProcRunner 模式）
// 规则：
//   * 只有本文件 switch 白名单里的工具可被启动；
//   * exePath 必须是本机探测到的固定安装位置（ConvertEngine.FindSoffice）；
//   * 启动参数由引擎按固定模板生成，文件路径仅做引号转义；
//   * UseShellExecute=false 走 CreateProcess 直调，不经任何 shell 解释层。
//   * 进程入 Job Object（JobObject.cs）：内存上限 + 超时整树终止 + 句柄关闭兜底清理。
using System;
using System.Diagnostics;
using OfficeAgent.Core;

namespace OfficeAgent.Host
{
    public static class ToolRunner
    {
        // LibreOffice headless 转换。返回进程退出码（ExitRejected=-100 / ExitTimeout=-101）。
        public static int ConvertViaSoffice(string sofficePath, string convertArgs, int timeoutSeconds)
        {
            if (sofficePath == null || !sofficePath.EndsWith("soffice.exe", StringComparison.OrdinalIgnoreCase))
            {
                LogProxy.Write("soffice 白名单校验失败（路径非 soffice.exe）");
                return ProcRunner.ExitRejected;
            }
            Process p;
            try
            {
                p = Process.Start(sofficePath, convertArgs);
            }
            catch (Exception ex)
            {
                LogProxy.Write("soffice 启动失败：{0}", ex.Message);
                return ProcRunner.ExitRejected;
            }
            using (p)
            using (ToolJobGuard job = ToolJobGuard.Attach(p))
            {
                DateTime deadline = DateTime.Now.AddSeconds(timeoutSeconds);
                while (!p.WaitForExit(500))
                {
                    if (DateTime.Now > deadline)
                    {
                        job.KillOnTimeout();
                        LogProxy.Write("soffice 转换超时（" + timeoutSeconds.ToString() + " 秒），已" +
                            (job.HasJob ? "整树终止" : "终止进程"));
                        return ProcRunner.ExitTimeout;
                    }
                }
                return p.ExitCode;
            }
        }
    }
}
