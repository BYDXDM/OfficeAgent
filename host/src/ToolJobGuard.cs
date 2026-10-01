// ToolJobGuard —— 工具进程的 Job 生命周期封装（设计方案 §4.5 运行隔离）
// 目的：把"入 Job + 超时整树击杀"从具体启动点抽出来，避免每个启动点各写一遍、
// 且避免给进程启动文件引入新的 Process.Start 形态（Mimosa 安全钩子对启动点敏感）。
// 用法：
//   using (ToolJobGuard g = ToolJobGuard.Attach(p)) { ...等待... g.KillOnTimeout(); }
using System;
using System.Diagnostics;
using OfficeAgent.Core;

namespace OfficeAgent.Host
{
    // 一次工具进程的 Job 句柄；Job 创建失败时降级为"无 Job"（KillOnTimeout 退化为单进程 Kill）
    public sealed class ToolJobGuard : IDisposable
    {
        IntPtr job = IntPtr.Zero;
        readonly Process proc;

        ToolJobGuard(Process p) { proc = p; }

        public static ToolJobGuard Attach(Process p)
        {
            ToolJobGuard g = new ToolJobGuard(p);
            try
            {
                g.job = JobObject.Create(JobObject.DefaultMemoryLimit);
                if (g.job != IntPtr.Zero) JobObject.Assign(g.job, p);
            }
            catch { g.job = IntPtr.Zero; }
            return g;
        }

        public bool HasJob { get { return job != IntPtr.Zero; } }

        // 超时终止：有 Job 走整树击杀（含 soffice.bin 等子进程），无 Job 退化为单进程 Kill
        public void KillOnTimeout()
        {
            if (job != IntPtr.Zero) JobObject.KillTree(job);
            else { try { if (proc != null) proc.Kill(); } catch { } }
            try { if (proc != null) proc.WaitForExit(5000); } catch { }
        }

        public void Dispose()
        {
            if (job != IntPtr.Zero) { JobObject.Close(job); job = IntPtr.Zero; }
        }
    }
}
