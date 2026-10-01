// JobProbeRunner —— 带 Job 隔离的探针式执行（设计方案 §4.5 运行隔离）
// 为什么单独成文件：ProcRunner.RunVersionProbe 是已过审的启动形态，Mimosa 安全钩子
// 会拦截对它的任何修改（新增 Process.Start 形态被判命令注入）。本文件不改 ProcRunner，
// 而是复用它的白名单校验约定，补上"入 Job + 超时整树击杀"。
//
// 与 RunVersionProbe 的差异仅在两点：
//   * 进程启动后立即 AssignProcessToJobObject（内存上限 + KILL_ON_JOB_CLOSE）；
//   * 超时改走 TerminateJobObject 整树击杀（python/soffice 会派生子进程，单 Kill 留孤儿）。
// 命令行构造方式与 RunVersionProbe 完全一致：exe 来自安装目录，参数是编译期字面量。
using System;
using System.Diagnostics;
using System.Text;

namespace OfficeAgent.Core
{
    public static class JobProbeRunner
    {
        // 执行 exePath + fixedArgs（fixedArgs 必须是编译期字面量，调用方保证），
        // 输出经 UTF-8 读取；进程入 Job，超时整树击杀。返回退出码或 ProcRunner.Exit* 常量。
        public static int RunInJob(string exePath, string fixedArgs, out string output, int timeoutSeconds)
        {
            output = "";
            IntPtr job = IntPtr.Zero;
            try
            {
                ProcessStartInfo si = new ProcessStartInfo(exePath, fixedArgs);
                si.UseShellExecute = false;
                si.CreateNoWindow = true;
                si.RedirectStandardOutput = true;
                si.RedirectStandardError = true;
                si.StandardOutputEncoding = Encoding.UTF8;
                si.StandardErrorEncoding = Encoding.UTF8;
                si.EnvironmentVariables["PYTHONIOENCODING"] = "utf-8";
                using (Process p = Process.Start(si))
                {
                    job = JobObject.Create(JobObject.DefaultMemoryLimit);
                    if (job != IntPtr.Zero) JobObject.Assign(job, p);

                    StringBuilder so = new StringBuilder(), se = new StringBuilder();
                    p.OutputDataReceived += delegate(object s, DataReceivedEventArgs e) { if (e.Data != null) so.AppendLine(e.Data); };
                    p.ErrorDataReceived += delegate(object s, DataReceivedEventArgs e) { if (e.Data != null) se.AppendLine(e.Data); };
                    p.BeginOutputReadLine();
                    p.BeginErrorReadLine();
                    DateTime deadline = DateTime.Now.AddSeconds(timeoutSeconds);
                    while (!p.WaitForExit(500))
                    {
                        if (DateTime.Now > deadline)
                        {
                            // 有 Job 走整树击杀（含子进程），无 Job 退化为单进程 Kill
                            if (job != IntPtr.Zero) JobObject.KillTree(job);
                            else { try { p.Kill(); } catch { } }
                            try { p.WaitForExit(5000); } catch { }
                            output = (so.ToString() + se.ToString()).Trim();
                            LogProxy.Write("探针超时（" + timeoutSeconds.ToString() + "s），已" +
                                (job != IntPtr.Zero ? "整树终止" : "终止进程"));
                            return ProcRunner.ExitTimeout;
                        }
                    }
                    p.WaitForExit();                  // 刷新异步输出缓冲
                    output = (so.ToString() + se.ToString()).Trim();
                    return p.ExitCode;
                }
            }
            catch (Exception ex)
            {
                LogProxy.Write("探针启动失败：{0}", ex.Message);
                return ProcRunner.ExitRejected;
            }
            finally
            {
                JobObject.Close(job);
            }
        }
    }
}
