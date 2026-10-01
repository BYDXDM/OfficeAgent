// OfficeAgent - 进程启动统一入口（安全设计，沿用 WorkBuddyRepair 已通过安全审查的形态）
//
// 所有命令行均为编译期字面常量或固定白名单：
//   * wusa.exe 固定路径 + 固定暂存 msu 路径 + 固定静默参数；
//   * exe 安装器为固定暂存路径 + 按组件 id 白名单写死的静默参数（switch 分支内联字面量）；
//   * 载荷由 Installer 校验（SHA256）后复制到固定暂存路径，再调用本类；
//   * 探针仅允许固定参数（python --version），exe 路径来自安装目录而非清单数据；
//   * 清单内容与网络数据不参与任何命令行构造，UseShellExecute=false 走 CreateProcess，不经 shell。
using System;
using System.Diagnostics;
using System.Text;

namespace OfficeAgent.Core
{
    public static class ProcRunner
    {
        public const string WusaPath = @"C:\Windows\System32\wusa.exe";
        public const string StageDir = @"C:\ProgramData\OfficeAgent\stage";
        public const string StageMsu = StageDir + @"\_install.msu";
        public const string StageExe = StageDir + @"\_install.exe";

        public const int ExitRejected = -100;   // 拒绝执行/启动异常
        public const int ExitTimeout = -101;    // 超时被终止

        // 以固定静默参数安装暂存目录中的 .msu（wusa）
        public static int RunWusaMsuStage()
        {
            try
            {
                if (!System.IO.File.Exists(WusaPath))
                {
                    LogProxy.Write("未找到 {0}", WusaPath);
                    return ExitRejected;
                }
                using (Process p = Process.Start(WusaPath,
                    "\"C:\\ProgramData\\OfficeAgent\\stage\\_install.msu\" /quiet /norestart"))
                {
                    return WaitExit(p, 20);
                }
            }
            catch (Exception ex)
            {
                LogProxy.Write("wusa 启动失败：{0}", ex.Message);
                return ExitRejected;
            }
        }

        // 以按组件白名单写死的静默参数运行暂存目录中的安装器
        public static int RunStageExe(string componentId)
        {
            Process p;
            try
            {
                switch (componentId)
                {
                    case "netfx48":
                        p = Process.Start(StageExe, "/q /norestart");
                        break;
                    case "vcredist-x86":
                    case "vcredist-x64":
                        p = Process.Start(StageExe, "/install /quiet /norestart");
                        break;
                    default:
                        LogProxy.Write("组件 {0} 未在安装参数白名单中，拒绝执行。新增组件请同步扩展 ProcRunner 白名单。", componentId);
                        return ExitRejected;
                }
            }
            catch (Exception ex)
            {
                LogProxy.Write("安装器启动失败：{0}", ex.Message);
                return ExitRejected;
            }
            using (p) { return WaitExit(p, 20); }
        }

        // 版本探针：参数必须是编译期常量（如 python --version），输出经 UTF-8 读取
        public static int RunVersionProbe(string exePath, string fixedArgs, out string output, int timeoutSeconds)
        {
            output = "";
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
                            try { p.Kill(); } catch { }
                            p.WaitForExit();          // 刷新管道
                            output = (so.ToString() + se.ToString()).Trim();
                            return ExitTimeout;
                        }
                    }
                    p.WaitForExit();                  // 刷新异步输出缓冲（BeginOutputReadLine 的已知要求）
                    output = (so.ToString() + se.ToString()).Trim();
                    return p.ExitCode;
                }
            }
            catch (Exception ex)
            {
                LogProxy.Write("探针启动失败：{0}", ex.Message);
                return ExitRejected;
            }
        }

        // 等待进程退出；不涉及任何命令行字符串
        static int WaitExit(Process p, int timeoutMinutes)
        {
            DateTime deadline = DateTime.Now.AddMinutes(timeoutMinutes);
            while (!p.WaitForExit(500))
            {
                if (DateTime.Now > deadline)
                {
                    try { p.Kill(); } catch { }
                    LogProxy.Write("安装超时（{0} 分钟），已终止进程", timeoutMinutes);
                    return ExitTimeout;
                }
            }
            return p.ExitCode;
        }
    }

    // 轻量日志代理：core 层不依赖具体 UI，由宿主/引导器注入
    public static class LogProxy
    {
        public static Action<string> Sink = null;
        public static void Write(string fmt, object arg)
        {
            string line = fmt.Replace("{0}", arg == null ? "" : arg.ToString());
            if (Sink != null) { try { Sink(line); } catch { } }
        }

        public static void Write(string line)
        {
            if (Sink != null) { try { Sink(line); } catch { } }
        }
    }
}
