// OfficeAgent 引导器入口（.NET 3.5 / Win7 RTM 自带 3.5.1，零前置）
// 参数：/report-only  仅检测并生成报告后退出（exit 0=全绿, 2=有缺失）
//       /fix          检测 + 自动补全免管理员项 + 复检 + 报告后退出
//       /adminfix     提权会话内补全（含需管理员项）
//       /console      尝试附加控制台输出
//       /root:<path>  显式指定安装根目录
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using OfficeAgent.Core;

namespace OfficeAgent.Boot
{
    internal static class Program
    {
        public static string Root = "";
        public static bool AdminFix = false;

        [DllImport("kernel32.dll")]
        static extern bool AllocConsole();
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool AttachConsole(int processId);
        [DllImport("kernel32.dll")]
        static extern uint SetErrorMode(uint uMode);
        const uint SEM_FAILCRITICALERRORS = 0x0001;
        const int ATTACH_PARENT_PROCESS = -1;

        [STAThread]
        static int Main(string[] args)
        {
            // 裸机上 python/LO 探针可能在 VC++ 安装前执行（缺失 api-ms-win-crt-*.dll），
            // Windows 默认会弹模态加载错误框阻塞检测线程；安装器类进程统一抑制该对话框，
            // 让探针直接收到加载失败退出码（126），由既有"未知/缺失"分支处理。
            try { SetErrorMode(SEM_FAILCRITICALERRORS); } catch { }

            bool reportOnly = false, console = false, fixMode = false, forceUi = false;
            foreach (string raw in args)
            {
                string a = raw.ToLowerInvariant();
                if (a == "/report-only") reportOnly = true;
                else if (a == "/console") console = true;
                else if (a == "/adminfix") AdminFix = true;
                else if (a == "/fix") fixMode = true;
                else if (a == "/ui") forceUi = true;   // 强制显示自检窗口（排查问题时用）
                else if (a.StartsWith("/root:")) Root = raw.Substring(6).Trim('"');
            }
            if (Root == null || Root.Length == 0) Root = EnvDetect.FindRoot();

            if (console)
            {
                try { AttachConsole(ATTACH_PARENT_PROCESS); } catch { }
                try { AllocConsole(); } catch { }
            }
            try { Console.OutputEncoding = Encoding.UTF8; } catch { }

            Log.Init();
            LogProxy.Sink = new Action<string>(Log.Line);
            Log.Line("OfficeAgent 引导器 v0.1.0 (M0)  root=" + Root);

            List<DetectItem> items = EnvDetect.DetectAll(Root);

            if (fixMode)
            {
                foreach (DetectItem it in items)
                {
                    if (it.FixKey == null || it.State == DetectState.Ok) continue;
                    if (it.NeedAdmin && !EnvDetect.IsAdmin()) { Log.Line("跳过(需管理员): " + it.Name); continue; }
                    string err = Installer.DoFix(it, Root);
                    Log.Line(err == null ? "完成: " + it.Name : "失败: " + it.Name + " — " + err);
                }
                items = EnvDetect.DetectAll(Root);
                return FinishReport(items, reportOnly);
            }

            if (reportOnly)
            {
                return FinishReport(items, true);
            }

            // 组件齐全 → 直接启动主程序，不再弹自检窗口（用户要求：第一次正常启动后跳过自检）。
            // 有缺失项或启动失败才回到自检界面；/ui 可强制显示。
            if (!forceUi)
            {
                bool missingAny = false;
                foreach (DetectItem it in items) { if (it.State == DetectState.Missing) { missingAny = true; break; } }
                if (!missingAny)
                {
                    Log.Line("组件齐全，直接启动主程序（加 /ui 可显示自检窗口）");
                    bool usedLite;
                    string launchErr = BootLaunch.LaunchHost(Root, out usedLite);
                    if (launchErr == null) return 0;
                    Log.Line(launchErr + "，转入自检界面");
                }
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new BootForm(items));
            return 0;
        }

        static int FinishReport(List<DetectItem> items, bool consoleOut)
        {
            string report = Report.Build(items, Root);
            string path = Report.Save(report);
            if (consoleOut) Log.Line(report);
            Log.Line("报告: " + path);
            foreach (DetectItem it in items)
            {
                if (it.State == DetectState.Missing) return 2;
            }
            return 0;
        }
    }

    // 引导日志：文件（UTF-8）+ 事件（UI 订阅）+ 控制台尽力输出
    public static class Log
    {
        static object gate = new object();
        static string file = null;
        public static event Action<string> Emitted;

        public static void Init()
        {
            try
            {
                string dir = Path.Combine(Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeAgent"), "logs");
                Directory.CreateDirectory(dir);
                file = Path.Combine(dir, "bootstrap-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".txt");
                File.AppendAllText(file, "=== OfficeAgent 引导日志 " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " ===" + Environment.NewLine, Encoding.UTF8);
            }
            catch { file = null; }
        }

        public static void Line(string s)
        {
            lock (gate)
            {
                string stamped = DateTime.Now.ToString("HH:mm:ss") + "  " + s;
                if (Emitted != null) { try { Emitted(stamped); } catch { } }
                try { Console.WriteLine(stamped); } catch { }
                if (file != null) { try { File.AppendAllText(file, stamped + Environment.NewLine, Encoding.UTF8); } catch { } }
            }
        }
    }
}
