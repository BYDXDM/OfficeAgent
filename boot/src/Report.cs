// 检测报告生成与落盘
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeAgent.Core;

namespace OfficeAgent.Boot
{
    public static class Report
    {
        public static string Build(List<DetectItem> items, string root)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("OfficeAgent 环境检测报告");
            sb.AppendLine("时间: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            sb.AppendLine("root: " + root);
            sb.AppendLine("系统: " + Environment.OSVersion.VersionString + (EnvDetect.Is64OS() ? " (x64)" : " (x86)"));
            sb.AppendLine("管理员: " + (EnvDetect.IsAdmin() ? "是" : "否"));
            sb.AppendLine(new string('-', 72));
            foreach (DetectItem it in items)
            {
                sb.AppendLine(string.Format("[{0}] {1}{2}",
                    EnvDetect.StateText(it.State), it.Name, it.NeedAdmin ? "  (需管理员)" : ""));
                sb.AppendLine("      " + it.Detail);
            }
            sb.AppendLine(new string('-', 72));
            int missing = 0;
            foreach (DetectItem it in items) { if (it.State == DetectState.Missing) missing++; }
            sb.AppendLine(missing == 0 ? "结论: P0 组件齐备" : ("结论: " + missing + " 项缺失"));
            return sb.ToString();
        }

        public static string Save(string content)
        {
            string dir = Path.Combine(Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeAgent"), "logs");
            string path = null;
            try
            {
                Directory.CreateDirectory(dir);
                path = Path.Combine(dir, "report-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".txt");
                File.WriteAllText(path, content, Encoding.UTF8);
                File.WriteAllText(Path.Combine(dir, "last-report.txt"), content, Encoding.UTF8);
            }
            catch { }
            return path ?? "(报告写入失败)";
        }
    }
}
