// MergeTemplate —— 报表汇总模板（设计方案 §6.2）：子公司/银行格式配置一次永久复用
// 目标列 ↔ 源列名映射（各文件列序可不同，按表头名匹配）；金额列参与合并合计勾稽。
// 存 %LOCALAPPDATA%\OfficeAgent\merge-templates.json
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace OfficeAgent.Host
{
    public class MergeTemplate
    {
        public string Name = "";
        public string SheetSpec = "0";       // 各文件统一 sheet 序号(1起)或名称
        public int HeaderRow = 1;
        public string[] TargetCols = new string[0];   // 底稿列名
        public string[] SrcCols = new string[0];      // 对应源列名/序号
        public string[] AmountCols = new string[0];   // 数值列（参与合计勾稽），用目标列名

        public static MergeTemplate FromForm(string name, string sheetSpec, int headerRow,
            string[] targets, string[] srcs, string[] amounts)
        {
            MergeTemplate t = new MergeTemplate();
            t.Name = name;
            t.SheetSpec = sheetSpec;
            t.HeaderRow = headerRow;
            t.TargetCols = targets;
            t.SrcCols = srcs;
            t.AmountCols = amounts;
            return t;
        }
    }

    public static class MergeTemplateStore
    {
        static string FilePath()
        {
            return Path.Combine(Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeAgent"),
                "merge-templates.json");
        }

        public static List<MergeTemplate> LoadAll()
        {
            List<MergeTemplate> list = new List<MergeTemplate>();
            try
            {
                if (!File.Exists(FilePath())) return list;
                string json = File.ReadAllText(FilePath(), Encoding.UTF8);
                foreach (Dictionary<string, string> o in MiniJson.ParseObjects(json))
                {
                    MergeTemplate t = new MergeTemplate();
                    t.Name = MiniJson.Get(o, "name");
                    if (t.Name.Length == 0) continue;
                    t.SheetSpec = MiniJson.GetOr(o, "sheetSpec", "0");
                    t.HeaderRow = MiniJson.GetInt(o, "headerRow", 1);
                    t.TargetCols = Split(MiniJson.Get(o, "targetCols"));
                    t.SrcCols = Split(MiniJson.Get(o, "srcCols"));
                    t.AmountCols = Split(MiniJson.Get(o, "amountCols"));
                    list.Add(t);
                }
            }
            catch { }
            return list;
        }

        public static MergeTemplate Get(string name)
        {
            foreach (MergeTemplate t in LoadAll())
            {
                if (string.Equals(t.Name, name, StringComparison.OrdinalIgnoreCase)) return t;
            }
            return null;
        }

        public static void Upsert(MergeTemplate t)
        {
            List<MergeTemplate> list = LoadAll();
            for (int i = 0; i < list.Count; i++)
            {
                if (string.Equals(list[i].Name, t.Name, StringComparison.OrdinalIgnoreCase)) { list.RemoveAt(i); break; }
            }
            list.Add(t);
            SaveAll(list);
        }

        public static bool Delete(string name)
        {
            List<MergeTemplate> list = LoadAll();
            for (int i = 0; i < list.Count; i++)
            {
                if (string.Equals(list[i].Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    list.RemoveAt(i);
                    SaveAll(list);
                    return true;
                }
            }
            return false;
        }

        static void SaveAll(List<MergeTemplate> list)
        {
            string dir = Path.GetDirectoryName(FilePath());
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
            StringBuilder sb = new StringBuilder();
            sb.Append("[\n");
            for (int i = 0; i < list.Count; i++)
            {
                MergeTemplate t = list[i];
                sb.Append("  {\"name\":\"").Append(MiniJson.Esc(t.Name)).Append("\",");
                sb.Append("\"sheetSpec\":\"").Append(MiniJson.Esc(t.SheetSpec)).Append("\",");
                sb.Append("\"headerRow\":").Append(t.HeaderRow).Append(",");
                sb.Append("\"targetCols\":\"").Append(MiniJson.Esc(Join(t.TargetCols))).Append("\",");
                sb.Append("\"srcCols\":\"").Append(MiniJson.Esc(Join(t.SrcCols))).Append("\",");
                sb.Append("\"amountCols\":\"").Append(MiniJson.Esc(Join(t.AmountCols))).Append("\"}");
                if (i < list.Count - 1) sb.Append(",");
                sb.Append("\n");
            }
            sb.Append("]\n");
            File.WriteAllText(FilePath(), sb.ToString(), new UTF8Encoding(false));
        }

        static string[] Split(string s)
        {
            List<string> parts = new List<string>();
            if (s != null)
            {
                foreach (string p in s.Split('\u0001'))
                {
                    string t = p.Trim();
                    if (t.Length > 0) parts.Add(t);
                }
            }
            return parts.ToArray();
        }

        static string Join(string[] arr)
        {
            StringBuilder sb = new StringBuilder();
            foreach (string s in arr) { if (sb.Length > 0) sb.Append('\u0001'); sb.Append(s); }
            return sb.ToString();
        }
    }
}
