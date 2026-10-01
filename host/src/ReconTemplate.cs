// ReconTemplate —— 核对映射模板记忆（设计方案 §6.6：手工配置一次，"上月对账再来一次"直接复用）
// 存 %LOCALAPPDATA%\OfficeAgent\recon-templates.json；列引用一律存「列名/序号字符串」，
// 复用时按新文件表头重新解析 → 同模板文件换月可用。
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace OfficeAgent.Host
{
    public class ReconTemplate
    {
        public string Name = "";
        public string FileA = "";        // 最近一次使用的文件（可空）
        public string FileB = "";
        public string SheetA = "0";
        public string SheetB = "0";
        public int HeaderRowA = 1;
        public int HeaderRowB = 1;
        public string KeyColsA = "";     // 逗号分隔（列名或 1-based 序号）
        public string KeyColsB = "";
        public string DebitA = "";
        public string CreditA = "";
        public string DebitB = "";
        public string CreditB = "";
        public double Tolerance = 0.01;
        public bool SqueezeKey = true;
        public bool FillMergedCells = true;
        public bool IncludeMatched = false;

        public ReconMapping ToMapping()
        {
            ReconMapping m = new ReconMapping();
            m.SheetA = SheetA; m.SheetB = SheetB;
            m.HeaderRowA = HeaderRowA; m.HeaderRowB = HeaderRowB;
            m.KeyColsA = SplitCols(KeyColsA);
            m.KeyColsB = SplitCols(KeyColsB);
            m.DebitA = DebitA ?? ""; m.CreditA = CreditA ?? "";
            m.DebitB = DebitB ?? ""; m.CreditB = CreditB ?? "";
            m.Tolerance = Tolerance;
            m.SqueezeKey = SqueezeKey;
            m.FillMergedCells = FillMergedCells;
            m.IncludeMatched = IncludeMatched;
            return m;
        }

        static string[] SplitCols(string s)
        {
            List<string> parts = new List<string>();
            if (s != null)
            {
                foreach (string p in s.Split(','))
                {
                    string t = p.Trim();
                    if (t.Length > 0) parts.Add(t);
                }
            }
            return parts.ToArray();
        }

        public static ReconTemplate FromMapping(string name, ReconMapping m, string fileA, string fileB)
        {
            ReconTemplate t = new ReconTemplate();
            t.Name = name;
            t.FileA = fileA ?? ""; t.FileB = fileB ?? "";
            t.SheetA = m.SheetA; t.SheetB = m.SheetB;
            t.HeaderRowA = m.HeaderRowA; t.HeaderRowB = m.HeaderRowB;
            t.KeyColsA = JoinCols(m.KeyColsA); t.KeyColsB = JoinCols(m.KeyColsB);
            t.DebitA = m.DebitA; t.CreditA = m.CreditA;
            t.DebitB = m.DebitB; t.CreditB = m.CreditB;
            t.Tolerance = m.Tolerance;
            t.SqueezeKey = m.SqueezeKey;
            t.FillMergedCells = m.FillMergedCells;
            t.IncludeMatched = m.IncludeMatched;
            return t;
        }

        static string JoinCols(string[] cols)
        {
            StringBuilder sb = new StringBuilder();
            foreach (string c in cols) { if (sb.Length > 0) sb.Append(","); sb.Append(c); }
            return sb.ToString();
        }
    }

    public static class ReconTemplateStore
    {
        public const string LastName = "__last";   // 自动覆盖的"上次使用"

        static string FilePath()
        {
            return Path.Combine(Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeAgent"),
                "recon-templates.json");
        }

        public static List<ReconTemplate> LoadAll()
        {
            List<ReconTemplate> list = new List<ReconTemplate>();
            try
            {
                if (!File.Exists(FilePath())) return list;
                string json = File.ReadAllText(FilePath(), Encoding.UTF8);
                List<Dictionary<string, string>> objs = ParseObjects(json);
                foreach (Dictionary<string, string> o in objs)
                {
                    ReconTemplate t = new ReconTemplate();
                    t.Name = Get(o, "name");
                    if (t.Name.Length == 0) continue;
                    t.FileA = Get(o, "fileA"); t.FileB = Get(o, "fileB");
                    t.SheetA = GetOr(o, "sheetA", "0"); t.SheetB = GetOr(o, "sheetB", "0");
                    t.HeaderRowA = GetInt(o, "headerRowA", 1);
                    t.HeaderRowB = GetInt(o, "headerRowB", 1);
                    t.KeyColsA = Get(o, "keyColsA"); t.KeyColsB = Get(o, "keyColsB");
                    t.DebitA = Get(o, "debitA"); t.CreditA = Get(o, "creditA");
                    t.DebitB = Get(o, "debitB"); t.CreditB = Get(o, "creditB");
                    t.Tolerance = GetDouble(o, "tolerance", 0.01);
                    t.SqueezeKey = GetBool(o, "squeezeKey", true);
                    t.FillMergedCells = GetBool(o, "fillMerged", true);
                    t.IncludeMatched = GetBool(o, "includeMatched", false);
                    list.Add(t);
                }
            }
            catch { }
            return list;
        }

        public static ReconTemplate Get(string name)
        {
            foreach (ReconTemplate t in LoadAll())
            {
                if (string.Equals(t.Name, name, StringComparison.OrdinalIgnoreCase)) return t;
            }
            return null;
        }

        public static void Upsert(ReconTemplate t)
        {
            List<ReconTemplate> list = LoadAll();
            for (int i = 0; i < list.Count; i++)
            {
                if (string.Equals(list[i].Name, t.Name, StringComparison.OrdinalIgnoreCase)) { list.RemoveAt(i); break; }
            }
            list.Add(t);
            SaveAll(list);
        }

        public static bool Delete(string name)
        {
            List<ReconTemplate> list = LoadAll();
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

        static void SaveAll(List<ReconTemplate> list)
        {
            string dir = Path.GetDirectoryName(FilePath());
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
            StringBuilder sb = new StringBuilder();
            sb.Append("[\n");
            for (int i = 0; i < list.Count; i++)
            {
                ReconTemplate t = list[i];
                sb.Append("  {\"name\":\"").Append(Esc(t.Name)).Append("\",");
                sb.Append("\"fileA\":\"").Append(Esc(t.FileA)).Append("\",\"fileB\":\"").Append(Esc(t.FileB)).Append("\",");
                sb.Append("\"sheetA\":\"").Append(Esc(t.SheetA)).Append("\",\"sheetB\":\"").Append(Esc(t.SheetB)).Append("\",");
                sb.Append("\"headerRowA\":").Append(t.HeaderRowA).Append(",\"headerRowB\":").Append(t.HeaderRowB).Append(",");
                sb.Append("\"keyColsA\":\"").Append(Esc(t.KeyColsA)).Append("\",\"keyColsB\":\"").Append(Esc(t.KeyColsB)).Append("\",");
                sb.Append("\"debitA\":\"").Append(Esc(t.DebitA)).Append("\",\"creditA\":\"").Append(Esc(t.CreditA)).Append("\",");
                sb.Append("\"debitB\":\"").Append(Esc(t.DebitB)).Append("\",\"creditB\":\"").Append(Esc(t.CreditB)).Append("\",");
                sb.Append("\"tolerance\":").Append(t.Tolerance.ToString("0.########", CultureInfo.InvariantCulture)).Append(",");
                sb.Append("\"squeezeKey\":").Append(t.SqueezeKey ? "true" : "false").Append(",");
                sb.Append("\"fillMerged\":").Append(t.FillMergedCells ? "true" : "false").Append(",");
                sb.Append("\"includeMatched\":").Append(t.IncludeMatched ? "true" : "false").Append("}");
                if (i < list.Count - 1) sb.Append(",");
                sb.Append("\n");
            }
            sb.Append("]\n");
            File.WriteAllText(FilePath(), sb.ToString(), new UTF8Encoding(false));
        }

        static string Esc(string s)
        {
            if (s == null) return "";
            StringBuilder sb = new StringBuilder();
            foreach (char ch in s)
            {
                switch (ch)
                {
                    case '\\': sb.Append("\\\\"); break;
                    case '"': sb.Append("\\\""); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default: sb.Append(ch); break;
                }
            }
            return sb.ToString();
        }

        // ---------- 极简 JSON 对象数组解析（扁平 schema；字符串值支持转义） ----------

        static List<Dictionary<string, string>> ParseObjects(string json)
        {
            List<Dictionary<string, string>> result = new List<Dictionary<string, string>>();
            if (json == null) return result;
            int i = 0, n = json.Length;
            while (i < n)
            {
                if (json[i] != '{') { i++; continue; }
                Dictionary<string, string> obj = new Dictionary<string, string>();
                i++;
                while (i < n && json[i] != '}')
                {
                    if (json[i] != '"') { i++; continue; }
                    string key = ReadString(json, ref i);
                    while (i < n && json[i] != ':') i++;
                    i++;    // 跳过 ':'
                    while (i < n && (json[i] == ' ' || json[i] == '\t')) i++;
                    string val;
                    if (i < n && json[i] == '"') val = ReadString(json, ref i);
                    else
                    {
                        int end = i;
                        while (end < n && json[end] != ',' && json[end] != '}') end++;
                        val = json.Substring(i, end - i).Trim();
                        i = end;
                    }
                    if (key.Length > 0 && !obj.ContainsKey(key)) obj[key] = val;
                    while (i < n && json[i] != ',' && json[i] != '}') i++;
                    if (i < n && json[i] == ',') i++;
                }
                if (i < n) i++;    // 跳过 '}'
                if (obj.Count > 0) result.Add(obj);
            }
            return result;
        }

        static string ReadString(string json, ref int i)
        {
            StringBuilder sb = new StringBuilder();
            i++;    // 跳过开头引号
            while (i < json.Length && json[i] != '"')
            {
                if (json[i] == '\\' && i + 1 < json.Length)
                {
                    i++;
                    char c = json[i];
                    if (c == 'n') sb.Append('\n');
                    else if (c == 'r') sb.Append('\r');
                    else if (c == 't') sb.Append('\t');
                    else if (c == 'u' && i + 4 < json.Length)
                    {
                        int code;
                        if (int.TryParse(json.Substring(i + 1, 4), NumberStyles.HexNumber,
                            CultureInfo.InvariantCulture, out code))
                            sb.Append((char)code);
                        i += 4;
                    }
                    else sb.Append(c);
                }
                else sb.Append(json[i]);
                i++;
            }
            i++;    // 跳过结尾引号
            return sb.ToString();
        }

        static string Get(Dictionary<string, string> o, string key)
        {
            string v;
            return o.TryGetValue(key, out v) ? v : "";
        }

        static string GetOr(Dictionary<string, string> o, string key, string def)
        {
            string v;
            return o.TryGetValue(key, out v) && v.Length > 0 ? v : def;
        }

        static int GetInt(Dictionary<string, string> o, string key, int def)
        {
            string v = Get(o, key);
            int n;
            return int.TryParse(v, out n) ? n : def;
        }

        static double GetDouble(Dictionary<string, string> o, string key, double def)
        {
            string v = Get(o, key);
            double d;
            return double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out d) ? d : def;
        }

        static bool GetBool(Dictionary<string, string> o, string key, bool def)
        {
            string v = Get(o, key);
            return v.Length == 0 ? def : v == "true";
        }
    }
}
