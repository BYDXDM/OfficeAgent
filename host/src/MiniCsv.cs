// MiniCsv —— CSV 读写（编码探测 UTF-8/GB18030、RFC4180 引号规则）
// 会计现实：银行/业务系统导出的 CSV 常见 GBK 与 UTF-8 混杂，写出默认 UTF-8 BOM（Excel 双击不乱码）。
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OfficeAgent.Host
{
    public static class MiniCsv
    {
        public static string DetectRead(string path, out Encoding used)
        {
            byte[] all = File.ReadAllBytes(path);
            if (all.Length >= 3 && all[0] == 0xEF && all[1] == 0xBB && all[2] == 0xBF)
            {
                used = new UTF8Encoding(false);
                return new UTF8Encoding(false).GetString(all, 3, all.Length - 3);
            }
            if (all.Length >= 2 && all[0] == 0xFF && all[1] == 0xFE)
            {
                used = Encoding.Unicode;
                return Encoding.Unicode.GetString(all, 2, all.Length - 2);
            }
            try
            {
                Encoding strict = new UTF8Encoding(false, true);
                string s = strict.GetString(all);
                used = new UTF8Encoding(false);
                return s;
            }
            catch
            {
                used = Encoding.GetEncoding("GB18030");
                return used.GetString(all);
            }
        }

        public static List<string[]> Parse(string text)
        {
            List<string[]> rows = new List<string[]>();
            List<string> row = new List<string>();
            StringBuilder cur = new StringBuilder();
            bool inQuotes = false;
            int len = text.Length;
            for (int i = 0; i < len; i++)
            {
                char c = text[i];
                if (inQuotes)
                {
                    if (c == '"')
                    {
                        if (i + 1 < len && text[i + 1] == '"') { cur.Append('"'); i++; }
                        else inQuotes = false;
                    }
                    else cur.Append(c);
                }
                else
                {
                    if (c == '"' && cur.Length == 0) inQuotes = true;
                    else if (c == ',') { row.Add(cur.ToString()); cur.Length = 0; }
                    else if (c == '\r') { if (i + 1 < len && text[i + 1] == '\n') i++; EndRow(rows, row, cur); }
                    else if (c == '\n') EndRow(rows, row, cur);
                    else cur.Append(c);
                }
            }
            if (cur.Length > 0 || row.Count > 0) EndRow(rows, row, cur);
            return rows;
        }

        static void EndRow(List<string[]> rows, List<string> row, StringBuilder cur)
        {
            row.Add(cur.ToString());
            cur.Length = 0;
            rows.Add(row.ToArray());
            row.Clear();
        }

        public static string Write(string path, List<string[]> rows)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                foreach (string[] row in rows)
                {
                    for (int i = 0; i < row.Length; i++)
                    {
                        if (i > 0) sb.Append(',');
                        sb.Append(QuoteIfNeeded(row[i]));
                    }
                    sb.Append("\r\n");
                }
                File.WriteAllText(path, sb.ToString(), new UTF8Encoding(true));
                return null;
            }
            catch (Exception ex) { return ex.Message; }
        }

        static string QuoteIfNeeded(string s)
        {
            if (s == null) return "";
            if (s.IndexOf(',') >= 0 || s.IndexOf('"') >= 0 || s.IndexOf('\n') >= 0 || s.IndexOf('\r') >= 0)
            {
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            }
            return s;
        }
    }
}
