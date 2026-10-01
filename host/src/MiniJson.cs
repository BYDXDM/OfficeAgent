// MiniJson —— 极简扁平 JSON 解析（对象数组 / 单对象；值域为字符串或标量）
// 供 ReconTemplateStore / MergeTemplateStore / 审计行读取等共用；不支持嵌套数组对象。
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OfficeAgent.Host
{
    internal static class MiniJson
    {
        // 解析顶层 [...] 中的所有 { ... } 对象（也容忍无外层括号的连续对象）
        public static List<Dictionary<string, string>> ParseObjects(string json)
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

        public static string ReadString(string json, ref int i)
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

        public static string Esc(string s)
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

        public static string Get(Dictionary<string, string> o, string key)
        {
            string v;
            return o.TryGetValue(key, out v) && v != null ? v : "";
        }

        public static string GetOr(Dictionary<string, string> o, string key, string def)
        {
            string v = Get(o, key);
            return v.Length > 0 ? v : def;
        }

        public static int GetInt(Dictionary<string, string> o, string key, int def)
        {
            int n;
            return int.TryParse(Get(o, key), out n) ? n : def;
        }

        public static double GetDouble(Dictionary<string, string> o, string key, double def)
        {
            double d;
            return double.TryParse(Get(o, key), NumberStyles.Float, CultureInfo.InvariantCulture, out d) ? d : def;
        }

        public static bool GetBool(Dictionary<string, string> o, string key, bool def)
        {
            string v = Get(o, key);
            return v.Length == 0 ? def : v == "true";
        }
    }
}
