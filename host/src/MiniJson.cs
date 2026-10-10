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

        // 严格 JSON 语法校验：合法返回 null，否则返回 "位置 N: 说明"（N 为 1-based 字符下标）。
        // 存在的理由：工具 schema / WebSearch 工具是**手写的 JSON 字符串**，漏写一个反斜杠
        // 就会让请求体出现非法转义（真实事故：`\表` → 网关 400 "Unrecognized character escape"），
        // 服务端只回错误堆栈、指不到源头。自测里先把这份 JSON 验一遍，这类 bug 就出不了门。
        public static string Validate(string json)
        {
            if (json == null) return "输入为 null";
            int i = 0;
            SkipWs(json, ref i);
            string err = ScanValue(json, ref i);
            if (err != null) return err;
            SkipWs(json, ref i);
            if (i != json.Length) return "位置 " + (i + 1) + ": 顶层值之后还有多余内容";
            return null;
        }

        static void SkipWs(string s, ref int i)
        {
            while (i < s.Length && (s[i] == ' ' || s[i] == '\t' || s[i] == '\r' || s[i] == '\n')) i++;
        }

        static string ScanValue(string s, ref int i)
        {
            if (i >= s.Length) return "位置 " + (i + 1) + ": 值缺失";
            char c = s[i];
            if (c == '{') return ScanObject(s, ref i);
            if (c == '[') return ScanArray(s, ref i);
            if (c == '"') return ScanString(s, ref i);
            if (c == 't') return ScanLiteral(s, ref i, "true");
            if (c == 'f') return ScanLiteral(s, ref i, "false");
            if (c == 'n') return ScanLiteral(s, ref i, "null");
            if (c == '-' || (c >= '0' && c <= '9')) return ScanNumber(s, ref i);
            return "位置 " + (i + 1) + ": 非法字符 '" + c + "'";
        }

        static string ScanLiteral(string s, ref int i, string lit)
        {
            if (i + lit.Length > s.Length || string.CompareOrdinal(s, i, lit, 0, lit.Length) != 0)
                return "位置 " + (i + 1) + ": 期望 " + lit;
            i += lit.Length;
            return null;
        }

        static string ScanObject(string s, ref int i)
        {
            i++;    // 跳过 '{'
            SkipWs(s, ref i);
            if (i < s.Length && s[i] == '}') { i++; return null; }
            while (true)
            {
                SkipWs(s, ref i);
                if (i >= s.Length || s[i] != '"') return "位置 " + (i + 1) + ": 对象的键必须是字符串";
                string e1 = ScanString(s, ref i);
                if (e1 != null) return e1;
                SkipWs(s, ref i);
                if (i >= s.Length || s[i] != ':') return "位置 " + (i + 1) + ": 键之后缺少 ':'";
                i++;
                SkipWs(s, ref i);
                string e2 = ScanValue(s, ref i);
                if (e2 != null) return e2;
                SkipWs(s, ref i);
                if (i < s.Length && s[i] == ',') { i++; continue; }
                if (i < s.Length && s[i] == '}') { i++; return null; }
                return "位置 " + (i + 1) + ": 对象里期望 ',' 或 '}'";
            }
        }

        static string ScanArray(string s, ref int i)
        {
            i++;    // 跳过 '['
            SkipWs(s, ref i);
            if (i < s.Length && s[i] == ']') { i++; return null; }
            while (true)
            {
                SkipWs(s, ref i);
                string e = ScanValue(s, ref i);
                if (e != null) return e;
                SkipWs(s, ref i);
                if (i < s.Length && s[i] == ',') { i++; continue; }
                if (i < s.Length && s[i] == ']') { i++; return null; }
                return "位置 " + (i + 1) + ": 数组里期望 ',' 或 ']'";
            }
        }

        static string ScanString(string s, ref int i)
        {
            i++;    // 跳过开引号
            while (true)
            {
                if (i >= s.Length) return "字符串未闭合";
                char c = s[i];
                if (c == '"') { i++; return null; }
                if (c == '\\')
                {
                    if (i + 1 >= s.Length) return "位置 " + (i + 1) + ": 反斜杠后缺少转义字符";
                    char e = s[i + 1];
                    if (e == '"' || e == '\\' || e == '/' || e == 'b' || e == 'f'
                        || e == 'n' || e == 'r' || e == 't')
                    {
                        i += 2;
                        continue;
                    }
                    if (e == 'u')
                    {
                        if (i + 5 >= s.Length) return "位置 " + (i + 1) + ": \\u 后面需要 4 位十六进制";
                        for (int k = i + 2; k <= i + 5; k++)
                        {
                            char h = s[k];
                            bool hex = (h >= '0' && h <= '9') || (h >= 'a' && h <= 'f') || (h >= 'A' && h <= 'F');
                            if (!hex) return "位置 " + (k + 1) + ": \\u 后面不是十六进制数字（'" + h + "'）";
                        }
                        i += 6;
                        continue;
                    }
                    return "位置 " + (i + 1) + ": 非法转义 \\" + e
                        + "（JSON 只允许 \\\" \\\\ \\/ \\b \\f \\n \\r \\t \\uXXXX；要表达字面反斜杠必须写两个）";
                }
                if (c < ' ') return "位置 " + (i + 1) + ": 字符串里出现未转义的控制字符（0x"
                    + ((int)c).ToString("x2") + "）";
                i++;
            }
        }

        static string ScanNumber(string s, ref int i)
        {
            int start = i;
            if (i < s.Length && s[i] == '-') i++;
            int digits = 0;
            while (i < s.Length && s[i] >= '0' && s[i] <= '9') { i++; digits++; }
            if (i < s.Length && s[i] == '.')
            {
                i++;
                while (i < s.Length && s[i] >= '0' && s[i] <= '9') { i++; digits++; }
            }
            if (i < s.Length && (s[i] == 'e' || s[i] == 'E'))
            {
                int save = i;
                i++;
                if (i < s.Length && (s[i] == '+' || s[i] == '-')) i++;
                int ed = 0;
                while (i < s.Length && s[i] >= '0' && s[i] <= '9') { i++; ed++; }
                if (ed == 0) i = save;    // 不是指数，回退，交给上层报错
            }
            if (digits == 0) return "位置 " + (start + 1) + ": 数字格式非法";
            return null;
        }
    }
}
