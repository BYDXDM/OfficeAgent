// JsonVal —— 极简递归 JSON 解析器（create_formula_workbook / excel_formula_reference 专用）。
// 为什么不用 MiniJson：MiniJson.ParseObjects 只支持一层扁平对象（skill.json 的历史约束），
// 而工作簿的 sheets 规格是嵌套结构（数组里套对象、对象里套数组）。
// 红线：C# 3.0 语法、纯 BCL、输入是模型给的参数（只作数据用，绝不参与命令行/shell 构造）。
// 值域映射：object → Dictionary<string,object>；array → List<object>；
//           string → string；number → double；true/false → bool；null → null。
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OfficeAgent.Host
{
    public static class JsonVal
    {
        // 解析失败抛 ArgumentException（带位置）；调用方以"参数不合法"话术回给模型。
        public static object Parse(string json)
        {
            if (json == null) throw new ArgumentException("JSON 为空");
            int i = 0;
            object v = ParseValue(json, ref i);
            SkipWs(json, ref i);
            if (i != json.Length) throw new ArgumentException("JSON 结尾有多余内容（位置 " + i + "）");
            return v;
        }

        public static Dictionary<string, object> ParseObject(string json)
        {
            object v = Parse(json);
            Dictionary<string, object> d = v as Dictionary<string, object>;
            if (d == null) throw new ArgumentException("应为 JSON 对象");
            return d;
        }

        public static List<object> ParseArray(string json)
        {
            object v = Parse(json);
            List<object> a = v as List<object>;
            if (a == null) throw new ArgumentException("应为 JSON 数组");
            return a;
        }

        // ---------- 取值辅助（类型不符返回默认值而非抛异常，调用方好写） ----------

        public static string Str(Dictionary<string, object> d, string key)
        {
            if (d == null || !d.ContainsKey(key)) return null;
            if (d[key] is string) return (string)d[key];
            if (d[key] is double) return ((double)d[key]).ToString("0.##########", CultureInfo.InvariantCulture);
            if (d[key] is bool) return (bool)d[key] ? "true" : "false";
            return null;
        }

        public static double Num(Dictionary<string, object> d, string key, double def)
        {
            if (d == null || !d.ContainsKey(key)) return def;
            if (d[key] is double) return (double)d[key];
            double p;
            if (d[key] is string && double.TryParse((string)d[key], NumberStyles.Any,
                CultureInfo.InvariantCulture, out p)) return p;
            return def;
        }

        public static bool Bool(Dictionary<string, object> d, string key, bool def)
        {
            if (d == null || !d.ContainsKey(key)) return def;
            if (d[key] is bool) return (bool)d[key];
            return def;
        }

        public static List<object> List(Dictionary<string, object> d, string key)
        {
            if (d == null || !d.ContainsKey(key)) return null;
            return d[key] as List<object>;
        }

        public static Dictionary<string, object> Obj(Dictionary<string, object> d, string key)
        {
            if (d == null || !d.ContainsKey(key)) return null;
            return d[key] as Dictionary<string, object>;
        }

        // ---------- 解析实现 ----------

        static void SkipWs(string s, ref int i)
        {
            while (i < s.Length && (s[i] == ' ' || s[i] == '\t' || s[i] == '\r' || s[i] == '\n')) i++;
        }

        static object ParseValue(string s, ref int i)
        {
            SkipWs(s, ref i);
            if (i >= s.Length) throw new ArgumentException("JSON 意外结束");
            char c = s[i];
            if (c == '{') return ParseObj(s, ref i);
            if (c == '[') return ParseArr(s, ref i);
            if (c == '"') return ParseStr(s, ref i);
            if (s.IndexOf("true", i, StringComparison.Ordinal) == i) { i += 4; return true; }
            if (s.IndexOf("false", i, StringComparison.Ordinal) == i) { i += 5; return false; }
            if (s.IndexOf("null", i, StringComparison.Ordinal) == i) { i += 4; return null; }
            return ParseNum(s, ref i);
        }

        static Dictionary<string, object> ParseObj(string s, ref int i)
        {
            Dictionary<string, object> d = new Dictionary<string, object>();
            i++;   // {
            SkipWs(s, ref i);
            if (i < s.Length && s[i] == '}') { i++; return d; }
            while (true)
            {
                SkipWs(s, ref i);
                if (i >= s.Length || s[i] != '"') throw new ArgumentException("对象键应为字符串（位置 " + i + "）");
                string key = ParseStr(s, ref i);
                SkipWs(s, ref i);
                if (i >= s.Length || s[i] != ':') throw new ArgumentException("缺少冒号（位置 " + i + "）");
                i++;
                d[key] = ParseValue(s, ref i);
                SkipWs(s, ref i);
                if (i >= s.Length) throw new ArgumentException("对象未闭合");
                if (s[i] == ',') { i++; continue; }
                if (s[i] == '}') { i++; return d; }
                throw new ArgumentException("对象内应为 , 或 }（位置 " + i + "）");
            }
        }

        static List<object> ParseArr(string s, ref int i)
        {
            List<object> a = new List<object>();
            i++;   // [
            SkipWs(s, ref i);
            if (i < s.Length && s[i] == ']') { i++; return a; }
            while (true)
            {
                a.Add(ParseValue(s, ref i));
                SkipWs(s, ref i);
                if (i >= s.Length) throw new ArgumentException("数组未闭合");
                if (s[i] == ',') { i++; continue; }
                if (s[i] == ']') { i++; return a; }
                throw new ArgumentException("数组内应为 , 或 ]（位置 " + i + "）");
            }
        }

        static string ParseStr(string s, ref int i)
        {
            if (s[i] != '"') throw new ArgumentException("应为字符串（位置 " + i + "）");
            i++;
            StringBuilder sb = new StringBuilder();
            while (i < s.Length)
            {
                char c = s[i];
                if (c == '"') { i++; return sb.ToString(); }
                if (c == '\\')
                {
                    i++;
                    if (i >= s.Length) break;
                    char e = s[i];
                    if (e == 'n') sb.Append('\n');
                    else if (e == 'r') sb.Append('\r');
                    else if (e == 't') sb.Append('\t');
                    else if (e == 'b') sb.Append('\b');
                    else if (e == 'f') sb.Append('\f');
                    else if (e == 'u' && i + 4 < s.Length)
                    {
                        int code;
                        if (int.TryParse(s.Substring(i + 1, 4), NumberStyles.HexNumber,
                            CultureInfo.InvariantCulture, out code)) sb.Append((char)code);
                        i += 4;
                    }
                    else sb.Append(e);
                    i++;
                    continue;
                }
                sb.Append(c);
                i++;
            }
            throw new ArgumentException("字符串未闭合");
        }

        static double ParseNum(string s, ref int i)
        {
            int start = i;
            while (i < s.Length && (char.IsDigit(s[i]) || s[i] == '-' || s[i] == '+' || s[i] == '.' ||
                                    s[i] == 'e' || s[i] == 'E')) i++;
            double v;
            if (double.TryParse(s.Substring(start, i - start), NumberStyles.Any,
                CultureInfo.InvariantCulture, out v)) return v;
            throw new ArgumentException("非法数值（位置 " + start + "）");
        }
    }
}
