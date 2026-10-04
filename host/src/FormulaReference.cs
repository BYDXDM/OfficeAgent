// FormulaReference —— excel_formula_reference 工具：内置 Excel 公式大全的只读查询。
// 数据文件 templates\excel-formulas.json（随安装包分发，离线可用）；
// 用户频繁问公式，模型先查这里拿到标准语法与示例再回答，不靠记忆。
// 只读工具：不写文件、不登记产物；重放无害（AgentLoop.IsMutatingTool 不收录）。
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeAgent.Core;

namespace OfficeAgent.Host
{
    public static class FormulaReference
    {
        class Entry
        {
            public string Cat = "";      // 分类（数学/统计/逻辑/文本/日期/查找引用/财务会计）
            public string Name = "";     // 公式名（如 VLOOKUP）
            public string Syntax = "";   // 语法
            public string Desc = "";     // 中文说明
            public string Example = "";  // 示例
        }

        static List<Entry> cache = null;
        static string loadErr = null;

        static bool Load()
        {
            if (cache != null) return true;
            if (loadErr != null) return false;
            try
            {
                string path = Path.Combine(Path.Combine(EnvDetect.FindRoot(), "templates"), "excel-formulas.json");
                if (!File.Exists(path)) { loadErr = "公式库文件缺失: " + path; return false; }
                List<object> arr = JsonVal.ParseArray(File.ReadAllText(path, Encoding.UTF8));
                cache = new List<Entry>();
                for (int i = 0; i < arr.Count; i++)
                {
                    Dictionary<string, object> d = arr[i] as Dictionary<string, object>;
                    if (d == null) continue;
                    Entry e = new Entry();
                    e.Cat = NullAsEmpty(JsonVal.Str(d, "cat"));
                    e.Name = NullAsEmpty(JsonVal.Str(d, "name"));
                    e.Syntax = NullAsEmpty(JsonVal.Str(d, "syntax"));
                    e.Desc = NullAsEmpty(JsonVal.Str(d, "desc"));
                    e.Example = NullAsEmpty(JsonVal.Str(d, "example"));
                    if (e.Name.Length > 0) cache.Add(e);
                }
                return true;
            }
            catch (Exception ex)
            {
                loadErr = "公式库解析失败: " + ex.Message;
                return false;
            }
        }

        static string NullAsEmpty(string s) { return s == null ? "" : s; }

        static bool Contains(string hay, string needle)
        {
            if (hay == null || needle == null || needle.Length == 0) return false;
            return hay.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        // 查询入口：keyword（模糊匹配名称/语法/说明）、category（精确分类）、name（公式名）至少给一个。
        // 全空 → 返回分类概览。命中 0 条不算失败（ok=true），让模型换个词重查。
        public static string Lookup(string keyword, string category, string name, out bool ok)
        {
            ok = true;
            if (!Load()) { ok = false; return "公式库不可用: " + loadErr; }
            string kw = (keyword == null ? "" : keyword.Trim());
            string cat = (category == null ? "" : category.Trim());
            string nm = (name == null ? "" : name.Trim());

            if (kw.Length == 0 && cat.Length == 0 && nm.Length == 0)
            {
                List<string> cats = new List<string>();
                List<int> counts = new List<int>();
                for (int i = 0; i < cache.Count; i++)
                {
                    int idx = cats.IndexOf(cache[i].Cat);
                    if (idx < 0) { cats.Add(cache[i].Cat); counts.Add(1); }
                    else counts[idx] = counts[idx] + 1;
                }
                StringBuilder sb = new StringBuilder();
                sb.Append("内置 Excel 公式库共 ").Append(cache.Count).Append(" 条，分类：");
                for (int i = 0; i < cats.Count; i++)
                {
                    if (i > 0) sb.Append("、");
                    sb.Append(cats[i]).Append("(").Append(counts[i]).Append(")");
                }
                sb.Append("。请带 keyword（关键词）或 category（分类）或 name（公式名）查询。");
                return sb.ToString();
            }

            List<Entry> hits = new List<Entry>();
            for (int i = 0; i < cache.Count; i++)
            {
                Entry e = cache[i];
                if (cat.Length > 0 && string.Compare(e.Cat, cat, StringComparison.OrdinalIgnoreCase) != 0) continue;
                if (nm.Length > 0 && !Contains(e.Name, nm) && !Contains(e.Syntax, nm)) continue;
                if (kw.Length > 0 && !Contains(e.Name, kw) && !Contains(e.Syntax, kw) && !Contains(e.Desc, kw) &&
                    !Contains(e.Cat, kw) && !Contains(e.Example, kw)) continue;
                hits.Add(e);
            }

            if (hits.Count == 0)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append("未找到匹配的公式（keyword=").Append(kw.Length == 0 ? "-" : kw);
                sb.Append(" category=").Append(cat.Length == 0 ? "-" : cat);
                sb.Append(" name=").Append(nm.Length == 0 ? "-" : nm).Append("）。");
                sb.Append("可换个关键词，或先不带参数查看全部分类。");
                return sb.ToString();
            }

            const int MaxShow = 15;
            StringBuilder outSb = new StringBuilder();
            outSb.Append("命中 ").Append(hits.Count).Append(" 条");
            if (hits.Count > MaxShow) outSb.Append("（显示前 ").Append(MaxShow).Append(" 条，请加更精确的 keyword）");
            outSb.Append("：\n");
            for (int i = 0; i < hits.Count && i < MaxShow; i++)
            {
                Entry e = hits[i];
                outSb.Append("\n【").Append(e.Cat).Append("】").Append(e.Name);
                if (e.Syntax.Length > 0) outSb.Append("\n  语法: ").Append(e.Syntax);
                if (e.Desc.Length > 0) outSb.Append("\n  说明: ").Append(e.Desc);
                if (e.Example.Length > 0) outSb.Append("\n  示例: ").Append(e.Example);
            }
            return outSb.ToString();
        }
    }
}
