// MaskEngine —— 隐私分级与脱敏（设计方案 §7.2）
//   L0 全本地：不触网，模型对话禁用（纯规则功能可用）
//   L1 脱敏出网（默认）：发给模型的文本先做规则脱敏 —— 手机号/身份证/银行卡/税号/邮箱
//   L2 全量出网：用户显式选择
// 脱敏只作用于"出网内容"；本机文件、输出与审计不受影响。命中规则列表交审计记录。
// 保守优先：无法识别时宁可多脱敏（15 位老身份证按卡号规则兜底，覆盖面优先于精确）。
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OfficeAgent.Host
{
    public static class MaskEngine
    {
        public class Rule
        {
            public string Name;
            public Regex Re;
            public MatchEvaluator Eval;
        }

        static List<Rule> rules = null;

        static string KeepEnds(string s, int head, int tail)
        {
            if (s.Length <= head + tail) return new string('*', s.Length);
            return s.Substring(0, head) + new string('*', s.Length - head - tail) + s.Substring(s.Length - tail);
        }

        static void EnsureRules()
        {
            if (rules != null) return;
            rules = new List<Rule>();

            // 手机号（大陆 11 位）
            rules.Add(new Rule
            {
                Name = "手机号",
                Re = new Regex("(?<!\\d)1[3-9]\\d{9}(?!\\d)"),
                Eval = delegate(Match m) { return KeepEnds(m.Value, 3, 2); }
            });
            // 身份证 18 位（含尾 X）
            rules.Add(new Rule
            {
                Name = "身份证",
                Re = new Regex("(?<!\\d)\\d{17}[0-9Xx](?!\\d)"),
                Eval = delegate(Match m) { return KeepEnds(m.Value, 4, 2); }
            });
            // 银行卡 13-19 位（借记/信用卡常见段）
            rules.Add(new Rule
            {
                Name = "银行卡",
                Re = new Regex("(?<!\\d)\\d{13,19}(?!\\d)"),
                Eval = delegate(Match m) { return KeepEnds(m.Value, 4, 4); }
            });
            // 统一社会信用代码 / 税号（18 位大写字母数字，剔除易混 I/O/S/V/Z）
            rules.Add(new Rule
            {
                Name = "税号",
                Re = new Regex("(?<![0-9A-Z])[0-9A-HJ-NPQRTUWXY]{18}(?![0-9A-Z])"),
                Eval = delegate(Match m) { return KeepEnds(m.Value, 4, 3); }
            });
            // 身份证 15 位（老号段，卡号规则之后兜底）
            rules.Add(new Rule
            {
                Name = "身份证",
                Re = new Regex("(?<!\\d)\\d{15}(?!\\d)"),
                Eval = delegate(Match m) { return KeepEnds(m.Value, 4, 2); }
            });
            // 邮箱
            rules.Add(new Rule
            {
                Name = "邮箱",
                Re = new Regex("[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\\.[A-Za-z]{2,}"),
                Eval = delegate(Match m)
                {
                    string v = m.Value;
                    int at = v.IndexOf('@');
                    string local = at > 2 ? v.Substring(0, 2) + "***" : "***";
                    return local + v.Substring(at);
                }
            });
        }

        // 返回脱敏后文本；hitsOut 收集命中的规则名（去重，供审计）
        public static string Mask(string text, List<string> hitsOut)
        {
            EnsureRules();
            if (text == null || text.Length == 0) return text;
            string result = text;
            foreach (Rule r in rules)
            {
                string before = result;
                result = r.Re.Replace(result, r.Eval);
                if (hitsOut != null && before != result && !hitsOut.Contains(r.Name)) hitsOut.Add(r.Name);
            }
            return result;
        }

        public static string LevelName(int level)
        {
            switch (level)
            {
                case 0: return "L0 全本地（不出网）";
                case 1: return "L1 脱敏出网（默认）";
                case 2: return "L2 全量出网";
                default: return "L" + level;
            }
        }

        // 无头自测：/masktest。退出码 0=全过，2=有失败
        public static int SelfTest()
        {
            // [输入, 期望(1=脱敏后不含原文敏感段, 0=不应被改动)]；脱敏期望值直接给定
            string[][] cases = new string[][] {
                new string[] { "客户电话 13812345678 麻烦回电", "客户电话 138******78 麻烦回电" },
                new string[] { "身份证 11010119900307861X 已核对", "身份证 1101************1X 已核对" },
                new string[] { "打款到 6222020200112233445", "打款到 6222***********3445" },
                new string[] { "税号 91330106MA2ABC123X 上传", "税号 9133***********23X 上传" },
                new string[] { "邮箱 zhang.san@corp.com.cn 收件", "邮箱 zh***@corp.com.cn 收件" },
                new string[] { "金额 5000.00 与 1,234.56 元无需脱敏", null },
                new string[] { "日期 2026-08-20 与凭证号 A00012345 无需脱敏", null },
                new string[] { "行号 123456789012 (12位数字) 不动", null }
            };
            int failed = 0;
            foreach (string[] c in cases)
            {
                List<string> hits = new List<string>();
                string got = Mask(c[0], hits);
                bool ok = c[1] == null ? got == c[0] : got == c[1];   // null 期望 = 原文不动
                if (!ok) failed++;
                Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + c[0] + "  =>  " + got +
                    (hits.Count > 0 ? "  [" + string.Join(",", hits.ToArray()) + "]" : ""));
            }
            Console.WriteLine(failed == 0 ? "ALL PASS" : (failed + " FAILED"));
            return failed == 0 ? 0 : 2;
        }
    }
}
