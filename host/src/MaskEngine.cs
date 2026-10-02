// MaskEngine —— 隐私分级与脱敏（设计方案 §7.2）
//   L0 全本地：不触网，模型对话禁用（纯规则功能可用）
//   L1 脱敏出网（默认）：发给模型的文本先做规则脱敏 —— 手机号/身份证/银行卡/税号/邮箱
//   L2 全量出网：用户显式选择
// 脱敏只作用于"出网内容"；本机文件、输出与审计不受影响。命中规则列表交审计记录。
// 保守优先：无法识别时宁可多脱敏（15 位老身份证按卡号规则兜底，覆盖面优先于精确）。
using System;
using System.Collections.Generic;
using System.Text;
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

        // 带分隔符号码的脱敏：先按"去掉分隔符"的位数校验（不符则原样返回，避免误伤日期等），
        // 再逐字符替换——保留原分隔符与长度，只把中间字符变成 '*'。
        // 为什么保留分隔符：脱敏后的文本仍要能被人读懂结构（如 6222 **** **** 3445），
        // 且不改动长度便于与原文对照。
        static string MaskSeparated(string s, int head, int tail)
        {
            if (s == null) return "";
            // 统计纯数字（含身份证尾 X）位数
            int digits = 0;
            foreach (char c in s)
            {
                if (c >= '0' && c <= '9') digits++;
                else if (c == 'X' || c == 'x') digits++;
            }
            if (digits < head + tail) return s;      // 位数太少，不像号码 → 不动

            // 按"分组"掩码：整组替换成等长 '*'，保留首组与末组。
            // 为什么按组而非按数字序号：19 位卡号按 4 分组是 4+4+4+4+3，
            // 若按"保留前 4 位 + 后 4 位数字"掩码，会得到 "6222 **** **** ***3 445"
            // —— 末组被切掉一个字符，视觉上不整齐、也不符合卡号书写惯例。
            // 按组掩码得到 "6222 **** **** **** 445"，既覆盖了中间全部数字，也可读。
            string[] parts = System.Text.RegularExpressions.Regex.Split(s, "([\\s-])");
            // Regex.Split 带捕获组时，分隔符也会作为元素出现在结果里（奇数下标）
            int groupCount = 0;
            for (int i = 0; i < parts.Length; i++) { if (parts[i].Length > 0 && parts[i] != " " && parts[i] != "-") groupCount++; }
            if (groupCount < 2) return KeepEnds(s, head, tail);   // 没有分组 → 退回逐字符掩码

            StringBuilder out2 = new StringBuilder(s.Length);
            int gi = 0;
            for (int i = 0; i < parts.Length; i++)
            {
                string p = parts[i];
                if (p.Length == 0) continue;
                if (p == " " || p == "-") { out2.Append(p); continue; }
                gi++;
                // 只有两组时，首尾各保留会一个都不掩（如 "138 5678"）→ 退回逐字符掩码。
                // 三组（如手机号 138 1234 5678）首尾都留只掩了 4 位，仍偏少，
                // 因此对分组数 <= 3 的情况改用逐字符掩码（按 head/tail 数字位数精确控制）。
                if (gi == 1 || gi == groupCount)
                {
                    if (groupCount <= 3) { out2.Append(p); continue; }   // 先原样占位，稍后由逐字符逻辑覆盖
                    out2.Append(p);       // 保留首组与末组
                }
                else
                {
                    for (int k = 0; k < p.Length; k++) out2.Append('*');   // 中间整组掩掉
                }
            }
            string grouped = out2.ToString();
            if (groupCount <= 3) return MaskByDigit(s, head, tail, digits);
            return grouped;
        }

        // 逐字符掩码：按"数字序号"决定是否掩掉，保留分隔符原样。
        // 用于分组不足以整组掩码的场合（手机号 138 1234 5678 等）。
        static string MaskByDigit(string s, int head, int tail, int digits)
        {
            StringBuilder sb = new StringBuilder(s.Length);
            int seen = 0;
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                bool isDigit = (c >= '0' && c <= '9') || c == 'X' || c == 'x';
                if (!isDigit) { sb.Append(c); continue; }
                seen++;
                bool mask = seen > head && seen <= digits - tail;
                sb.Append(mask ? '*' : c);
            }
            return sb.ToString();
        }

        static void EnsureRules()
        {
            if (rules != null) return;
            rules = new List<Rule>();

            // ★ 带分隔符的号码必须先于"连续数字"规则处理。
            //   实测缺口（本轮发现）：银行卡习惯每 4 位加空格（6222 0202 0011 2233 445），
            //   手机号常写成 138-1234-5678 / 138 1234 5678，身份证也可能带空格——
            //   原规则只认连续数字，这些格式**完全不脱敏、原样出网**，等于隐私分级失效。
            //   实现：把分隔符纳入匹配，但在评估器里按"去掉分隔符后的纯数字位数"判定，
            //   位数不符就原样返回（避免误伤 "2026 08 15" 这类日期）。
            //   分隔符只允许空格与半角横线（连字符/短横），且各段长度必须像号码分组。
            // ★ 顺序很重要，两个方向都踩过坑：
            //   · 身份证规则排在银行卡前 → 会把 "6222 0202 0011 2233 445" 的前 16 位当身份证
            //     吃掉，剩下 " 445" 裸露（实测）。
            //   · 银行卡规则排在身份证前 → 会把 "1101 0119 9003 0786 1X" 当银行卡，
            //     末组 "0786" 不掩、且类型误判（实测）。
            //   判据：身份证 18 位，分组形态固定为 4-4-4-4-2（末段 2 位含校验位）。
            //   银行卡常见 4-4-4-4 或 4-4-4-4-3。**末段 2 位**是身份证的特征，
            //   据此可避开银行卡；但 4-4-4-4-2 与 4-4-4-4 共享前四段，
            //   故身份证规则必须用**锚定末尾 2 位**的精确形态，且排在银行卡之前。
            rules.Add(new Rule
            {
                Name = "身份证",   // 带分隔：1101 0119 9003 0786 1X（4-4-4-4-2）
                Re = new Regex("(?<![\\dA-Za-z])\\d{4}[\\s-]\\d{4}[\\s-]\\d{4}[\\s-]\\d{4}[\\s-][0-9Xx]{2}(?![\\dA-Za-z])"),
                Eval = delegate(Match m) { return MaskSeparated(m.Value, 4, 2); }
            });
            rules.Add(new Rule
            {
                Name = "身份证",   // 带分隔：1101 0119 9003 078 61X（4-4-4-3-3）
                Re = new Regex("(?<![\\dA-Za-z])\\d{4}[\\s-]\\d{4}[\\s-]\\d{4}[\\s-]\\d{3}[\\s-][0-9Xx]{3}(?![\\dA-Za-z])"),
                Eval = delegate(Match m) { return MaskSeparated(m.Value, 4, 2); }
            });
            rules.Add(new Rule
            {
                Name = "银行卡",   // 带分隔：6222 0202 0011 2233 445 / 6222-0202-0011-2233-445
                // 段长用 3~4 位而非固定 4 位：19 位卡号常分组为 4+4+4+4+3，
                // 末段只有 3 位。若写死 {4}，正则会在 "2233" 后停下、漏掉 "445"（实测踩到，
                // 表现为 "6222 **** **** 2233 445" 中间 8 位未被掩掉）。
                Re = new Regex("(?<!\\d)\\d{4}[\\s-]\\d{3,4}(?:[\\s-]?\\d{3,4}){0,3}(?!\\d)"),
                Eval = delegate(Match m) { return MaskSeparated(m.Value, 4, 4); }
            });
            rules.Add(new Rule
            {
                Name = "手机号",   // 带分隔：138-1234-5678 / 138 1234 5678
                Re = new Regex("(?<!\\d)1[3-9]\\d[\\s-]?\\d{4}[\\s-]?\\d{4}(?!\\d)"),
                Eval = delegate(Match m) { return MaskSeparated(m.Value, 3, 2); }
            });

            // 手机号（大陆 11 位，连续）
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
                new string[] { "行号 123456789012 (12位数字) 不动", null },
                // ★ 带分隔符的号码（本轮修的真实缺口）：银行卡/手机号/身份证的常见书写方式
                //   每 4 位加空格或横线。原实现只认连续数字，这些格式会被**原样发给模型**，
                //   等于隐私分级失效。以下用例钉住该行为，防止回归。
                new string[] { "卡号 6222 0202 0011 2233 445", "卡号 6222 **** **** **** 445" },
                new string[] { "手机 138-1234-5678", "手机 138-****-**78" },
                new string[] { "手机 138 1234 5678", "手机 138 **** **78" },
                new string[] { "身份证 1101 0119 9003 0786 1X", "身份证 1101 **** **** **** 1X" }
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
