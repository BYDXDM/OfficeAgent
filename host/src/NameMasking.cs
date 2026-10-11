// NameMasking —— 会话级姓名假名层（0.9.13）
// 场景：考勤/frp 模板解析出的真实员工姓名不出本机。所有发给模型的文本先把姓名换成
// 「员工01/员工02…」（编号一次会话内稳定），模型产出的「员工01」在显示/落档边界
// 再还原成真名（Restore）——模型全程只见假名，最终答复/会话存档/产物用真名。
// 名单来自 FrpReport 的考勤解析（名字列几何定位），不用通用中文姓名识别——
// "马上/王干/张力/万一"都是普通词，正则猜姓名误报不可接受。
// 线程安全：AgentTools 工具线程与 UI 线程（BuildFileContext/AgentDone）都会调，统一 lock。
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OfficeAgent.Host
{
    public static class NameMasking
    {
        static readonly object gate = new object();
        static readonly List<string> names = new List<string>();

        public static bool HasNames { get { lock (gate) { return names.Count > 0; } } }

        public static void Reset() { lock (gate) { names.Clear(); } }

        // 名单合并注册：已登记的保持编号不变，新名字追加编号（同会话多份文件/多轮分析一致）
        public static void Register(IEnumerable<string> roster)
        {
            if (roster == null) return;
            lock (gate)
            {
                foreach (string n in roster)
                {
                    string t = (n ?? "").Trim();
                    if (t.Length == 0) continue;
                    if (names.IndexOf(t) >= 0) continue;
                    names.Add(t);
                }
            }
        }

        // 当前名单快照（会话存档用，顺序即编号）
        public static List<string> Roster() { lock (gate) { return new List<string>(names); } }

        static string Pseudo(int i) { return "员工" + (i + 1).ToString("00"); }

        // 真名 → 员工NN。长名先替换（短名可能是长名的子串，如"李奎"⊂"李奎远"）。
        public static string Mask(string text)
        {
            if (text == null || text.Length == 0) return text;
            lock (gate)
            {
                if (names.Count == 0) return text;
                List<int> order = new List<int>();
                for (int i = 0; i < names.Count; i++) order.Add(i);
                order.Sort(delegate(int a, int b) { return names[b].Length.CompareTo(names[a].Length); });
                foreach (int i in order) text = text.Replace(names[i], Pseudo(i));
                return text;
            }
        }

        // 员工NN → 真名。整词匹配（容忍"员工 1"式空格与 1/2 位数字）；
        // 只还原名单内的编号，模型编造的员工号原样保留。
        public static string Restore(string text)
        {
            if (text == null || text.Length == 0) return text;
            lock (gate)
            {
                if (names.Count == 0) return text;
                return Regex.Replace(text, "员工\\s*(\\d{1,2})", delegate(Match mm)
                {
                    int n;
                    if (!int.TryParse(mm.Groups[1].Value, out n)) return mm.Value;
                    if (n < 1 || n > names.Count) return mm.Value;
                    return names[n - 1];
                });
            }
        }
    }
}
