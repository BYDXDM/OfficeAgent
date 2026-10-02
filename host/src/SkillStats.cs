// SkillStats —— 技能调用成功率台账（规划层第三期）
//
// 目的：把"哪个动作历史上真的跑得通"固化成数据，供两处使用：
//   ① 失败换路（SkillToolBridge.Alternatives）：同分候选里优先推成功率高的；
//   ② 系统提示（ChatHint）：把"反复失败的动作"标出来，劝模型别浪费预算去试。
//
// ★ 为什么是"成功率"而不是"耗时排序"：
//   第二期实测暴露的真实成本是**失败与重试**（起一次 Python sidecar ~650ms，
//   失败后模型还要再想一轮、再调一次），而不是单次调用的毫秒差。
//   排序按成功率能直接减少失败次数；按耗时排序只在都成功时有意义。
//
// ★ 为什么落盘而不是只放内存：
//   成功率的价值来自**跨会话累积**。进程内统计在"每次启动都是第一次"的场景下
//   永远是 0/0，等于没做。故落到 %LOCALAPPDATA%\OfficeAgent\skill-stats.json。
//
// ★ 为什么必须**失败也要记**、且不设"成功就清零"的乐观策略：
//   环境类失败（缺 VC++/LibreOffice）修好后是会转好的，故用**全部历史**计数，
//   但只在样本数足够（MinSamples）时才拿成功率排序——避免"试了一次失败"就永久拉黑。
//
// 读写约定：零依赖扁平 JSON（MiniJson 不支持嵌套，见设计文档 §1.5）。
//          格式 = {"skill|action":{"ok":N,"fail":N}, ...}
//          写入走临时文件 + 替换，避免中途崩溃留下半截文件。
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OfficeAgent.Host
{
    public static class SkillStats
    {
        // 至少要有这么多次调用，成功率才被用于排序/告警。
        // 取 2：一次失败可能只是路径写错，两次仍失败才值得提示模型。
        public const int MinSamples = 2;

        // 成功率低于该值即视为"这个动作历史表现差"，在提示里点名。
        //
        // ★ 边界语义：**严格小于**才点名（恰好 50% 不点名）。
        //   理由是踩过的坑：初版自检想验"差动作被点名"，随手造了 2 成功 2 失败（正好 50%），
        //   结果没被点名就以为代码有 bug。实际是**测试用例选在了边界上**，语义本身没写清楚。
        //   现在把这条钉进注释与自检：50% = "各占一半，还谈不上差"，不该去打扰模型；
        //   低于 50%（失败多于成功）才是真的值得提醒。自检里用 1/2 而不是 2/2 构造差动作。
        public const double BadRate = 0.5;

        public class Entry
        {
            public int Ok;
            public int Fail;
            public int Total { get { return Ok + Fail; } }
            public double Rate { get { return Total == 0 ? 0.0 : (double)Ok / Total; } }
        }

        static Dictionary<string, Entry> table = null;
        static readonly object gate = new object();

        public static string FilePath()
        {
            return Path.Combine(AuditLog.Dir(), "skill-stats.json");
        }

        static string Key(string skillId, string action)
        {
            return (skillId == null ? "" : skillId) + "|" + (action == null ? "" : action);
        }

        // ---------- 载入 / 保存 ----------

        // 扁平编码（**不是**嵌套 JSON）：
        //   {"rows":"skill|action|ok|fail;skill|action|ok|fail;..."}
        //
        // ★ 为什么不用嵌套对象 {"a|b":{"ok":1}}：MiniJson 不支持嵌套，且**静默出错**。
        //   实测喂 {"acct-tools|vat":{"ok":3,"fail":1},"xlsx-ops|inspect":{"ok":5,"fail":0}}：
        //     第 1 个对象 = { "acct-tools|vat": "{\"ok\":3", "fail": "1" }
        //        ——内层对象被截成半截字符串，且 "fail" 键**窜到了外层**（与外层同名键冲突时还会被丢弃）
        //     第 2 个对象 = { "ok": "5", "fail": "0" }
        //        ——数组第二项被拍平成独立顶层对象，其 skill 归属**彻底丢失**
        //   也就是说嵌套写法会把统计**读成错的数据**且不报错。设计文档 §1.5 已把这条钉为红线，
        //   故这里沿用与 skill.json 相同的扁平编码：字段用 '|' 分隔、记录用 ';' 分隔。
        //   动作名里不含 '|' 或 ';'（skill.json 的分隔符约定已保证这一点），故编码无歧义。
        static Dictionary<string, Entry> Load()
        {
            Dictionary<string, Entry> t = new Dictionary<string, Entry>();
            try
            {
                string f = FilePath();
                if (!File.Exists(f)) return t;
                string json = File.ReadAllText(f, Encoding.UTF8);
                List<Dictionary<string, string>> objs = MiniJson.ParseObjects(json);
                if (objs.Count == 0) return t;
                string rows = MiniJson.Get(objs[0], "rows");
                if (rows == null || rows.Trim().Length == 0) return t;
                foreach (string raw in rows.Split(';'))
                {
                    string seg = raw.Trim();
                    if (seg.Length == 0) continue;
                    string[] parts = seg.Split('|');
                    if (parts.Length < 4) continue;   // skill|action|ok|fail
                    int okv, failv;
                    if (!int.TryParse(parts[2].Trim(), out okv) || okv < 0) continue;
                    if (!int.TryParse(parts[3].Trim(), out failv) || failv < 0) continue;
                    string k = Key(parts[0].Trim(), parts[1].Trim());
                    Entry e;
                    if (!t.TryGetValue(k, out e)) { e = new Entry(); t[k] = e; }
                    e.Ok = okv; e.Fail = failv;
                }
            }
            catch { }
            return t;
        }

        static Dictionary<string, Entry> Table()
        {
            lock (gate)
            {
                if (table == null) table = Load();
                return table;
            }
        }

        // 测试用：丢弃内存缓存，强制下次重新读盘
        public static void InvalidateCache()
        {
            lock (gate) { table = null; }
        }

        // 落盘：先写**本进程独有**的临时文件，再用 File.Replace 原子替换。
        //
        // ★ 为什么不能用固定的 ".tmp" 名 + Delete + Move（初版写法，两个真缺陷）：
        //   ① 固定名跨进程互撞：GUI 与 CLI 自测同时记一笔时，两个进程写同一个 .tmp，
        //      后写的覆盖先写的，先写的再 Move 就得到一个**半截文件**或直接抛异常。
        //   ② Delete 与 Move 之间存在窗口：进程若在此刻退出/被杀，台账直接**整个消失**。
        //      成功率是"累积"才有价值的东西，丢一次等于用户的历史全没了。
        //   File.Replace 是原子的（失败时保留原文件），临时名带进程号与线程号故不会互撞。
        //   若目标尚不存在（首次写入），Replace 会抛错，故回退到 Move。
        static void Save(Dictionary<string, Entry> t)
        {
            try
            {
                string f = FilePath();
                string dir = Path.GetDirectoryName(f);
                if (dir != null && dir.Length > 0 && !Directory.Exists(dir)) Directory.CreateDirectory(dir);
                // 扁平编码：skill|action|ok|fail，记录间用 ';'
                StringBuilder rows = new StringBuilder();
                foreach (KeyValuePair<string, Entry> kv in t)
                {
                    if (kv.Value.Total == 0) continue;
                    if (rows.Length > 0) rows.Append(";");
                    rows.Append(kv.Key).Append("|").Append(kv.Value.Ok).Append("|").Append(kv.Value.Fail);
                }
                StringBuilder sb = new StringBuilder();
                sb.Append("{\"schema\":1,\"rows\":\"")
                  .Append(MiniJson.Esc(rows.ToString())).Append("\"}");
                // 进程+线程唯一，跨进程并发也不会写到同一个临时文件
                string tmp = f + "." + System.Diagnostics.Process.GetCurrentProcess().Id +
                    "." + System.Threading.Thread.CurrentThread.ManagedThreadId + ".tmp";
                File.WriteAllText(tmp, sb.ToString(), new UTF8Encoding(false));
                if (File.Exists(f))
                {
                    // 原子替换：新内容就位或原文件保留，不存在"两边都没有"的窗口
                    File.Replace(tmp, f, null);
                }
                else
                {
                    File.Move(tmp, f);
                }
            }
            catch
            {
                // 记录失败绝不能影响技能调用本身（这是"锦上添花"的台账）
            }
        }

        // ---------- 记录 ----------

        // 记一次调用结果。失败与成功都记（失败才是本台账要减少的对象）。
        //
        // ★ 递增前必须**重新读盘**，不能信进程内缓存：
        //   本表是"读-改-写"整个文件。若两个进程各自持有自己的快照（GUI 跑了一次技能、
        //   同时 CLI 自测也跑了一次），双方都会用**自己那份旧表**覆盖对方的结果——
        //   后写的那个把对方的调用**整个抹掉**。抽查时表现为"明明跑过，统计里没有"。
        //   成功率是低写入频率（每次技能调用一笔）、高读取价值的场景，
        //   每次重读的成本可忽略，换来的是不会静默丢数据。
        //   注意：这仍不是严格的事务（没有跨进程文件锁），极端并发下仍可能丢**单笔**；
        //   但对"成功率排序"这个用途，量级正确即可，没必要引入审计链那样的锁开销。
        public static void Record(string skillId, string action, bool ok)
        {
            try
            {
                string k = Key(skillId, action);
                Dictionary<string, Entry> fresh = Load();   // 重读，避免覆盖别的进程写入
                Entry e;
                if (!fresh.TryGetValue(k, out e)) { e = new Entry(); fresh[k] = e; }
                if (ok) e.Ok++; else e.Fail++;
                lock (gate)
                {
                    table = fresh;      // 缓存与刚落盘的内容保持一致
                }
                Save(fresh);
            }
            catch { }
        }

        // ---------- 查询 ----------

        // 取某动作的统计；无记录时返回 null（调用方据此不排序——**没数据不等于差**）
        public static Entry Get(string skillId, string action)
        {
            Entry e;
            if (Table().TryGetValue(Key(skillId, action), out e)) return e;
            return null;
        }

        // 成功率，样本不足时返回 -1（= 未知，排序时排在中性位置）
        public static double Rate(string skillId, string action)
        {
            Entry e = Get(skillId, action);
            if (e == null || e.Total < MinSamples) return -1.0;
            return e.Rate;
        }

        // 历史表现差（样本足够且成功率低于阈值）的动作清单，供提示点名。
        // 只返回 actionCount 个以内，避免提示膨胀。
        public static List<string[]> BadActions(string skillId, int max)
        {
            List<string[]> bad = new List<string[]>();
            if (max <= 0) return bad;
            foreach (KeyValuePair<string, Entry> kv in Table())
            {
                int bar = kv.Key.IndexOf('|');
                if (bar <= 0) continue;
                string sid = kv.Key.Substring(0, bar);
                string act = kv.Key.Substring(bar + 1);
                if (sid != skillId) continue;
                Entry e = kv.Value;
                if (e.Total < MinSamples) continue;
                if (e.Rate >= BadRate) continue;
                bad.Add(new string[] { act, e.Ok.ToString(), e.Fail.ToString() });
            }
            // 稳定排序：失败次数多的在前（同等成功率时更值得提醒）
            bad.Sort(delegate(string[] a, string[] b) { return b[2].CompareTo(a[2]); });
            if (bad.Count > max) bad.RemoveRange(max, bad.Count - max);
            return bad;
        }

        // 全部统计的简要文本（自检与"技能统计"诊断用）
        public static string Describe()
        {
            Dictionary<string, Entry> t = Table();
            if (t.Count == 0) return "（暂无技能调用统计）";
            List<string> lines = new List<string>();
            foreach (KeyValuePair<string, Entry> kv in t)
            {
                if (kv.Value.Total == 0) continue;
                lines.Add(kv.Key + " ok=" + kv.Value.Ok + " fail=" + kv.Value.Fail +
                    " rate=" + (kv.Value.Rate * 100).ToString("0") + "%");
            }
            lines.Sort();
            StringBuilder sb = new StringBuilder();
            sb.Append("共 ").Append(lines.Count).Append(" 个动作有调用记录：");
            foreach (string l in lines) sb.Append("\n  ").Append(l);
            return sb.ToString();
        }

        // 清空统计（自检用；GUI 不暴露——用户没有理由要清掉这个）
        public static void ResetForTest()
        {
            lock (gate) { table = new Dictionary<string, Entry>(); Save(table); }
        }
    }
}
