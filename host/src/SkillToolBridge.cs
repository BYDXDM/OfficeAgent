// SkillToolBridge —— 把技能注册表投影成模型可见的 function-calling 工具（规划层第一期）
//
// 背景：此前技能只能被 IntentRouter 的规则命中后执行，模型侧完全看不见技能
//      （AgentTools.SchemasJson 只声明 7 个通用工具）。本类补上这座桥：
//      skill.json 的 actions/actionParams → OpenAI function-calling schema。
//
// 编码约定（**必须与 skills/*/skill.json 保持一致**）：
//   actions       = "动作名|描述;动作名|描述;..."      分隔符：动作间 ';'，名与描述间 '|'
//   actionParams  = "动作名:参数名:类型:必填:描述;..." 分隔符：':'，多项间 ';'
//   类型取值      = string | number | file[] | bool
//   必填取值      = required | optional
//   ⚠️ 描述文本中不得出现 ';' '|' ':'（会与分隔符冲突）。
//      选用扁平编码的原因：MiniJson 不支持嵌套，嵌套结构会被静默解析成错误数据
//      （实测 {"actions":[{"name":"x"}]} 会把内层键提升到外层并提前闭合对象）。
//
// 命名规则：工具名 = "skill_" + skillId + "_" + action，其中 skillId/action 里的 '-' '.'
// 统一替换为 '_'（下划线）。原始 id/action 由 _map 反查，绝不从工具名反推。
//
// 安全：模型给的参数只作数据用——全部写进 request.json（文件），
//      绝不参与任何命令行拼接（命令行是编译期字面量 "-m oa_skill_main"）。
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OfficeAgent.Host
{
    // 一个可调用动作的完整描述（清单解析结果）
    public class SkillActionSpec
    {
        public string SkillId = "";
        public string SkillDir = "";
        public string Action = "";
        public string Desc = "";
        public bool Degraded = false;
        public string DegradedReason = "";
        // 参数：name / type / required / desc
        public List<string[]> Params = new List<string[]>();
    }

    public static class SkillToolBridge
    {
        // 工具名前缀：AgentTools.Dispatch / AgentLoop.IsMutatingTool 据此识别技能调用。
        // 集中在此定义，避免多处硬编码漂移。
        public const string Prefix = "skill_";

        // 投影结果缓存（与 SkillSystem 的场景缓存同寿命：skill.json 变更会 InvalidateCache）
        static List<SkillActionSpec> cache = null;
        static string cacheRoot = null;

        public static void InvalidateCache() { cache = null; cacheRoot = null; }

        static string ToolName(string skillId, string action)
        {
            return Prefix + Sanitize(skillId) + "_" + Sanitize(action);
        }

        static string Sanitize(string s)
        {
            if (s == null) return "";
            StringBuilder sb = new StringBuilder();
            foreach (char c in s)
            {
                // 工具名只保留 [A-Za-z0-9_]：'-' '.' 等一律换 '_'（多数端点对 name 有字符集约束）
                bool ok = (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9') || c == '_';
                sb.Append(ok ? c : '_');
            }
            return sb.ToString();
        }

        // ---------- 清单解析 ----------

        // 解析单个技能的 actions/actionParams。无 actions 段时返回空表
        // （向后兼容：老技能不带 actions，行为与从前一致——只能被规则路由命中）。
        public static List<SkillActionSpec> ParseActions(SkillRegistryEntry entry)
        {
            List<SkillActionSpec> list = new List<SkillActionSpec>();
            if (entry == null || entry.Path == null || entry.Path.Length == 0) return list;
            string file = Path.Combine(entry.Path, "skill.json");
            if (!File.Exists(file)) return list;

            string json;
            try { json = File.ReadAllText(file, Encoding.UTF8); }
            catch { return list; }

            List<Dictionary<string, string>> objs = MiniJson.ParseObjects(json);
            if (objs.Count == 0) return list;
            Dictionary<string, string> o = objs[0];

            string actions = MiniJson.Get(o, "actions");
            if (actions == null || actions.Trim().Length == 0) return list;

            // 先建动作表，再挂参数（参数可能引用尚未出现的动作？不允许——只挂已声明的）
            Dictionary<string, SkillActionSpec> byAction = new Dictionary<string, SkillActionSpec>();
            foreach (string raw in actions.Split(';'))
            {
                string seg = raw.Trim();
                if (seg.Length == 0) continue;
                int bar = seg.IndexOf('|');
                SkillActionSpec sp = new SkillActionSpec();
                sp.SkillId = entry.Id;
                sp.SkillDir = entry.Path;
                sp.Degraded = entry.State == SkillConst.StateDegraded;
                sp.DegradedReason = entry.DegradedReason ?? "";
                if (bar < 0) { sp.Action = seg.Trim(); sp.Desc = ""; }
                else { sp.Action = seg.Substring(0, bar).Trim(); sp.Desc = seg.Substring(bar + 1).Trim(); }
                if (sp.Action.Length == 0) continue;
                if (byAction.ContainsKey(sp.Action)) continue;   // 重复动作名：取首个
                byAction[sp.Action] = sp;
                list.Add(sp);
            }

            string specs = MiniJson.Get(o, "actionParams");
            if (specs != null && specs.Trim().Length > 0)
            {
                foreach (string raw in specs.Split(';'))
                {
                    string seg = raw.Trim();
                    if (seg.Length == 0) continue;
                    string[] parts = seg.Split(':');
                    if (parts.Length < 4) continue;
                    string act = parts[0].Trim();
                    SkillActionSpec sp;
                    if (!byAction.TryGetValue(act, out sp)) continue;   // 引用未声明动作：忽略
                    string pname = parts[1].Trim();
                    string ptype = parts[2].Trim();
                    string preq = parts[3].Trim();
                    // 描述里可能含 ':'，第 5 段之后全部拼回
                    StringBuilder pd = new StringBuilder();
                    for (int i = 4; i < parts.Length; i++)
                    {
                        if (pd.Length > 0) pd.Append(':');
                        pd.Append(parts[i]);
                    }
                    if (pname.Length == 0) continue;
                    if (ptype != "string" && ptype != "number" && ptype != "file[]" && ptype != "bool") ptype = "string";
                    if (preq != "required") preq = "optional";
                    sp.Params.Add(new string[] { pname, ptype, preq, pd.ToString().Trim() });
                }
            }
            return list;
        }

        // ---------- 投影（供 AgentTools.SchemasJson 调用） ----------

        public static List<SkillActionSpec> Collect(string root)
        {
            if (cache != null && cacheRoot == root) return cache;
            List<SkillActionSpec> all = new List<SkillActionSpec>();
            List<SkillRegistryEntry> regs = SkillSystem.CachedScan(root);
            foreach (SkillRegistryEntry e in regs)
            {
                // 禁用技能不投影；degraded 仍投影（让模型知道"有这个能力但环境缺组件"，
                // 比让它完全看不见、只能瞎猜要好）。
                if (e.State == SkillConst.StateDisabled) continue;
                List<SkillActionSpec> acts = ParseActions(e);
                foreach (SkillActionSpec sp in acts) all.Add(sp);
            }
            cache = all;
            cacheRoot = root;
            return all;
        }

        // 生成 function-calling schema 数组（**不含**外层 []，由调用方拼接）
        public static string BuildSchemas(string root)
        {
            List<SkillActionSpec> all = Collect(root);
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < all.Count; i++)
            {
                SkillActionSpec sp = all[i];
                if (i > 0) sb.Append(",");
                sb.Append("{\"type\":\"function\",\"function\":{\"name\":\"")
                  .Append(MiniJson.Esc(ToolName(sp.SkillId, sp.Action)))
                  .Append("\",\"description\":\"").Append(MiniJson.Esc(Describe(sp)))
                  .Append("\",\"parameters\":{\"type\":\"object\",\"properties\":{");
                for (int k = 0; k < sp.Params.Count; k++)
                {
                    string[] p = sp.Params[k];
                    if (k > 0) sb.Append(",");
                    sb.Append("\"").Append(MiniJson.Esc(p[0])).Append("\":{");
                    // file[] 对模型表现为"单个绝对路径字符串"更稳妥（部分端点不认数组）
                    sb.Append("\"type\":\"string\",\"description\":\"")
                      .Append(MiniJson.Esc(DescribeParam(p)))
                      .Append("\"}");
                }
                sb.Append("},\"required\":[");
                bool first = true;
                for (int k = 0; k < sp.Params.Count; k++)
                {
                    string[] p = sp.Params[k];
                    if (p[2] != "required") continue;
                    if (!first) sb.Append(",");
                    first = false;
                    sb.Append("\"").Append(MiniJson.Esc(p[0])).Append("\"");
                }
                sb.Append("]}}}");
            }
            return sb.ToString();
        }

        static string Describe(SkillActionSpec sp)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(sp.Desc);
            if (sb.Length == 0) sb.Append(sp.SkillId).Append(" 的 ").Append(sp.Action).Append(" 操作");
            if (sp.Degraded && sp.DegradedReason.Length > 0)
                sb.Append("（注意：当前环境可能缺少依赖——").Append(sp.DegradedReason).Append("）");
            return sb.ToString();
        }

        static string DescribeParam(string[] p)
        {
            string d = p[3];
            string extra;
            if (p[1] == "file[]") extra = "（文件的绝对路径）";
            else if (p[1] == "number") extra = "（数字）";
            else if (p[1] == "bool") extra = "（true/false）";
            else extra = "";
            if (d.Length == 0) return extra.Length > 0 ? extra.Substring(1, extra.Length - 2) : "";
            return d + extra;
        }

        // ---------- 失败换路 ----------

        // 找出与失敗动作"语义相近"的替代动作（规划层第二期 · 失败换路）。
        // 相似度 = 两个技能 skill.json 里 scenarios 场景词的重叠个数。
        //
        // ★ 跨技能候选优先于同技能其他动作。
        //   初版给同技能动作固定 100 分以"优先推荐"，实测是个缺陷：技能动辄 6~10 个动作，
        //   一个 10 动作的技能会把所有名额用同技能的兄弟动作占满，而真正能解决问题的
        //   跨技能替代（如 acct-tools/bank-recon 失败后改用 bank-recon/recon）永远排不进来。
        //   同技能动作语义其实很近（都是同一领域），但**能力往往重叠**，换过去多半还是失败；
        //   跨技能替代才是"换条路"。故跨技能按重叠分排前，同技能动作只作为兜底补足名额。
        // 返回最多 limit 条，按得分降序；无候选时返回空表（调用方据此不改写错误文本）。
        public static List<SkillActionSpec> Alternatives(string root, string skillId, string action, int limit)
        {
            List<SkillActionSpec> result = new List<SkillActionSpec>();
            if (limit <= 0) return result;
            List<SkillActionSpec> all = Collect(root);

            string[] want = ScenarioWords(root, skillId);
            List<string> crossScored = new List<string>();   // 跨技能候选（按重叠分）
            List<string> sameScored = new List<string>();    // 同技能兄弟动作（兜底）
            foreach (SkillActionSpec sp in all)
            {
                if (sp.SkillId == skillId && sp.Action == action) continue;   // 排除自身
                if (sp.SkillId == skillId) { sameScored.Add(sp.Action); continue; }
                int score = 0;
                string[] have = ScenarioWords(root, sp.SkillId);
                foreach (string w in want)
                {
                    if (w.Length == 0) continue;
                    foreach (string h in have)
                    {
                        if (h.Length == 0) continue;
                        if (string.Equals(w, h, StringComparison.OrdinalIgnoreCase)) { score++; break; }
                    }
                }
                if (score <= 0) continue;
                // 分数字段补零到 4 位，保证字典序 = 数值序
                crossScored.Add(score.ToString("D4") + "\u0001" + sp.SkillId + "\u0001" + sp.Action);
            }
            crossScored.Sort();
            crossScored.Reverse();

            List<string> ordered = new List<string>();
            ordered.AddRange(crossScored);
            // 跨技能候选不足时才用同技能兄弟动作补足
            foreach (string a in sameScored)
            {
                if (ordered.Count >= limit) break;
                ordered.Add("0000\u0001" + skillId + "\u0001" + a);
            }

            foreach (string line in ordered)
            {
                if (result.Count >= limit) break;
                string[] parts = line.Split('\u0001');
                if (parts.Length < 3) continue;
                foreach (SkillActionSpec sp in all)
                {
                    if (sp.SkillId == parts[1] && sp.Action == parts[2]) { result.Add(sp); break; }
                }
            }
            return result;
        }

        // 场景词缓存（技能 id → scenarios 词数组）：换路是失败路径才走，但同一回合可能多次触发
        static Dictionary<string, string[]> scenarioCache = null;
        static string scenarioCacheRoot = null;

        static string[] ScenarioWords(string root, string skillId)
        {
            if (scenarioCache == null || scenarioCacheRoot != root)
            {
                scenarioCache = new Dictionary<string, string[]>();
                scenarioCacheRoot = root;
                foreach (SkillRegistryEntry e in SkillSystem.CachedScan(root))
                {
                    List<string> words = new List<string>();
                    if (e.Scenarios != null)
                    {
                        foreach (string w in e.Scenarios.Split(','))
                        {
                            string t = w.Trim();
                            if (t.Length > 0) words.Add(t);
                        }
                    }
                    scenarioCache[e.Id] = words.ToArray();
                }
            }
            string[] r;
            if (scenarioCache.TryGetValue(skillId, out r)) return r;
            return new string[0];
        }

        // 把失败结果改写成"带替代建议"的文本，供回灌给模型。
        // 只在确实找到候选时才追加建议——没有候选就原样返回，不给模型无效信息。
        public static string WithAlternatives(string root, string skillId, string action, string failureText, int limit)
        {
            List<SkillActionSpec> alts = Alternatives(root, skillId, action, limit);
            if (alts.Count == 0) return failureText;
            StringBuilder sb = new StringBuilder();
            sb.Append(failureText);
            sb.Append("\n可以改用以下替代工具重试：");
            for (int i = 0; i < alts.Count; i++)
            {
                SkillActionSpec sp = alts[i];
                sb.Append("\n  - ").Append(ToolName(sp.SkillId, sp.Action));
                if (sp.Desc.Length > 0) sb.Append("（").Append(sp.Desc).Append("）");
            }
            sb.Append("\n若替代方案也不适用，请直接说明失败原因，不要反复重试同一个工具。");
            return sb.ToString();
        }

        // 清空换路缓存（skill.json 变更时与动作缓存一起失效）
        public static void InvalidateAlternatives()
        {
            scenarioCache = null;
            scenarioCacheRoot = null;
        }

        // ---------- 反查与执行 ----------

        // 工具名 → 动作规格。找不到返回 null
        public static SkillActionSpec Resolve(string root, string toolName)
        {
            if (toolName == null || !toolName.StartsWith(Prefix, StringComparison.Ordinal)) return null;
            foreach (SkillActionSpec sp in Collect(root))
            {
                if (ToolName(sp.SkillId, sp.Action) == toolName) return sp;
            }
            return null;
        }

        // 执行一次技能调用。args 是模型给参数字典（已解析）。
        // 返回结果文本；out ok 表示成败。data 里的产物路径会追加进 products。
        public static string Invoke(string root, SkillActionSpec sp, Dictionary<string, string> args,
            List<string> products, out bool ok)
        {
            ok = false;
            if (sp == null) return "未找到对应的技能动作。";
            try
            {
                // 必填校验：缺参数直接返回结构化错误让模型补，而不是让 Python 端报晦涩错误
                List<string> missing = new List<string>();
                foreach (string[] p in sp.Params)
                {
                    if (p[2] != "required") continue;
                    string v;
                    if (args == null || !args.TryGetValue(p[0], out v) || v == null || v.Trim().Length == 0)
                        missing.Add(p[0]);
                }
                if (missing.Count > 0)
                {
                    StringBuilder m = new StringBuilder("缺少必填参数: ");
                    for (int i = 0; i < missing.Count; i++)
                    {
                        if (i > 0) m.Append("、");
                        m.Append(missing[i]);
                    }
                    m.Append("。请补齐后重新调用 ").Append(ToolName(sp.SkillId, sp.Action)).Append("。");
                    return m.ToString();
                }

                // 组装 request.json：**参数全部走文件**
                // 约定（与 skills/*/main.py 一致）：顶层 "task" + 顶层参数（不是嵌套 params）
                StringBuilder req = new StringBuilder();
                req.Append("{\"task\":\"").Append(MiniJson.Esc(sp.SkillId)).Append("\"");
                req.Append(",\"action\":\"").Append(MiniJson.Esc(sp.Action)).Append("\"");

                // 输入文件：file[] 类参数聚合成 inputs 数组（技能脚本从 inputs 取文件）
                List<string> inputs = new List<string>();
                foreach (string[] p in sp.Params)
                {
                    if (p[1] != "file[]") continue;
                    string v;
                    if (args == null || !args.TryGetValue(p[0], out v) || v == null) continue;
                    // 允许模型用 ';' 或 '|' 分隔多个文件
                    string[] parts = v.Split(new char[] { ';', '|' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (string one in parts)
                    {
                        string t = one.Trim();
                        if (t.Length > 0 && !inputs.Contains(t)) inputs.Add(t);
                    }
                }
                if (inputs.Count > 0)
                {
                    req.Append(",\"inputs\":[");
                    for (int i = 0; i < inputs.Count; i++)
                    {
                        if (i > 0) req.Append(",");
                        req.Append("\"").Append(MiniJson.Esc(inputs[i])).Append("\"");
                    }
                    req.Append("]");
                }

                // 其余参数按类型落顶层（number 不加引号，其余加引号）
                if (args != null)
                {
                    foreach (string[] p in sp.Params)
                    {
                        if (p[1] == "file[]") continue;   // 已进 inputs
                        string v;
                        if (!args.TryGetValue(p[0], out v) || v == null) continue;
                        v = v.Trim();
                        if (v.Length == 0) continue;
                        req.Append(",\"").Append(MiniJson.Esc(p[0])).Append("\":");
                        if (p[1] == "number" || p[1] == "bool") req.Append(v);
                        else req.Append("\"").Append(MiniJson.Esc(v)).Append("\"");
                    }
                }
                req.Append("}");

                SkillRunResult r = SkillRunner.Run(root, sp.SkillDir, req.ToString(), 120);
                ok = r.Ok;
                AuditLog.Record("skill_tool", sp.SkillId + "/" + sp.Action +
                    "; ok=" + (r.Ok ? "true" : "false") + "; ms=" + r.ElapsedMs);

                StringBuilder outText = new StringBuilder();
                outText.Append(r.Message);
                // 产物路径回传（若有）：让模型知道下一步可以引用哪个文件
                string prod = ExtractOutPath(r.DataJson);
                if (prod != null && prod.Length > 0 && products != null && !products.Contains(prod))
                    products.Add(prod);
                if (prod != null && prod.Length > 0)
                    outText.Append("\n产物文件: ").Append(prod);
                return outText.ToString();
            }
            catch (Exception ex)
            {
                ok = false;
                return "技能 " + sp.SkillId + "/" + sp.Action + " 执行失败：" +
                    (ex.Message == null ? ex.GetType().Name : ex.Message);
            }
        }

        // 从技能返回的 data JSON 里取产物路径（data.out / data.output / data.path）。
        // 用 MiniJson 扁平解析：技能回传的 data 是扁平对象，够用。
        static string ExtractOutPath(string dataJson)
        {
            if (dataJson == null || dataJson.Trim().Length == 0) return "";
            try
            {
                List<Dictionary<string, string>> objs = MiniJson.ParseObjects(dataJson);
                if (objs.Count == 0) return "";
                string[] keys = new string[] { "out", "output", "path", "file" };
                foreach (string k in keys)
                {
                    string v = MiniJson.Get(objs[0], k);
                    if (v != null && v.Trim().Length > 0 && File.Exists(v.Trim())) return v.Trim();
                }
            }
            catch { }
            return "";
        }
    }
}
