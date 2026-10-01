// IntentRouter —— 规则优先的自然语言动作识别（不调用模型、不执行副作用）
// 输入文件来源：消息文本中的路径 token（正则抽取，确定性）+ 会话最近文件上下文。
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using OfficeAgent.Core;

namespace OfficeAgent.Host
{
    public static class IntentRouter
    {
        static readonly Regex PathToken = new Regex(
            "[A-Za-z]:\\\\[^\\s\"'<>|?*！？（）【】「」《》()\\[\\]:]+");

        public static bool TryCreate(string text, IList<string> contextFiles, out ActionPlan plan)
        {
            plan = new ActionPlan();
            string s = text == null ? "" : text.Trim();
            if (s.Length == 0) return false;
            string skillId;
            ActionKind kind = DetectKind(s, out skillId);
            if (kind == ActionKind.None) return false;
            plan.Kind = kind;
            plan.State = ActionState.AwaitingConfirmation;
            if (kind == ActionKind.Skill) plan.SkillId = skillId == null ? "" : skillId;
            foreach (string f in ExtractPaths(s))
            {
                if (!plan.Inputs.Contains(f)) plan.Inputs.Add(f);
            }
            foreach (string f in contextFiles)
            {
                if (f != null && f.Length > 0 && File.Exists(f) && !plan.Inputs.Contains(f)) plan.Inputs.Add(f);
            }
            if (plan.Inputs.Count > 50) plan.Inputs.RemoveRange(50, plan.Inputs.Count - 50);
            if (kind == ActionKind.Recon) BuildRecon(s, plan);
            else if (kind == ActionKind.Merge) BuildMerge(s, plan);
            else if (kind == ActionKind.Invoice) BuildInvoice(s, plan);
            else BuildSkill(plan);
            plan.RequiresOverwriteConfirmation = plan.OutputPath.Length > 0 && File.Exists(plan.OutputPath);
            return true;
        }

        // 从文本抽取存在的文件路径：盘符开头，吃掉除空白与常用标点外的连续字符（允许中文路径）
        public static List<string> ExtractPaths(string text)
        {
            List<string> found = new List<string>();
            if (text == null || text.Length == 0) return found;
            foreach (Match m in PathToken.Matches(text))
            {
                // 句读尾巴（中英文句号/逗号/顿号/分号）从 token 尾部剥掉；路径内部的点号保留
                string p = m.Value.TrimEnd('.', '，', '。', '；', '、', ',', ';');
                try
                {
                    if (p.Length > 3 && File.Exists(p) && !found.Contains(p)) found.Add(p);
                }
                catch { }
            }
            return found;
        }

        // 返回动作类型；命中 python 技能时经 skillId 带出其 id（替代此前的 [ThreadStatic] 偷传值）
        static ActionKind DetectKind(string s, out string skillId)
        {
            skillId = "";
            if (HasAny(s, new string[] { "核对", "对账", "勾稽", "流水" })) return ActionKind.Recon;
            if (HasAny(s, new string[] { "汇总", "归集", "合并报表" })) return ActionKind.Merge;
            if (HasAny(s, new string[] { "发票", "提取票", "识别票" })) return ActionKind.Invoice;
            // 技能场景词（skill.json scenarios，缓存扫描）：
            //   builtin 技能 → entry 映射引擎任务；python38 技能 → Skill（经 SkillRunner 执行）
            foreach (SkillRegistryEntry e in SkillSystem.CachedScan(EnvDetect.FindRoot()))
            {
                if (e.State == SkillConst.StateDisabled) continue;
                ActionKind k = ActionKind.None;
                if (e.Runtime == SkillConst.RuntimeBuiltin) k = KindOfEntry(e.Entry);
                else if (e.Runtime == SkillConst.RuntimePython38) k = ActionKind.Skill;
                if (k == ActionKind.None) continue;
                foreach (string sc in e.Scenarios.Split(','))
                {
                    string w = sc.Trim();
                    if (w.Length > 0 && s.IndexOf(w, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        if (k == ActionKind.Skill) skillId = e.Id;
                        return k;
                    }
                }
            }
            return ActionKind.None;
        }

        static ActionKind KindOfEntry(string entry)
        {
            if (entry == "recon") return ActionKind.Recon;
            if (entry == "merge") return ActionKind.Merge;
            if (entry == "invoices") return ActionKind.Invoice;
            return ActionKind.None;
        }

        static void BuildRecon(string s, ActionPlan p)
        {
            p.Summary = "准备进行两表核对";
            p.Recon = new ReconMapping();
            p.Recon.Tolerance = 0.01;
            p.OutputPath = DefaultOutput(p.Inputs, "核对");
            List<string> missing = new List<string>();
            if (p.Inputs.Count < 2) missing.Add("两个输入文件");
            if (Contains(s, "容差")) p.Summary += "（已识别核对意图；容差需在详细参数页确认）";
            else p.Summary += "（默认容差 ±0.01）";
            p.Missing = missing.ToArray();
        }

        static void BuildMerge(string s, ActionPlan p)
        {
            p.Summary = "准备进行多文件报表汇总";
            p.OutputPath = DefaultOutput(p.Inputs, "汇总底稿");
            p.Missing = p.Inputs.Count < 2 ? new string[] { "至少两个输入文件" } : new string[] { "汇总模板或目标列映射" };
        }

        static void BuildInvoice(string s, ActionPlan p)
        {
            p.Summary = "准备提取 PDF 发票字段";
            p.OutputPath = DefaultOutput(p.Inputs, "发票清单");
            p.Missing = p.Inputs.Count == 0 ? new string[] { "至少一个 PDF 文件" } : new string[0];
        }

        static void BuildSkill(ActionPlan p)
        {
            p.Summary = "准备执行技能「" + p.SkillId + "」";
            p.OutputPath = "";
            p.Missing = p.Inputs.Count == 0 ? new string[] { "至少一个输入文件" } : new string[0];
        }

        static string DefaultOutput(List<string> files, string prefix)
        {
            if (files == null || files.Count == 0) return "";
            string dir = Path.GetDirectoryName(Path.GetFullPath(files[0]));
            return Path.Combine(dir, prefix + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");
        }

        static bool HasAny(string s, string[] words)
        {
            foreach (string w in words) if (s.IndexOf(w, StringComparison.OrdinalIgnoreCase) >= 0) return true;
            return false;
        }
        static bool Contains(string s, string w) { return s.IndexOf(w, StringComparison.OrdinalIgnoreCase) >= 0; }
    }
}
