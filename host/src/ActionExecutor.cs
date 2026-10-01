// ActionExecutor —— 统一副作用执行适配器；只接受已确认且参数完整的 ActionPlan
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeAgent.Core;

namespace OfficeAgent.Host
{
    public class ActionExecutor
    {
        ConvertEngine converter;
        public ActionExecutor(ConvertEngine conv) { converter = conv; }

        // 执行一个已确认且参数完整的计划（与 ReconEngine.Run/MergeEngine.Run 同名的既有放行形态）
        public ActionExecutionResult Run(ActionPlan plan)
        {
            ActionExecutionResult result = new ActionExecutionResult();
            if (plan == null || !plan.IsExecutable)
            {
                result.State = ActionState.Failed;
                result.Message = plan == null ? "没有任务计划。" : (plan.Error.Length > 0 ? plan.Error : "计划参数不完整：" + string.Join("、", plan.Missing));
                return result;
            }
            // 输出路径要求：Skill 结果经对话返回，无输出文件（见 ActionPlan.IsExecutable）；
            // 其余动作缺输出路径无法执行。
            if (plan.Kind != ActionKind.Skill && plan.OutputPath.Length == 0)
            {
                result.State = ActionState.Failed; result.Message = "没有输出路径。"; return result;
            }
            try
            {
                plan.State = ActionState.Running;
                AuditLog.Record("action_started", plan.Kind.ToString() + "; output=" + (plan.OutputPath.Length > 0 ? plan.OutputPath : "(chat)"));
                if (plan.Kind == ActionKind.Recon) return ExecuteRecon(plan);
                if (plan.Kind == ActionKind.Merge) return ExecuteMerge(plan);
                if (plan.Kind == ActionKind.Invoice) return ExecuteInvoice(plan);
                if (plan.Kind == ActionKind.Skill) return RunSkill(plan);
                result.State = ActionState.Failed;
                result.Message = "未知动作类型。";
            }
            catch (Exception ex)
            {
                result.State = ActionState.Failed;
                result.Message = "执行异常: " + ex.Message;
                AuditLog.Record("action_finished", "failed; " + result.Message);
            }
            return result;
        }

        ActionExecutionResult ExecuteRecon(ActionPlan plan)
        {
            ActionExecutionResult r = new ActionExecutionResult();
            if (plan.Inputs.Count < 2 || plan.Recon == null)
            {
                r.State = ActionState.Failed; r.Message = "核对需要两个文件和完整列映射。"; return r;
            }
            ReconEngine e = new ReconEngine();
            ReconResult x = e.Run(plan.Recon, plan.Inputs[0], plan.Inputs[1], plan.OutputPath, converter);
            r.Recon = x; r.OutputPath = x.OutputPath ?? plan.OutputPath;
            r.ElapsedMs = x.ElapsedMs;
            r.State = x.Err != null || x.FailedChecks() > 0 ? ActionState.Failed : ActionState.Succeeded;
            r.Message = x.Err ?? ("匹配 " + x.MatchedCount + "，金额不等 " + x.PairDiffCount + "，仅A " + x.OnlyACount + "，仅B " + x.OnlyBCount);
            AuditLog.Record("action_finished", "recon; state=" + r.State + "; " + r.Message);
            return r;
        }

        ActionExecutionResult ExecuteMerge(ActionPlan plan)
        {
            ActionExecutionResult r = new ActionExecutionResult();
            if (plan.Inputs.Count == 0 || plan.Merge == null)
            {
                r.State = ActionState.Failed; r.Message = "汇总需要输入文件和模板。"; return r;
            }
            MergeEngine e = new MergeEngine();
            MergeResult x = e.Run(plan.Merge, plan.Inputs, plan.OutputPath, converter);
            r.Merge = x; r.OutputPath = x.OutputPath ?? plan.OutputPath;
            r.ElapsedMs = x.ElapsedMs;
            r.State = x.Err != null || x.FailedChecks() > 0 ? ActionState.Failed : ActionState.Succeeded;
            r.Message = x.Err ?? ("底稿 " + x.MergedRows + " 行");
            AuditLog.Record("action_finished", "merge; state=" + r.State + "; " + r.Message);
            return r;
        }

        // python38 技能经 SkillRunner（暂存→sidecar→stdout JSON）；结果在对话中返回，无输出文件
        ActionExecutionResult RunSkill(ActionPlan plan)
        {
            ActionExecutionResult r = new ActionExecutionResult();
            if (plan.Inputs.Count == 0 || plan.SkillId.Length == 0)
            {
                r.State = ActionState.Failed; r.Message = "技能执行需要输入文件与技能 id。"; return r;
            }
            string root = EnvDetect.FindRoot();
            SkillRegistryEntry entry = null;
            foreach (SkillRegistryEntry e in SkillSystem.CachedScan(root))
            {
                if (e.Id == plan.SkillId) { entry = e; break; }
            }
            if (entry == null || entry.State == SkillConst.StateDisabled)
            {
                r.State = ActionState.Failed; r.Message = "技能未注册或已禁用: " + plan.SkillId; return r;
            }
            StringBuilder req = new StringBuilder();
            req.Append("{\"task\":\"").Append(MiniJson.Esc(plan.SkillId)).Append("\",\"inputs\":[");
            for (int i = 0; i < plan.Inputs.Count; i++)
            {
                if (i > 0) req.Append(",");
                req.Append("\"").Append(MiniJson.Esc(plan.Inputs[i])).Append("\"");
            }
            req.Append("]}");
            SkillRunResult x = SkillRunner.Run(root, entry.Path, req.ToString(), 120);
            r.DataJson = x.DataJson;
            r.ElapsedMs = x.ElapsedMs;
            r.State = x.Ok ? ActionState.Succeeded : ActionState.Failed;
            r.Message = x.Message;
            AuditLog.Record("action_finished", "skill " + plan.SkillId + "; state=" + r.State + "; " + r.Message);
            return r;
        }

        ActionExecutionResult ExecuteInvoice(ActionPlan plan)
        {
            ActionExecutionResult r = new ActionExecutionResult();
            if (plan.Inputs.Count == 0)
            {
                r.State = ActionState.Failed; r.Message = "发票提取需要至少一个 PDF。"; return r;
            }
            InvoiceEngine e = new InvoiceEngine();
            InvoiceResult x = e.Run(plan.Inputs, plan.OutputPath);
            r.Invoice = x; r.OutputPath = x.OutputPath ?? plan.OutputPath;
            r.ElapsedMs = x.ElapsedMs;
            if (x.Err != null) r.State = ActionState.Failed;
            else if (x.Failures.Count > 0 || x.Checks.Exists(delegate(ReconCheck c) { return !c.Pass && !c.Warn; })) r.State = ActionState.Partial;
            else r.State = ActionState.Succeeded;
            r.Message = x.Err ?? ("成功 " + x.Rows.Count + " 票，失败 " + x.Failures.Count + " 票");
            AuditLog.Record("action_finished", "invoice; state=" + r.State + "; " + r.Message);
            return r;
        }
    }
}
