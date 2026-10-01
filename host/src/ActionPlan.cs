// ActionPlan —— 自然语言任务的确定性计划模型（M3 第一阶段）
// 只描述动作与参数，不执行文件操作、不调用 LLM。
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace OfficeAgent.Host
{
    public enum ActionKind { None, Recon, Merge, Invoice, Skill }
    public enum ActionState { Draft, AwaitingConfirmation, Running, Succeeded, Partial, Failed, Cancelled }

    public class ActionPlan
    {
        public ActionKind Kind = ActionKind.None;
        public ActionState State = ActionState.Draft;
        public List<string> Inputs = new List<string>();
        public string TemplateName = "";
        public string SkillId = "";              // Kind == Skill：技能 id（SkillSystem 注册表）
        public ReconMapping Recon = null;
        public MergeTemplate Merge = null;
        public string OutputPath = "";
        public string Summary = "";
        public string[] Missing = new string[0];
        public bool RequiresOverwriteConfirmation = false;
        public string Error = "";
        public long ElapsedMs;

        public bool IsExecutable
        {
            get
            {
                if (Kind == ActionKind.None || Error.Length > 0) return false;
                if (Missing != null && Missing.Length > 0) return false;
                // Skill 结果经对话返回，无输出文件要求；其余动作必须有输出路径
                if (Kind == ActionKind.Skill) return SkillId.Length > 0;
                return OutputPath.Length > 0;
            }
        }

        public string DisplayText()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(Summary.Length == 0 ? Kind.ToString() : Summary).Append("\n");
            sb.Append("输入：");
            for (int i = 0; i < Inputs.Count; i++)
            {
                if (i > 0) sb.Append("、");
                sb.Append(Path.GetFileName(Inputs[i]));
            }
            sb.Append("\n输出：").Append(OutputPath.Length == 0 ? "待确定" : OutputPath);
            if (TemplateName.Length > 0) sb.Append("\n模板：").Append(TemplateName);
            if (Recon != null) sb.Append("\n容差：±").Append(Recon.Tolerance.ToString("0.####", CultureInfo.InvariantCulture));
            if (Missing.Length > 0)
            {
                sb.Append("\n还缺：");
                for (int i = 0; i < Missing.Length; i++) { if (i > 0) sb.Append("、"); sb.Append(Missing[i]); }
            }
            if (RequiresOverwriteConfirmation) sb.Append("\n输出文件已存在，确认后覆盖。");
            return sb.ToString();
        }
    }

    public class ActionExecutionResult
    {
        public ActionState State;
        public string Message = "";
        public string OutputPath = "";
        public string DataJson = "";     // Skill 结果 JSON
        public long ElapsedMs;
        public ReconResult Recon;
        public MergeResult Merge;
        public InvoiceResult Invoice;
    }
}
