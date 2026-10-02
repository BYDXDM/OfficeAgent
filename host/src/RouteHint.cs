// RouteHint —— 把规则路由的命中结果变成给模型的**提示**（规划层第三期）
//
// 问题（第三期要解决的"准不准"，而非第二期的"能不能"）：
//   IntentRouter 是规则匹配，只认 scenarios 场景词的**字面包含**。
//   一旦用户换一种说法（"上个月的银行对账单帮着看看对不对得上"），
//   规则要么完全不命中，要么命中一个**语义相近但错误的**技能，而模型对此一无所知——
//  它只看到用户那句话，于是凭自己的判断挑工具，挑错时也没有任何信号提示它"这里可能是对账"。
//
// 方案：把规则的判断结果作为**提示**塞进用户载荷，而不是作为**指令**去替模型做决定。
//   命中 → "看起来你想做 X，对应技能是 skill_a_b"（帮模型少绕路）
//   未命中 → 不产生任何文本（不打扰）
//
// ★ 为什么是"提示"而不是"路由"：
//   第一期已把技能投影成模型可直接调用的工具，模型自己有能力选对。
//   规则的价值是**低成本先验**，不是**决策**。若把规则结果当成强制指令，
//   规则一错就全错（模型没有推翻的余地），这比让模型自己判断更糟。
//   故措辞一律用"可能/看起来/供参考"，并显式声明"以你的判断为准"。
//
// 安全：本类只做**字符串分析**，不碰文件、不执行任何技能、不调用模型。
//      提示文本里的路径来自用户消息本身（IntentRouter.ExtractPaths 只返回 File.Exists 为真的项），
//      不是外部注入面；且它只进 user 载荷，绝不参与命令行拼接。
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeAgent.Host
{
    public static class RouteHint
    {
        // 生成提示文本。无把握时返回空串（调用方据此不追加任何内容）。
        //
        // plan          : IntentRouter 的命中结果（Kind 为 None 时返回空串）
        // matchedSkill  : 规则命中的技能 id（Kind != Skill 时为空）
        // userText      : 用户原话（用于判断"命中得有多勉强"）
        public static string Build(ActionPlan plan, string userText)
        {
            if (plan == null || plan.Kind == ActionKind.None) return "";

            StringBuilder sb = new StringBuilder();
            sb.Append("\n\n[系统路由提示·仅供参考] ");
            sb.Append("从你的话里识别到可能的任务类型：");

            if (plan.Kind == ActionKind.Skill)
            {
                string skill = plan.SkillId == null ? "" : plan.SkillId.Trim();
                if (skill.Length == 0) return "";   // 技能 id 都没解析出来，提示没有价值
                sb.Append("技能「").Append(skill).Append("」的领域。");
                // 具名推荐：把该技能的动作名列出来（不给参数——参数看工具 schema）
                List<string> acts = ActionNames(skill);
                if (acts.Count > 0)
                {
                    sb.Append("该技能下有这些动作可用：");
                    for (int i = 0; i < acts.Count; i++)
                    {
                        if (i > 0) sb.Append("、");
                        sb.Append(acts[i]);
                    }
                    sb.Append("。");
                }
            }
            else
            {
                switch (plan.Kind)
                {
                    case ActionKind.Recon: sb.Append("两表核对（银行流水/日记账对账、勾稽）。"); break;
                    case ActionKind.Merge: sb.Append("多文件报表汇总（归集、合并报表）。"); break;
                    case ActionKind.Invoice: sb.Append("PDF 发票字段提取。"); break;
                    default: sb.Append(plan.Kind.ToString()).Append("。"); break;
                }
            }

            // 输入文件提示：规则只抽了**确实存在**的路径，这是模型自己列目录也要花一跳才能拿到的信息
            if (plan.Inputs != null && plan.Inputs.Count > 0)
            {
                sb.Append("这句话里提到的、确认存在的文件：");
                int n = plan.Inputs.Count > 5 ? 5 : plan.Inputs.Count;
                for (int i = 0; i < n; i++)
                {
                    if (i > 0) sb.Append("；");
                    sb.Append(plan.Inputs[i]);
                }
                if (plan.Inputs.Count > n) sb.Append("（还有 ").Append(plan.Inputs.Count - n).Append(" 个）");
                sb.Append("。");
            }

            // 缺参提示：规则知道"要做什么但还缺什么"，这对模型很有用——但同样只是提示
            if (plan.Missing != null && plan.Missing.Length > 0)
            {
                sb.Append("按规则判断，这个任务还缺：");
                for (int i = 0; i < plan.Missing.Length; i++)
                {
                    if (i > 0) sb.Append("、");
                    sb.Append(plan.Missing[i]);
                }
                sb.Append("。");
            }

            // 免责与优先级：明确告诉模型这是先验而非结论。
            // 这一段是"提示"与"指令"的分界线，不可省略——否则规则误判时模型会盲从。
            sb.Append("这**只是规则匹配的先验**，可能判错：如果与你对用户意图的理解不符，")
              .Append("以你自己的判断为准，按实际需要选工具。")
              .Append("不要向用户复述本提示，也不要因为本提示而跳过去确认文件是否存在。");
            if (plan.Kind == ActionKind.Skill)
            {
                sb.Append("若决定用该技能，直接调用对应的 skill_ 工具；若规则判错了，用别的工具即可。");
            }
            return sb.ToString();
        }

        // 某技能的动作名列表（按 skill.json 声明顺序）。取不到时返回空表。
        static List<string> ActionNames(string skillId)
        {
            List<string> names = new List<string>();
            try
            {
                foreach (SkillActionSpec sp in SkillToolBridge.Collect(OfficeAgent.Core.EnvDetect.FindRoot()))
                {
                    if (sp.SkillId == skillId) names.Add(sp.Action);
                }
            }
            catch { }
            return names;
        }
    }
}
