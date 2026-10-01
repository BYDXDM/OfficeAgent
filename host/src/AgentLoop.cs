// AgentLoop —— 会话模型的工具调用主循环（ChatPanel 与 CLI /agenttest 共用）
// 协议：OpenAI function calling；端点不支持 tools 时自动降级为纯对话重试一次。
// 工具白名单见 AgentTools（读文件/列目录/下载/转换/环境修复）；模型参数只作数据用。
using System;
using System.Collections.Generic;
using System.Text;
using OfficeAgent.Core;

namespace OfficeAgent.Host
{
    public static class AgentLoop
    {
        public class Result
        {
            public string FinalText = "";    // 最终答复（Error 非空时无意义）
            public string ToolLog = "";      // 工具调用记录（每行一条）
            public bool UsedTools;
            public string Error = "";        // 非空 = 请求失败
            public string FirstError = "";   // 降级前的原始错误（诊断用）
            public List<string> Products = new List<string>();   // 本回合产出的文件（附件卡片）
            public int Hops;                 // 实际轮数
        }

        const int MaxHops = 10;   // 常规工具轮次上限；到达后还有一次"收尾跳"强制汇总

        // onToolLog：每执行完一个工具回调一次（累计日志）；可为 null（CLI 场景）。
        // history：既往对话（role/content，可为 null）——**必须传**，否则模型每轮都是"第一次对话"。
        // onPlanTool：task_plan 工具的 UI 回调（ChatPanel 计划栏）；null 时给通用文本应答
        public static Result Run(LlmClient client, AppConfig config, string sysPrompt, string userPayload,
            Action<string> onToolLog, Func<string, string> onPlanTool,
            System.Collections.Generic.IEnumerable<LlmTurn> history)
        {
            Result r = new Result();
            string tools = AgentTools.SchemasJson(config != null && config.PluginPlan);
            List<string> products = new List<string>();
            ConvertEngine conv = new ConvertEngine();
            conv.SofficePath = ConvertEngine.FindSoffice(EnvDetect.FindRoot());

            List<LlmTurn> request = new List<LlmTurn>();
            if (history != null)
            {
                foreach (LlmTurn t in history) request.Add(t);
            }
            request.Add(new LlmTurn("user", userPayload));
            StringBuilder log = new StringBuilder();

            try
            {
                // 循环上限是 MaxHops+1：第 MaxHops+1 轮是收尾跳（强制模型汇总，不再执行工具）
                for (int hop = 0; hop < MaxHops + 1; hop++)
                {
                    bool wrapUp = hop == MaxHops;
                    if (wrapUp)
                    {
                        // 收尾跳：以用户身份要求汇总。tools 保持声明（剥掉会因消息里已有
                        // tool 结果被服务端拒绝），靠指令约束模型不再调工具。
                        request.Add(new LlmTurn("user",
                            "工具调用轮次已达上限。请基于以上工具的执行结果，直接给出最终答复：" +
                            "任务完成了什么、产物保存在哪里、还有哪些没完成及原因。不要调用任何工具。"));
                    }
                    r.Hops = hop + 1;
                    LlmReply reply = client.ChatRaw(LlmClient.BuildMessagesJson(sysPrompt, request), tools, true);
                    // 仅在"第一跳且尚未执行任何工具"时允许降级：一旦消息里已有工具结果，
                    // 剥掉 tools 重发会让请求自相矛盾（tool 消息引用未声明的函数）→ 服务端必拒
                    if (reply.Error.Length > 0 && hop == 0 && tools != null && LooksLikeToolsRejected(reply.Error))
                    {
                        r.FirstError = reply.Error;   // 保留原始错误供诊断（降级后不可见）
                        tools = null;   // 端点不支持 function calling → 降级纯对话重试一次
                        reply = client.ChatRaw(LlmClient.BuildMessagesJson(sysPrompt, request), null, true);
                    }
                    if (reply.Error.Length > 0) { r.Error = reply.Error; return r; }
                    if (!reply.HasToolCalls) { r.FinalText = reply.Content; r.Products = products; return r; }
                    if (wrapUp)
                    {
                        // 收尾跳模型仍要调工具：不再迁就，把已执行步骤如实汇报
                        r.FinalText = "这个任务在限定的 " + MaxHops + " 轮工具调用内没有完成，最后一步没执行。" +
                            "已执行的步骤：\n" + log.ToString() +
                            "\n\n建议：把任务说得更具体（文件放哪、输出叫什么），或拆成小步重试；已生成的产物不受影响。";
                        r.Products = products;
                        return r;
                    }

                    // 回显 assistant 工具调用消息 → 执行 → 回传结果
                    request.Add(new LlmTurn("assistant_raw", reply.RawMessageJson));
                    foreach (string[] call in reply.ToolCalls)
                    {
                        r.UsedTools = true;
                        bool ok = true;
                        string result;
                        if (call[1] == "task_plan")
                        {
                            // 计划栏插件：有 UI 回调走界面；无界面（CLI）记审计并给通用应答
                            if (onPlanTool != null) result = onPlanTool(call[2]);
                            else { AuditLog.Record("task_plan", (call[2] ?? "").Replace("\"", "")); result = "计划已记录（当前界面不展示计划栏）"; }
                        }
                        else
                        {
                            result = AgentTools.Dispatch(call[1], call[2], config, conv, products, out ok);
                        }
                        request.Add(new LlmTurn("tool@" + call[0], result));
                        if (log.Length > 0) log.Append("\n");
                        log.Append("🔧 ").Append(call[1]).Append(ok ? " ✓" : " ✗");
                        AuditLog.Record("tool_call", call[1] + "; ok=" + ok);
                        if (onToolLog != null)
                        {
                            try { onToolLog(log.ToString()); } catch { }
                        }
                    }
                }
                return r;   // 理论不可达：收尾跳必返回（答复或汇总）
            }
            catch (Exception ex)
            {
                r.Error = "Agent 循环异常: " + ex.Message;
                return r;
            }
        }

        static bool LooksLikeToolsRejected(string err)
        {
            string e = (err ?? "").ToLowerInvariant();
            return e.Contains("http 400") || e.Contains("http 422") || e.Contains("tools") ||
                   (e.Contains("tool") && e.Contains("not")) || e.Contains("未知参数") || e.Contains("unknown");
        }
    }
}
