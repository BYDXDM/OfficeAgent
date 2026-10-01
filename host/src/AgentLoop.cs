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

        const int MaxHops = 12;   // 常规工具轮次上限；到达后还有一次"收尾跳"强制汇总
                                  // (10→12：真实任务里"列目录→确认→转换→建表"常需 6~9 跳，
                                  //  留出余量避免第 10 跳就收尾。过量试探由重复调用检测兜底。)

        // 工具结果回传截断上限：模型只看到前 N 字符，尾部提示可用工具读更多。
        // 目的：单次 read_text_file/list_directory 的大结果会把上下文吃满，
        // 导致后续轮次模型"看不清重点"而反复试探。
        const int MaxToolResultChars = 12000;

        // onToolLog：每执行完一个工具回调一次（累计日志）；可为 null（CLI 场景）。
        // history：既往对话（role/content，可为 null）——**必须传**，否则模型每轮都是"第一次对话"。
        // onPlanTool：task_plan 工具的 UI 回调（ChatPanel 计划栏）；null 时给通用文本应答
        // onDelta：最终答复的流式回调（累计文本）；null = 不用流式（CLI 场景）
        public static Result Run(LlmClient client, AppConfig config, string sysPrompt, string userPayload,
            Action<string> onToolLog, Func<string, string> onPlanTool,
            System.Collections.Generic.IEnumerable<LlmTurn> history)
        {
            return Run(client, config, sysPrompt, userPayload, onToolLog, onPlanTool, history, null);
        }

        public static Result Run(LlmClient client, AppConfig config, string sysPrompt, string userPayload,
            Action<string> onToolLog, Func<string, string> onPlanTool,
            System.Collections.Generic.IEnumerable<LlmTurn> history, Action<string> onDelta)
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
                // 重复调用台账：记录"工具名+参数"出现次数，用于打断无效试探循环
                Dictionary<string, int> callCount = new Dictionary<string, int>();
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
                    if (!reply.HasToolCalls)
                    {
                        // 这一跳就是最终答复（模型不再要工具）。
                        // 若开了流式：用同样的 messages 再发一次 stream:true，让用户看到文字逐段出现。
                        // 代价：仅在最后一跳多一次请求；失败则直接用已拿到的文本，不影响正确性。
                        if (onDelta != null && reply.Content != null && reply.Content.Length > 0)
                        {
                            string serr;
                            string sbody = client.ChatFinalStream(
                                LlmClient.BuildMessagesJson(sysPrompt, request), onDelta, out serr);
                            r.FinalText = (sbody != null && sbody.Length > 0) ? sbody : reply.Content;
                        }
                        else r.FinalText = reply.Content;
                        r.Products = products;
                        return r;
                    }
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
                            // 重复调用检测：同一工具+同一参数第二次出现时，不再重复执行，
                            // 直接把"你刚做过同样的调用"告诉模型，逼它换策略或收尾。
                            // （只读类工具重复无副作用但也无意义；写类工具重复执行更危险。）
                            string sig = (call[1] ?? "") + "|" + (call[2] ?? "");
                            int seenTimes = 0;
                            if (callCount.ContainsKey(sig)) { seenTimes = callCount[sig]; callCount[sig] = seenTimes + 1; }
                            else { callCount[sig] = 1; }
                            if (seenTimes >= 1)
                            {
                                ok = false;
                                result = "你已经用完全相同的参数调用过 " + call[1] +
                                    " 了，结果在上面，不要重复调用。" +
                                    "请换一种做法（换参数/换路径/换工具），或者如果信息已经够了就直接给出最终答复。";
                            }
                            else
                            {
                                result = AgentTools.Dispatch(call[1], call[2], config, conv, products, out ok);
                                result = TruncToolResult(result);
                            }
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

        // 工具结果截断：超长结果只回传前 MaxToolResultChars 字符。
        // 保留头部（通常含路径/表头等关键信息），尾部说明如何取更多。
        static string TruncToolResult(string result)
        {
            if (result == null) return "";
            if (result.Length <= MaxToolResultChars) return result;
            return result.Substring(0, MaxToolResultChars) +
                "\n…[结果过长已截断，共 " + result.Length + " 字符]。" +
                "如需其余内容，请用更精确的参数（如指定具体文件/子目录）重新调用，不要重复读同一份。";
        }

        static bool LooksLikeToolsRejected(string err)
        {
            string e = (err ?? "").ToLowerInvariant();
            return e.Contains("http 400") || e.Contains("http 422") || e.Contains("tools") ||
                   (e.Contains("tool") && e.Contains("not")) || e.Contains("未知参数") || e.Contains("unknown");
        }
    }
}
