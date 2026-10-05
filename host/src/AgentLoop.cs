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
            public bool ThinkingStripped;    // 最终答复里剥掉过模型的思考段（诊断用）
        }

        const int MaxHops = 12;   // 常规工具轮次上限；到达后还有一次"收尾跳"强制汇总
                                  // (10→12：真实任务里"列目录→确认→转换→建表"常需 6~9 跳，
                                  //  留出余量避免第 10 跳就收尾。过量试探由重复调用检测兜底。)
                                  // ★ 第二期起这不再是"主"判据：加权预算（HopBudget）会先于它触收尾。
                                  //   保留它是**硬兜底**——万一预算计数出 bug（成本算成 0），
                                  //   这里仍能保证循环终止。

        // 真实跳数的硬上限：加权预算按等效值算，读型半价意味着实际跳数可能超过 MaxHops。
        // 该值 = 等效上限 / 最小单位成本，向上取整后留余量，纯粹防死循环。
        const int HardHopLimit = 48;

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
            // 技能工具（规划层第一期）：把 skills 注册表投影成额外的 function-calling 工具，
            // 让模型能直接调用会计/Excel/Word 技能，而不是只会用 CSV 文本糊表。
            // 有动作才拼接：空串时保持原样，避免产生 "[...,]" 这种非法 JSON。
            string skillTools = AgentTools.SkillSchemasJson(EnvDetect.FindRoot());
            if (skillTools != null && skillTools.Length > 0 && tools != null && tools.Length > 0)
                tools = tools.Substring(0, tools.Length - 1) + "," + skillTools + "]";
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
                // 回合级产物登记（第二期）：每次工具产出文件后登记，并在结果里附上清单，
                // 让后续步骤能引用"产物N"而不是从上下文里抄路径。
                ArtifactRegistry.Reset();
                // 加权预算（第二期）：读型半价、技能/写型全价、task_plan 免费。
                // 达到等效上限即进入收尾跳；HardHopLimit 是防死循环的硬兜底。
                HopBudget.State budget = new HopBudget.State();
                bool pseudoRetried = false;   // 伪工具调用只纠正一次，避免死循环
                for (int hop = 0; hop < HardHopLimit; hop++)
                {
                    bool wrapUp = HopBudget.ShouldWrapUp(budget) || hop >= MaxHops;
                    if (wrapUp && hop > 0)
                    {
                        // 收尾跳：以用户身份要求汇总。tools 保持声明（剥掉会因消息里已有
                        // tool 结果被服务端拒绝），靠指令约束模型不再调工具。
                        request.Add(new LlmTurn("user",
                            "工具调用轮次已达上限（" + HopBudget.Describe(budget) + "）。" +
                            "请基于以上工具的执行结果，直接给出最终答复：" +
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
                        // ★ 反"伪工具调用"（0.9.7，修用户实测的"产物丢失"）
                        //   现象：模型把工具调用**写成代码块**（```python / analyze_attendance）
                        //   而不是真正发起 function call，然后凭空续写结果——答复里给了完整
                        //   表格和保存路径，但审计日志无写文件记录、目标目录为空，文件从未生成。
                        //   系统提示已要求"一律直接发起工具调用"，但约束不住，所以这里做一次
                        //   **强制纠正重试**：识别到伪调用就把原答复回灌 + 明确要求改用真正的
                        //   function call 重来，让循环继续，而不是就此收尾。
                        if (!pseudoRetried && LooksLikePseudoToolCall(reply.Content))
                        {
                            pseudoRetried = true;
                            r.UsedTools = true;
                            request.Add(new LlmTurn("assistant", reply.Content == null ? "" : reply.Content));
                            request.Add(new LlmTurn("user",
                                "你刚才把工具调用写成了**代码块**（例如 python 代码块里写 analyze_attendance），" +
                                "那不是真正的调用，系统不会执行它，你后面写的结果是凭空编的。\n" +
                                "请立刻改用**真正的工具调用**重新执行这个任务：不要输出代码示例、不要描述步骤、" +
                                "不要凭记忆编造结果。如果确实不需要工具，就直接给出基于已知事实的答复。"));
                            log.AppendLine("· 检测到伪工具调用（写成代码块），已要求改用真正的 function call 重试");
                            continue;
                        }
                        // 这一跳就是最终答复（模型不再要工具）。
                        // 流式策略：只有当本回合**已经跑过工具**（长任务，用户等待久、逐字显示有实感）
                        // 且配置开启时才复查一次 stream:true。
                        // 理由：流式复现要**再发一次请求**，而单轮问答是最常见场景，
                        // 让它每次多花一倍 token/时间不划算（用户明确对成本敏感）。
                        bool wantStream = onDelta != null && r.UsedTools &&
                                          (config == null || config.StreamFinal) &&
                                          reply.Content != null && reply.Content.Length > 0;
                        if (wantStream)
                        {
                            string serr;
                            string sbody = client.ChatFinalStream(
                                LlmClient.BuildMessagesJson(sysPrompt, request), onDelta, out serr);
                            // 流式那次是独立采样：为空说明这次没拿到内容，回退到已得的文本。
                            r.FinalText = (sbody != null && sbody.Length > 0) ? sbody : reply.Content;
                        }
                        else r.FinalText = reply.Content;
                        // 剥掉模型写在 content 里的思考段（部分端点如 glm-4.5-air 会以
                        // </think> 收尾并重写正式答复；原样展示会先给用户一段自相矛盾的草稿）。
                        bool stripped;
                        r.FinalText = ModelText.Clean(r.FinalText, out stripped);
                        if (stripped) r.ThinkingStripped = true;
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
                        // 三态：正常执行 / 主动跳过（重复调用）/ 失败。
                        // 跳过不能算失败——否则日志渲染成 ✗、审计写 ok=False，语义就错了。
                        bool skipped = false;
                        string result;
                        // 加权计费：读到工具名就记账（跳过/失败也计——它们同样消耗了一次模型轮次，
                        // 而且失败的写型调用往往还要重试，把它计费能更早触收尾，避免反复空转）
                        budget.Spent += HopBudget.Cost(call[1], budget);
                        if (call[1] == "task_plan")
                        {
                            // 计划栏插件：有 UI 回调走界面；无界面（CLI）记审计并给通用应答
                            if (onPlanTool != null) result = onPlanTool(call[2]);
                            else { AuditLog.Record("task_plan", (call[2] ?? "").Replace("\"", "")); result = "计划已记录（当前界面不展示计划栏）"; }
                        }
                        else
                        {
                            // 重复调用检测：同一工具+同一参数第二次出现时的处理。
                            // 分隔符用 \u0001（不可见控制字符）而非 "|"：后者不可单射——
                            // ("a","b|c") 与 ("a|b","c") 会拼出同一个 "a|b|c" 而误判为重复调用。
                            // 模型的 argsJson 里确实可能出现 "|"（如 create_presentation 的 outline）。
                            //
                            // ★ 只对"写型"工具拒绝重复，读型工具允许重放：
                            //   读型（list_directory/read_text_file）的结果**会随时间变化**——
                            //   同一回合里先列目录、再建表/转换/下载，然后重列同一目录是完全正当的
                            //   （要确认新产物落盘了）。若无条件跳过，模型会拿到"你以为你列过了"的
                            //   假信息，反而错过刚生成的文件。
                            //   写型（create_*/convert_*/download_*）重放才会真的产生副作用或重复劳动，
                            //   这才是该拦的对象。
                            string sig = (call[1] ?? "") + "\u0001" + (call[2] ?? "");
                            int seenTimes = 0;
                            if (callCount.ContainsKey(sig)) { seenTimes = callCount[sig]; callCount[sig] = seenTimes + 1; }
                            else { callCount[sig] = 1; }
                            if (seenTimes >= 1 && IsMutatingTool(call[1]))
                            {
                                skipped = true;
                                result = "你已经用完全相同的参数调用过 " + call[1] +
                                    " 了，结果在上面，不要重复调用。" +
                                    "请换一种做法（换参数/换路径/换工具），或者如果信息已经够了就直接给出最终答复。";
                            }
                            else
                            {
                                result = AgentTools.Dispatch(call[1], call[2], config, conv, products, out ok);
                                // 失败换路（第二期）：技能调用失败时，按 scenarios 场景词重叠
                                // 找出语义相近的替代工具一并回灌，给模型一次换路机会。
                                // 只对技能做——内置通用工具的失败多半是路径/权限问题，换个工具没用。
                                if (!ok && call[1] != null && call[1].StartsWith(SkillToolBridge.Prefix, StringComparison.Ordinal))
                                {
                                    SkillActionSpec failedSpec = SkillToolBridge.Resolve(EnvDetect.FindRoot(), call[1]);
                                    if (failedSpec != null)
                                    {
                                        result = SkillToolBridge.WithAlternatives(
                                            EnvDetect.FindRoot(), failedSpec.SkillId, failedSpec.Action, result, 2);
                                    }
                                }
                                result = TruncToolResult(result);
                                // 读型工具重放时明确告知：这次是重新读取，结果可能与上面不同
                                if (seenTimes >= 1)
                                {
                                    result = "（注意：这是对同一目标的重新读取，内容可能与上面那次不同）\n" + result;
                                }
                            }
                        }
                        // 产物登记：本跳若产出新文件（AgentTools 把路径追加进 products），
                        // 登记成别名并把清单附在结果后面，供后续步骤引用。
                        if (products.Count > 0)
                        {
                            string newest = products[products.Count - 1];
                            string alias = ArtifactRegistry.Add(newest);
                            if (alias.Length > 0)
                            {
                                result += "\n（本文件已登记为 " + alias + "）";
                            }
                            string listing = ArtifactRegistry.Describe();
                            if (listing.Length > 0) result += "\n" + listing;
                        }
                        request.Add(new LlmTurn("tool@" + call[0], result));
                        if (log.Length > 0) log.Append("\n");
                        log.Append("🔧 ").Append(call[1]).Append(skipped ? " ↷跳过" : (ok ? " ✓" : " ✗"));
                        // 审计值用小写 true/false（bool.ToString() 会给 True/False，
                        // 与既有审计记录约定不一致，也会让 grep "ok=false" 漏掉）。
                        AuditLog.Record("tool_call", call[1] + "; " +
                            (skipped ? "skipped=duplicate" : ("ok=" + (ok ? "true" : "false"))));
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

        // 工具是否"写型"（有副作用/产出文件）。
        // 写型工具重放会真的重复干活或覆盖产物，应当拦；读型工具重放是正当的
        //（目标内容可能在上一次调用之后变了），只提示不拦。
        static bool IsMutatingTool(string name)
        {
            if (name == null) return false;
            // 技能工具一律视为写型：技能可能产出文件/写工作表，重放有副作用，
            // 无法从工具名可靠区分读写（同一技能的不同 action 语义不同），故保守拦截。
            if (name.StartsWith(SkillToolBridge.Prefix, StringComparison.Ordinal)) return true;
            // 与 AgentTools 的白名单一一对应；新增工具时同步这里，否则默认按"读型"放行（更安全）
            return name == "create_spreadsheet" || name == "create_presentation"
                || name == "create_formula_workbook" || name == "convert_document" || name == "download_file"
                || name == "delete_file" || name == "repair_environment";
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

        // 判定"伪工具调用"：答复里出现**代码块**，且块内出现已知工具名。
        // 刻意只看代码块内部——正常答复里提到工具名（如"我用 create_formula_workbook 建的"）
        // 不该被误判。返回 true 表示模型像是"想调工具却写成了示例代码"。
        internal static bool LooksLikePseudoToolCall(string content)
        {
            if (content == null || content.Length == 0) return false;
            List<string> tools = AgentTools.KnownToolNames();
            if (tools.Count == 0) return false;
            int i = 0;
            while (true)
            {
                int s = content.IndexOf("```", i, StringComparison.Ordinal);
                if (s < 0) return false;
                int e = content.IndexOf("```", s + 3, StringComparison.Ordinal);
                string block = e > s ? content.Substring(s + 3, e - s - 3) : content.Substring(s + 3);
                foreach (string t in tools)
                {
                    if (t != null && t.Length > 2 && block.IndexOf(t, StringComparison.OrdinalIgnoreCase) >= 0) return true;
                }
                if (e < 0) return false;
                i = e + 3;
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
