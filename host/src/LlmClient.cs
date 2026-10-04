// LlmClient —— OpenAI 兼容端点客户端（/models 拉取 + /chat/completions 对话 + 工具调用）
// 出网统一走 HostGuard 校验；不自动跟随重定向（手动复检）；TLS 1.2 显式启用。
// Agent 能力：messages 数组（多轮历史）、function tools（AgentTools 白名单）、
// 智谱 web_search 联网工具（仅 bigmodel/zhipu 端点启用）、max_tokens 4096（1024 会截断长回复）。
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;

namespace OfficeAgent.Host
{
    // 一条待发送消息；Role 为 "assistant_raw" 时 Content 是原始 assistant 消息对象 JSON
    // （工具调用回传需要原样回显）；Role 形如 "tool@<call_id>" 表示工具结果。
    public class LlmTurn
    {
        public string Role;
        public string Content;
        public LlmTurn(string role, string content) { Role = role; Content = content; }
    }

    // 一次模型回复
    public class LlmReply
    {
        public string Content = "";                  // 助手文本（可能为空）
        public bool HasToolCalls;                    // 是否请求调用工具
        public List<string[]> ToolCalls;             // 每项 [id, name, argumentsJson]
        public string RawMessageJson = "";           // 原始 message 对象（回显用）
        public string Error = "";                    // 非空 = 请求失败
    }

    public class LlmClient
    {
        public string BaseUrl;
        public string ApiKey;
        public string Model;
        public bool AllowLan;
        public bool EnableWebSearch = false;   // 联网搜索（仅智谱系端点生效，见 WebSearchToolsJson）

        public LlmClient(AppConfig c)
        {
            BaseUrl = NormalizeBase(c == null ? "" : c.BaseUrl);
            ApiKey = c == null ? null : c.GetKey();
            Model = c == null ? "" : c.Model;
            AllowLan = c != null && c.AllowLan;
            EnableWebSearch = c != null && c.WebSearch;
        }

        // 预置常用模型（/models 拉取失败时也有得选；设置页与会话页内联下拉共用）。
        // 2026-10 用户定制：日常只用这五个（其余模型走「从服务刷新模型列表」或手填）。
        public static readonly string[] ModelPresets = new string[] {
            "deepseek-flash",          // DeepSeek-V4.1-Flash（官方 id=deepseek-flash，1M 上下文）
            "glm-4.5-air",
            "glm-4.1v-thinking-flash", // GLM-4.1V-Thinking 免费版（bigmodel 实际 id 带 -flash 后缀）
            "gpt-5.6-luna",
            "gpt-6-luna" };

        // 预置模型所属的官方 OpenAI 兼容端点。内联下拉跨供应商切换模型时自动带出地址：
        // 密钥按服务（host）分开保存（见 AppConfig），从根上杜绝"拿 A 家的密钥向 B 家发请求"。
        // 返回 null = 未知/自建网关模型，只切模型名不动地址。
        public static string BaseUrlForModel(string model)
        {
            string m = (model ?? "").Trim().ToLowerInvariant();
            if (m.Length == 0) return null;
            if (m.StartsWith("glm-", StringComparison.Ordinal)) return "https://open.bigmodel.cn/api/paas/v4";
            if (m.StartsWith("deepseek-", StringComparison.Ordinal)) return "https://api.deepseek.com";
            if (m.StartsWith("qwen-", StringComparison.Ordinal)) return "https://dashscope.aliyuncs.com/compatible-mode/v1";
            if (m.StartsWith("gpt-", StringComparison.Ordinal)) return "https://api.openai.com/v1";
            return null;
        }

        static string NormalizeBase(string url)
        {
            if (url == null) return "";
            url = url.Trim();
            // 用户习惯照抄官方文档的完整端点（…/chat/completions、…/models）——剥掉后缀，
            // BaseUrl 只保留到 /v1、/v4 这一层，避免二次拼接出 /chat/completions/chat/completions → 404
            string lower = url.ToLowerInvariant();
            string[] suffixes = new string[] { "/chat/completions", "/models", "/completions", "/embeddings" };
            bool changed = true;
            while (changed)
            {
                changed = false;
                foreach (string sfx in suffixes)
                {
                    if (lower.EndsWith(sfx, StringComparison.Ordinal))
                    {
                        url = url.Substring(0, url.Length - sfx.Length);
                        lower = lower.Substring(0, lower.Length - sfx.Length);
                        while (lower.EndsWith("/")) { url = url.Substring(0, url.Length - 1); lower = lower.Substring(0, lower.Length - 1); }
                        changed = true;
                    }
                }
            }
            return url;
        }

        // 智谱系端点才注入 web_search 工具（OpenAI 等端点会 400 拒绝未知工具类型）
        public bool EndpointSupportsWebSearch()
        {
            string b = (BaseUrl ?? "").ToLowerInvariant();
            return b.Contains("bigmodel.cn") || b.Contains("zhipu");
        }

        static string WebSearchToolsJson()
        {
            return "[{\"type\":\"web_search\",\"web_search\":{\"enable\":true}}]";
        }

        static HttpWebRequest Request(string url, string method, string apiKey, int timeoutMs)
        {
            try { ServicePointManager.SecurityProtocol |= (SecurityProtocolType)3072; } catch { } // Tls12
            try { ServicePointManager.Expect100Continue = false; } catch { }   // 防 POST 大 body 与朴素服务端互相等待死锁
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            req.Method = method;
            req.Timeout = timeoutMs;
            req.ReadWriteTimeout = timeoutMs;
            req.AllowAutoRedirect = false;   // 重定向手动复检
            req.KeepAlive = false;           // 一请求一连接，配合服务端 Connection: close
            req.UserAgent = "OfficeAgent/0.4";
            if (apiKey != null && apiKey.Length > 0)
                req.Headers[HttpRequestHeader.Authorization] = "Bearer " + apiKey;
            return req;
        }

        // 手动跟随重定向：每一跳重新过 HostGuard（防 302 跳内网）
        string Send(string url, string method, string body, int timeoutMs, out int status)
        {
            status = 0;
            string current = url;
            for (int hop = 0; hop < 3; hop++)
            {
                string guardErr = HostGuard.Check(current, AllowLan);
                if (guardErr != null) throw new InvalidOperationException("安全守卫：" + guardErr);
                HttpWebRequest req = Request(current, method, ApiKey, timeoutMs);
                if (body != null)
                {
                    byte[] data = Encoding.UTF8.GetBytes(body);
                    req.ContentType = "application/json";
                    req.ContentLength = data.Length;
                    using (Stream s = req.GetRequestStream()) s.Write(data, 0, data.Length);
                }
                try
                {
                    using (HttpWebResponse resp = (HttpWebResponse)req.GetResponse())
                    {
                        status = (int)resp.StatusCode;
                        using (StreamReader sr = new StreamReader(resp.GetResponseStream(), Encoding.UTF8))
                            return sr.ReadToEnd();
                    }
                }
                catch (WebException we)
                {
                    HttpWebResponse resp = we.Response as HttpWebResponse;
                    if (resp != null && (int)resp.StatusCode >= 300 && (int)resp.StatusCode < 400)
                    {
                        string loc = resp.Headers[HttpResponseHeader.Location];
                        if (loc == null) throw new InvalidOperationException("重定向缺少 Location");
                        current = new Uri(new Uri(current), loc).ToString();
                        status = (int)resp.StatusCode;
                        continue;
                    }
                    if (resp != null)
                    {
                        using (StreamReader sr = new StreamReader(resp.GetResponseStream(), Encoding.UTF8))
                        {
                            string detail = sr.ReadToEnd();
                            throw new InvalidOperationException("HTTP " + (int)resp.StatusCode + ": " + TrimMsg(detail));
                        }
                    }
                    throw new InvalidOperationException("网络错误: " + we.Message);
                }
            }
            throw new InvalidOperationException("重定向次数过多");
        }

        static string TrimMsg(string s)
        {
            if (s == null) return "";
            s = s.Trim();
            if (s.Length > 300) s = s.Substring(0, 300) + "…";
            return s.Replace("\n", " ");
        }

        // 拉取模型列表（GET /models），返回模型 id 列表
        public List<string> ListModels(out string err)
        {
            err = null;
            try
            {
                int status;
                string json = Send(BaseUrl + "/models", "GET", null, 30000, out status);
                List<string> ids = JsonFindStrings(json, "id");
                if (ids.Count == 0) { err = "响应中未找到模型列表（确认是 OpenAI 兼容端点）"; return null; }
                return ids;
            }
            catch (InvalidOperationException ex)
            {
                if (ex.Message.StartsWith("HTTP 404", StringComparison.Ordinal))
                    err = "该服务不支持模型列表接口（HTTP 404）。不影响使用：可直接在模型框输入名称，或从预置下拉选择。";
                else err = ex.Message;
                return null;
            }
            catch (Exception ex) { err = ex.Message; return null; }
        }

        // ---------- messages 构造 ----------

        // system + 历史 + 消息对列表 → messages 数组 JSON。
        // "assistant_raw" 原样回显；"tool@<id>" 转 {"role":"tool","tool_call_id":..,"content":..}
        public static string BuildMessagesJson(string systemPrompt, List<LlmTurn> turns)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("[");
            bool first = true;
            if (systemPrompt != null && systemPrompt.Length > 0)
            {
                sb.Append("{\"role\":\"system\",\"content\":\"").Append(JsonEscape(systemPrompt)).Append("\"}");
                first = false;
            }
            foreach (LlmTurn t in turns)
            {
                if (!first) sb.Append(",");
                first = false;
                if (t.Role == "assistant_raw")
                {
                    sb.Append(t.Content == null ? "{}" : t.Content);
                }
                else if (t.Role != null && t.Role.StartsWith("tool@", StringComparison.Ordinal))
                {
                    sb.Append("{\"role\":\"tool\",\"tool_call_id\":\"").Append(JsonEscape(t.Role.Substring(5)))
                      .Append("\",\"content\":\"").Append(JsonEscape(t.Content == null ? "" : t.Content)).Append("\"}");
                }
                else
                {
                    sb.Append("{\"role\":\"").Append(JsonEscape(t.Role)).Append("\",\"content\":\"")
                      .Append(JsonEscape(t.Content == null ? "" : t.Content)).Append("\"}");
                }
            }
            sb.Append("]");
            return sb.ToString();
        }

        // ---------- 核心请求（非流式，支持工具） ----------

        // 最近一次发出的请求体/响应体（诊断用；/agenttest 打印路径）
        public static string LastRequestBody = "";
        public static string LastResponseBody = "";

        // toolsJson：function 工具数组 JSON（null=不传）；includeWebSearch=true 时（智谱系）附加 web_search。
        public LlmReply ChatRaw(string messagesJson, string toolsJson, bool includeWebSearch)
        {
            LlmReply r = new LlmReply();
            try
            {
                if (Model == null || Model.Length == 0) throw new InvalidOperationException("未选择模型");
                StringBuilder body = new StringBuilder();
                body.Append("{\"model\":\"").Append(JsonEscape(Model)).Append("\",\"messages\":").Append(messagesJson);
                bool wantSearch = includeWebSearch && EndpointSupportsWebSearch();
                if (toolsJson != null && wantSearch)
                {
                    // web_search 并入 tools 数组（智谱官方形态）；不再发顶层错形参数（会干扰工具调用）
                    string all = toolsJson.TrimEnd(']') + "," + WebSearchToolsJson().Substring(1);
                    body.Append(",\"tools\":").Append(all).Append(",\"tool_choice\":\"auto\"");
                }
                else if (toolsJson != null)
                    body.Append(",\"tools\":").Append(toolsJson).Append(",\"tool_choice\":\"auto\"");
                else if (wantSearch)
                    body.Append(",\"tools\":").Append(WebSearchToolsJson());
                body.Append(",\"temperature\":0.3,\"max_tokens\":4096,\"stream\":false}");
                LastRequestBody = body.ToString();
                int status;
                string json = Send(BaseUrl + "/chat/completions", "POST", body.ToString(), 180000, out status);
                // 定位 choices[0].message：先找 "message" token，再从其后第一个 '{' 提取括号配平对象。
                // （此前直接把 token 位置当对象起点 → 提取失败 → 整个响应被回显 → 服务端 1214 角色为空）
                int mi = json.IndexOf("\"message\"", StringComparison.Ordinal);
                LastResponseBody = json;
                string msg = mi >= 0 ? ExtractJsonObject(json, json.IndexOf('{', mi)) : null;
                if (msg == null) throw new InvalidOperationException("响应中未找到 message 对象");
                r.Content = JsonGetString(msg, "content");
                if (r.Content == null) r.Content = "";
                // 工具调用解析（只认 function 类型；web_search 等服务端自执行类型不本地分发）
                ParseToolCallsInto(msg, r);
                // 重建干净的 assistant 消息用于回显：只保留 role/content/tool_calls，
                // 剥掉 reasoning_content（glm-4.5 推理模型会拒绝或浪费上下文）与 index 等服务端字段
                StringBuilder asst = new StringBuilder();
                asst.Append("{\"role\":\"assistant\",\"content\":\"").Append(JsonEscape(r.Content)).Append("\"");
                if (r.HasToolCalls)
                {
                    asst.Append(",\"tool_calls\":[");
                    for (int i = 0; i < r.ToolCalls.Count; i++)
                    {
                        if (i > 0) asst.Append(",");
                        // arguments 必须回填为 JSON 字符串（内容里的路径反斜杠需重新转义，否则请求体非法）
                        asst.Append("{\"id\":\"").Append(JsonEscape(r.ToolCalls[i][0]))
                          .Append("\",\"type\":\"function\",\"function\":{\"name\":\"").Append(JsonEscape(r.ToolCalls[i][1]))
                          .Append("\",\"arguments\":\"").Append(JsonEscape(r.ToolCalls[i][2])).Append("\"}}");
                    }
                    asst.Append("]");
                }
                asst.Append("}");
                r.RawMessageJson = asst.ToString();
                if (r.Content.Length == 0 && !r.HasToolCalls)
                {
                    // 联网搜索等场景模型可能返回空文本（结果在服务端已注入下一轮）——不当错误，交由上层提示
                    r.Content = "（模型本轮没有返回文字内容，可能联网搜索无结果。请换个问法，或到设置里关闭「联网搜索」重试。）";
                }
                return r;
            }
            catch (Exception ex)
            {
                r.Error = ex.Message;
                return r;
            }
        }

        // 工具调用解析（internal 供 /toolidtest 无头回归）。
        // 为什么不再"从每个 function token 往回找最近的 id"：并行工具调用时，部分模型/网关
        // 会把 id 写在 function 之后、干脆缺 id，甚至两个调用复用同一个 id——回溯找法会抓错
        // （串到别的字段）或产出重复 id。这种请求原样回显上去，严格校验的端点（DeepSeek 等）
        // 直接拒绝：HTTP 400 "Duplicate value for 'tool_call_id' of call_00_xxx in message[N]"
        // ——用户侧的表现就是"一调用插件/技能就报错"。
        // 这里把 tool_calls 数组的每个条目括号配平成独立对象，id/name/arguments 全部取条目内部，
        // 最后做 id 归一化：缺失 → 补唯一值；重复 → 重写为唯一值。
        // 回显消息（RawMessageJson）与 tool@id 消息都从归一化后的 ToolCalls 构建，请求必然自洽。
        internal static void ParseToolCallsInto(string msg, LlmReply r)
        {
            int tc = msg.IndexOf("\"tool_calls\"", StringComparison.Ordinal);
            if (tc < 0) return;
            int arr = msg.IndexOf('[', tc);
            if (arr < 0) return;
            r.ToolCalls = new List<string[]>();
            List<string> seenIds = new List<string>();
            int i = arr;
            while (true)
            {
                int close = msg.IndexOf(']', i + 1);
                int b = msg.IndexOf('{', i + 1);
                if (b < 0 || (close >= 0 && close < b)) break;   // 数组结束
                string entry = ExtractJsonObject(msg, b);
                if (entry == null) break;
                i = b + entry.Length - 1;   // 连同嵌套一起跳过本条目
                string id = JsonGetString(entry, "id");
                string name = null, args = null;
                int fi = entry.IndexOf("\"function\"", StringComparison.Ordinal);
                if (fi >= 0)
                {
                    string fobj = ExtractJsonObject(entry, entry.IndexOf('{', fi));
                    if (fobj != null)
                    {
                        name = JsonGetString(fobj, "name");
                        args = JsonGetString(fobj, "arguments");
                    }
                }
                if (name == null || name.Length == 0) continue;   // 非 function 类型（web_search 等）不本地分发
                if (args == null) args = "{}";
                if (id == null || id.Length == 0 || seenIds.Contains(id))
                    id = "call_oa_" + Guid.NewGuid().ToString("N").Substring(0, 16);
                seenIds.Add(id);
                r.ToolCalls.Add(new string[] { id, name, args });
            }
            r.HasToolCalls = r.ToolCalls.Count > 0;
        }

        // 从 startIdx 处的 '{' 起提取括号配平的 JSON 对象（跳过字符串字面量）
        static string ExtractJsonObject(string json, int braceStart)
        {
            if (json == null || braceStart < 0 || braceStart >= json.Length || json[braceStart] != '{') return null;
            int depth = 0; bool inStr = false; bool esc = false;
            for (int i = braceStart; i < json.Length; i++)
            {
                char c = json[i];
                if (inStr)
                {
                    if (esc) esc = false;
                    else if (c == '\\') esc = true;
                    else if (c == '"') inStr = false;
                    continue;
                }
                if (c == '"') inStr = true;
                else if (c == '{') depth++;
                else if (c == '}')
                {
                    depth--;
                    if (depth == 0) return json.Substring(braceStart, i - braceStart + 1);
                }
            }
            return null;
        }

        // 最终答复流式：把已构造好的 messages JSON 发出去，stream:true 逐段回调累计文本。
        // 用于 agent 循环「模型不再调工具、要给出最终答复」的那一跳——工具轮次仍非流式。
        // 返回：累计文本；出错或端点不支持流式时返回 null 并置 err（调用方回退到非流式 ChatRaw）。
        // 注意：本方法不发 tools，纯文本答复，因此不会出现 tool_calls 分片拼接问题。
        public string ChatFinalStream(string messagesJson, Action<string> onDelta, out string err)
        {
            err = null;
            try
            {
                if (Model == null || Model.Length == 0) { err = "未选择模型"; return null; }
                StringBuilder body = new StringBuilder();
                body.Append("{\"model\":\"").Append(JsonEscape(Model)).Append("\",\"messages\":").Append(messagesJson);
                body.Append(",\"temperature\":0.3,\"max_tokens\":4096,\"stream\":true}");
                LastRequestBody = body.ToString();

                string url = BaseUrl + "/chat/completions";
                string guardErr = HostGuard.Check(url, AllowLan);
                if (guardErr != null) { err = "安全守卫：" + guardErr; return null; }

                StringBuilder acc = new StringBuilder();
                HttpWebRequest req = Request(url, "POST", ApiKey, 180000);
                byte[] data = Encoding.UTF8.GetBytes(body.ToString());
                req.ContentType = "application/json";
                req.ContentLength = data.Length;
                using (Stream s = req.GetRequestStream()) s.Write(data, 0, data.Length);
                using (HttpWebResponse resp = (HttpWebResponse)req.GetResponse())
                {
                    if ((int)resp.StatusCode >= 300) { err = "HTTP " + (int)resp.StatusCode; return null; }
                    using (StreamReader sr = new StreamReader(resp.GetResponseStream(), Encoding.UTF8))
                    {
                        string line;
                        while ((line = sr.ReadLine()) != null)
                        {
                            if (!line.StartsWith("data:")) continue;
                            string payload = line.Substring(5).Trim();
                            if (payload.Length == 0) continue;
                            if (payload == "[DONE]") break;
                            string delta = JsonGetString(payload, "content");
                            if (delta != null && delta.Length > 0)
                            {
                                acc.Append(delta);
                                if (onDelta != null) { try { onDelta(acc.ToString()); } catch { } }
                            }
                            else if (acc.Length == 0)
                            {
                                // 有的网关把错误塞在流里的 message 字段
                                string emsg = JsonGetString(payload, "message");
                                if (emsg != null && emsg.Length > 0) { err = "服务端返回: " + TrimMsg(emsg); return null; }
                            }
                        }
                    }
                }
                if (acc.Length == 0) { err = "流式响应为空"; return null; }   // 调用方回退非流式
                return acc.ToString();
            }
            catch (Exception ex)
            {
                err = ex.Message;
                return null;
            }
        }

        // 兼容旧调用（连通性测试、列映射建议）：单轮、不带工具
        public string Chat(string userText, string systemPrompt, out string err)
        {
            List<LlmTurn> turns = new List<LlmTurn>();
            turns.Add(new LlmTurn("user", userText));
            LlmReply r = ChatRaw(BuildMessagesJson(systemPrompt, turns), null, false);
            err = r.Error;
            return r.Error.Length > 0 ? null : r.Content;
        }

        // 流式对话（SSE）：逐段回调 onDelta(累计文本)；服务端不支持流式时自动回退一次性请求。
        // 说明：本方法只用于「最终答复」这一跳——带 tools 的工具轮次一律走非流式 ChatRaw，
        // 因为流式分片里的 tool_calls 需要跨片拼接，收益低而风险高。
        public string ChatStream(string userText, string systemPrompt, Action<string> onDelta, out string err)
        {
            err = null;
            try
            {
                if (Model == null || Model.Length == 0) throw new InvalidOperationException("未选择模型");
                StringBuilder body = new StringBuilder();
                body.Append("{\"model\":\"").Append(JsonEscape(Model)).Append("\",\"messages\":[");
                if (systemPrompt != null && systemPrompt.Length > 0)
                    body.Append("{\"role\":\"system\",\"content\":\"").Append(JsonEscape(systemPrompt)).Append("\"},");
                body.Append("{\"role\":\"user\",\"content\":\"").Append(JsonEscape(userText)).Append("\"}],");
                body.Append("\"temperature\":0.3,\"max_tokens\":4096,\"stream\":true}");

                string url = BaseUrl + "/chat/completions";
                string guardErr = HostGuard.Check(url, AllowLan);
                if (guardErr != null) throw new InvalidOperationException("安全守卫：" + guardErr);

                try
                {
                    StringBuilder acc = new StringBuilder();
                    HttpWebRequest req = Request(url, "POST", ApiKey, 180000);
                    byte[] data = Encoding.UTF8.GetBytes(body.ToString());
                    req.ContentType = "application/json";
                    req.ContentLength = data.Length;
                    using (Stream s = req.GetRequestStream()) s.Write(data, 0, data.Length);
                    using (HttpWebResponse resp = (HttpWebResponse)req.GetResponse())
                    {
                        if ((int)resp.StatusCode >= 300) throw new InvalidOperationException("HTTP " + (int)resp.StatusCode);
                        using (StreamReader sr = new StreamReader(resp.GetResponseStream(), Encoding.UTF8))
                        {
                            string line;
                            while ((line = sr.ReadLine()) != null)
                            {
                                if (line.StartsWith("data:"))
                                {
                                    string payload = line.Substring(5).Trim();
                                    if (payload == "[DONE]") break;
                                    if (payload.Length == 0) continue;
                                    string delta = JsonGetString(payload, "content");
                                    if (delta != null && delta.Length > 0)
                                    {
                                        acc.Append(delta);
                                        if (onDelta != null) { try { onDelta(acc.ToString()); } catch { } }
                                    }
                                    string emsg = JsonGetString(payload, "message");
                                    if (emsg != null && acc.Length == 0) throw new InvalidOperationException("服务端返回: " + TrimMsg(emsg));
                                }
                            }
                        }
                    }
                    if (acc.Length == 0)
                    {
                        // 服务端忽略 stream:true（200 返回整段 JSON）→ 回退一次性请求
                        if (onDelta != null) { try { onDelta("…改为整段接收"); } catch { } }
                        return Chat(userText, systemPrompt, out err);
                    }
                    return acc.ToString();
                }
                catch (WebException)
                {
                    // 流式不可用（老网关/非兼容端点）→ 回退一次性请求
                    if (onDelta != null) { try { onDelta("…改为整段接收"); } catch { } }
                    return Chat(userText, systemPrompt, out err);
                }
            }
            catch (Exception ex) { err = ex.Message; return null; }
        }

        static string JsonEscape(string s)
        {
            StringBuilder sb = new StringBuilder();
            foreach (char ch in s)
            {
                switch (ch)
                {
                    case '\\': sb.Append("\\\\"); break;
                    case '"': sb.Append("\\\""); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default: sb.Append(ch); break;
                }
            }
            return sb.ToString();
        }

        // 极简 JSON 字符串值提取（扁平场景足够：找 "key":"value"，从后往前取嵌套里最后一个）
        static string JsonGetString(string json, string key)
        {
            if (json == null) return null;
            string pat = "\"" + key + "\":";
            int i = json.LastIndexOf(pat, StringComparison.Ordinal);
            if (i < 0) i = json.IndexOf(pat, StringComparison.Ordinal);
            if (i < 0) return null;
            i += pat.Length;
            while (i < json.Length && json[i] == ' ') i++;
            if (i >= json.Length || json[i] != '"') return null;
            i++;
            StringBuilder sb = new StringBuilder();
            while (i < json.Length && json[i] != '"')
            {
                if (json[i] == '\\' && i + 1 < json.Length)
                {
                    i++;
                    char c = json[i];
                    if (c == 'n') sb.Append('\n');
                    else if (c == 'r') sb.Append('\r');
                    else if (c == 't') sb.Append('\t');
                    else if (c == 'u' && i + 4 < json.Length)
                    {
                        int code;
                        if (int.TryParse(json.Substring(i + 1, 4), System.Globalization.NumberStyles.HexNumber,
                            System.Globalization.CultureInfo.InvariantCulture, out code))
                            sb.Append((char)code);
                        i += 4;
                    }
                    else sb.Append(c);
                }
                else sb.Append(json[i]);
                i++;
            }
            return sb.ToString();
        }

        static List<string> JsonFindStrings(string json, string key)
        {
            List<string> result = new List<string>();
            if (json == null) return result;
            string pat = "\"" + key + "\":\"";
            int i = 0;
            while ((i = json.IndexOf(pat, i, StringComparison.Ordinal)) >= 0)
            {
                i += pat.Length;
                StringBuilder sb = new StringBuilder();
                while (i < json.Length && json[i] != '"')
                {
                    if (json[i] == '\\' && i + 1 < json.Length) { i++; sb.Append(json[i]); }
                    else sb.Append(json[i]);
                    i++;
                }
                string v = sb.ToString();
                if (v.Length > 0 && !result.Contains(v)) result.Add(v);
                i++;
            }
            return result;
        }
    }
}
