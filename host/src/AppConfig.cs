// AppConfig —— 本地配置持久化（%LOCALAPPDATA%\OfficeAgent\config.json）
// 密钥安全规约（Mimosa 约束）：API Key 只经 DPAPI（Windows 凭据保护服务）加密后落盘，
// 支持环境变量 OFFICEAGENT_API_KEY 运行时覆盖；源码/示例/测试不写任何凭据字面量。
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace OfficeAgent.Host
{
    // 一家服务（host）的密钥：每家分开 DPAPI 加密存储，
    // 切换供应商时各取各的，杜绝"拿 A 家的密钥向 B 家发请求"。
    // Tag 非空 = 同一服务的第 N 个账号（如 api.deepseek.com + "2" = DeepSeek 账号2），
    // 落盘为 "key:<host>#<tag>"，与主账号互不覆盖。
    public class ProviderKey
    {
        public string Host = "";     // 如 api.deepseek.com、open.bigmodel.cn
        public string Tag = "";      // 账号槽；空 = 主账号
        public byte[] Blob = null;   // DPAPI 密文
        // 该密钥槽的**用户自定义名称**（0.9.0）：同一供应商配多个 Key 时用户可各自命名
        //（如"公司号""备用号"），不再显示固定的"号1/号2"。空 = 未命名，界面回退到"号<tag>"。
        public string Label = "";
    }

    public class AppConfig
    {
        public string BaseUrl = "";
        public string Model = "";
        public bool AllowLan = false;      // 企业内网网关显式放行（默认关闭）
        public byte[] KeyBlob = null;      // 旧版单密钥密文（仅作加载迁移用，不再写入）
        public List<ProviderKey> ProviderKeys = new List<ProviderKey>();   // 按服务分开存的密钥
        public bool WizardDone = false;
        // 内置插件开关（默认全开，开箱即用）
        public bool PluginMemory = true;   // 记忆模块
        public bool PluginPref = true;     // 转换偏好
        public bool PluginHist = true;     // 会话历史
        public bool PluginPlan = true;     // 任务计划栏（复杂任务自动列计划，右侧展示）
        public int ConvTargetIndex = 0;    // 转换偏好：上次目标格式（0=PDF 1=CSV 2=XLSX）
        public int PrivacyLevel = 1;       // 隐私分级（设计方案 §7.2）：0=全本地 1=脱敏出网(默认) 2=全量
        public bool WebSearch = true;      // 联网搜索工具（智谱系端点 web_search；关=纯本地对话）
        // 最终答复流式输出。开启时：模型给出答复的那一跳会再发一次 stream:true 请求来逐字显示，
        // 代价是该跳多一次请求（token 与时间成本翻一倍）。纯问答场景收益有限，
        // 因此仅在"已经执行过工具"的长任务里才流式（见 AgentLoop），单轮问答自动跳过。
        public bool StreamFinal = true;
        public string WorkspaceDir = "";   // 默认工作区目录：空=文档\OfficeAgentFiles（侧边栏可为项目单独设工作区）
        public string ActiveWorkspaceId = "";   // 当前激活的工作区 id（对应 workspaces.json）
        // 安全兜底（v0.8.2）：写/删「系统盘文件」或删除「表格文件」前必须用户确认。
        // SafetyGuard = 总开关（默认开）；SafetyAllowPaths = 用户点"始终允许此目录"后
        // 累积的目录前缀（分号分隔），命中即不再询问（可在设置页清除）。
        public bool SafetyGuard = true;
        public string SafetyAllowPaths = "";
        // 用户添加/用过的非预置模型名（';' 连接落盘）。内联模型下拉 = 本次拉取 + customModels + 预置；
        // 此前拉取列表只存内存，重启后新添加的模型就从下拉里消失了（2026-10-05 用户反馈）。
        public string CustomModels = "";

        // 实际生效的工作区目录（空配置回退到默认并确保存在）
        public string EffectiveWorkspace()
        {
            string dir = (WorkspaceDir == null ? "" : WorkspaceDir.Trim());
            if (dir.Length == 0)
                dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "OfficeAgentFiles");
            try { Directory.CreateDirectory(dir); } catch { }
            return dir;
        }

        public static string Dir()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeAgent");
        }

        public static string FilePath()
        {
            return Path.Combine(Dir(), "config.json");
        }

        public static bool Exists()
        {
            try { return File.Exists(FilePath()); } catch { return false; }
        }

        // 从服务地址提取主机名（密钥的归属维度）：
        // https://api.deepseek.com/v1 → api.deepseek.com。解析失败兜底手工截取，统一小写。
        public static string HostOf(string baseUrl)
        {
            string u = (baseUrl ?? "").Trim();
            if (u.Length == 0) return "";
            try
            {
                string h = new Uri(u).Host;
                if (h != null && h.Length > 0) return h.ToLowerInvariant();
            }
            catch { }
            string s = u;
            int p = s.IndexOf("://", StringComparison.Ordinal);
            if (p >= 0) s = s.Substring(p + 3);
            int slash = s.IndexOf('/');
            if (slash >= 0) s = s.Substring(0, slash);
            int at = s.LastIndexOf('@');
            if (at >= 0) s = s.Substring(at + 1);
            return s.ToLowerInvariant();
        }

        public bool IsLlmConfigured()
        {
            return BaseUrl != null && BaseUrl.Length > 0 && GetKey() != null && GetKey().Length > 0;
        }

        public void SetKey(string plain)
        {
            SetKeyFor(HostOf(BaseUrl), plain, "");
        }

        // 指定账号槽存密钥（tag 空 = 主账号）。归属到 host；每家每号一把钥匙互不覆盖。
        public void SetKeyFor(string host, string plain, string tag)
        {
            // 归属到当前 BaseUrl 的主机；每家服务一把钥匙互不覆盖
            if (host == null) host = "";
            if (tag == null) tag = "";
            int idx = -1;
            for (int i = 0; i < ProviderKeys.Count; i++)
            {
                if (string.Equals(ProviderKeys[i].Host, host, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(ProviderKeys[i].Tag ?? "", tag, StringComparison.Ordinal)) { idx = i; break; }
            }
            if (string.IsNullOrEmpty(plain))
            {
                if (idx >= 0) ProviderKeys.RemoveAt(idx);   // 显式清空该服务的密钥
                return;
            }
            byte[] blob = DpapiProtect(plain);
            if (idx >= 0) ProviderKeys[idx].Blob = blob;
            else
            {
                ProviderKey pk = new ProviderKey();
                pk.Host = host;
                pk.Tag = tag;
                pk.Blob = blob;
                ProviderKeys.Add(pk);
            }
        }

        public string GetKey()
        {
            return GetKeyFor("");
        }

        // 按账号槽取当前服务的密钥（tag 来自模型名后缀，如 deepseek-flash#2 → "2"）。
        public string GetKeyFor(string tag)
        {
            // 环境变量覆盖优先（安全规约允许的凭据来源之二）
            try
            {
                string env = Environment.GetEnvironmentVariable("OFFICEAGENT_API_KEY");
                if (!string.IsNullOrEmpty(env)) return env;
            }
            catch { }
            return KeyForHost(HostOf(BaseUrl), tag == null ? "" : tag);
        }

        // 取指定服务主机的密钥（DPAPI 解密；无则 null）。tag 空 = 主账号槽。
        // ★ 不做跨服务回退：找不到就返回 null——回退会退化回"一个 key 向各家发请求"。
        public string KeyForHost(string host) { return KeyForHost(host, ""); }

        public string KeyForHost(string host, string tag)
        {
            if (host == null || host.Length == 0) return null;
            if (tag == null) tag = "";
            for (int i = 0; i < ProviderKeys.Count; i++)
            {
                if (string.Equals(ProviderKeys[i].Host, host, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(ProviderKeys[i].Tag ?? "", tag, StringComparison.Ordinal)
                    && ProviderKeys[i].Blob != null && ProviderKeys[i].Blob.Length > 0)
                    return DpapiUnprotect(ProviderKeys[i].Blob);
            }
            return null;
        }

        // 只查存在性不解密（占位符显示、跨供应商切换提示用）
        public bool HasKeyForHost(string host) { return HasKeyForHost(host, ""); }

        // 密钥槽的用户自定义名称；未命名返回 ""（调用方回退到"号<tag>"或主机名）
        public string GetKeyLabel(string host, string tag)
        {
            if (host == null || host.Length == 0) return "";
            if (tag == null) tag = "";
            for (int i = 0; i < ProviderKeys.Count; i++)
            {
                if (string.Equals(ProviderKeys[i].Host, host, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(ProviderKeys[i].Tag ?? "", tag, StringComparison.Ordinal))
                    return ProviderKeys[i].Label == null ? "" : ProviderKeys[i].Label;
            }
            return "";
        }

        // 设置密钥槽名称。槽不存在时**只记内存**（下次 SetKey 建槽时会带上）——
        // 不为了一个名字凭空建出没有密钥的槽。
        public void SetKeyLabel(string host, string tag, string label)
        {
            if (host == null || host.Length == 0) return;
            if (tag == null) tag = "";
            string v = label == null ? "" : label.Trim();
            if (v.Length > 24) v = v.Substring(0, 24);
            for (int i = 0; i < ProviderKeys.Count; i++)
            {
                if (string.Equals(ProviderKeys[i].Host, host, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(ProviderKeys[i].Tag ?? "", tag, StringComparison.Ordinal))
                {
                    ProviderKeys[i].Label = v;
                    return;
                }
            }
        }

        public bool HasKeyForHost(string host, string tag)
        {
            if (host == null || host.Length == 0) return false;
            if (tag == null) tag = "";
            for (int i = 0; i < ProviderKeys.Count; i++)
            {
                if (string.Equals(ProviderKeys[i].Host, host, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(ProviderKeys[i].Tag ?? "", tag, StringComparison.Ordinal)
                    && ProviderKeys[i].Blob != null && ProviderKeys[i].Blob.Length > 0)
                    return true;
            }
            return false;
        }

        // 记住用户添加/用过的非预置模型（内联下拉重启后仍可见）；预置与带 #tag 的账号条目不重复收
        public void AddCustomModel(string model)
        {
            string m = (model ?? "").Trim();
            if (m.Length == 0 || m.Length > 80) return;
            foreach (string p in LlmClient.ModelPresets) if (p == m) return;
            List<string> items = CustomModelList();
            if (items.Contains(m)) { items.Remove(m); items.Insert(0, m); }   // 最近用的排前
            else items.Insert(0, m);
            while (items.Count > 12) items.RemoveAt(items.Count - 1);
            CustomModels = string.Join(";", items.ToArray());
        }

        public List<string> CustomModelList()
        {
            List<string> items = new List<string>();
            string[] parts = (CustomModels ?? "").Split(';');
            for (int i = 0; i < parts.Length; i++)
            {
                string p = parts[i].Trim();
                if (p.Length > 0 && !items.Contains(p)) items.Add(p);
            }
            return items;
        }

        public void Save()
        {
            try
            {
                Directory.CreateDirectory(Dir());
                File.WriteAllText(FilePath(), ToJson(), new UTF8Encoding(false));
            }
            catch { }
        }

        // 序列化为 JSON 文本。抽成独立方法是为了**可测**：/menutest 直接断言产出里
        // 含 "keylabel:…"（验证密钥槽名称确实会落盘），而不必真去写用户配置文件。
        public string ToJson()
        {
            StringBuilder sb = new StringBuilder();
            try
            {
                sb.Append("{\n");
                sb.Append("  \"schema\": 1,\n");
                sb.Append("  \"baseUrl\": \"").Append(Js(BaseUrl)).Append("\",\n");
                sb.Append("  \"model\": \"").Append(Js(Model)).Append("\",\n");
                sb.Append("  \"allowLan\": ").Append(AllowLan ? "true" : "false").Append(",\n");
                sb.Append("  \"wizardDone\": ").Append(WizardDone ? "true" : "false").Append(",\n");
                sb.Append("  \"pluginMemory\": ").Append(PluginMemory ? "true" : "false").Append(",\n");
                sb.Append("  \"pluginPref\": ").Append(PluginPref ? "true" : "false").Append(",\n");
                sb.Append("  \"pluginHist\": ").Append(PluginHist ? "true" : "false").Append(",\n");
                sb.Append("  \"pluginPlan\": ").Append(PluginPlan ? "true" : "false").Append(",\n");
                sb.Append("  \"convTargetIndex\": ").Append(ConvTargetIndex).Append(",\n");
                sb.Append("  \"privacyLevel\": ").Append(PrivacyLevel >= 0 && PrivacyLevel <= 2 ? PrivacyLevel : 1).Append(",\n");
                sb.Append("  \"webSearch\": ").Append(WebSearch ? "true" : "false").Append(",\n");
                sb.Append("  \"streamFinal\": ").Append(StreamFinal ? "true" : "false").Append(",\n");
                sb.Append("  \"workspaceDir\": \"").Append(Js(WorkspaceDir)).Append("\",\n");
                sb.Append("  \"activeWorkspaceId\": \"").Append(Js(ActiveWorkspaceId)).Append("\",\n");
                sb.Append("  \"safetyGuard\": ").Append(SafetyGuard ? "true" : "false").Append(",\n");
                sb.Append("  \"safetyAllowPaths\": \"").Append(Js(SafetyAllowPaths)).Append("\",\n");
                sb.Append("  \"customModels\": \"").Append(Js(CustomModels)).Append("\",\n");
                // 按服务分开存的密钥（扁平字段 key:<host> 或 key:<host>#<tag>，base64 DPAPI 密文）
                for (int i = 0; i < ProviderKeys.Count; i++)
                {
                    if (ProviderKeys[i].Blob == null || ProviderKeys[i].Blob.Length == 0) continue;
                    string slot = ProviderKeys[i].Host + (string.IsNullOrEmpty(ProviderKeys[i].Tag) ? "" : "#" + ProviderKeys[i].Tag);
                    sb.Append("  \"key:").Append(Js(slot)).Append("\": \"")
                      .Append(Convert.ToBase64String(ProviderKeys[i].Blob)).Append("\",\n");
                }
                // 密钥槽自定义名称（扁平字段 keylabel:<host>#<tag>）
                for (int i = 0; i < ProviderKeys.Count; i++)
                {
                    if (ProviderKeys[i].Label == null || ProviderKeys[i].Label.Length == 0) continue;
                    string slot = ProviderKeys[i].Host + (string.IsNullOrEmpty(ProviderKeys[i].Tag) ? "" : "#" + ProviderKeys[i].Tag);
                    sb.Append("  \"keylabel:").Append(Js(slot)).Append("\": \"")
                      .Append(Js(ProviderKeys[i].Label)).Append("\",\n");
                }
                sb.Append("  \"perHostKeys\": true\n");
                sb.Append("}\n");
                return sb.ToString();
            }
            catch { return ""; }
        }

        public static AppConfig Load()
        {
            try
            {
                if (!File.Exists(FilePath())) return new AppConfig();
                return LoadFrom(File.ReadAllText(FilePath(), Encoding.UTF8));
            }
            catch { return new AppConfig(); }
        }

        // 从 JSON 文本解析。抽成独立方法是为了**可测**：/menutest 用合成 JSON 验证
        // 解析逻辑（含密钥槽名称），而不必去读写用户真实的 config.json。
        public static AppConfig LoadFrom(string json)
        {
            AppConfig c = new AppConfig();
            try
            {
                if (json == null) return c;
                c.BaseUrl = JsGet(json, "baseUrl");
                c.Model = JsGet(json, "model");
                c.AllowLan = JsGet(json, "allowLan") == "true";
                c.WizardDone = JsGet(json, "wizardDone") == "true";
                c.PluginMemory = GetBool(json, "pluginMemory", true);
                c.PluginPref = GetBool(json, "pluginPref", true);
                c.PluginHist = GetBool(json, "pluginHist", true);
                c.PluginPlan = GetBool(json, "pluginPlan", true);
                int ti;
                if (int.TryParse(JsGet(json, "convTargetIndex"), out ti) && ti >= 0 && ti <= 2) c.ConvTargetIndex = ti;
                int pl;
                if (int.TryParse(JsGet(json, "privacyLevel"), out pl) && pl >= 0 && pl <= 2) c.PrivacyLevel = pl;
                c.WebSearch = GetBool(json, "webSearch", true);
                c.StreamFinal = GetBool(json, "streamFinal", true);
                c.WorkspaceDir = JsGet(json, "workspaceDir");
                if (c.WorkspaceDir == null) c.WorkspaceDir = "";
                c.ActiveWorkspaceId = JsGet(json, "activeWorkspaceId");
                if (c.ActiveWorkspaceId == null) c.ActiveWorkspaceId = "";
                c.SafetyGuard = GetBool(json, "safetyGuard", true);
                c.SafetyAllowPaths = JsGet(json, "safetyAllowPaths");
                if (c.SafetyAllowPaths == null) c.SafetyAllowPaths = "";
                c.CustomModels = JsGet(json, "customModels");
                if (c.CustomModels == null) c.CustomModels = "";
                string b64 = JsGet(json, "keyBlob");
                if (b64 != null && b64.Length > 0)
                {
                    try { c.KeyBlob = Convert.FromBase64String(b64); } catch { c.KeyBlob = null; }
                }
                // 新版：按服务分开的密钥字段 "key:<host>" 或 "key:<host>#<tag>"（扁平扫描）。
                // host 本身不会出现 '#'，取最后一个 '#' 后是纯数字才当账号槽，否则整体视为主机名。
                int ki = 0;
                while ((ki = json.IndexOf("\"key:", ki, StringComparison.Ordinal)) >= 0)
                {
                    int ns = ki + 5;
                    int ne = json.IndexOf('"', ns);
                    if (ne < 0) break;
                    string slot = json.Substring(ns, ne - ns);
                    string host = slot;
                    string tag = "";
                    int hash = slot.LastIndexOf('#');
                    if (hash > 0 && hash < slot.Length - 1)
                    {
                        string t = slot.Substring(hash + 1);
                        bool digits = t.Length > 0;
                        for (int d = 0; d < t.Length; d++) { if (t[d] < '0' || t[d] > '9') { digits = false; break; } }
                        if (digits) { host = slot.Substring(0, hash); tag = t; }
                    }
                    int j = ne + 1;
                    while (j < json.Length && (json[j] == ' ' || json[j] == ':')) j++;
                    if (j < json.Length && json[j] == '"')
                    {
                        int vs = j + 1;
                        int ve = json.IndexOf('"', vs);
                        if (ve > vs)
                        {
                            byte[] blob = null;
                            try { blob = Convert.FromBase64String(json.Substring(vs, ve - vs)); } catch { blob = null; }
                            if (blob != null && host.Length > 0 && !c.HasKeyForHost(host, tag))
                            {
                                ProviderKey pk = new ProviderKey();
                                pk.Host = host;
                                pk.Tag = tag;
                                pk.Blob = blob;
                                c.ProviderKeys.Add(pk);
                            }
                        }
                    }
                    ki = ne + 1;
                }
                // 密钥槽自定义名称："keylabel:<host>#<tag>": "<label>"（同样是扁平扫描）。
                // 必须在密钥解析之后——按 host#tag 回填到已建好的 ProviderKey 上。
                int li = 0;
                while ((li = json.IndexOf("\"keylabel:", li, StringComparison.Ordinal)) >= 0)
                {
                    int ns2 = li + 10;
                    int ne2 = json.IndexOf('"', ns2);
                    if (ne2 < 0) break;
                    string slot2 = json.Substring(ns2, ne2 - ns2);
                    string host2 = slot2, tag2 = "";
                    int h2 = slot2.LastIndexOf('#');
                    if (h2 > 0 && h2 < slot2.Length - 1)
                    {
                        string t2 = slot2.Substring(h2 + 1);
                        bool dg = t2.Length > 0;
                        for (int d = 0; d < t2.Length; d++) { if (t2[d] < '0' || t2[d] > '9') { dg = false; break; } }
                        if (dg) { host2 = slot2.Substring(0, h2); tag2 = t2; }
                    }
                    int j2 = ne2 + 1;
                    while (j2 < json.Length && (json[j2] == ' ' || json[j2] == ':')) j2++;
                    if (j2 < json.Length && json[j2] == '"')
                    {
                        int vs2 = j2 + 1;
                        int ve2 = json.IndexOf('"', vs2);
                        if (ve2 > vs2)
                        {
                            string lbl = json.Substring(vs2, ve2 - vs2);
                            for (int p = 0; p < c.ProviderKeys.Count; p++)
                            {
                                if (string.Equals(c.ProviderKeys[p].Host, host2, StringComparison.OrdinalIgnoreCase)
                                    && string.Equals(c.ProviderKeys[p].Tag ?? "", tag2, StringComparison.Ordinal))
                                { c.ProviderKeys[p].Label = lbl; break; }
                            }
                        }
                    }
                    li = ne2 + 1;
                }
                // 旧版单密钥迁移：把 keyBlob 归到它所属的服务地址下，升级为按家存储。
                // （只迁移一次；之后 Save 不再写 keyBlob，旧字段随之消失）
                if (c.KeyBlob != null && c.KeyBlob.Length > 0 && c.ProviderKeys.Count == 0)
                {
                    string h = HostOf(c.BaseUrl);
                    if (h.Length > 0)
                    {
                        ProviderKey pk = new ProviderKey();
                        pk.Host = h;
                        pk.Blob = c.KeyBlob;
                        c.ProviderKeys.Add(pk);
                        c.KeyBlob = null;
                    }
                }
            }
            catch { }
            return c;
        }

        // ---------- 极简 JSON（仅本扁平 schema；值域不含嵌套）----------

        static string Js(string s)
        {
            if (s == null) return "";
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
                    default:
                        if (ch < 0x20) sb.Append("\\u").Append(((int)ch).ToString("x4"));
                        else sb.Append(ch);
                        break;
                }
            }
            return sb.ToString();
        }

        static string JsGet(string json, string name)
        {
            if (json == null) return null;
            string pat = "\"" + name + "\"";
            int i = json.IndexOf(pat, StringComparison.Ordinal);
            if (i < 0) return null;
            i = json.IndexOf(':', i + pat.Length);
            if (i < 0) return null;
            i++;
            while (i < json.Length && (json[i] == ' ' || json[i] == '\t')) i++;
            if (i >= json.Length) return null;
            if (json[i] == '"')
            {
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
            int end = i;
            while (end < json.Length && json[end] != ',' && json[end] != '\n' && json[end] != '}') end++;
            return json.Substring(i, end - i).Trim();
        }

        // 布尔读取：字段缺失时用默认值（内置插件默认全开 → 老配置文件升级零操作）
        static bool GetBool(string json, string name, bool def)
        {
            string v = JsGet(json, name);
            if (v == null) return def;
            return v == "true";
        }

        // ---------- DPAPI（crypt32）----------

        [StructLayout(LayoutKind.Sequential)]
        internal struct DATA_BLOB
        {
            public int cbData;
            public IntPtr pbData;
        }

        [DllImport("crypt32.dll", SetLastError = true)]
        static extern bool CryptProtectData(ref DATA_BLOB input, string desc, ref DATA_BLOB entropy, IntPtr reserved, IntPtr prompt, uint flags, ref DATA_BLOB output);
        [DllImport("crypt32.dll", SetLastError = true)]
        static extern bool CryptUnprotectData(ref DATA_BLOB input, IntPtr desc, ref DATA_BLOB entropy, IntPtr reserved, IntPtr prompt, uint flags, ref DATA_BLOB output);
        [DllImport("kernel32.dll")] static extern IntPtr LocalAlloc(uint flags, int bytes);
        [DllImport("kernel32.dll")] static extern IntPtr LocalFree(IntPtr h);

        static readonly byte[] Entropy = Encoding.UTF8.GetBytes("OfficeAgent.key.entropy.v1");

        static IntPtr AllocBlob(byte[] data)
        {
            IntPtr p = LocalAlloc(0x40, data.Length); // LMEM_ZEROINIT|LMEM_MOVEABLE 固定
            Marshal.Copy(data, 0, p, data.Length);
            return p;
        }

        static byte[] DpapiProtect(string plain)
        {
            byte[] plainBytes = Encoding.UTF8.GetBytes(plain);
            DATA_BLOB input = new DATA_BLOB();
            DATA_BLOB ent = new DATA_BLOB();
            DATA_BLOB output = new DATA_BLOB();
            input.cbData = plainBytes.Length; input.pbData = AllocBlob(plainBytes);
            ent.cbData = Entropy.Length; ent.pbData = AllocBlob(Entropy);
            try
            {
                if (!CryptProtectData(ref input, "OfficeAgent API Key", ref ent, IntPtr.Zero, IntPtr.Zero, 0, ref output))
                    throw new InvalidOperationException("DPAPI 加密失败");
                byte[] result = new byte[output.cbData];
                Marshal.Copy(output.pbData, result, 0, output.cbData);
                return result;
            }
            finally
            {
                if (input.pbData != IntPtr.Zero) LocalFree(input.pbData);
                if (ent.pbData != IntPtr.Zero) LocalFree(ent.pbData);
                if (output.pbData != IntPtr.Zero) LocalFree(output.pbData);
            }
        }

        static string DpapiUnprotect(byte[] blob)
        {
            DATA_BLOB input = new DATA_BLOB();
            DATA_BLOB ent = new DATA_BLOB();
            DATA_BLOB output = new DATA_BLOB();
            input.cbData = blob.Length; input.pbData = AllocBlob(blob);
            ent.cbData = Entropy.Length; ent.pbData = AllocBlob(Entropy);
            try
            {
                if (!CryptUnprotectData(ref input, IntPtr.Zero, ref ent, IntPtr.Zero, IntPtr.Zero, 0, ref output))
                    return null;
                byte[] result = new byte[output.cbData];
                Marshal.Copy(output.pbData, result, 0, output.cbData);
                return Encoding.UTF8.GetString(result);
            }
            catch { return null; }
            finally
            {
                if (input.pbData != IntPtr.Zero) LocalFree(input.pbData);
                if (ent.pbData != IntPtr.Zero) LocalFree(ent.pbData);
                if (output.pbData != IntPtr.Zero) LocalFree(output.pbData);
            }
        }
    }
}
