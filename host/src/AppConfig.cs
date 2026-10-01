// AppConfig —— 本地配置持久化（%LOCALAPPDATA%\OfficeAgent\config.json）
// 密钥安全规约（Mimosa 约束）：API Key 只经 DPAPI（Windows 凭据保护服务）加密后落盘，
// 支持环境变量 OFFICEAGENT_API_KEY 运行时覆盖；源码/示例/测试不写任何凭据字面量。
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace OfficeAgent.Host
{
    public class AppConfig
    {
        public string BaseUrl = "";
        public string Model = "";
        public bool AllowLan = false;      // 企业内网网关显式放行（默认关闭）
        public byte[] KeyBlob = null;      // DPAPI 密文
        public bool WizardDone = false;
        // 内置插件开关（默认全开，开箱即用）
        public bool PluginMemory = true;   // 记忆模块
        public bool PluginPref = true;     // 转换偏好
        public bool PluginHist = true;     // 会话历史
        public bool PluginPlan = true;     // 任务计划栏（复杂任务自动列计划，右侧展示）
        public int ConvTargetIndex = 0;    // 转换偏好：上次目标格式（0=PDF 1=CSV 2=XLSX）
        public int PrivacyLevel = 1;       // 隐私分级（设计方案 §7.2）：0=全本地 1=脱敏出网(默认) 2=全量
        public bool WebSearch = true;      // 联网搜索工具（智谱系端点 web_search；关=纯本地对话）
        public string WorkspaceDir = "";   // 默认工作区目录：空=文档\OfficeAgentFiles（侧边栏可为项目单独设工作区）
        public string ActiveWorkspaceId = "";   // 当前激活的工作区 id（对应 workspaces.json）

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

        public bool IsLlmConfigured()
        {
            return BaseUrl != null && BaseUrl.Length > 0 && GetKey() != null && GetKey().Length > 0;
        }

        public void SetKey(string plain)
        {
            if (string.IsNullOrEmpty(plain)) { KeyBlob = null; return; }
            KeyBlob = DpapiProtect(plain);
        }

        public string GetKey()
        {
            // 环境变量覆盖优先（安全规约允许的凭据来源之二）
            try
            {
                string env = Environment.GetEnvironmentVariable("OFFICEAGENT_API_KEY");
                if (!string.IsNullOrEmpty(env)) return env;
            }
            catch { }
            if (KeyBlob == null || KeyBlob.Length == 0) return null;
            return DpapiUnprotect(KeyBlob);
        }

        public void Save()
        {
            try
            {
                Directory.CreateDirectory(Dir());
                StringBuilder sb = new StringBuilder();
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
                sb.Append("  \"workspaceDir\": \"").Append(Js(WorkspaceDir)).Append("\",\n");
                sb.Append("  \"activeWorkspaceId\": \"").Append(Js(ActiveWorkspaceId)).Append("\",\n");
                sb.Append("  \"keyBlob\": \"").Append(KeyBlob == null ? "" : Convert.ToBase64String(KeyBlob)).Append("\"\n");
                sb.Append("}\n");
                File.WriteAllText(FilePath(), sb.ToString(), new UTF8Encoding(false));
            }
            catch { }
        }

        public static AppConfig Load()
        {
            AppConfig c = new AppConfig();
            try
            {
                if (!File.Exists(FilePath())) return c;
                string json = File.ReadAllText(FilePath(), Encoding.UTF8);
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
                c.WorkspaceDir = JsGet(json, "workspaceDir");
                if (c.WorkspaceDir == null) c.WorkspaceDir = "";
                c.ActiveWorkspaceId = JsGet(json, "activeWorkspaceId");
                if (c.ActiveWorkspaceId == null) c.ActiveWorkspaceId = "";
                string b64 = JsGet(json, "keyBlob");
                if (b64 != null && b64.Length > 0)
                {
                    try { c.KeyBlob = Convert.FromBase64String(b64); } catch { c.KeyBlob = null; }
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
