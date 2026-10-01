// SkillSystem —— 技能体系骨架（设计方案 §4：发现 / 清单校验 / 注册表 / 能力表）
// 本期范围（M3 二期）：
//   * skill.json 清单解析与 schema 校验（扁平 JSON，数组用逗号分隔，兼容 MiniJson）
//   * 三级扫描根：内置 <root>\skills → 用户 %LOCALAPPDATA%\OfficeAgent\skills → 团队共享（config 留位）
//   * 注册表 skills-registry.json（同 id 取 host 兼容的最高版本；被禁用保留但路由不可见）
//   * 能力表（libreoffice / office-com / ocr）→ 技能 depsTools 匹配（"?" 后缀 = 缺失可降级）
//   * 执行桥接（python sidecar / builtin 分发到引擎）留待下期；本骨架保证发现与元数据正确
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace OfficeAgent.Host
{
    // 技能体系字面量的唯一定义处（清单类/注册表类都在 SkillSystem 之外，故独立成类）。
    // 用常量替代散落各文件的裸字符串：拼错即编译报错，改值只需一处。
    public static class SkillConst
    {
        public const string RuntimeBuiltin = "builtin";
        public const string RuntimePython38 = "python38";
        public const string RuntimeDotnet = "dotnet";

        public const string StateEnabled = "enabled";
        public const string StateDisabled = "disabled";
        public const string StateDegraded = "degraded";

        public const string SourceBuiltin = "builtin";
        public const string SourceUser = "user";
        public const string SourceTeam = "team";
    }

    public class SkillManifest
    {
        public string Id = "";
        public string Name = "";
        public string Version = "1.0.0";
        public string Host = "1.0.0";        // 要求的宿主最低版本
        public string Runtime = SkillConst.RuntimeBuiltin;   // builtin | python38 | dotnet（本期仅 builtin 会被执行）
        public string Entry = "";            // builtin: 引擎任务名（recon/merge/invoices）
        public string[] Scenarios = new string[0];
        public string[] DepsTools = new string[0];   // "libreoffice" 必需 / "libreoffice?" 可选
        public string Dir = "";              // 技能目录
        public string Source = "";           // builtin | user | team
        public List<string> Problems = new List<string>();

        // 声明但当前执行桥无法放行的能力（网络/Office COM）。SkillRunner.Run 会据此拒绝执行，
        // 避免"声明了权限"的技能在无沙箱下跑起来（此前两字段解析后无人读取，形同虚设）。
        public bool Network = false;
        public bool OfficeCom = false;
        public string Signature = "";                 // 团队级验签（§4.3）：团队共享根未落地，暂不验
        public bool SignatureVerified = false;        // 仅团队根会置 true；builtin/user 不验签
        public string[] Unenforced = new string[0];   // 声明了但执行期会拒绝的能力名

        public bool Valid { get { return Problems.Count == 0 && Id.Length > 0; } }
    }

    public class SkillRegistryEntry
    {
        public string Id = "";
        public string Version = "";
        public string Path = "";
        public string Source = "";
        public string State = SkillConst.StateEnabled;     // enabled | disabled | degraded
        public string Scenarios = "";
        public string Runtime = "";
        public string Entry = "";            // builtin: 引擎任务名（recon/merge/invoices）
        public string DegradedReason = "";
    }

    public static class SkillSystem
    {
        public const string HostVersion = "1.0.0";   // 宿主版本（host 兼容判定）

        // ---------- 清单解析 ----------

        public static SkillManifest ParseManifest(string dir)
        {
            SkillManifest m = new SkillManifest();
            m.Dir = dir;
            string file = Path.Combine(dir, "skill.json");
            if (!File.Exists(file))
            {
                m.Problems.Add("缺少 skill.json");
                return m;
            }
            string json;
            try { json = File.ReadAllText(file, Encoding.UTF8); }
            catch (Exception ex) { m.Problems.Add("skill.json 读取失败: " + ex.Message); return m; }

            List<Dictionary<string, string>> objs = MiniJson.ParseObjects(json);
            if (objs.Count == 0) { m.Problems.Add("skill.json 解析失败（无对象）"); return m; }
            Dictionary<string, string> o = objs[0];

            m.Id = MiniJson.Get(o, "id").Trim();
            m.Name = MiniJson.Get(o, "name").Trim();
            m.Version = MiniJson.GetOr(o, "version", "1.0.0");
            m.Host = MiniJson.GetOr(o, "host", "1.0.0");
            m.Runtime = MiniJson.GetOr(o, "runtime", SkillConst.RuntimeBuiltin).Trim();
            m.Entry = MiniJson.Get(o, "entry").Trim();
            m.Scenarios = SplitList(MiniJson.Get(o, "scenarios"));
            m.DepsTools = SplitList(MiniJson.Get(o, "depsTools"));
            m.Network = MiniJson.Get(o, "network") == "true";
            m.OfficeCom = MiniJson.Get(o, "officeCom") == "true";
            m.Signature = MiniJson.Get(o, "signature");

            // 声明了网络/COM 的技能：执行桥当前不支持放行（无沙箱代理），记录下来供注册表与 UI 提示
            List<string> unenforced = new List<string>();
            if (m.Network) unenforced.Add("network");
            if (m.OfficeCom) unenforced.Add("officeCom");
            m.Unenforced = unenforced.ToArray();

            if (m.Id.Length == 0) m.Problems.Add("缺少 id");
            else if (!IsSafeId(m.Id)) m.Problems.Add("id 含非法字符（仅限字母数字-_.）");
            if (m.Name.Length == 0) m.Problems.Add("缺少 name");
            if (!IsSemver(m.Version)) m.Problems.Add("version 不是 semver: " + m.Version);
            if (!IsSemver(m.Host)) m.Problems.Add("host 不是 semver: " + m.Host);
            if (CompareSemver(m.Host, HostVersion) > 0)
                m.Problems.Add("要求宿主 >= " + m.Host + "，当前 " + HostVersion);
            if (m.Runtime != SkillConst.RuntimeBuiltin && m.Runtime != SkillConst.RuntimePython38 && m.Runtime != SkillConst.RuntimeDotnet)
                m.Problems.Add("未知 runtime: " + m.Runtime);
            if (m.Entry.Length == 0) m.Problems.Add("缺少 entry");
            return m;
        }

        static bool IsSafeId(string id)
        {
            foreach (char c in id)
            {
                bool ok = (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9')
                    || c == '-' || c == '_' || c == '.';
                if (!ok) return false;
            }
            return id.Length > 0;
        }

        static bool IsSemver(string s)
        {
            if (s == null || s.Length == 0) return false;
            string[] parts = s.Split('.');
            if (parts.Length < 2 || parts.Length > 4) return false;
            for (int i = 0; i < parts.Length; i++)
            {
                int n;
                if (!int.TryParse(parts[i], out n) || n < 0) return false;
            }
            return true;
        }

        // -1 = a<b；0 相等；1 = a>b
        public static int CompareSemver(string a, string b)
        {
            int[] va = SemverNums(a), vb = SemverNums(b);
            for (int i = 0; i < 4; i++)
            {
                if (va[i] != vb[i]) return va[i] < vb[i] ? -1 : 1;
            }
            return 0;
        }

        static int[] SemverNums(string s)
        {
            int[] r = new int[] { 0, 0, 0, 0 };
            if (s == null) return r;
            string[] parts = s.Split('.');
            for (int i = 0; i < 4 && i < parts.Length; i++)
            {
                int n;
                if (int.TryParse(parts[i], out n)) r[i] = n;
            }
            return r;
        }

        static string[] SplitList(string s)
        {
            List<string> parts = new List<string>();
            if (s != null)
            {
                foreach (string p in s.Split(','))
                {
                    string t = p.Trim();
                    if (t.Length > 0) parts.Add(t);
                }
            }
            return parts.ToArray();
        }

        // ---------- 能力表 ----------

        public class Capabilities
        {
            public bool LibreOffice;
            public bool OfficeCom;
            public bool Ocr;

            // 工具名 → 能力判定（唯一定义处）。未知工具名一律 false：拼错依赖名必须暴露，
            // 不能静默放行（否则技能声明的依赖形同虚设）。
            public bool Has(string tool)
            {
                if (tool == "libreoffice") return LibreOffice;
                if (tool == "office-com") return OfficeCom;
                if (tool == "ocr") return Ocr;
                return false;
            }
        }

        // Office COM 探测按需缓存：Type.GetTypeFromProgID 会触碰注册表与 COM，
        // 在 IntentRouter 热路径（每条消息）里做会明显卡顿，故首次访问才探测并记住结果。
        static bool? officeComCache = null;

        static bool DetectOfficeCom()
        {
            if (officeComCache == null)
            {
                try { officeComCache = Type.GetTypeFromProgID("Excel.Application") != null; }
                catch { officeComCache = false; }
            }
            return officeComCache.Value;
        }

        public static Capabilities DetectCapabilities(string root, ConvertEngine conv)
        {
            Capabilities c = new Capabilities();
            c.LibreOffice = conv != null && conv.SofficePath != null;
            c.OfficeCom = DetectOfficeCom();
            c.Ocr = false;   // OCR 组件未随包（tesseract 可选，见设计方案 §6.4）
            return c;
        }

        // 轻量能力表：只做文件存在性判断（无 COM 探测），供热路径使用。
        // 注册表/路由只需要 deps 满足性，COM 是否可用由执行期再确认。
        public static Capabilities DetectCapabilitiesFast(string root)
        {
            Capabilities c = new Capabilities();
            c.LibreOffice = ConvertEngine.FindSoffice(root) != null;
            c.OfficeCom = false;   // 热路径不探测 COM（保守：依赖 office-com 的技能标 degraded）
            c.Ocr = false;
            return c;
        }

        // 工具依赖满足性："?" 后缀可选。返回 null=满足，否则降级原因。
        // 未知工具名（不在 Capabilities 能力表内）视为不满足 —— 拼写错误必须显式暴露，
        // 而不是静默放行；需要"环境可能没有、没有也能跑"的依赖请写 "name?" 显式声明可选。
        public static string DepsMissing(string[] depsTools, Capabilities cap)
        {
            if (depsTools == null) return null;
            foreach (string raw in depsTools)
            {
                string t = raw.Trim();
                if (t.Length == 0) continue;
                bool optional = t.EndsWith("?", StringComparison.Ordinal);
                string name = optional ? t.Substring(0, t.Length - 1) : t;
                if (name.Length == 0) continue;
                bool have = cap != null && cap.Has(name);
                if (!have && !optional) return "缺少工具依赖: " + name;
            }
            return null;
        }

        // ---------- 扫描与注册 ----------

        // 三级扫描根（§4.3）：builtin（低）→ user → team 共享（高，优先级最高）
        // team 根来自 config 的 teamSkillsRoot（默认空 = 未配置，等同"留位"）。
        public static List<string> ScanRoots(string root)
        {
            List<string> roots = new List<string>();
            roots.Add(Path.Combine(root, "skills"));   // builtin（低优先）
            string user = Path.Combine(Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeAgent"), "skills");
            roots.Add(user);                            // user（中优先）
            roots.Add(TeamRoot());                      // team（高优先；未配置时为空串，扫描时跳过）
            return roots;
        }

        // 团队共享根：环境变量 OfficeAgentTeamSkills 优先，否则读 config;空串 = 未配置
        public static string TeamRoot()
        {
            try
            {
                string env = Environment.GetEnvironmentVariable("OfficeAgentTeamSkills");
                if (env != null && env.Trim().Length > 0) return env.Trim();
            }
            catch { }
            return "";
        }

        // 路由用缓存：每进程扫描一次（技能场景词喂 IntentRouter，避免每条消息扫盘）。
        // 用轻量能力探测（只查文件存在性，不做 COM），保证热路径不被 COM 探测拖慢。
        // 修掉了此前传 conv=null 导致 LibreOffice 恒 false、依赖 LO 的技能被误判 degraded 的问题。
        static List<SkillRegistryEntry> scenarioCache = null;
        static FileSystemWatcher watcher = null;

        public static List<SkillRegistryEntry> CachedScan(string root)
        {
            if (scenarioCache == null)
            {
                scenarioCache = Scan(root, null, null, true, true);
                EnsureWatcher(root);
            }
            return scenarioCache;
        }

        // §4.3 目录变更监视：技能根有增删改就让缓存失效（装新技能无需重启）
        static void EnsureWatcher(string root)
        {
            if (watcher != null) return;
            try
            {
                watcher = new FileSystemWatcher();
                List<string> roots = ScanRoots(root);
                foreach (string r in roots)
                {
                    if (r != null && r.Length > 0 && Directory.Exists(r)) { watcher.Path = r; break; }
                }
                if (watcher.Path == null || watcher.Path.Length == 0) { watcher = null; return; }
                watcher.IncludeSubdirectories = true;
                watcher.NotifyFilter = NotifyFilters.DirectoryName | NotifyFilters.FileName | NotifyFilters.LastWrite;
                watcher.Changed += delegate(object s, FileSystemEventArgs e) { NotifyChange(e.Name); };
                watcher.Created += delegate(object s, FileSystemEventArgs e) { NotifyChange(e.Name); };
                watcher.Deleted += delegate(object s, FileSystemEventArgs e) { NotifyChange(e.Name); };
                watcher.Renamed += delegate(object s, RenamedEventArgs e) { NotifyChange(e.Name); };
                watcher.EnableRaisingEvents = true;
            }
            catch { watcher = null; }
        }

        // 只对 skill.json 的变动失效缓存（技能目录内其它文件写入不影响注册表）
        static void NotifyChange(string name)
        {
            if (name == null) return;
            if (name.EndsWith("skill.json", StringComparison.OrdinalIgnoreCase)) InvalidateCache();
        }

        // 目录变更后失效缓存（装/卸技能、改 skill.json 后调用）
        public static void InvalidateCache() { scenarioCache = null; }

        // 扫描全部根 → 清单校验 → 能力匹配 → 注册表（同 id 取版本最高者）
        public static List<SkillRegistryEntry> Scan(string root, ConvertEngine conv)
        {
            return Scan(root, conv, null, true, false);
        }

        // userRootOverride：自测用，指向临时目录以免污染真实用户技能根；
        // writeRegistry=false 时不落盘（自测不写用户注册表）。
        // fast = true 时用轻量能力探测（无 COM），供热路径复用。
        public static List<SkillRegistryEntry> Scan(string root, ConvertEngine conv, string userRootOverride, bool writeRegistry)
        {
            return Scan(root, conv, userRootOverride, writeRegistry, false);
        }

        public static List<SkillRegistryEntry> Scan(string root, ConvertEngine conv, string userRootOverride,
            bool writeRegistry, bool fast)
        {
            Capabilities cap = fast ? DetectCapabilitiesFast(root) : DetectCapabilities(root, conv);
            Dictionary<string, SkillRegistryEntry> best = new Dictionary<string, SkillRegistryEntry>();
            List<string> states = LoadStates();
            // (路径, 来源) 显式成对 —— 不做路径后缀猜测（安装目录可能恰好也叫 OfficeAgent）
            List<KeyValuePair<string, string>> roots = new List<KeyValuePair<string, string>>();
            List<string> scanRoots = ScanRoots(root);
            roots.Add(new KeyValuePair<string, string>(scanRoots[0], SkillConst.SourceBuiltin));
            roots.Add(new KeyValuePair<string, string>(
                userRootOverride != null ? userRootOverride : scanRoots[1], SkillConst.SourceUser));
            // 团队共享根（§4.3）：优先级最高；未配置（空串）时跳过
            if (userRootOverride == null && scanRoots.Count > 2 && scanRoots[2] != null && scanRoots[2].Length > 0)
                roots.Add(new KeyValuePair<string, string>(scanRoots[2], SkillConst.SourceTeam));
            foreach (KeyValuePair<string, string> rootPair in roots)
            {
                string scanRoot = rootPair.Key;
                if (!Directory.Exists(scanRoot)) continue;
                string source = rootPair.Value;
                DirectoryInfo[] dirs;
                try { dirs = new DirectoryInfo(scanRoot).GetDirectories(); }
                catch { continue; }
                foreach (DirectoryInfo d in dirs)
                {
                    SkillManifest m = ParseManifest(d.FullName);
                    if (!m.Valid) continue;
                    if (best.ContainsKey(m.Id) && CompareSemver(best[m.Id].Version, m.Version) >= 0) continue;

                    SkillRegistryEntry e = new SkillRegistryEntry();
                    e.Id = m.Id; e.Version = m.Version; e.Path = d.FullName; e.Source = source;
                    e.Runtime = m.Runtime;
                    e.Entry = m.Entry;
                    e.Scenarios = Join(m.Scenarios);
                    string missing = DepsMissing(m.DepsTools, cap);
                    if (missing != null) { e.State = SkillConst.StateDegraded; e.DegradedReason = missing; }
                    else e.State = SkillConst.StateEnabled;
                    best[m.Id] = e;
                }
            }
            List<SkillRegistryEntry> list = new List<SkillRegistryEntry>(best.Values);
            foreach (SkillRegistryEntry e in list)
            {
                if (states.Contains(SkillConst.StateDisabled + ":" + e.Id)) e.State = SkillConst.StateDisabled;
            }
            if (writeRegistry) SaveRegistry(list);
            return list;
        }

        static string Join(string[] arr)
        {
            StringBuilder sb = new StringBuilder();
            foreach (string s in arr) { if (sb.Length > 0) sb.Append(","); sb.Append(s); }
            return sb.ToString();
        }

        static string RegistryFile()
        {
            return Path.Combine(Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeAgent"),
                "skills-registry.json");
        }

        static void SaveRegistry(List<SkillRegistryEntry> list)
        {
            try
            {
                string dir = Path.GetDirectoryName(RegistryFile());
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                StringBuilder sb = new StringBuilder();
                sb.Append("[\n");
                for (int i = 0; i < list.Count; i++)
                {
                    SkillRegistryEntry e = list[i];
                    sb.Append("  {\"id\":\"").Append(MiniJson.Esc(e.Id)).Append("\",");
                    sb.Append("\"version\":\"").Append(MiniJson.Esc(e.Version)).Append("\",");
                    sb.Append("\"path\":\"").Append(MiniJson.Esc(e.Path)).Append("\",");
                    sb.Append("\"source\":\"").Append(MiniJson.Esc(e.Source)).Append("\",");
                    sb.Append("\"state\":\"").Append(MiniJson.Esc(e.State)).Append("\",");
                    sb.Append("\"scenarios\":\"").Append(MiniJson.Esc(e.Scenarios)).Append("\",");
                    sb.Append("\"runtime\":\"").Append(MiniJson.Esc(e.Runtime)).Append("\",");
                    sb.Append("\"entry\":\"").Append(MiniJson.Esc(e.Entry)).Append("\",");
                    sb.Append("\"degradedReason\":\"").Append(MiniJson.Esc(e.DegradedReason)).Append("\"}");
                    if (i < list.Count - 1) sb.Append(",");
                    sb.Append("\n");
                }
                sb.Append("]\n");
                File.WriteAllText(RegistryFile(), sb.ToString(), new UTF8Encoding(false));
            }
            catch { }
        }

        // 用户级禁用清单（随注册表文件旁存一份 ids；禁用状态跨扫描保留）
        static string StatesFile()
        {
            return Path.Combine(Path.GetDirectoryName(RegistryFile()), "skill-states.json");
        }

        static List<string> LoadStates()
        {
            List<string> states = new List<string>();
            try
            {
                if (!File.Exists(StatesFile())) return states;
                string json = File.ReadAllText(StatesFile(), Encoding.UTF8);
                foreach (Dictionary<string, string> o in MiniJson.ParseObjects(json))
                {
                    string id = MiniJson.Get(o, "id");
                    string state = MiniJson.GetOr(o, "state", SkillConst.StateEnabled);
                    if (id.Length > 0 && state == SkillConst.StateDisabled) states.Add(SkillConst.StateDisabled + ":" + id);
                }
            }
            catch { }
            return states;
        }

        public static void SetEnabled(string skillId, bool enabled)
        {
            try
            {
                List<Dictionary<string, string>> rows = new List<Dictionary<string, string>>();
                string file = StatesFile();
                if (File.Exists(file))
                {
                    foreach (Dictionary<string, string> o in MiniJson.ParseObjects(File.ReadAllText(file, Encoding.UTF8)))
                    {
                        if (MiniJson.Get(o, "id") != skillId) rows.Add(o);
                    }
                }
                if (!enabled)
                {
                    Dictionary<string, string> row = new Dictionary<string, string>();
                    row["id"] = skillId;
                    row["state"] = SkillConst.StateDisabled;
                    rows.Add(row);
                }
                StringBuilder sb = new StringBuilder();
                sb.Append("[\n");
                for (int i = 0; i < rows.Count; i++)
                {
                    sb.Append("  {\"id\":\"").Append(MiniJson.Esc(rows[i]["id"]))
                      .Append("\",\"state\":\"").Append(MiniJson.Esc(rows[i]["state"])).Append("\"}");
                    if (i < rows.Count - 1) sb.Append(",");
                    sb.Append("\n");
                }
                sb.Append("]\n");
                File.WriteAllText(file, sb.ToString(), new UTF8Encoding(false));
            }
            catch { }
        }
    }
}
