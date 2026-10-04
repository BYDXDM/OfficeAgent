// AgentTools —— 会话模型可调用的白名单工具（function calling）
// 设计约束：
//   * 只有本文件注册的 5 个工具；模型给的参数只作数据用，绝不参与任何命令行/shell 构造；
//   * download_file 出网前经 HostGuard 校验（仅 http/https，拒绝内网/环回/保留地址），每跳重定向复检；
//   * read_text_file 只读文本类扩展名，限制大小；原始文件一律只读；
//   * repair_environment 经 RepairLauncher 启动同包引导器提权变体（UAC 由系统弹）。
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using OfficeAgent.Core;

namespace OfficeAgent.Host
{
    public static class AgentTools
    {
        // 工作区目录（MainForm/ChatPanel 在加载/修改配置时同步到这里）。
        // 建表/建 PPT 给相对路径时落在这里；下载文件也存这里。
        public static string WorkspaceRoot = "";
        // 最近一次成功产出的文件路径（附件卡片用）
        public static string LastProduct = "";

        // 同名产物的"已改名"通知：UI 注入。
        // 参数 = 用户原本想要的文件全路径（已存在、未被覆盖）。
        // ⚠️ 实现必须是**非阻塞**的（典型做法：BeginInvoke 到 UI 线程后在气泡里加一条提示）。
        // 绝不能在这里弹模态对话框：本方法在 agent 的后台线程上被调用
        // （ChatPanel.DoSend → new Thread → AgentLoop.Run → CreateSpreadsheet → 这里），
        // 后台线程弹模态框会自建消息泵（行为不可预期），且用户不点就会一直挂住 worker 线程。
        // 可空：CLI/无界面时不做任何通知。
        public static Action<string> OnOverwriteAvoided = null;

        // 产物路径避让：目标已存在时不覆盖，自动加 (2)/(3) 后缀，并通知 UI。
        // 这是纯函数式的"绝不丢数据"策略：旧文件永远保留，新产物总有一个不重名的落点。
        // 之所以不问用户：询问需要阻塞后台线程（见 OnOverwriteAvoided 注释）。
        // 用户若要覆盖，删掉旧文件后让 agent 重做即可。
        static string AvoidOverwrite(string p)
        {
            if (p == null || !File.Exists(p)) return p;
            string dir = Path.GetDirectoryName(p);
            string baseName = Path.GetFileNameWithoutExtension(p);
            string ext = Path.GetExtension(p);
            string chosen = null;
            for (int i = 2; i < 1000; i++)
            {
                string cand = Path.Combine(dir, baseName + "(" + i + ")" + ext);
                if (!File.Exists(cand)) { chosen = cand; break; }
            }
            if (chosen == null)
                chosen = Path.Combine(dir, baseName + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ext);
            // 通知是"锦上添花"：回调自身抛异常绝不能影响产物落盘
            if (OnOverwriteAvoided != null)
            {
                try { OnOverwriteAvoided(p); } catch { }
            }
            AuditLog.Record("file_write", "avoid_overwrite existing=" + p + " -> " + chosen);
            return chosen;
        }
        // 工具 schema（OpenAI function calling 格式；智谱等兼容端点同格式）
        public static string SchemasJson()
        {
            return SchemasJson(false);
        }

        // task_plan 计划栏插件的 schema（内置插件页可关）
        public const string TaskPlanSchema =
            "{\"type\":\"function\",\"function\":{\"name\":\"task_plan\",\"description\":\"复杂任务的计划清单（界面右侧计划栏向用户展示进度）。遇到 3 步以上的任务先 start 列出步骤，每完成一步用 done 勾掉，全部完成后用 finish。\",\"parameters\":{\"type\":\"object\",\"properties\":{\"action\":{\"type\":\"string\",\"description\":\"start=创建计划；done=勾选已完成步骤；finish=全部完成\"},\"data\":{\"type\":\"string\",\"description\":\"start: 步骤列表（每行一步）；done: 步骤序号或包含的关键词\"}},\"required\":[\"action\"]}}}";

        public static string SchemasJson(bool includePlan)
        {
            string s = "[" +
                "{\"type\":\"function\",\"function\":{\"name\":\"read_text_file\",\"description\":\"读取本地文件内容（文本类 txt/md/csv/log/json/xml/代码等；也支持 PDF——自动提取文字层）。xlsx/xls/doc/ppt 请用 convert_document 转换后再读。\",\"parameters\":{\"type\":\"object\",\"properties\":{\"path\":{\"type\":\"string\",\"description\":\"文件的绝对路径\"}},\"required\":[\"path\"]}}}," +
                "{\"type\":\"function\",\"function\":{\"name\":\"list_directory\",\"description\":\"列出某个文件夹里的文件和子文件夹。\",\"parameters\":{\"type\":\"object\",\"properties\":{\"path\":{\"type\":\"string\",\"description\":\"文件夹的绝对路径\"}},\"required\":[\"path\"]}}}," +
                "{\"type\":\"function\",\"function\":{\"name\":\"download_file\",\"description\":\"从 http/https 网址下载文件，保存到工作区目录并返回保存路径。\",\"parameters\":{\"type\":\"object\",\"properties\":{\"url\":{\"type\":\"string\",\"description\":\"下载链接（http/https）\"},\"filename\":{\"type\":\"string\",\"description\":\"可选：保存的文件名\"}},\"required\":[\"url\"]}}}," +
                "{\"type\":\"function\",\"function\":{\"name\":\"convert_document\",\"description\":\"把 office 文档转格式：doc/docx/ppt/pptx/xls/xlsx 转 pdf，xlsx 转 csv，csv 转 xlsx。\",\"parameters\":{\"type\":\"object\",\"properties\":{\"input\":{\"type\":\"string\",\"description\":\"输入文件绝对路径\"},\"target\":{\"type\":\"string\",\"description\":\"目标格式：pdf 或 csv 或 xlsx\"}},\"required\":[\"input\",\"target\"]}}}," +
                "{\"type\":\"function\",\"function\":{\"name\":\"create_spreadsheet\",\"description\":\"创建全新的 Excel 表格（.xlsx）。用 csv 参数提供表格内容：标准 CSV 文本，第一行是表头，用 \\n 表示换行。数字会自动识别为数值。path 给文件名（相对路径）时会保存到工作区目录。\",\"parameters\":{\"type\":\"object\",\"properties\":{\"path\":{\"type\":\"string\",\"description\":\"输出的 .xlsx 路径（文件名则存到工作区）\"},\"csv\":{\"type\":\"string\",\"description\":\"表格内容（CSV 文本，第一行表头）\"}},\"required\":[\"path\",\"csv\"]}}}," +
                "{\"type\":\"function\",\"function\":{\"name\":\"create_presentation\",\"description\":\"创建全新的 PPT 演示文稿（.pptx）。用 outline 参数提供每页内容：页与页之间用 ;; 分隔，每页格式为 标题|要点1;要点2;要点3。path 给文件名（相对路径）时会保存到工作区目录。\",\"parameters\":{\"type\":\"object\",\"properties\":{\"path\":{\"type\":\"string\",\"description\":\"输出的 .pptx 路径（文件名则存到工作区）\"},\"outline\":{\"type\":\"string\",\"description\":\"每页内容：标题|要点1;要点2 ;; 下一页标题|要点\"}},\"required\":[\"path\",\"outline\"]}}}," +
                "{\"type\":\"function\",\"function\":{\"name\":\"create_formula_workbook\",\"description\":\"生成带公式的 Excel 工作簿（.xlsx），用户填数即自动计算。优先用模板：template=payroll 工资表标准套账（社保/公积金/个税全公式，数据行CSV列序:姓名,部门,基本工资,岗位津贴,加班费）；template=vat 增值税台账（CSV列序:日期,摘要,类型(只填销项/进项),金额(不含税),税率）；template=ledger 流水账（CSV列序:日期,摘要,类别,收入,支出）。模板自带汇总页，改明细汇总自动变；可选 params 覆盖参数（工资表 pensionRate/medicalRate/unemploymentRate/housingFundRate/taxThreshold/blankRows，流水账 openingBalance），如 pensionRate=0.08;housingFundRate=0.12。自由定制用 sheets+summary（规格见参数说明）。\",\"parameters\":{\"type\":\"object\",\"properties\":{\"path\":{\"type\":\"string\",\"description\":\"输出的 .xlsx 路径（文件名则存到工作区）\"},\"template\":{\"type\":\"string\",\"description\":\"payroll|vat|ledger 三选一\"},\"rows\":{\"type\":\"string\",\"description\":\"模板数据行 CSV 文本（列序见模板说明；首行是列名时会被自动忽略）\"},\"params\":{\"type\":\"string\",\"description\":\"可选：参数覆盖 key=value;分号分隔\"},\"sheets\":{\"type\":\"string\",\"description\":\"自由模式（与 template 二选一）：sheets 数组 JSON 文本，每项 {name:表名, header:[列1,列2], rows:[[a,1],[b,2]], formulaCols:[{col:F, formula:=D{r}-E{r}}], blankRows:50, totalRow:true, widths:[10,20]}；{r} 代表当前行号，公式以 = 开头\"},\"summary\":{\"type\":\"string\",\"description\":\"可选（配合 sheets）：汇总页配置 JSON 文本 {source:明细表名, groupCol:C, labelHeader:类别, sumCols:[{col:D, header:收入},{col:E, header:支出}]}——按分组列 SUMIF 自动生成汇总，改明细汇总自动变\"}},\"required\":[\"path\"]}}}," +
                "{\"type\":\"function\",\"function\":{\"name\":\"excel_formula_reference\",\"description\":\"查询内置 Excel 公式大全（语法+中文说明+示例，离线）。用户问 Excel 公式怎么写、怎么算个税/折旧/条件求和时，先调用本工具查标准语法再回答，不要凭记忆给出可能出错的公式。不带参数时返回全部分类概览。\",\"parameters\":{\"type\":\"object\",\"properties\":{\"keyword\":{\"type\":\"string\",\"description\":\"关键词（模糊匹配名称/语法/说明，如：折旧、查找、求和、个税、账龄）\"},\"category\":{\"type\":\"string\",\"description\":\"精确分类：数学与三角/统计/逻辑/文本/日期与时间/查找与引用/财务会计/其他实用\"},\"name\":{\"type\":\"string\",\"description\":\"公式名（如 VLOOKUP、SUMIF）\"}}}}}," +
                "{\"type\":\"function\",\"function\":{\"name\":\"repair_environment\",\"description\":\"启动环境修复器（弹 UAC 提权），检测并离线安装缺失的系统组件（KB/.NET/VC++/Python/LibreOffice）。用户需在 UAC 与引导器窗口中确认。\",\"parameters\":{\"type\":\"object\",\"properties\":{}}}}" +
                "]";
            if (includePlan) s = s.Substring(0, s.Length - 1) + "," + TaskPlanSchema + "]";
            return s;
        }

        // 技能工具 schema（规划层第一期）：把技能注册表投影成模型可见的工具。
        // 返回**不含**外层 [] 的片段；无技能时返回空串（调用方不得多插逗号）。
        // 单独成方法而不是并进 SchemasJson：SchemasJson 有多处调用者，
        // 让"是否附带技能"成为显式选择，避免自检/降级路径意外拿到一堆技能工具。
        public static string SkillSchemasJson(string root)
        {
            try { return SkillToolBridge.BuildSchemas(root); }
            catch { return ""; }   // 投影失败绝不能连累基础工具可用性
        }

        // 分发执行。返回结果文本；ok=false 表示工具执行失败（文本里带原因）。
        // products：非 null 时收集本回合产出的文件路径（会话附件卡片用）
        public static string Dispatch(string name, string argsJson, AppConfig cfg, ConvertEngine conv, List<string> products, out bool ok)
        {
            ok = true;
            try
            {
                Dictionary<string, string> a = ParseArgs(argsJson);
                switch (name)
                {
                    case "read_text_file": return ReadTextFile(GetStr(a, "path"));
                    case "list_directory": return ListDir(GetStr(a, "path"));
                    case "download_file": {
                        string r1 = DownloadFile(GetStr(a, "url"), GetStr(a, "filename"), cfg, out ok);
                        if (ok && products != null) products.Add(LastProduct);
                        return r1; }
                    case "convert_document": {
                        string r2 = ConvertDoc(GetStr(a, "input"), GetStr(a, "target"), conv, out ok);
                        if (ok && products != null) products.Add(LastProduct);
                        return r2; }
                    case "create_spreadsheet": {
                        string r3 = CreateSpreadsheet(GetStr(a, "path"), GetStr(a, "csv"), out ok);
                        if (ok && products != null) products.Add(LastProduct);
                        return r3; }
                    case "create_presentation": {
                        string r4 = CreatePresentation(GetStr(a, "path"), GetStr(a, "outline"), out ok);
                        if (ok && products != null) products.Add(LastProduct);
                        return r4; }
                    case "create_formula_workbook": {
                        // 嵌套规格（sheets/summary）由 FormulaWorkbook 内的 JsonVal 解析，不走扁平 ParseArgs
                        string r5 = FormulaWorkbook.Create(argsJson, out ok);
                        if (ok && products != null) products.Add(LastProduct);
                        return r5; }
                    case "excel_formula_reference":
                        return FormulaReference.Lookup(GetStr(a, "keyword"), GetStr(a, "category"), GetStr(a, "name"), out ok);
                    case "repair_environment": return RepairEnvironment();
                    default:
                        // 技能工具（skill_<id>_<action>）：投影自技能注册表，见 SkillToolBridge。
                        // 参数只进 request.json（文件），命令行保持编译期字面量。
                        if (name != null && name.StartsWith(SkillToolBridge.Prefix, StringComparison.Ordinal))
                        {
                            string root = EnvDetect.FindRoot();
                            SkillActionSpec sp = SkillToolBridge.Resolve(root, name);
                            if (sp == null)
                            {
                                ok = false;
                                return "未知技能工具: " + name + "。请只用当前声明的工具。";
                            }
                            return SkillToolBridge.Invoke(root, sp, a, products, out ok);
                        }
                        ok = false;
                        return "未知工具: " + name;
                }
            }
            catch (Exception ex)
            {
                ok = false;
                return ExplainToolError(name, ex);
            }
        }

        // 工具异常 → 人话：讲清「为什么失败 + 可以怎么办」，而不是抛原始异常串。
        // 常见根因：缺 VC++ 运行库（Python/LibreOffice 起不来）、Office COM 未安装或挂死、
        //           文件被占用、权限不足。
        static string ExplainToolError(string tool, Exception ex)
        {
            string raw = ex == null ? "" : (ex.Message == null ? ex.GetType().Name : ex.Message);
            string hint = "";
            string low = (raw ?? "").ToLowerInvariant();
            if (low.Contains("msvcp") || low.Contains("vcruntime") || low.Contains("msvcr") ||
                low.Contains("api-ms-win-crt") || low.Contains("dll") && low.Contains("not found"))
            {
                hint = "看起来缺 Visual C++ 运行库。可以让 agent 调用 repair_environment 工具，" +
                       "或在引导器里点「安装缺失组件」补装 VC++ 后重试。";
            }
            else if (low.Contains("0x80070005") || low.Contains("access") && low.Contains("denied") ||
                     low.Contains("拒绝访问"))
            {
                hint = "文件或目录没有写权限。请确认目标文件没有被 Excel/WPS 打开，" +
                       "或换一个有写权限的输出目录后重试。";
            }
            else if (low.Contains("being used") || low.Contains("被占用") || low.Contains("sharing violation"))
            {
                hint = "文件正被其他程序占用（通常是 Excel/WPS 打开了它）。关闭后重试即可。";
            }
            else if (low.Contains("ocr") || low.Contains("扫描") || low.Contains("no text"))
            {
                hint = "该 PDF 可能没有文字层（扫描件）。可以先用其他工具转成图片，再让我看图识别。";
            }
            else if (low.Contains("out of memory") || low.Contains("内存"))
            {
                hint = "文件太大导致内存不足。建议先拆分文件，或改用核对/汇总功能分批处理。";
            }
            if (hint.Length == 0)
                return "工具 " + tool + " 执行失败：" + raw + "。可以换个参数或换一种做法重试；若反复失败请把这句话发给开发者。";
            return "工具 " + tool + " 执行失败：" + raw + "\n可能的原因与建议：" + hint;
        }

        static Dictionary<string, string> ParseArgs(string argsJson)
        {
            if (argsJson == null || argsJson.Trim().Length == 0) return new Dictionary<string, string>();
            List<Dictionary<string, string>> objs = MiniJson.ParseObjects(argsJson);
            if (objs.Count > 0) return objs[0];
            return new Dictionary<string, string>();
        }

        static string GetStr(Dictionary<string, string> a, string key)
        {
            if (a == null || !a.ContainsKey(key)) return "";
            string v = a[key];
            return v == null ? "" : v.Trim();
        }

        // ---------- read_text_file ----------

        static readonly string[] TextExts = new string[] {
            ".txt", ".md", ".csv", ".log", ".json", ".xml", ".ini", ".cfg", ".conf",
            ".py", ".ps1", ".bat", ".cmd", ".yml", ".yaml", ".sql", ".html", ".htm", ".tsv" };

        // 会话页附文件上下文用的公开入口
        public static string ReadTextFilePublic(string path, out bool ok)
        {
            ok = false;
            if (path == null || path.Length == 0 || !File.Exists(path)) return "";
            string ext = (Path.GetExtension(path) ?? "").ToLowerInvariant();
            foreach (string e in TextExts) { if (e == ext) { ok = true; break; } }
            if (!ok) return "";
            ok = true;
            return ReadTextFile(path);
        }

        static string ReadTextFile(string path)
        {
            if (path.Length == 0) return "缺少 path 参数";
            if (!File.Exists(path)) return "文件不存在: " + path;
            string ext = (Path.GetExtension(path) ?? "").ToLowerInvariant();
            // PDF 走 pdfium 文本层提取（无需外部阅读器）；扫描件无文字层时给出明确提示
            if (ext == ".pdf")
            {
                string root = EnvDetect.FindRoot();
                string perr;
                string ptxt = PdfTextReader.Read(path, root, out perr);
                if (ptxt == null)
                    return "读取 PDF 失败：" + (perr == null ? "未知原因" : perr)
                        + "。可尝试用 convert_document 转换为 txt/pdf。";
                return ptxt;
            }
            bool isText = false;
            foreach (string e in TextExts) { if (e == ext) { isText = true; break; } }
            if (!isText) return "不是文本类文件（" + ext + "）。表格请用 convert_document 转换后再读，或直接拖进会话窗口。";
            FileInfo fi = new FileInfo(path);
            const int MaxBytes = 64 * 1024;
            int n = (int)Math.Min(fi.Length, MaxBytes);
            byte[] raw = new byte[n];
            using (FileStream fs = File.OpenRead(path)) { fs.Read(raw, 0, n); }   // 只读，不改动原文件
            string text = DecodeText(raw);
            string trunc = fi.Length > MaxBytes ? "\n…[文件过大，仅显示前 64KB]…" : "";
            return "文件 " + path + "（" + fi.Length + " 字节）内容：\n" + text + trunc;
        }

        static string DecodeText(byte[] raw)
        {
            if (raw.Length >= 3 && raw[0] == 0xEF && raw[1] == 0xBB && raw[2] == 0xBF)
                return new UTF8Encoding(false).GetString(raw, 3, raw.Length - 3);
            if (raw.Length >= 2 && raw[0] == 0xFF && raw[1] == 0xFE)
                return Encoding.Unicode.GetString(raw, 2, raw.Length - 2);
            try
            {
                Encoding strict = new UTF8Encoding(false, true);
                return strict.GetString(raw);
            }
            catch
            {
                try { return Encoding.GetEncoding(936).GetString(raw); }
                catch { return Encoding.Default.GetString(raw); }
            }
        }

        // ---------- list_directory ----------

        static string ListDir(string path)
        {
            if (path.Length == 0) return "缺少 path 参数";
            if (!Directory.Exists(path)) return "文件夹不存在: " + path;
            StringBuilder sb = new StringBuilder();
            sb.Append("目录 ").Append(path).Append("：\n");
            try
            {
                string[] dirs = Directory.GetDirectories(path);
                int shown = 0;
                foreach (string d in dirs)
                {
                    if (shown >= 100) { sb.Append("…（子目录超过 100 个，已截断）\n"); break; }
                    sb.Append("[目录] ").Append(Path.GetFileName(d)).Append("\n");
                    shown++;
                }
                string[] files = Directory.GetFiles(path);
                shown = 0;
                foreach (string f in files)
                {
                    if (shown >= 200) { sb.Append("…（文件超过 200 个，已截断）\n"); break; }
                    FileInfo fi = new FileInfo(f);
                    sb.Append(Path.GetFileName(f)).Append("  （").Append(fi.Length).Append(" 字节）\n");
                    shown++;
                }
            }
            catch (Exception ex) { return "读取目录失败: " + ex.Message; }
            return sb.ToString();
        }

        // ---------- download_file ----------

        static string DownloadFile(string url, string filename, AppConfig cfg, out bool ok)
        {
            ok = false;
            if (url == null || url.Trim().Length == 0) return "缺少 url 参数";
            url = url.Trim();
            string lower = url.ToLowerInvariant();
            if (!lower.StartsWith("http://", StringComparison.Ordinal) && !lower.StartsWith("https://", StringComparison.Ordinal))
                return "仅支持 http/https 链接";
            // HostGuard：拒绝内网/环回/保留地址（AllowLan 仅企业自建网关显式放行）
            string guard = HostGuard.Check(url, cfg != null && cfg.AllowLan);
            if (guard != null) return "安全守卫拒绝该地址: " + guard;

            // 下载落盘：优先工作区目录，未配置时回退 文档\OfficeAgentDownloads
            string dir = WorkspaceRoot;
            if (dir == null || dir.Length == 0)
                dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "OfficeAgentDownloads");
            Directory.CreateDirectory(dir);
            string name = filename != null && filename.Trim().Length > 0 ? SanitizeName(filename.Trim()) : SanitizeName(UriFileName(url));
            if (name.Length == 0) name = "download_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".bin";
            string dst = Path.Combine(dir, name);
            if (File.Exists(dst)) dst = Path.Combine(dir,
                Path.GetFileNameWithoutExtension(name) + "_" + DateTime.Now.ToString("HHmmss") + Path.GetExtension(name));

            try
            {
                try { ServicePointManager.SecurityProtocol |= (SecurityProtocolType)3072; } catch { }
                string current = url;
                long total = 0;
                for (int hop = 0; hop < 4; hop++)
                {
                    string g = HostGuard.Check(current, cfg != null && cfg.AllowLan);
                    if (g != null) return "重定向后安全守卫拒绝: " + g;
                    HttpWebRequest req = (HttpWebRequest)WebRequest.Create(current);
                    req.Method = "GET";
                    req.Timeout = 60000;
                    req.ReadWriteTimeout = 120000;
                    req.AllowAutoRedirect = false;
                    req.KeepAlive = false;
                    req.UserAgent = "OfficeAgent/0.4";
                    HttpWebResponse resp;
                    try { resp = (HttpWebResponse)req.GetResponse(); }
                    catch (WebException we)
                    {
                        HttpWebResponse er = we.Response as HttpWebResponse;
                        if (er != null && (int)er.StatusCode >= 300 && (int)er.StatusCode < 400)
                        {
                            string loc = er.Headers[HttpResponseHeader.Location];
                            if (loc == null) return "重定向缺少 Location";
                            current = new Uri(new Uri(current), loc).ToString();
                            continue;
                        }
                        return "下载失败: " + we.Message;
                    }
                    using (resp)
                    {
                        if ((int)resp.StatusCode >= 300) return "HTTP " + (int)resp.StatusCode;
                        long len = resp.ContentLength;
                        if (len > 200L * 1024 * 1024) return "文件超过 200MB 上限，拒绝下载";
                        using (Stream s = resp.GetResponseStream())
                        using (FileStream fs = new FileStream(dst, FileMode.Create, FileAccess.Write))
                        {
                            byte[] buf = new byte[65536];
                            while (true)
                            {
                                int n = s.Read(buf, 0, buf.Length);
                                if (n <= 0) break;
                                total += n;
                                if (total > 200L * 1024 * 1024)
                                {
                                    fs.Close();
                                    try { File.Delete(dst); } catch { }
                                    return "文件超过 200MB 上限，已中止";
                                }
                                fs.Write(buf, 0, n);
                            }
                        }
                    }
                    break;
                }
                ok = true;
                LastProduct = dst;
                AuditLog.Record("file_write", "agent_download " + dst + " (" + total + " bytes)");
                return "已下载保存到: " + dst + "（" + total + " 字节）";
            }
            catch (Exception ex)
            {
                try { if (File.Exists(dst)) File.Delete(dst); } catch { }
                return "下载失败: " + ex.Message;
            }
        }

        static string UriFileName(string url)
        {
            try
            {
                string p = new Uri(url).AbsolutePath;
                string f = p.Substring(p.LastIndexOf('/') + 1);
                return Uri.UnescapeDataString(f);
            }
            catch { return ""; }
        }

        static string SanitizeName(string name)
        {
            if (name == null) return "";
            StringBuilder sb = new StringBuilder();
            foreach (char c in name)
            {
                bool bad = c == '\\' || c == '/' || c == ':' || c == '*' || c == '?' || c == '"' || c == '<' || c == '>' || c == '|' || c < 0x20;
                if (!bad) sb.Append(c);
            }
            return sb.ToString().Trim();
        }

        // ---------- convert_document ----------

        static string ConvertDoc(string input, string target, ConvertEngine conv, out bool ok)
        {
            ok = false;
            if (input == null || input.Length == 0 || !File.Exists(input)) return "输入文件不存在: " + input;
            string t = (target == null ? "pdf" : target.Trim().ToLowerInvariant());
            ConvTarget ct;
            if (t == "pdf") ct = ConvTarget.Pdf;
            else if (t == "csv") ct = ConvTarget.Csv;
            else if (t == "xlsx") ct = ConvTarget.Xlsx;
            else return "不支持的目标格式: " + t + "（仅 pdf/csv/xlsx）";
            if (conv == null) return "转换引擎未就绪（内部状态异常）。请让用户改用「批量转换」页手动转换，或调用 repair_environment 修复环境后重试。";
            string outPath;
            string err = conv.Convert(input, ct, out outPath);
            if (err != null) return ExplainConvertError(input, t, err);
            ok = true;
            LastProduct = outPath;
            return "转换完成，输出文件: " + outPath;
        }

        // 转换失败 → 按格式给出可操作建议（xlsx→pdf 依赖 LibreOffice/Office，是最常见的失败点）
        static string ExplainConvertError(string input, string target, string err)
        {
            string hint;
            if (target == "pdf")
                hint = "转换 PDF 需要 LibreOffice 或本机 Office。请在引导器里确认 LibreOffice 已安装（缺失组件会显示为待补全）；" +
                       "若本机有 Office 但仍失败，可能是 Office 隐藏会话挂死，可改用 LibreOffice 重试。";
            else if (target == "csv" || target == "xlsx")
                hint = "请确认源文件是有效的表格且没有加密/受损；若文件正被 Excel/WPS 打开，请先关闭再重试。";
            else
                hint = "请确认源文件格式正确且未被其他程序占用。";
            return "转换失败：" + err + "\n可能的原因与建议：" + hint;
        }

        // ---------- create_spreadsheet / create_presentation ----------

        // 输出路径围栏（建表/建 PPT/带公式工作簿共用）：逐段拒绝 ".."、扩展名白名单。
        // 相对路径按工作区目录解析（模型常给文件名不给全路径，此前一律拒绝体验差）。
        internal static string SafeOutputPath(string path, string ext, out string err)
        {
            err = null;
            string p = (path == null ? "" : path.Trim());
            if (p.Length == 0) { err = "缺少输出路径"; return null; }
            if (!Path.IsPathRooted(p))
            {
                string root = WorkspaceRoot;
                if (root == null || root.Length == 0)
                    root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "OfficeAgentFiles");
                p = Path.Combine(root, p);
            }
            foreach (string seg in p.Replace('/', '\\').Split('\\'))
            {
                if (seg == "..") { err = "输出路径不允许包含 .."; return null; }
            }
            try { p = Path.GetFullPath(p); }
            catch (Exception ex) { err = "输出路径非法: " + ex.Message; return null; }
            if (!p.EndsWith(ext, StringComparison.OrdinalIgnoreCase)) { err = "输出路径必须是 " + ext + " 文件"; return null; }
            try { string parent = Path.GetDirectoryName(p); if (parent != null && parent.Length > 0) Directory.CreateDirectory(parent); } catch { }
            // 同名文件已存在：自动改名（不改名会静默毁掉用户旧产物）；
            // 不问用户，因为询问要阻塞 agent 的后台线程——详见 AvoidOverwrite 上方注释。
            p = AvoidOverwrite(p);
            return p;
        }

        // CSV 文本 → xlsx（MiniCsv 解析 + MiniXlsxWrite 写带表头样式的表格）
        public static string CreateSpreadsheet(string path, string csv, out bool ok)
        {
            ok = false;
            string fenceErr;
            string p = SafeOutputPath(path, ".xlsx", out fenceErr);
            if (p == null) return fenceErr;
            List<string[]> rows;
            try { rows = MiniCsv.Parse(csv == null ? "" : csv); }
            catch (Exception ex) { return "表格内容解析失败: " + ex.Message; }
            if (rows.Count == 0) return "表格内容为空";
            ReportSheet sheet = new ReportSheet("Sheet1");
            for (int ri = 0; ri < rows.Count; ri++)
            {
                List<ReportCell> row = new List<ReportCell>();
                foreach (string cell in rows[ri])
                {
                    if (ri == 0) { row.Add(ReportCell.S(cell)); continue; }
                    double num;
                    string t = (cell == null ? "" : cell.Trim());
                    if (t.Length > 0 && double.TryParse(t, out num)) row.Add(ReportCell.N(num, "n"));
                    else row.Add(ReportCell.S(t));
                }
                sheet.Rows.Add(row);
            }
            sheet.FreezeRows = 1;
            List<ReportSheet> sheets = new List<ReportSheet>();
            sheets.Add(sheet);
            string saveErr = MiniXlsxWrite.Save(p, sheets);
            if (saveErr != null) return "生成失败: " + saveErr;
            ok = true;
            LastProduct = p;
            AuditLog.Record("file_write", "agent_make_xlsx " + p);
            return "已生成 Excel 表格（" + rows.Count + " 行）: " + p;
        }

        // outline → pptx（模板替换式，见 PptWriter）
        public static string CreatePresentation(string path, string outline, out bool ok)
        {
            ok = false;
            string fenceErr;
            string p = SafeOutputPath(path, ".pptx", out fenceErr);
            if (p == null) return fenceErr;
            if (outline == null || outline.Trim().Length == 0) return "演示文稿内容为空";
            List<string[]> slides = new List<string[]>();
            foreach (string page in outline.Split(new string[] { ";;" }, StringSplitOptions.RemoveEmptyEntries))
            {
                string seg = page.Trim();
                if (seg.Length == 0) continue;
                string title = seg, body = "";
                int bar = seg.IndexOf('|');
                if (bar >= 0) { title = seg.Substring(0, bar); body = seg.Substring(bar + 1); }
                body = body.Replace(";", "\n").Replace("；", "\n");
                slides.Add(new string[] { title.Trim(), body });
            }
            if (slides.Count == 0) return "没有可用的幻灯片内容";
            string root = EnvDetect.FindRoot();
            // 模板目录：优先 templates\make-ppt\（新位置，与 skills\ 区分开——它只是模板，
            // 不是可执行技能）；回退 skills\make-ppt\ 以兼容升级前的旧安装。
            string template = Path.Combine(Path.Combine(root, "templates"), Path.Combine("make-ppt", "template.pptx"));
            if (!File.Exists(template))
                template = Path.Combine(Path.Combine(root, "skills"), Path.Combine("make-ppt", "template.pptx"));
            string saveErr = PptWriter.Save(template, p, slides);
            if (saveErr != null) return "生成失败: " + saveErr;
            ok = true;
            LastProduct = p;
            AuditLog.Record("file_write", "agent_make_ppt " + p);
            return "已生成 PPT（" + slides.Count + " 页）: " + p;
        }

        // ---------- repair_environment ----------

        static string RepairEnvironment()
        {
            AuditLog.Record("action_started", "repair_environment via agent tool");
            string err = RepairLauncher.Start(EnvDetect.FindRoot());
            if (err != null) return err;
            return "已弹出环境修复器（系统会先弹 UAC 提权确认）。请在引导器窗口点击「安装缺失组件」，完成后回来告诉我结果。";
        }
    }
}
