// OfficeAgent 主程序入口（.NET Framework 4.8 / Win7 SP1+）
// 参数：/selftest  控制台输出环境检测后退出（M0 冒烟用）
//       /convert <in> <pdf|csv|xlsx>            无头转换（三级总线）
//       /renderpdf <in.pdf> <out.png>           无头渲染（pdfium）
//       /recon <A> <B> [选项]                   两表核对（M2，详见 RunRecon）
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using OfficeAgent.Core;

namespace OfficeAgent.Host
{
    internal static class Program
    {
        [DllImport("kernel32.dll")]
        static extern bool AllocConsole();
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool AttachConsole(int processId);
        const int ATTACH_PARENT_PROCESS = -1;

        [DllImport("user32.dll")]
        static extern bool SetProcessDPIAware();

        [STAThread]
        static int Main(string[] args)
        {
            bool selftest = false;
            string convertInput = null, convertTarget = null, renderInput = null, renderOutput = null;
            bool guardTest = false;
            bool maskTest = false;
            bool intentTest = false;
            bool suggestTest = false;
            bool caretTest = false;
            string gridTestInput = null;
            bool skillTest = false;
            string[] skillRunArgs = null;
            string[] agentTestArgs = null;
            string[] reconArgs = null;
            string[] mergeArgs = null;
            string[] invoiceArgs = null;
            string pdfTextInput = null;

            for (int i = 0; i < args.Length; i++)
            {
                string a = args[i].ToLowerInvariant();
                if (a == "/selftest") selftest = true;
                else if (a == "/guardtest") guardTest = true;
                else if (a == "/masktest") maskTest = true;
                else if (a == "/intents") intentTest = true;
                else if (a == "/suggesttest") suggestTest = true;
                else if (a == "/carettest") caretTest = true;
                else if (a == "/gridtest" && i + 1 < args.Length) { gridTestInput = args[i + 1]; i += 1; }
                else if (a == "/skilltest") skillTest = true;
                else if (a == "/skillrun")
                {
                    List<string> rest = new List<string>();
                    for (int j = i + 1; j < args.Length; j++) rest.Add(args[j]);
                    skillRunArgs = rest.ToArray();
                    break;
                }
                else if (a == "/agenttest")
                {
                    List<string> rest = new List<string>();
                    for (int j = i + 1; j < args.Length; j++) rest.Add(args[j]);
                    agentTestArgs = rest.ToArray();
                    break;
                }
                else if (a == "/convert" && i + 2 < args.Length) { convertInput = args[i + 1]; convertTarget = args[i + 2]; i += 2; }
                else if (a == "/renderpdf" && i + 2 < args.Length) { renderInput = args[i + 1]; renderOutput = args[i + 2]; i += 2; }
                else if (a == "/recon")
                {
                    List<string> rest = new List<string>();
                    for (int j = i + 1; j < args.Length; j++) rest.Add(args[j]);
                    reconArgs = rest.ToArray();
                    break;
                }
                else if (a == "/pdftext" && i + 1 < args.Length) { pdfTextInput = args[i + 1]; i += 1; }
                else if (a == "/merge" || a == "/invoices")
                {
                    List<string> rest = new List<string>();
                    for (int j = i + 1; j < args.Length; j++) rest.Add(args[j]);
                    if (a == "/merge") mergeArgs = rest.ToArray(); else invoiceArgs = rest.ToArray();
                    break;
                }
            }

            // CLI 无头链同样落审计（设计方案 §7.3：每个动作可审计；GUI 的 app_start 在 MainForm）
            if (selftest || guardTest || maskTest || intentTest || suggestTest || skillTest || caretTest
                || gridTestInput != null
                || skillRunArgs != null || agentTestArgs != null || reconArgs != null || mergeArgs != null
                || invoiceArgs != null || convertInput != null || renderInput != null
                || pdfTextInput != null)
            {
                AuditLog.Record("app_start", "cli");
            }

            if (guardTest)
            {
                try { AttachConsole(ATTACH_PARENT_PROCESS); } catch { }
                try { AllocConsole(); } catch { }
                try { Console.OutputEncoding = Encoding.UTF8; } catch { }
                return RunGuardTest();
            }

            if (maskTest)
            {
                try { AttachConsole(ATTACH_PARENT_PROCESS); } catch { }
                try { AllocConsole(); } catch { }
                try { Console.OutputEncoding = Encoding.UTF8; } catch { }
                return RunMaskTest();
            }

            if (intentTest)
            {
                try { AttachConsole(ATTACH_PARENT_PROCESS); } catch { }
                try { AllocConsole(); } catch { }
                try { Console.OutputEncoding = Encoding.UTF8; } catch { }
                return RunIntents();
            }

            if (suggestTest)
            {
                try { AttachConsole(ATTACH_PARENT_PROCESS); } catch { }
                try { AllocConsole(); } catch { }
                try { Console.OutputEncoding = Encoding.UTF8; } catch { }
                return RunSuggestTest();
            }

            if (gridTestInput != null)
            {
                try { AttachConsole(ATTACH_PARENT_PROCESS); } catch { }
                try { AllocConsole(); } catch { }
                try { Console.OutputEncoding = Encoding.UTF8; } catch { }
                return RunGridTest(gridTestInput);
            }

            if (caretTest)
            {
                try { AttachConsole(ATTACH_PARENT_PROCESS); } catch { }
                try { AllocConsole(); } catch { }
                try { Console.OutputEncoding = Encoding.UTF8; } catch { }
                return RunCaretTest();
            }

            if (skillTest)
            {
                try { AttachConsole(ATTACH_PARENT_PROCESS); } catch { }
                try { AllocConsole(); } catch { }
                try { Console.OutputEncoding = Encoding.UTF8; } catch { }
                return RunSkillTest();
            }

            if (skillRunArgs != null)
            {
                try { AttachConsole(ATTACH_PARENT_PROCESS); } catch { }
                try { AllocConsole(); } catch { }
                try { Console.OutputEncoding = Encoding.UTF8; } catch { }
                return RunSkillRun(skillRunArgs);
            }

            if (agentTestArgs != null)
            {
                try { AttachConsole(ATTACH_PARENT_PROCESS); } catch { }
                try { AllocConsole(); } catch { }
                try { Console.OutputEncoding = Encoding.UTF8; } catch { }
                return RunAgentTest(agentTestArgs);
            }

            if (reconArgs != null)
            {
                try { AttachConsole(ATTACH_PARENT_PROCESS); } catch { }
                try { AllocConsole(); } catch { }
                try { Console.OutputEncoding = Encoding.UTF8; } catch { }
                return RunRecon(reconArgs);
            }

            if (mergeArgs != null)
            {
                try { AttachConsole(ATTACH_PARENT_PROCESS); } catch { }
                try { AllocConsole(); } catch { }
                try { Console.OutputEncoding = Encoding.UTF8; } catch { }
                return RunMerge(mergeArgs);
            }

            if (invoiceArgs != null)
            {
                try { AttachConsole(ATTACH_PARENT_PROCESS); } catch { }
                try { AllocConsole(); } catch { }
                try { Console.OutputEncoding = Encoding.UTF8; } catch { }
                return RunInvoices(invoiceArgs);
            }

            if (pdfTextInput != null)
            {
                try { AttachConsole(ATTACH_PARENT_PROCESS); } catch { }
                try { AllocConsole(); } catch { }
                try { Console.OutputEncoding = Encoding.UTF8; } catch { }
                return RunPdfText(pdfTextInput);
            }

            if (selftest)
            {
                try { AttachConsole(ATTACH_PARENT_PROCESS); } catch { }
                try { AllocConsole(); } catch { }
                try { Console.OutputEncoding = Encoding.UTF8; } catch { }
                string root = EnvDetect.FindRoot();
                Console.WriteLine("root=" + root);
                foreach (DetectItem it in EnvDetect.DetectAll(root))
                {
                    Console.WriteLine(string.Format("[{0}] {1}  {2}", EnvDetect.StateText(it.State), it.Name, it.Detail));
                }
                return 0;
            }

            if (convertInput != null)
            {
                try { AttachConsole(ATTACH_PARENT_PROCESS); } catch { }
                try { AllocConsole(); } catch { }
                try { Console.OutputEncoding = Encoding.UTF8; } catch { }
                return RunConvert(convertInput, convertTarget);
            }

            if (renderInput != null)
            {
                try { AttachConsole(ATTACH_PARENT_PROCESS); } catch { }
                try { AllocConsole(); } catch { }
                try { Console.OutputEncoding = Encoding.UTF8; } catch { }
                return RunRenderPdf(renderInput, renderOutput);
            }

            bool createdNew;
            using (Mutex mutex = new Mutex(true, "OfficeAgent.Host.Instance", out createdNew))
            {
                if (!createdNew) return 0;
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                try { SetProcessDPIAware(); } catch { }
                Application.Run(new MainForm());
            }
            return 0;
        }

        // HostGuard 无头自测：scheme/环回/私网/保留地址 拒绝 + 公网 IP 放行 + allowLan 放行
        // 光标定位无头回归（/carettest）：直接驱动真实 TextBox + CaretHelper。
        // 回归目标：修复前 EM_POSFROMCHAR 失败被当成坐标，SetCaretPos 落在 (0,-1)，
        // 表现为"打完字光标瞬移到第一个字符之前"——断言每处插入点的 X 必须随索引单调不减。
        // xlsx 网格读取无头回归（/gridtest <xlsx>）：断言表头保留、行数正确、行间不重复。
        // 回归目标：StreamRows 曾在循环内每节点重复提交同一行，LoadGrid 又复用 cell 缓冲，
        // 结果是表头丢失、所有行都等于最后一行（预览页表现为只剩一列、数据重复）。
        static int RunGridTest(string xlsxPath)
        {
            if (!System.IO.File.Exists(xlsxPath)) { Console.WriteLine("FAIL: 文件不存在 " + xlsxPath); return 2; }
            XlsxBook book = XlsxBook.Open(xlsxPath);
            int totalRows, usedCols;
            string[,] g = book.LoadGrid(0, 500, 64, out totalRows, out usedCols);
            book.Dispose();

            int rows = g.GetLength(0);
            Console.WriteLine("总行数=" + rows);
            Console.WriteLine("列数=" + usedCols);
            Console.WriteLine("表头行=" + (g[0, 0] == null ? "" : g[0, 0]) + "|" + (usedCols > 1 && g[0, 1] != null ? g[0, 1] : ""));

            int failed = 0;
            // 表头必须在第一行（曾整行丢失）
            bool headerOk = g[0, 0] != null && g[0, 0] == "科目";
            if (!headerOk) failed++;
            Console.WriteLine((headerOk ? "[OK]  " : "[FAIL] ") + "首行为表头（科目）实际=\"" + (g[0, 0] ?? "") + "\"");

            // 行数必须等于源文件 3 行（表头 + 2 数据行）
            bool countOk = rows == 3;
            if (!countOk) failed++;
            Console.WriteLine((countOk ? "[OK]  " : "[FAIL] ") + "行数=3 实际=" + rows);

            // 各数据行内容必须互不相同（曾全部等于最后一行）
            bool distinct = rows >= 3
                && (g[1, 0] ?? "") == "库存现金"
                && (g[2, 0] ?? "") == "银行存款";
            if (!distinct) failed++;
            Console.WriteLine((distinct ? "[OK]  " : "[FAIL] ") + "数据行区分 第2行=\"" + (rows > 1 ? g[1, 0] : "")
                + "\" 第3行=\"" + (rows > 2 ? g[2, 0] : "") + "\"");

            Console.WriteLine(failed == 0 ? "ALL PASS" : (failed + " FAILED"));
            return failed == 0 ? 0 : 2;
        }

        static int RunCaretTest()
        {
            int failed = 0;
            Form host = new Form();
            host.ShowInTaskbar = false;
            host.FormBorderStyle = FormBorderStyle.None;
            host.StartPosition = FormStartPosition.Manual;
            host.Location = new Point(-4000, -4000);   // 屏幕外，避免抢前台
            host.Size = new Size(700, 90);
            TextBox box = new TextBox();
            box.Multiline = true;
            box.BorderStyle = BorderStyle.None;
            box.Font = new Font("Microsoft YaHei UI", 9.75F);
            box.ScrollBars = ScrollBars.Vertical;
            box.Location = new Point(12, 8);
            box.Size = new Size(676, 50);
            host.Controls.Add(box);
            host.Show();
            Application.DoEvents();

            Font font = box.Font;

            string[] samples = new string[] {
                "a", "abc", "hello world", "中文测试内容",
                "line1\r\nline2", "mixed 中文 english 混杂"
            };

            foreach (string s in samples)
            {
                box.Text = s;
                Application.DoEvents();
                bool allValid = true;
                bool perLineMonotonic = true;
                // 逐行段检查 X 单调：跨行时 X 合法地回到行首，不能整体判单调
                int prevX = -1;
                int prevY = -1;
                for (int i = 0; i <= s.Length; i++)
                {
                    box.SelectionStart = i;
                    int x, y;
                    if (!CaretHelper.TryInsertionPoint(box, font, out x, out y)) { allValid = false; continue; }
                    if (x < 0 || y < 0) allValid = false;
                    if (y != prevY) { prevX = -1; prevY = y; }   // 换行 → 重置基线
                    if (x < prevX) perLineMonotonic = false;
                    prevX = x;
                }
                bool ok = allValid && perLineMonotonic;
                if (!ok) failed++;
                Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "文本 \"" + s.Replace("\r\n", "\\n")
                    + "\" len=" + s.Length + "  坐标有效=" + allValid + "  行内单调=" + perLineMonotonic);
            }

            // 关键用例：行尾（索引 == 长度）必须有可信坐标
            // —— 这正是原 bug 的触发点（EM_POSFROMCHAR(len) 返回 -1）
            box.Text = "abc";
            box.SelectionStart = 3;
            int ex, ey;
            bool endOk = CaretHelper.TryInsertionPoint(box, font, out ex, out ey) && ex > 0 && ey >= 0;
            if (!endOk) failed++;
            Console.WriteLine((endOk ? "[OK]  " : "[FAIL] ") + "行尾插入点 len=3 -> (" + ex + "," + ey + ")  要求 x>0");

            // 多行：第二行首个字符（索引 7）的 Y 必须大于第一行，且 X 回到行首
            box.Text = "line1\r\nline2";
            box.SelectionStart = 1;
            int l1x, l1y;
            CaretHelper.TryInsertionPoint(box, font, out l1x, out l1y);
            box.SelectionStart = 7;          // \r\n 之后的 'l'
            int l2x, l2y;
            CaretHelper.TryInsertionPoint(box, font, out l2x, out l2y);
            bool mlOk = l2y > l1y && l2x <= l1x;
            if (!mlOk) failed++;
            Console.WriteLine((mlOk ? "[OK]  " : "[FAIL] ") + "多行 第1行(x=" + l1x + ",y=" + l1y
                + ") → 第2行首(x=" + l2x + ",y=" + l2y + ")  要求 y 增大且 x 回到行首");

            // 对比：索引 0 与行尾必须落在不同 X（bug 时两者都塌到 0）
            box.SelectionStart = 0;
            int sx, sy;
            CaretHelper.TryInsertionPoint(box, font, out sx, out sy);
            bool distinct = ex > sx;
            if (!distinct) failed++;
            Console.WriteLine((distinct ? "[OK]  " : "[FAIL] ") + "行首 x=" + sx + " < 行尾 x=" + ex);

            // 空文本不应产生负坐标
            box.Text = "";
            box.SelectionStart = 0;
            int nx, ny;
            bool emptyOk = CaretHelper.TryInsertionPoint(box, font, out nx, out ny) && nx >= 0 && ny >= 0;
            if (!emptyOk) failed++;
            Console.WriteLine((emptyOk ? "[OK]  " : "[FAIL] ") + "空文本 -> (" + nx + "," + ny + ")  要求非负");

            // 真实 SetCaretPos 往返：验证写入系统的坐标与计算值一致
            box.Text = "abcdef";
            box.SelectionStart = 6;
            int cx, cy;
            CaretHelper.TryInsertionPoint(box, font, out cx, out cy);
            CaretHelper.Move(box, font);
            NativePoint got = new NativePoint();
            bool roundTrip = GetCaretPos(ref got) && got.X == cx + 1 && got.Y == cy;
            if (!roundTrip) failed++;
            Console.WriteLine((roundTrip ? "[OK]  " : "[FAIL] ") + "SetCaretPos 往返 期望=(" + (cx + 1) + "," + cy + ") 实际=(" + got.X + "," + got.Y + ")");

            host.Close();
            Console.WriteLine(failed == 0 ? "ALL PASS" : (failed + " FAILED"));
            return failed == 0 ? 0 : 2;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct NativePoint { public int X; public int Y; }
        [DllImport("user32.dll")]
        static extern bool GetCaretPos(ref NativePoint p);

        static int RunGuardTest()
        {
            string[][] cases = new string[][] {
                new string[] { "ftp://api.example.com/v1",        "false", "REJECT" },
                new string[] { "http://localhost:8080/v1",        "false", "REJECT" },
                new string[] { "http://127.0.0.1:8080/v1",        "false", "REJECT" },
                new string[] { "https://10.0.0.5/v1",             "false", "REJECT" },
                new string[] { "https://192.168.1.10/v1",         "false", "REJECT" },
                new string[] { "https://172.16.0.9/v1",           "false", "REJECT" },
                new string[] { "https://169.254.3.3/v1",          "false", "REJECT" },
                new string[] { "https://[::1]/v1",                "false", "REJECT" },
                new string[] { "https://8.8.8.8/v1",              "false", "ALLOW" },
                new string[] { "http://127.0.0.1:8080/v1",        "true",  "ALLOW" },
                new string[] { "https://192.168.1.10/v1",         "true",  "ALLOW" },
                new string[] { "https://api.example.com/v1",      "false", "ANY" }
            };
            int failed = 0;
            foreach (string[] c in cases)
            {
                string r = HostGuard.Check(c[0], c[1] == "true");
                string verdict = r == null ? "ALLOW" : "REJECT";
                bool ok = c[2] == "ANY" || verdict == c[2];
                if (!ok) failed++;
                Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + c[0] + "  allowLan=" + c[1] + "  => " + verdict + (r == null ? "" : "  (" + r + ")"));
            }
            Console.WriteLine(failed == 0 ? "ALL PASS" : (failed + " FAILED"));
            return failed == 0 ? 0 : 2;
        }

        // 无头转换冒烟：OfficeAgent.exe /convert <input> <pdf|csv|xlsx>
        static int RunConvert(string input, string targetText)
        {
            ConvTarget target;
            string t = (targetText ?? "").ToLowerInvariant();
            if (t == "pdf") target = ConvTarget.Pdf;
            else if (t == "csv") target = ConvTarget.Csv;
            else if (t == "xlsx") target = ConvTarget.Xlsx;
            else { Console.WriteLine("未知目标: " + targetText); return 3; }

            string root = EnvDetect.FindRoot();
            ConvertEngine engine = new ConvertEngine();
            engine.SofficePath = ConvertEngine.FindSoffice(root);
            Console.WriteLine("soffice=" + (engine.SofficePath ?? "(未探测到)"));
            string outPath;
            string err = engine.Convert(input, target, out outPath);
            if (err == null)
            {
                Console.WriteLine("OK: " + outPath + "  (" + new System.IO.FileInfo(outPath).Length + " bytes)");
                return 0;
            }
            Console.WriteLine("FAIL: " + err);
            return 2;
        }

    // 无头渲染冒烟：OfficeAgent.exe /renderpdf <in.pdf> <out.png>
    static int RunRenderPdf(string pdfPath, string pngPath)
    {
        try
        {
            string root = EnvDetect.FindRoot();
            Pdfium.EnsureInit(root);
            string err;
            using (Pdfium.PdfDoc doc = Pdfium.PdfDoc.Open(pdfPath, out err))
            {
                if (doc == null) { Console.WriteLine("FAIL: " + err); return 2; }
                Console.WriteLine("pages=" + doc.PageCount);
                using (Bitmap bm = doc.RenderPage(0, 120.0))
                {
                    bm.Save(pngPath, System.Drawing.Imaging.ImageFormat.Png);
                    Console.WriteLine("OK: " + pngPath + "  (" + new System.IO.FileInfo(pngPath).Length + " bytes)");
                }
            }
            return 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine("FAIL: " + ex.Message);
            return 2;
        }
    }

    // 两表核对（M2）：OfficeAgent.exe /recon <A> <B> [--keya .. --keyb ..] [--debita .. --credita ..]
    //   列参数可用 1-based 序号或表头名；键列可逗号分隔多列。
    //   [--sheetA|--sheetB 序号或名称] [--hdrA|--hdrB 表头行] [--tol 0.01] [--out 差异表.xlsx]
    //   [--all 附匹配明细页] [--save 模板名] [--tpl 模板名]
    // 退出码：0=完成（含差异）；2=错误或勾稽校验失败；3=用法错误
    static int RunRecon(string[] args)
    {
        AuditLog.Record("action_started", "recon");
        string fileA = null, fileB = null;
        Dictionary<string, string> opt = new Dictionary<string, string>();
        for (int i = 0; i < args.Length; i++)
        {
            string a = args[i];
            if (a == "--all") { opt["all"] = "1"; continue; }   // 无值旗标
            if (a.StartsWith("--", StringComparison.Ordinal) && i + 1 < args.Length)
            {
                opt[a.Substring(2).ToLowerInvariant()] = args[i + 1];
                i++;
            }
            else if (fileA == null) fileA = a;
            else if (fileB == null) fileB = a;
        }
        if (fileA == null || fileB == null)
        {
            Console.WriteLine("用法: OfficeAgent.exe /recon <文件A> <文件B> [--keya 账号 --keyb 对方账号]");
            Console.WriteLine("      [--debita 收入 --credita 支出 --debitb 借方 --creditb 贷方] [--tol 0.01]");
            Console.WriteLine("      [--sheetA 1 --sheetB 1] [--hdrA 1 --hdrB 1] [--out 核对.xlsx] [--all]");
            Console.WriteLine("      [--save 模板名] [--tpl 模板名]");
            return 3;
        }

        ReconMapping map;
        if (opt.ContainsKey("tpl"))
        {
            ReconTemplate t = ReconTemplateStore.Get(opt["tpl"]);
            if (t == null) { Console.WriteLine("FAIL: 模板不存在: " + opt["tpl"]); return 2; }
            map = t.ToMapping();
            Console.WriteLine("模板: " + t.Name);
        }
        else map = new ReconMapping();

        if (opt.ContainsKey("sheeta")) map.SheetA = opt["sheeta"];
        if (opt.ContainsKey("sheetb")) map.SheetB = opt["sheetb"];
        if (opt.ContainsKey("hdra")) { int n; if (int.TryParse(opt["hdra"], out n)) map.HeaderRowA = n; }
        if (opt.ContainsKey("hdrb")) { int n; if (int.TryParse(opt["hdrb"], out n)) map.HeaderRowB = n; }
        if (opt.ContainsKey("keya")) map.KeyColsA = SplitTokens(opt["keya"]);
        if (opt.ContainsKey("keyb")) map.KeyColsB = SplitTokens(opt["keyb"]);
        if (opt.ContainsKey("debita")) map.DebitA = opt["debita"];
        if (opt.ContainsKey("credita")) map.CreditA = opt["credita"];
        if (opt.ContainsKey("debitb")) map.DebitB = opt["debitb"];
        if (opt.ContainsKey("creditb")) map.CreditB = opt["creditb"];
        if (opt.ContainsKey("tol"))
        {
            double d;
            if (!double.TryParse(opt["tol"], NumberStyles.Float, CultureInfo.InvariantCulture, out d) || d < 0)
            { Console.WriteLine("FAIL: --tol 非法: " + opt["tol"]); return 3; }
            map.Tolerance = d;
        }
        if (opt.ContainsKey("squeeze")) map.SqueezeKey = opt["squeeze"] != "0";
        if (opt.ContainsKey("all")) map.IncludeMatched = true;

        if (map.KeyColsA.Length == 0 || map.KeyColsB.Length == 0)
        {
            Console.WriteLine("FAIL: 未指定键列（--keya/--keyb）");
            return 3;
        }

        string outPath;
        if (opt.ContainsKey("out")) outPath = opt["out"];
        else
        {
            string dir = System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(fileA));
            string name = "核对_" + TrimBase(fileA, 24) + "_vs_" + TrimBase(fileB, 24) + ".xlsx";
            outPath = System.IO.Path.Combine(dir, name);
        }

        ConvertEngine conv = new ConvertEngine();
        conv.SofficePath = ConvertEngine.FindSoffice(EnvDetect.FindRoot());
        ReconEngine engine = new ReconEngine();
        ReconResult r = engine.Run(map, fileA, fileB, outPath, conv);

        if (r.Err != null) { Console.WriteLine("FAIL: " + r.Err); return 2; }

        Console.WriteLine("匹配=" + r.MatchedCount + "  金额不等=" + r.PairDiffCount +
            "  仅A=" + r.OnlyACount + "  仅B=" + r.OnlyBCount +
            "  (A侧 " + r.TotalA + " 行 / B侧 " + r.TotalB + " 行)");
        Console.WriteLine("合计: A=" + r.SumA.ToString("N2", CultureInfo.InvariantCulture) +
            "  B=" + r.SumB.ToString("N2", CultureInfo.InvariantCulture) +
            "  差=" + (r.SumA - r.SumB).ToString("N2", CultureInfo.InvariantCulture));
        foreach (ReconCheck c in r.Checks)
        {
            Console.WriteLine((c.Warn ? "[提示] " : (c.Pass ? "[✓] " : "[✗] ")) + c.Name + "  " + c.Detail);
        }
        if (r.ParseErrors.Count > 0)
        {
            foreach (ReconParseError e in r.ParseErrors)
                Console.WriteLine("[解析错误] " + e.Side + " 侧第 " + e.ExcelRow + " 行「" + e.ColName + "」: " + e.Raw);
        }
        Console.WriteLine("差异表: " + r.OutputPath + "  (" + r.ElapsedMs + " ms)");

        if (opt.ContainsKey("save"))
        {
            ReconTemplate t = ReconTemplate.FromMapping(opt["save"], map, fileA, fileB);
            ReconTemplateStore.Upsert(t);
            Console.WriteLine("模板已保存: " + opt["save"]);
        }
        ReconTemplate last = ReconTemplate.FromMapping(ReconTemplateStore.LastName, map, fileA, fileB);
        ReconTemplateStore.Upsert(last);

        AuditLog.Record("action_finished", "recon " + r.OutputPath);
        return r.FailedChecks() > 0 ? 2 : 0;
    }

    static string[] SplitTokens(string s)
    {
        List<string> parts = new List<string>();
        foreach (string p in s.Split(','))
        {
            string t = p.Trim();
            if (t.Length > 0) parts.Add(t);
        }
        return parts.ToArray();
    }

    static string TrimBase(string path, int max)
    {
        string b = System.IO.Path.GetFileNameWithoutExtension(path);
        if (b.Length > max) b = b.Substring(0, max);
        foreach (char c in System.IO.Path.GetInvalidFileNameChars()) b = b.Replace(c, '_');
        return b;
    }

    // PDF 文本层调试：OfficeAgent.exe /pdftext <in.pdf>
    static int RunPdfText(string pdfPath)
    {
        try
        {
            string root = EnvDetect.FindRoot();
            Pdfium.EnsureInit(root);
            string err;
            using (Pdfium.PdfDoc doc = Pdfium.PdfDoc.Open(pdfPath, out err))
            {
                if (doc == null) { Console.WriteLine("FAIL: " + err); return 2; }
                for (int p = 0; p < doc.PageCount; p++)
                {
                    List<string> lines = doc.ExtractLines(p, out err);
                    if (lines == null) { Console.WriteLine("FAIL: " + err); return 2; }
                    Console.WriteLine("---- page " + (p + 1) + " ----");
                    foreach (string ln in lines) Console.WriteLine(ln);
                }
                return 0;
            }
        }
        catch (Exception ex) { Console.WriteLine("FAIL: " + ex.Message); return 2; }
    }

    // 脱敏引擎无头自测：OfficeAgent.exe /masktest
    static int RunMaskTest()
    {
        return MaskEngine.SelfTest();
    }

    // 意图路由无头自测：OfficeAgent.exe /intents（用 fixtures 真实文件驱动）
    static int RunIntents()
    {
        string root = EnvDetect.FindRoot();
        string fix = System.IO.Path.Combine(root, "tests");
        fix = System.IO.Path.Combine(fix, "fixtures");
        string flow = System.IO.Path.Combine(System.IO.Path.Combine(fix, "recon"), "flow.csv");
        string ledger = System.IO.Path.Combine(System.IO.Path.Combine(fix, "recon"), "ledger.csv");
        string invoice = System.IO.Path.Combine(fix, "invoice-sample.pdf");

        int failed = 0;
        ActionPlan p;

        // 1. 动作词但无文件 → 识别 Recon 且缺输入
        bool ok = IntentRouter.TryCreate("帮我核对一下", new string[0], out p) && p.Kind == ActionKind.Recon && p.Missing.Length > 0 && p.Inputs.Count == 0;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "无文件核对 → Recon + 缺参数（inputs=" + p.Inputs.Count + " missing=" + p.Missing.Length + "）");

        // 2. 文本带两个路径 → Recon 双输入
        ok = IntentRouter.TryCreate("核对 " + flow + " 和 " + ledger, new string[0], out p)
            && p.Kind == ActionKind.Recon && p.Inputs.Count == 2;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "带路径核对 → inputs=" + p.Inputs.Count);

        // 3. 上下文文件兜底
        ok = IntentRouter.TryCreate("核对一下", new string[] { flow, ledger }, out p)
            && p.Kind == ActionKind.Recon && p.Inputs.Count == 2;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "上下文兜底 → inputs=" + p.Inputs.Count);

        // 4. 汇总意图
        ok = IntentRouter.TryCreate("把 " + flow + " 汇总一下", new string[0], out p) && p.Kind == ActionKind.Merge;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "汇总意图 → " + p.Kind);

        // 5. 发票意图 + 单文件
        ok = IntentRouter.TryCreate("提取 " + invoice + " 的发票", new string[0], out p)
            && p.Kind == ActionKind.Invoice && p.Inputs.Count == 1;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "发票意图 → inputs=" + p.Inputs.Count);

        // 6. 普通问句不触发
        ok = !IntentRouter.TryCreate("你好，今天天气怎么样", new string[] { flow }, out p);
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "普通问句不触发路由");

        // 7. 取消/确认词不走路由（路由层只认动作意图；确认词在 ChatPanel pendingPlan 层处理）
        ok = !IntentRouter.TryCreate("确认", new string[0], out p);
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "确认词不误触路由");

        // 8. 技能场景词 → Skill 计划（row-stat：统计）
        ok = IntentRouter.TryCreate("统计一下 " + flow, new string[0], out p)
            && p.Kind == ActionKind.Skill && p.SkillId == "row-stat";
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "技能场景词路由（" + p.Kind + " id=" + p.SkillId + "）");

        Console.WriteLine(failed == 0 ? "ALL PASS" : (failed + " FAILED"));
        return failed == 0 ? 0 : 2;
    }

    // 列映射建议解析自测：OfficeAgent.exe /suggesttest（离线，不走网络——直接喂"模型回答"样本）
    static int RunSuggestTest()
    {
        string root = EnvDetect.FindRoot();
        string fix = System.IO.Path.Combine(root, "tests");
        fix = System.IO.Path.Combine(fix, "fixtures");
        string flow = System.IO.Path.Combine(System.IO.Path.Combine(fix, "recon"), "flow.csv");
        string ledger = System.IO.Path.Combine(System.IO.Path.Combine(fix, "recon"), "ledger.csv");

        string[] ha = ColumnSuggest.ReadHeaders(flow, "", 1);
        string[] hb = ColumnSuggest.ReadHeaders(ledger, "", 1);
        int failed = 0;
        SuggestResult r;

        // 0. 表头读取
        bool ok = ha.Length == 6 && hb.Length == 5;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "表头读取 A=" + ha.Length + " B=" + hb.Length);

        // 1. 正常建议 → 通过且映射正确
        string good = "{\"keyA\":\"账号\",\"keyB\":\"对方账号\",\"debitA\":\"收入\",\"creditA\":\"支出\"," +
                      "\"debitB\":\"借方\",\"creditB\":\"贷方\",\"tolerance\":0.01,\"explain\":\"账号对账\"}";
        r = ColumnSuggest.ParseSuggestion(good, ha, hb);
        ok = r.Ok && r.Mapping.KeyColsA.Length == 1 && r.Mapping.KeyColsA[0] == "账号" &&
             r.Mapping.DebitA == "收入" && r.Mapping.CreditB == "贷方" && r.Mapping.Tolerance == 0.01;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "正常建议解析与映射");

        // 2. 建议包含表头外的列名 → 整体拒绝（白名单红线）
        string bad = "{\"keyA\":\"账号\",\"keyB\":\"工资\",\"debitA\":\"收入\",\"creditA\":\"支出\"," +
                     "\"debitB\":\"借方\",\"creditB\":\"贷方\",\"tolerance\":0.01,\"explain\":\"\"}";
        r = ColumnSuggest.ParseSuggestion(bad, ha, hb);
        ok = !r.Ok && r.Error.IndexOf("不在表头名单") >= 0;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "越权列名整体拒绝（" + r.Error + "）");

        // 3. 非 JSON 回答 → 拒绝
        r = ColumnSuggest.ParseSuggestion("抱歉，我无法完成该任务。", ha, hb);
        ok = !r.Ok;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "非 JSON 回答拒绝");

        // 4. 缺键列 → 拒绝
        string nokey = "{\"keyA\":\"\",\"keyB\":\"对方账号\",\"debitA\":\"收入\",\"creditA\":\"\"," +
                       "\"debitB\":\"借方\",\"creditB\":\"\",\"tolerance\":0.01,\"explain\":\"\"}";
        r = ColumnSuggest.ParseSuggestion(nokey, ha, hb);
        ok = !r.Ok;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "缺键列拒绝");

        // 5. 带代码围栏的回答也能解析
        string fenced = "```json\n" + good + "\n```";
        r = ColumnSuggest.ParseSuggestion(fenced, ha, hb);
        ok = r.Ok;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "代码围栏容错解析");

        Console.WriteLine(failed == 0 ? "ALL PASS" : (failed + " FAILED"));
        return failed == 0 ? 0 : 2;
    }

    // 技能执行：OfficeAgent.exe /skillrun <skillId> <输入文件...> --yes [--request '{"k":"v"}']
    // 组装 request.json（inputs=绝对路径数组），经 SkillRunner 暂存→python sidecar→stdout JSON。
    // 确认红线：本命令是内部/自动化入口，必须显式给 --yes 才执行（等价于 GUI 的"确认"），
    // 否则只打印计划并退出——非交互调用不能绕过 M3.2「用户显式确认才执行」。全链落审计。
    static int RunSkillRun(string[] args)
    {
        string skillId = null;
        List<string> inputs = new List<string>();
        string requestExtra = null;
        bool confirmed = false;
        for (int i = 0; i < args.Length; i++)
        {
            string a = args[i];
            if (a == "--yes" || a == "--confirm") confirmed = true;
            else if (a == "--request" && i + 1 < args.Length) { requestExtra = args[i + 1]; i++; }
            else if (skillId == null) skillId = a;
            else if (System.IO.File.Exists(a)) inputs.Add(System.IO.Path.GetFullPath(a));
        }
        if (skillId == null || inputs.Count == 0)
        {
            Console.WriteLine("用法: OfficeAgent.exe /skillrun <skillId> <输入文件...> --yes [--request JSON]");
            Console.WriteLine("      --yes = 显式确认执行（不带则只显示计划，不执行任何技能代码）");
            return 3;
        }

        // 未确认：只显示计划，零副作用（GUI 走 pendingPlan 气泡确认，CLI 走 --yes）
        if (!confirmed)
        {
            Console.WriteLine("计划（未执行）：技能 " + skillId);
            foreach (string f in inputs) Console.WriteLine("  输入: " + f);
            Console.WriteLine("回复计划已确认请重跑并加 --yes。未确认不执行任何技能代码。");
            return 2;
        }

        string root = EnvDetect.FindRoot();
        ConvertEngine conv = new ConvertEngine();
        conv.SofficePath = ConvertEngine.FindSoffice(root);
        List<SkillRegistryEntry> reg = SkillSystem.Scan(root, conv);
        SkillRegistryEntry found = null;
        foreach (SkillRegistryEntry e in reg) { if (e.Id == skillId) found = e; }
        if (found == null) { Console.WriteLine("FAIL: 技能未注册: " + skillId); return 2; }
        if (found.State == SkillConst.StateDisabled) { Console.WriteLine("FAIL: 技能已禁用"); return 2; }
        if (found.Runtime != SkillConst.RuntimePython38) { Console.WriteLine("FAIL: 本命令仅执行 python38 技能（" + found.Runtime + " 走对应引擎入口）"); return 2; }

        StringBuilder req = new StringBuilder();
        // task 用技能 id（与 ActionExecutor.RunSkill 及 skills/*/main.py 的 request.task 口径一致）
        req.Append("{\"task\":\"").Append(MiniJson.Esc(skillId)).Append("\",\"inputs\":[");
        for (int i = 0; i < inputs.Count; i++)
        {
            if (i > 0) req.Append(",");
            req.Append("\"").Append(MiniJson.Esc(inputs[i])).Append("\"");
        }
        req.Append("]");
        if (requestExtra != null && requestExtra.Length > 0) req.Append(",").Append(requestExtra.Trim('{', '}'));
        req.Append("}");

        // 审计：确认 → 开始 → 结束（与 GUI 的 plan_created/confirmation/action_* 同口径）
        AuditLog.Record("plan_created", "skill " + skillId + "; inputs=" + inputs.Count + "; cli");
        AuditLog.Record("confirmation", "skill " + skillId + "; cli --yes");
        AuditLog.Record("action_started", "skill " + skillId);
        SkillRunResult r = SkillRunner.Run(root, found.Path, req.ToString(), 120);
        AuditLog.Record("action_finished", "skill " + skillId + "; state=" + (r.Ok ? "Succeeded" : "Failed") + "; " + r.Message);
        Console.WriteLine((r.Ok ? "[OK] " : "[FAIL] ") + r.Message + "  (" + r.ElapsedMs + " ms, exit=" + r.ExitCode + ")");
        if (r.DataJson != null && r.DataJson.Length > 0) Console.WriteLine(r.DataJson);
        return r.Ok ? 0 : 2;
    }

    // Agent 工具链无头实测：OfficeAgent.exe /agenttest <问题...>
    // 走与 ChatPanel 相同的 AgentLoop（ChatRaw + 工具分发），验证 messages/tools/tool_calls 全链路。
    static int RunAgentTest(string[] args)
    {
        string question = args != null && args.Length > 0 ? string.Join(" ", args) : "你好";
        AppConfig config = AppConfig.Load();
        if (!config.IsLlmConfigured()) { Console.WriteLine("FAIL: 模型未配置（先在 GUI 设置或填 config.json）"); return 2; }
        LlmClient client = new LlmClient(config);
        Console.WriteLine("端点: " + client.BaseUrl + "  模型: " + client.Model + "  联网: " + config.WebSearch);
        string sys = "你是 OfficeAgent 的无头测试助手。需要读取本地文件时调用 read_text_file 工具。用简体中文回答。";
        System.Diagnostics.Stopwatch sw = System.Diagnostics.Stopwatch.StartNew();
        AgentLoop.Result r = AgentLoop.Run(client, config, sys, question, null, null, null);
        sw.Stop();
        if (r.ToolLog.Length > 0) Console.WriteLine("工具: " + r.ToolLog.Replace("\n", " | "));
        if (r.Error.Length > 0)
        {
            Console.WriteLine("FAIL: " + r.Error);
            try { string dump = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "oa_last_request.json"); System.IO.File.WriteAllText(dump, LlmClient.LastRequestBody, new UTF8Encoding(false)); Console.WriteLine("请求体已存: " + dump); } catch { }
            return 2;
        }
        try { System.IO.File.WriteAllText(System.IO.Path.Combine(System.IO.Path.GetTempPath(), "oa_last_response.json"), LlmClient.LastResponseBody, new UTF8Encoding(false)); } catch { }
        try { System.IO.File.WriteAllText(System.IO.Path.Combine(System.IO.Path.GetTempPath(), "oa_last_request.json"), LlmClient.LastRequestBody, new UTF8Encoding(false)); } catch { }
        if (r.FirstError.Length > 0) Console.WriteLine("首次错误(已降级): " + r.FirstError);
        Console.WriteLine("轮数: " + r.Hops + "  用时: " + (sw.ElapsedMilliseconds / 1000.0).ToString("0.0") + "s");
        Console.WriteLine("回复: " + r.FinalText);
        return 0;
    }

    // 技能体系骨架自测：OfficeAgent.exe /skilltest
    static int RunSkillTest()
    {
        int failed = 0;
        ConvertEngine conv = new ConvertEngine();
        conv.SofficePath = ConvertEngine.FindSoffice(EnvDetect.FindRoot());

        // 1. 内置示例技能解析
        string root = EnvDetect.FindRoot();
        SkillManifest m = SkillSystem.ParseManifest(System.IO.Path.Combine(System.IO.Path.Combine(root, "skills"), "bank-recon"));
        bool ok = m.Valid && m.Id == "bank-recon" && m.Runtime == SkillConst.RuntimeBuiltin && m.Entry == "recon" && m.Scenarios.Length == 4;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "内置技能解析（id=" + m.Id + " problems=" + m.Problems.Count + "）");

        // 2. 非法清单拒绝：缺 id / 坏 id / 坏 semver
        string badDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "oa-bad-skill");
        try
        {
            if (!System.IO.Directory.Exists(badDir)) System.IO.Directory.CreateDirectory(badDir);
            System.IO.File.WriteAllText(System.IO.Path.Combine(badDir, "skill.json"),
                "{\"name\":\"bad\",\"version\":\"abc\",\"runtime\":\"python3.9\",\"entry\":\"\"}",
                new UTF8Encoding(false));
            SkillManifest bad = SkillSystem.ParseManifest(badDir);
            ok = !bad.Valid && bad.Problems.Count >= 3;
            if (!ok) failed++;
            Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "坏清单拒绝（problems=" + bad.Problems.Count + "）");
        }
        finally { try { System.IO.Directory.Delete(badDir, true); } catch { } }

        // 3. host 兼容判定
        ok = SkillSystem.CompareSemver("1.0.0", "1.0.0") == 0 &&
             SkillSystem.CompareSemver("1.2.0", "1.1.9") > 0 &&
             SkillSystem.CompareSemver("0.9", "1.0.0") < 0;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "semver 比较");

        // 4. 能力表与可选依赖降级
        SkillSystem.Capabilities cap = SkillSystem.DetectCapabilities(root, conv);
        ok = SkillSystem.DepsMissing(new string[] { "libreoffice?" }, cap) == null;   // 可选依赖永不阻塞
        string hardErr = SkillSystem.DepsMissing(new string[] { "ocr" }, cap);        // 本机无 OCR → 必需依赖缺失
        ok = ok && (cap.Ocr ? hardErr == null : hardErr != null);
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "能力表（LO=" + cap.LibreOffice + " COM=" + cap.OfficeCom + " OCR=" + cap.Ocr + "）");

        // 5. 全根扫描 + 注册表落盘
        List<SkillRegistryEntry> reg = SkillSystem.Scan(root, conv);
        SkillRegistryEntry found = null;
        foreach (SkillRegistryEntry e in reg) { if (e.Id == "bank-recon") found = e; }
        ok = found != null && found.State != SkillConst.StateDisabled && found.Source == SkillConst.SourceBuiltin;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "扫描注册（" + reg.Count + " 技能，bank-recon state=" + (found == null ? "?" : found.State) + "）");

        // 6. 用户根覆盖 builtin（同 id 高版本胜出）—— 在临时 user 根内进行，不触碰真实用户技能目录
        string tmpUserSkills = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "oa-skilltest-user");
        string userSkillDir = System.IO.Path.Combine(tmpUserSkills, "bank-recon");
        try
        {
            if (System.IO.Directory.Exists(tmpUserSkills)) System.IO.Directory.Delete(tmpUserSkills, true);
            System.IO.Directory.CreateDirectory(userSkillDir);
            System.IO.File.WriteAllText(System.IO.Path.Combine(userSkillDir, "skill.json"),
                "{\"id\":\"bank-recon\",\"name\":\"银行流水核对(用户版)\",\"version\":\"1.3.0\",\"host\":\"1.0.0\"," +
                "\"runtime\":\"builtin\",\"entry\":\"recon\",\"scenarios\":\"表格核对,对账\"}",
                new UTF8Encoding(false));
            List<SkillRegistryEntry> reg2 = SkillSystem.Scan(root, conv, tmpUserSkills, false);
            SkillRegistryEntry f2 = null;
            foreach (SkillRegistryEntry e in reg2) { if (e.Id == "bank-recon") f2 = e; }
            ok = f2 != null && f2.Version == "1.3.0" && f2.Source == SkillConst.SourceUser;
            if (!ok) failed++;
            Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "用户根高版本覆盖（version=" + (f2 == null ? "?" : f2.Version) + " source=" + (f2 == null ? "?" : f2.Source) + "）");
        }
        finally { try { System.IO.Directory.Delete(tmpUserSkills, true); } catch { } }

        // 7. python38 技能端到端（row-stat：暂存→sidecar→stdout JSON）
        string csvInput = System.IO.Path.Combine(System.IO.Path.Combine(System.IO.Path.Combine(root, "tests"), "fixtures"), "gb18030.csv");
        SkillRunResult sr = SkillRunner.Run(root, System.IO.Path.Combine(System.IO.Path.Combine(root, "skills"), "row-stat"),
            "{\"task\":\"row-stat\",\"inputs\":[\"" + MiniJson.Esc(csvInput) + "\"]}", 120);
        ok = sr.Ok && sr.DataJson.IndexOf("gb18030") >= 0 && sr.DataJson.IndexOf("11110.5") >= 0;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "python 技能端到端（" + sr.Message + "）" +
            (ok ? "" : "  data=" + sr.DataJson));

        // 8. Skill 计划经 ActionExecutor 的门槛与执行（回归：此前 ActionExecutor 无条件要求 OutputPath，
        //    而 ActionPlan 允许 Skill 无输出文件 → GUI 确认后的技能计划到不了 RunSkill）
        ActionPlan sp = new ActionPlan();
        sp.Kind = ActionKind.Skill;
        sp.SkillId = "row-stat";
        sp.Inputs.Add(csvInput);
        sp.Missing = new string[0];
        sp.State = ActionState.AwaitingConfirmation;
        bool execOk = sp.IsExecutable;   // Skill 无 OutputPath 也应可执行
        ActionExecutor ax = new ActionExecutor(conv);
        ActionExecutionResult ar = ax.Run(sp);
        ok = execOk && ar.State == ActionState.Succeeded && ar.DataJson.IndexOf("11110.5") >= 0;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "Skill 计划经执行器（executable=" + execOk +
            " state=" + ar.State + " msg=" + ar.Message + "）");

        // 9. 建表/建 PPT 生成器直测（不依赖模型；回归用户实机报的"无法创建表格、ppt"）
        string genDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "oa_gentest");
        try { if (System.IO.Directory.Exists(genDir)) System.IO.Directory.Delete(genDir, true); } catch { }
        System.IO.Directory.CreateDirectory(genDir);
        bool genOk;
        string genMsg = AgentTools.CreateSpreadsheet(
            System.IO.Path.Combine(genDir, "t.xlsx"), "姓名,金额\n张三,1250.5\n李四,980", out genOk);
        bool xlsxOk = genOk && System.IO.File.Exists(System.IO.Path.Combine(genDir, "t.xlsx"));
        if (!xlsxOk) failed++;
        Console.WriteLine((xlsxOk ? "[OK]  " : "[FAIL] ") + "CreateSpreadsheet 直测（" + genMsg + "）");

        genMsg = AgentTools.CreatePresentation(
            System.IO.Path.Combine(genDir, "t.pptx"), "OfficeAgent 测试|要点一;要点二;;第二页|内容A;内容B", out genOk);
        bool pptxOk = genOk && System.IO.File.Exists(System.IO.Path.Combine(genDir, "t.pptx"));
        if (!pptxOk) failed++;
        Console.WriteLine((pptxOk ? "[OK]  " : "[FAIL] ") + "CreatePresentation 直测（" + genMsg + "）");

        // 9b. 同名产物避让（回归：agent 不能静默覆盖用户既有文件）
        // 无界面时 ConfirmOverwrite=null → 自动改名 t(2).xlsx，旧文件内容保持不动。
        AgentTools.ConfirmOverwrite = null;
        string keepPath = System.IO.Path.Combine(genDir, "keep.xlsx");
        System.IO.File.WriteAllText(keepPath, "OLD", System.Text.Encoding.UTF8);
        bool owOk;
        string owMsg = AgentTools.CreateSpreadsheet(
            System.IO.Path.Combine(genDir, "keep.xlsx"), "姓名,金额\n王五,1", out owOk);
        string keptText = "";
        try { keptText = System.IO.File.ReadAllText(keepPath, System.Text.Encoding.UTF8); } catch { }
        bool renamed = owOk && System.IO.File.Exists(System.IO.Path.Combine(genDir, "keep(2).xlsx"));
        bool preserved = keptText == "OLD";
        if (!renamed || !preserved) failed++;
        Console.WriteLine(((renamed && preserved) ? "[OK]  " : "[FAIL] ") +
            "同名产物避让（旧文件" + (preserved ? "保留" : "被改动") + "，新文件" +
            (renamed ? "改名为 keep(2).xlsx" : "未改名") + "）");

        Console.WriteLine("（产物保留在 " + genDir + " 供人工抽查，下次运行时清理）");

        Console.WriteLine(failed == 0 ? "ALL PASS" : (failed + " FAILED"));
        return failed == 0 ? 0 : 2;
    }

    // 报表汇总：OfficeAgent.exe /merge <文件...> --targets "科目,金额" --srcs "科目,金额" --amounts "金额" [--hdr 1] [--sheet 1] [--tpl 名称] [--out 底稿.xlsx]
    static int RunMerge(string[] args)
    {
        List<string> files = new List<string>();
        Dictionary<string, string> opt = new Dictionary<string, string>();
        for (int i = 0; i < args.Length; i++)
        {
            string a = args[i];
            if (a.StartsWith("--", StringComparison.Ordinal) && i + 1 < args.Length)
            {
                opt[a.Substring(2).ToLowerInvariant()] = args[i + 1];
                i++;
            }
            else files.Add(a);
        }
        if (files.Count == 0)
        {
            Console.WriteLine("用法: OfficeAgent.exe /merge <文件...> --targets \"科目,金额\" --srcs \"科目,金额\" --amounts \"金额\"");
            Console.WriteLine("      [--hdr 1] [--sheet 1] [--tpl 模板名] [--out 汇总底稿.xlsx]");
            return 3;
        }
        MergeTemplate tpl;
        if (opt.ContainsKey("tpl"))
        {
            tpl = MergeTemplateStore.Get(opt["tpl"]);
            if (tpl == null) { Console.WriteLine("FAIL: 模板不存在: " + opt["tpl"]); return 2; }
        }
        else
        {
            if (!opt.ContainsKey("targets") || !opt.ContainsKey("srcs"))
            {
                Console.WriteLine("FAIL: 需要模板（--tpl）或列映射（--targets/--srcs）");
                return 3;
            }
            tpl = MergeTemplate.FromForm("__cli", opt.ContainsKey("sheet") ? opt["sheet"] : "1",
                1, SplitTokens(opt["targets"]), SplitTokens(opt["srcs"]),
                opt.ContainsKey("amounts") ? SplitTokens(opt["amounts"]) : new string[0]);
        }
        int hdr;
        if (opt.ContainsKey("hdr") && int.TryParse(opt["hdr"], out hdr) && hdr >= 1) tpl.HeaderRow = hdr;

        string outPath;
        if (opt.ContainsKey("out")) outPath = opt["out"];
        else outPath = System.IO.Path.Combine(
            System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(files[0])),
            "汇总底稿.xlsx");

        ConvertEngine conv = new ConvertEngine();
        conv.SofficePath = ConvertEngine.FindSoffice(EnvDetect.FindRoot());
        MergeEngine engine = new MergeEngine();
        MergeResult r = engine.Run(tpl, files, outPath, conv);
        if (r.Err != null) { Console.WriteLine("FAIL: " + r.Err); return 2; }
        foreach (MergeFileResult fr in r.Files)
        {
            if (fr.Ok) Console.WriteLine("[OK] " + fr.File + "  rows=" + fr.Rows);
            else Console.WriteLine("[FAIL] " + fr.File + "  " + fr.Err);
        }
        foreach (ReconCheck c in r.Checks)
            Console.WriteLine((c.Warn ? "[提示] " : (c.Pass ? "[✓] " : "[✗] ")) + c.Name + "  " + c.Detail);
        Console.WriteLine("底稿: " + r.OutputPath + "  (" + r.ElapsedMs + " ms)");
        return r.FailedChecks() > 0 ? 2 : 0;
    }

    // 发票提取：OfficeAgent.exe /invoices <pdf或文件夹...> --out 发票清单.xlsx
    static int RunInvoices(string[] args)
    {
        List<string> files = new List<string>();
        Dictionary<string, string> opt = new Dictionary<string, string>();
        for (int i = 0; i < args.Length; i++)
        {
            string a = args[i];
            if (a.StartsWith("--", StringComparison.Ordinal) && i + 1 < args.Length)
            {
                opt[a.Substring(2).ToLowerInvariant()] = args[i + 1];
                i++;
            }
            else if (System.IO.Directory.Exists(a))
            {
                foreach (string f in System.IO.Directory.GetFiles(a, "*.pdf")) files.Add(f);
            }
            else files.Add(a);
        }
        if (files.Count == 0)
        {
            Console.WriteLine("用法: OfficeAgent.exe /invoices <pdf或文件夹...> --out 发票清单.xlsx");
            return 3;
        }
        string outPath;
        if (opt.ContainsKey("out")) outPath = opt["out"];
        else outPath = System.IO.Path.Combine(
            System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(files[0])), "发票清单.xlsx");

        InvoiceEngine engine = new InvoiceEngine();
        InvoiceResult r = engine.Run(files, outPath);
        if (r.Err != null) { Console.WriteLine("FAIL: " + r.Err); return 2; }
        foreach (InvoiceRow w in r.Rows)
            Console.WriteLine("[OK] " + System.IO.Path.GetFileName(w.File) +
                "  no=" + w.InvoiceNo + "  total=" + w.Total.ToString("N2", CultureInfo.InvariantCulture));
        foreach (InvoiceFailure f in r.Failures)
            Console.WriteLine("[跳过] " + System.IO.Path.GetFileName(f.File) + "  " + f.Reason);
        foreach (ReconCheck c in r.Checks)
            Console.WriteLine((c.Warn ? "[提示] " : (c.Pass ? "[✓] " : "[✗] ")) + c.Name + "  " + c.Detail);
        Console.WriteLine("清单: " + r.OutputPath + "  (" + r.ElapsedMs + " ms)");
        return 0;
    }
}
}
