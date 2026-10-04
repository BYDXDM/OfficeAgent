// OfficeAgent 主程序入口（.NET Framework 4.8 / Win7 SP1+）
// 参数：/selftest  控制台输出环境检测后退出（M0 冒烟用）
//       /convert <in> <pdf|csv|xlsx>            无头转换（三级总线）
//       /renderpdf <in.pdf> <out.png>           无头渲染（pdfium）
//       /recon <A> <B> [选项]                   两表核对（M2，详见 RunRecon）
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
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
            bool toolidTest = false;
            bool formulaTest = false;
            string gridTestInput = null;
            bool skillTest = false;
            bool auditTest = false;
            bool detectTest = false;
            bool bridgeTest = false;
            bool planTest = false;
        bool perfTest = false;
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
                else if (a == "/toolidtest") toolidTest = true;
                else if (a == "/formulatest") formulaTest = true;
                else if (a == "/gridtest" && i + 1 < args.Length) { gridTestInput = args[i + 1]; i += 1; }
                else if (a == "/skilltest") skillTest = true;
                else if (a == "/audittest") auditTest = true;
                else if (a == "/detecttest") detectTest = true;
                else if (a == "/bridgetest") bridgeTest = true;
                else if (a == "/plantest") planTest = true;
        else if (a == "/perftest") perfTest = true;
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
                || auditTest || detectTest || bridgeTest || planTest || perfTest || toolidTest || formulaTest
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

            if (toolidTest)
            {
                try { AttachConsole(ATTACH_PARENT_PROCESS); } catch { }
                try { AllocConsole(); } catch { }
                try { Console.OutputEncoding = Encoding.UTF8; } catch { }
                return RunToolIdTest();
            }

            if (formulaTest)
            {
                try { AttachConsole(ATTACH_PARENT_PROCESS); } catch { }
                try { AllocConsole(); } catch { }
                try { Console.OutputEncoding = Encoding.UTF8; } catch { }
                return RunFormulaTest();
            }

            if (skillTest)
            {
                try { AttachConsole(ATTACH_PARENT_PROCESS); } catch { }
                try { AllocConsole(); } catch { }
                try { Console.OutputEncoding = Encoding.UTF8; } catch { }
                return RunSkillTest();
            }

            if (auditTest)
            {
                try { AttachConsole(ATTACH_PARENT_PROCESS); } catch { }
                try { AllocConsole(); } catch { }
                try { Console.OutputEncoding = Encoding.UTF8; } catch { }
                return RunAuditTest();
            }

            if (detectTest)
            {
                try { AttachConsole(ATTACH_PARENT_PROCESS); } catch { }
                try { AllocConsole(); } catch { }
                try { Console.OutputEncoding = Encoding.UTF8; } catch { }
                return RunDetectTest();
            }

            if (bridgeTest)
            {
                try { AttachConsole(ATTACH_PARENT_PROCESS); } catch { }
                try { AllocConsole(); } catch { }
                try { Console.OutputEncoding = Encoding.UTF8; } catch { }
                return RunBridgeTest();
            }

            if (planTest)
            {
                try { AttachConsole(ATTACH_PARENT_PROCESS); } catch { }
                try { AllocConsole(); } catch { }
                try { Console.OutputEncoding = Encoding.UTF8; } catch { }
                return RunPlanTest();
            }

            if (perfTest)
            {
                try { AttachConsole(ATTACH_PARENT_PROCESS); } catch { }
                try { AllocConsole(); } catch { }
                try { Console.OutputEncoding = Encoding.UTF8; } catch { }
                return RunPerfTest();
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

        // /toolidtest —— tool_calls 解析无头回归（不连网）。
        // 背景：并行工具调用时，部分模型/网关会给出重复或缺失的 tool_call id；
        // 旧解析法"从 function 往回找最近 id"还会串号。原样回显后，严格校验的端点
        // 直接拒收：HTTP 400 "Duplicate value for 'tool_call_id' of call_00_xxx in message[N]"，
        // 用户侧表现就是"一调用插件/技能就报错"。本测试用构造的响应消息验证 ParseToolCallsInto。
        static int RunToolIdTest()
        {
            int failed = 0;
            Action<string, bool> check = delegate(string name, bool ok)
            {
                Console.WriteLine((ok ? "  ok    " : "  FAIL  ") + name);
                if (!ok) failed++;
            };

            // 1) 正常两条：id 各异，原样保留
            LlmReply r1 = new LlmReply();
            LlmClient.ParseToolCallsInto(
                "{\"role\":\"assistant\",\"content\":\"\",\"tool_calls\":[" +
                "{\"id\":\"call_00_AAA111\",\"type\":\"function\",\"function\":{\"name\":\"list_directory\",\"arguments\":\"{}\"}}," +
                "{\"id\":\"call_00_BBB222\",\"type\":\"function\",\"function\":{\"name\":\"read_text_file\",\"arguments\":\"{\\\"p\\\":\\\"a.txt\\\"}\"}}" +
                "]}", r1);
            check("正常两条全部解析", r1.ToolCalls != null && r1.ToolCalls.Count == 2 && r1.HasToolCalls);
            check("正常 id 原样保留", r1.ToolCalls[0][0] == "call_00_AAA111" && r1.ToolCalls[1][0] == "call_00_BBB222");
            check("正常 name/arguments 解析", r1.ToolCalls[0][1] == "list_directory" && r1.ToolCalls[1][1] == "read_text_file"
                && r1.ToolCalls[1][2].Contains("a.txt"));

            // 2) 模型复用同一 id（服务端报 Duplicate 的直接来源）→ 第二条改写为唯一值
            LlmReply r2 = new LlmReply();
            LlmClient.ParseToolCallsInto(
                "{\"tool_calls\":[" +
                "{\"id\":\"call_00_DUP\",\"type\":\"function\",\"function\":{\"name\":\"list_directory\",\"arguments\":\"{}\"}}," +
                "{\"id\":\"call_00_DUP\",\"type\":\"function\",\"function\":{\"name\":\"list_directory\",\"arguments\":\"{}\"}}" +
                "]}", r2);
            check("重复 id 数量保留", r2.ToolCalls != null && r2.ToolCalls.Count == 2);
            check("重复 id 第二条被改写为唯一", r2.ToolCalls[0][0] == "call_00_DUP"
                && r2.ToolCalls[1][0].StartsWith("call_oa_", StringComparison.Ordinal));

            // 3) 缺失 id → 补唯一值
            LlmReply r3 = new LlmReply();
            LlmClient.ParseToolCallsInto(
                "{\"tool_calls\":[" +
                "{\"type\":\"function\",\"function\":{\"name\":\"list_directory\",\"arguments\":\"{}\"}}," +
                "{\"type\":\"function\",\"function\":{\"name\":\"read_text_file\",\"arguments\":\"{}\"}}" +
                "]}", r3);
            check("缺失 id 自动补齐且唯一", r3.ToolCalls != null && r3.ToolCalls.Count == 2
                && r3.ToolCalls[0][0].Length > 0 && r3.ToolCalls[1][0].Length > 0
                && r3.ToolCalls[0][0] != r3.ToolCalls[1][0]);

            // 4) 参数串里带转义的 \"id\": 字样 → 不得串成条目 id（旧回溯法的串号场景）
            LlmReply r4 = new LlmReply();
            LlmClient.ParseToolCallsInto(
                "{\"tool_calls\":[" +
                "{\"id\":\"call_00_REAL1\",\"type\":\"function\",\"function\":{\"name\":\"create_spreadsheet\",\"arguments\":\"{\\\"rows\\\":[{\\\"id\\\":\\\"x1\\\"}]}\"}}," +
                "{\"type\":\"function\",\"function\":{\"name\":\"read_text_file\",\"arguments\":\"{\\\"body\\\":\\\"id=x2\\\"}\"}}" +
                "]}", r4);
            check("参数内 id 不污染条目 id", r4.ToolCalls != null && r4.ToolCalls.Count == 2
                && r4.ToolCalls[0][0] == "call_00_REAL1"
                && r4.ToolCalls[1][0].StartsWith("call_oa_", StringComparison.Ordinal));

            // 5) id 写在 function 之后（个别网关的乱序形态）
            LlmReply r5 = new LlmReply();
            LlmClient.ParseToolCallsInto(
                "{\"tool_calls\":[" +
                "{\"type\":\"function\",\"function\":{\"name\":\"list_directory\",\"arguments\":\"{}\"},\"id\":\"call_00_LATE\"}" +
                "]}", r5);
            check("乱序 id 仍取得到", r5.ToolCalls != null && r5.ToolCalls.Count == 1 && r5.ToolCalls[0][0] == "call_00_LATE");

            // 6) web_search 等非 function 条目跳过
            LlmReply r6 = new LlmReply();
            LlmClient.ParseToolCallsInto(
                "{\"tool_calls\":[" +
                "{\"id\":\"ws1\",\"type\":\"web_search\",\"web_search\":{\"search_query\":\"x\"}}," +
                "{\"id\":\"call_00_OK\",\"type\":\"function\",\"function\":{\"name\":\"list_directory\",\"arguments\":\"{}\"}}" +
                "]}", r6);
            check("非 function 条目跳过", r6.ToolCalls != null && r6.ToolCalls.Count == 1 && r6.ToolCalls[0][0] == "call_00_OK");

            // 7) 无 tool_calls / 空数组
            LlmReply r7 = new LlmReply();
            LlmClient.ParseToolCallsInto("{\"role\":\"assistant\",\"content\":\"你好\"}", r7);
            check("无 tool_calls 不误报", !r7.HasToolCalls && r7.ToolCalls == null);

            LlmReply r8 = new LlmReply();
            LlmClient.ParseToolCallsInto("{\"tool_calls\":[]}", r8);
            check("空数组不算工具调用", r8.ToolCalls != null && r8.ToolCalls.Count == 0 && !r8.HasToolCalls);

            Console.WriteLine(failed == 0 ? "toolidtest ALL PASS" : ("toolidtest FAILED=" + failed));
            return failed == 0 ? 0 : 2;
        }

        // /formulatest —— create_formula_workbook + excel_formula_reference 无头回归（不连网、不开 Excel）。
        // 验证：① JsonVal 嵌套解析；② 三模板生成的 xlsx 公式落位（zip 内 XML 直读）；
        //       ③ workbook.xml 带 calcPr fullCalcOnLoad（无缓存值公式打开即重算的前提）；
        //       ④ 自由模式 formulaCols/summary 的 SUMIF 跨表引用；⑤ 公式库查询与 schema JSON 合法性。
        static int RunFormulaTest()
        {
            int failed = 0;
            Action<string, bool> check = delegate(string name, bool cond)
            {
                Console.WriteLine((cond ? "  ok    " : "  FAIL  ") + name);
                if (!cond) failed++;
            };
            string dir = Path.Combine(Path.GetTempPath(), "oa_formulatest");
            try { Directory.CreateDirectory(dir); } catch { }
            if (!Directory.Exists(dir)) { Console.WriteLine("  FAIL  无法创建临时目录 " + dir); return 2; }

            // ① JsonVal：嵌套 + 转义 + 数字/布尔
            try
            {
                Dictionary<string, object> o = JsonVal.ParseObject(
                    "{\"name\":\"表A\",\"n\":2.5,\"ok\":true,\"rows\":[[\"a\",\"x=\\\"y\\\"\"],[]]}");
                check("JsonVal 嵌套解析", o["name"] as string == "表A" &&
                    (double)o["n"] == 2.5 && (bool)o["ok"] == true &&
                    (JsonVal.List(o, "rows") as List<object>).Count == 2);
            }
            catch (Exception ex) { check("JsonVal 嵌套解析（异常: " + ex.Message + "）", false); }

            // ② 工资表模板
            string payroll = Path.Combine(dir, "工资表-2026年1月.xlsx");
            bool ok1;
            string r1 = FormulaWorkbook.Create("{\"path\":\"" + payroll.Replace("\\", "\\\\") +
                "\",\"template\":\"payroll\",\"rows\":\"张三,财务,8000,1000,500\\n李四,销售,6000,800,0\\n王五,库房,5000,0,300\"}", out ok1);
            check("工资表生成成功", ok1 && File.Exists(payroll));
            if (ok1)
            {
                string wb = ReadZipEntry(payroll, "xl/workbook.xml");
                check("workbook.xml 含 calcPr fullCalcOnLoad", wb != null && wb.Contains("fullCalcOnLoad"));
                string s2 = ReadZipEntry(payroll, "xl/worksheets/sheet2.xml");
                check("工资表公式：社保引用参数!B5", s2 != null && s2.Contains("*参数!B5"));
                check("工资表公式：个税月度税率 IF 链", s2 != null && (s2.Contains("J2&lt;=12000") || s2.Contains("J3&lt;=12000")));
                check("工资表公式：应纳税所得额 MAX 守卫", s2 != null && s2.Contains("MAX(0,G2-参数!B7"));
                check("工资表合计行 SUM", s2 != null && s2.Contains("SUM(D2:D"));
                check("工资表空白行 IF 守卫（不显示 0）", s2 != null && s2.Contains("IF(B54="));
            }

            // ③ 增值税台账
            string vat = Path.Combine(dir, "增值税台账.xlsx");
            bool ok2;
            string r2 = FormulaWorkbook.Create("{\"path\":\"" + vat.Replace("\\", "\\\\") +
                "\",\"template\":\"vat\",\"rows\":\"2026-01-05,销售甲产品,销项,113000,0.13\\n2026-01-08,采购原材料,进项,56500,0.13\"}", out ok2);
            check("增值税台账生成成功", ok2 && File.Exists(vat));
            if (ok2)
            {
                string s2 = ReadZipEntry(vat, "xl/worksheets/sheet2.xml");
                check("增值税汇总 SUMIF 跨表引用", s2 != null && s2.Contains("SUMIF(台账!C2:C") && s2.Contains("台账!F2:F"));
                check("增值税应纳税额=销项-进项", s2 != null && s2.Contains("ROUND(B3-B5,2)"));
                string s1 = ReadZipEntry(vat, "xl/worksheets/sheet1.xml");
                check("台账税额=金额×税率", s1 != null && s1.Contains("ROUND(D2*E2,2)"));
            }

            // ④ 流水账
            string ledger = Path.Combine(dir, "流水账.xlsx");
            bool ok3;
            string r3 = FormulaWorkbook.Create("{\"path\":\"" + ledger.Replace("\\", "\\\\") +
                "\",\"template\":\"ledger\",\"params\":\"openingBalance=1000\",\"rows\":\"2026-01-01,收房租,房租收入,3000,0\\n2026-01-03,买菜,餐饮,0,120\"}", out ok3);
            check("流水账生成成功", ok3 && File.Exists(ledger));
            if (ok3)
            {
                string s2 = ReadZipEntry(ledger, "xl/worksheets/sheet2.xml");
                check("流水余额=期初+扩张区间SUM", s2 != null && s2.Contains("参数!B2+SUM($D$2:D2)") && s2.Contains("SUM($E$2:E"));
                string s3 = ReadZipEntry(ledger, "xl/worksheets/sheet3.xml");
                check("流水分类 SUMIF", s3 != null && s3.Contains("SUMIF(流水!C2:C"));
                check("流水期末结余公式", s3 != null && s3.Contains("ROUND(B2-B3+参数!B2,2)"));
            }

            // ⑤ 自由模式：formulaCols + summary
            string free = Path.Combine(dir, "自由表.xlsx");
            bool ok4;
            string spec = "{\"path\":\"" + free.Replace("\\", "\\\\") +
                "\",\"sheets\":[{\"name\":\"明细\",\"header\":[\"类别\",\"数量\",\"单价\",\"金额\"]," +
                "\"rows\":[[\"办公用品\",10,25],[\"交通费\",3,50]]," +
                "\"formulaCols\":[{\"col\":\"D\",\"formula\":\"=B{r}*C{r}\"}],\"blankRows\":10,\"totalRow\":true,\"widths\":[12,8,8,12]}]," +
                "\"summary\":\"{\\\"source\\\":\\\"明细\\\",\\\"groupCol\\\":\\\"A\\\",\\\"labelHeader\\\":\\\"类别\\\",\\\"sumCols\\\":[{\\\"col\\\":\\\"D\\\",\\\"header\\\":\\\"金额\\\"}]}\"}";
            string r4 = FormulaWorkbook.Create(spec, out ok4);
            check("自由模式生成成功", ok4 && File.Exists(free));
            if (ok4)
            {
                string s1 = ReadZipEntry(free, "xl/worksheets/sheet1.xml");
                check("自由公式列 {r} 展开为行号", s1 != null && s1.Contains("B2*C2") && s1.Contains("B12*C12"));
                string s2 = ReadZipEntry(free, "xl/worksheets/sheet2.xml");
                check("自由汇总 SUMIF('明细'!) 引用", s2 != null && s2.Contains("SUMIF('明细'!A2:A14,A2,'明细'!D2:D14)"));
            }

            // ⑥ 参数不合法的友好报错
            bool okBad;
            string bad = FormulaWorkbook.Create("{\"path\":\"" + Path.Combine(dir, "bad.xlsx").Replace("\\", "\\\\") +
                "\",\"template\":\"unknown\"}", out okBad);
            check("未知模板给友好报错", !okBad && bad.Contains("未知模板"));

            // ⑦ 公式库
            bool okc1, okc2, okc3, okc4;
            string cat = FormulaReference.Lookup(null, null, null, out okc1);
            check("公式库分类概览", okc1 && cat.Contains("财务会计") && cat.Contains("共 "));
            string hit = FormulaReference.Lookup(null, null, "VLOOKUP", out okc2);
            check("公式库按名查 VLOOKUP", okc2 && hit.Contains("VLOOKUP") && hit.Contains("查找区域"));
            string taxHit = FormulaReference.Lookup("个税", null, null, out okc3);
            check("公式库关键词查个税", okc3 && taxHit.Contains("速算"));
            string miss = FormulaReference.Lookup("不存在的公式xyz", null, null, out okc4);
            check("公式库未命中不算失败", okc4 && miss.Contains("未找到"));

            // ⑧ SchemasJson 含新工具且整体是合法 JSON
            try
            {
                object parsed = JsonVal.Parse(AgentTools.SchemasJson(true));
                List<object> arr = parsed as List<object>;
                bool hasFw = false, hasRef = false;
                foreach (object it in arr)
                {
                    Dictionary<string, object> d = it as Dictionary<string, object>;
                    if (d == null) continue;
                    Dictionary<string, object> f = JsonVal.Obj(d, "function");
                    string fn = f == null ? "" : JsonVal.Str(f, "name");
                    if (fn == "create_formula_workbook") hasFw = true;
                    if (fn == "excel_formula_reference") hasRef = true;
                }
                check("SchemasJson 是合法 JSON 且含两个新工具", arr != null && arr.Count >= 8 && hasFw && hasRef);
            }
            catch (Exception ex) { check("SchemasJson JSON 合法性（异常: " + ex.Message + "）", false); }

            Console.WriteLine(failed == 0 ? "formulatest ALL PASS" : ("formulatest FAILED=" + failed));
            return failed == 0 ? 0 : 2;
        }

        // 读 zip 内条目文本（自测专用；失败返回 null）
        static string ReadZipEntry(string xlsxPath, string entryName)
        {
            try
            {
                using (MiniZipFile zf = MiniZipFile.OpenRead(xlsxPath))
                using (StreamReader sr = new StreamReader(zf.OpenEntry(entryName), Encoding.UTF8))
                    return sr.ReadToEnd();
            }
            catch { return null; }
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
            // --request-file：从文件读参数 JSON。
            // 为什么需要它：命令行传复杂 JSON（含引号/中文/嵌套）时，shell 的引号转义极不可靠 ——
            // 实测 PowerShell 会把 \" 吃掉导致 JSON 断裂。技能参数一旦有嵌套结构（如 blocks 数组），
            // --request 基本不可用。文件传参不受 shell 影响，是唯一可靠的方式。
            else if (a == "--request-file" && i + 1 < args.Length)
            {
                string rf = args[i + 1]; i++;
                try { requestExtra = System.IO.File.ReadAllText(rf, Encoding.UTF8); }
                catch (Exception ex) { Console.WriteLine("读取 --request-file 失败: " + ex.Message); return 3; }
            }
            else if (skillId == null) skillId = a;
            else if (System.IO.File.Exists(a)) inputs.Add(System.IO.Path.GetFullPath(a));
        }
        if (skillId == null || inputs.Count == 0)
        {
            Console.WriteLine("用法: OfficeAgent.exe /skillrun <skillId> <输入文件...> --yes [--request JSON | --request-file <json文件>]");
            Console.WriteLine("      --yes = 显式确认执行（不带则只显示计划，不执行任何技能代码）");
            Console.WriteLine("      参数含嵌套结构时请用 --request-file（命令行转义不可靠）");
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
        if (requestExtra != null && requestExtra.Length > 0)
        {
            // 把附加参数并入 request 顶层。
            // 注意不能简单 Trim('{','}')：多行缩进的 JSON（如 --request-file 读进来的）
            // 末尾花括号独占一行，Trim 只去首尾字符会留下残缺的 "}\n"，导致解析报
            // "Extra data"。这里改为：能找到最外层 {} 就去掉它们并 Trim 空白；
            // 若内容不是对象形态（无大括号），则原样并入（兼容历史上直接传 "k":"v" 的用法）。
            string extra = requestExtra.Trim();
            int lb = extra.IndexOf('{');
            int rb = extra.LastIndexOf('}');
            if (lb >= 0 && rb > lb) extra = extra.Substring(lb + 1, rb - lb - 1).Trim();
            if (extra.Length > 0) req.Append(",").Append(extra);
        }
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

    // 审计哈希链并发自测：OfficeAgent.exe /audittest
    // 环境检测防抖动自测：OfficeAgent.exe /detecttest
    // 回归目标（审计发现的缺口）：
    //   ① 热修查询原先把"WMI 查询失败"与"确实没装"混为一谈（都返回 false → 都判 Missing）。
    //      低配 Win7 上 WMI 首次查询超时是常态 → 检测结果在两次启动间抖动 →
    //      修复窗口被反复唤醒（用户抱怨的"老是被打断"）。
    //   ② 磁盘检测阈值恰好卡在 2GB，可用空间在阈值上下浮动时会 Ok↔Missing 抖动。
    //      迟滞的高低水位算错过一次（5L/2 整数除法 → 高水位退化为 2GB，迟滞为零）。
    // 本用例把这些语义钉住，防止回归。
    static int RunDetectTest()
    {
        int failed = 0;

        // 1. 三态枚举存在且互异
        bool t1 = EnvDetect.HotfixState.Present != EnvDetect.HotfixState.Absent
               && EnvDetect.HotfixState.Absent != EnvDetect.HotfixState.Unavailable
               && EnvDetect.HotfixState.Present != EnvDetect.HotfixState.Unavailable;
        if (!t1) failed++;
        Console.WriteLine((t1 ? "[OK]  " : "[FAIL] ") + "热修三态枚举（Present/Absent/Unavailable）");

        // 2. 本机 WMI 可用时：不存在的补丁必须判 Absent（而不是 Unavailable）
        EnvDetect.HotfixState s = EnvDetect.QueryHotfix("KB_DOES_NOT_EXIST_9999999");
        bool t2 = s == EnvDetect.HotfixState.Absent;
        if (!t2) failed++;
        Console.WriteLine((t2 ? "[OK]  " : "[FAIL] ") + "不存在的补丁 → Absent（确认未装，非查不出来）实际=" + s);

        // 3. HasHotfix 向后兼容：只有 Present 才算 true
        bool t3 = !EnvDetect.HasHotfix("KB_DOES_NOT_EXIST_9999999");
        if (!t3) failed++;
        Console.WriteLine((t3 ? "[OK]  " : "[FAIL] ") + "HasHotfix 兼容语义（Absent→false）");

        // 4. 磁盘迟滞算术：**直接断言常量**，不依赖运行机的实际剩余空间。
        //    这里正是为了钉住那个真实犯过的错：初版写成 5L / 2 * 1024 * 1024 * 1024，
        //    C# 整数除法先算成 2 → 高水位退化为 2GiB，与低水位相等，迟滞宽度为 0、完全失效，
        //    而注释却声称 2.5GiB。（旧的断言只看"disk 项状态是否合法"，把高水位改回 2GB
        //    照样通过，抓不到这个回归——审计指出后改为断言算术本身。）
        const long GiB = 1024L * 1024 * 1024;
        bool t4a = EnvDetect.DiskLowWater == 2L * GiB;
        bool t4b = EnvDetect.DiskHighWater == 2684354560L;        // 2.5 GiB，写死数值防止"两边一起改错"
        bool t4c = EnvDetect.DiskHighWater > EnvDetect.DiskLowWater;   // 带宽必须 > 0，否则迟滞无意义
        bool t4 = t4a && t4b && t4c;
        if (!t4) failed++;
        Console.WriteLine((t4 ? "[OK]  " : "[FAIL] ") + "磁盘迟滞算术（低=" + EnvDetect.DiskLowWater +
            " 高=" + EnvDetect.DiskHighWater + " 带宽=" + (EnvDetect.DiskHighWater - EnvDetect.DiskLowWater) + " B）");

        // 4b. 磁盘项本身存在且状态合法（弱检查，仅保证检测没崩）
        List<DetectItem> items = EnvDetect.DetectAll(EnvDetect.FindRoot());
        DetectItem disk = null;
        foreach (DetectItem it in items) { if (it.Id == "disk") { disk = it; break; } }
        bool t4d = disk != null && (disk.State == DetectState.Ok || disk.State == DetectState.Missing);
        if (!t4d) failed++;
        Console.WriteLine((t4d ? "[OK]  " : "[FAIL] ") + "磁盘检测项存在且状态合法（" +
            (disk == null ? "未找到" : disk.State + " - " + disk.Detail) + "）");

        // 5. Unknown 不得被当成 Missing：统计一遍，确认没有把 Unknown 混入缺失
        int missing = 0, unknown = 0;
        foreach (DetectItem it in items)
        {
            if (it.State == DetectState.Missing) missing++;
            else if (it.State == DetectState.Unknown) unknown++;
        }
        Console.WriteLine("[INFO] 检测汇总：共 " + items.Count + " 项，缺失 " + missing + " 项，未知 " + unknown + " 项");

        // 6. 状态文案区分 Missing 与 Unknown（避免用户把"查不出来"误读为"缺了"）
        string sm = EnvDetect.StateText(DetectState.Missing);
        string su = EnvDetect.StateText(DetectState.Unknown);
        bool t6 = sm != su;
        if (!t6) failed++;
        Console.WriteLine((t6 ? "[OK]  " : "[FAIL] ") + "缺失/未知文案可区分（" + sm + " vs " + su + "）");

        Console.WriteLine(failed == 0 ? "ALL PASS" : (failed + " FAILED"));
        return failed == 0 ? 0 : 2;
    }

    // 回归目标：Record 过去只 lock 进程内 + 用进程内缓存的 lastSeq/lastHash，
    // 两个进程（GUI 开着同时跑 CLI）会各自算出同一 seq 并追加，产生**重复 seq 与分叉链**，
    // 于是 VerifyChain 把正常并发误报为"审计日志可能被篡改"。
    //
    // ⚠️ 覆盖面说明（不要说成"已验证跨进程"）：
    //   本用例在**单进程**内用 8 线程并发调 Record，断言"链可校验 / 无重复 seq / 无分叉"。
    //   但同进程的线程已被 AuditLog 内的 lock(gate) 串行化，audit.lock 这个**跨进程**文件锁
    //   在这里根本不会被竞争到 —— 也就是说，即便把整个文件锁删掉，本用例**照样 ALL PASS**。
    //   它真正验证的是"每次都重新 LoadTail()、不复用过期进程内缓存"这一必要成分
    //  （对跨进程正确性也必需），而不是跨进程锁本身。
    //   真跨进程验证靠实机演练：开 GUI 的同时跑 CLI，再用 VerifyChain 校验；本用例不覆盖。
    static int RunAuditTest()
    {
        int failed = 0;
        Console.WriteLine("审计文件: " + AuditLog.FilePath());

        // 1. 多线程并发写入（同进程内）

        System.Threading.Thread[] ts = new System.Threading.Thread[8];
        for (int i = 0; i < ts.Length; i++)
        {
            int id = i;
            ts[i] = new System.Threading.Thread(new System.Threading.ThreadStart(delegate()
            {
                for (int k = 0; k < 25; k++) AuditLog.Record("audittest", "t" + id + "-k" + k);
            }));
        }
        for (int i = 0; i < ts.Length; i++) ts[i].Start();
        for (int i = 0; i < ts.Length; i++) ts[i].Join();
        Console.WriteLine("[OK]   并发写入完成（8 线程 x 25 条）");

        // 2. 链校验必须通过
        int count;
        long broken;
        string err = AuditLog.VerifyChain(out count, out broken);
        bool chainOk = err == null;
        if (!chainOk) failed++;
        Console.WriteLine((chainOk ? "[OK]  " : "[FAIL] ") + "哈希链可校验（检查 " + count + " 条）" +
            (chainOk ? "" : "  err=" + err));

        // 3. 无重复 seq（读原始文件）
        List<string> lines = new List<string>();
        try
        {
            string[] raw = System.IO.File.ReadAllLines(AuditLog.FilePath(), Encoding.UTF8);
            foreach (string l in raw) { if (l != null && l.Trim().Length > 0) lines.Add(l); }
        }
        catch (Exception ex) { Console.WriteLine("[FAIL] 读取审计文件失败: " + ex.Message); failed++; }

        List<string> seqs = new List<string>();
        int dup = 0;
        foreach (string l in lines)
        {
            string s = Field(l, "seq");
            if (s == null) continue;
            if (seqs.Contains(s)) dup++;
            else seqs.Add(s);
        }
        bool noDup = dup == 0;
        if (!noDup) failed++;
        Console.WriteLine((noDup ? "[OK]  " : "[FAIL] ") + "无重复 seq（共 " + lines.Count + " 条，重复 " + dup + " 个）");

        // 4. 无分叉：同一 prev 不得被两条记录引用（除首条 GENESIS）
        Dictionary<string, int> prevCount = new Dictionary<string, int>();
        foreach (string l in lines)
        {
            string p = Field(l, "prev");
            if (p == null) continue;
            if (prevCount.ContainsKey(p)) prevCount[p] = prevCount[p] + 1;
            else prevCount[p] = 1;
        }
        int fork = 0;
        foreach (KeyValuePair<string, int> kv in prevCount)
        {
            if (kv.Value > 1) fork++;
        }
        bool noFork = fork == 0;
        if (!noFork) failed++;
        Console.WriteLine((noFork ? "[OK]  " : "[FAIL] ") + "哈希链无分叉（分叉点 " + fork + " 个）");

        Console.WriteLine(failed == 0 ? "ALL PASS" : (failed + " FAILED"));
        return failed == 0 ? 0 : 2;
    }

    // 极简字段取值（"key": 或 "key":" 形态；够本用例使用）
    static string Field(string line, string key)
    {
        string pat = "\"" + key + "\":";
        int i = line.IndexOf(pat, StringComparison.Ordinal);
        if (i < 0) return null;
        int pos = i + pat.Length;
        if (pos < line.Length && line[pos] == '"')
        {
            pos++;
            int end = line.IndexOf('"', pos);
            if (end < 0) return null;
            return line.Substring(pos, end - pos);
        }
        int e = pos;
        while (e < line.Length && (char.IsDigit(line[e]) || line[e] == '-')) e++;
        return e > pos ? line.Substring(pos, e - pos) : null;
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
        // 现在不再询问用户（询问需阻塞后台线程），一律自动改名并回调通知。
        string notified = "";
        AgentTools.OnOverwriteAvoided = delegate(string existing) { notified = existing; };
        string keepPath = System.IO.Path.Combine(genDir, "keep.xlsx");
        System.IO.File.WriteAllText(keepPath, "OLD", System.Text.Encoding.UTF8);
        bool owOk;
        string owMsg = AgentTools.CreateSpreadsheet(
            System.IO.Path.Combine(genDir, "keep.xlsx"), "姓名,金额\n王五,1", out owOk);
        string keptText = "";
        try { keptText = System.IO.File.ReadAllText(keepPath, System.Text.Encoding.UTF8); } catch { }
        bool renamed = owOk && System.IO.File.Exists(System.IO.Path.Combine(genDir, "keep(2).xlsx"));
        bool preserved = keptText == "OLD";
        bool notifiedOk = notified == keepPath;
        if (!renamed || !preserved || !notifiedOk) failed++;
        Console.WriteLine(((renamed && preserved && notifiedOk) ? "[OK]  " : "[FAIL] ") +
            "同名产物避让（旧文件" + (preserved ? "保留" : "被改动") + "，新文件" +
            (renamed ? "改名为 keep(2).xlsx" : "未改名") +
            "，通知回调" + (notifiedOk ? "已触发" : "未触发") + "）");
        AgentTools.OnOverwriteAvoided = null;   // 复位，避免影响后续用例

        // 9c. 避让通知回调抛异常时不得影响产物落盘（回调是"锦上添花"）
        AgentTools.OnOverwriteAvoided = delegate(string existing) { throw new InvalidOperationException("模拟通知失败"); };
        bool robOk = false;
        string robMsg = AgentTools.CreateSpreadsheet(
            System.IO.Path.Combine(genDir, "keep.xlsx"), "姓名,金额\n赵六,2", out robOk);
        AgentTools.OnOverwriteAvoided = null;
        bool robFile = System.IO.File.Exists(System.IO.Path.Combine(genDir, "keep(3).xlsx"));
        if (!robOk || !robFile) failed++;
        Console.WriteLine(((robOk && robFile) ? "[OK]  " : "[FAIL] ") +
            "避让通知异常不影响落盘（产物" + (robFile ? "已生成 keep(3).xlsx" : "未生成") + "）");

        Console.WriteLine("（产物保留在 " + genDir + " 供人工抽查，下次运行时清理）");

        // 11. 思考段剥离（回归：glm-4.5-air 把思考写进 content，以 </think> 收尾后重写答复）
        bool mtChanged;
        string thinkIn = "草稿：张三应发 8,800 元……</think>\n正式答复：张三应发 8,500 元。";
        string thinkOut = ModelText.Clean(thinkIn, out mtChanged);
        bool thinkOk = mtChanged && thinkOut.IndexOf("</think>") < 0
            && thinkOut.IndexOf("草稿") < 0 && thinkOut.IndexOf("8,500") >= 0
            && thinkOut.StartsWith("正式答复");
        // 取最后一个闭合标签：草稿1</think>改稿2</think>答案3 → 只留答案3
        bool mt2; string multi = ModelText.Clean("draft1</think>draft2</think>final", out mt2);
        thinkOk = thinkOk && mt2 && multi == "final";
        // 无标签原样返回（绝大多数模型走这条，行为零变化）
        bool mt3; string plain = ModelText.Clean("正常答复，没有标签。", out mt3);
        thinkOk = thinkOk && !mt3 && plain == "正常答复，没有标签。";
        // 只有开标签（被截断）→ 明确提示而不是把半截思考当答案
        bool mt4; string trunc = ModelText.Clean("<think>还没想完", out mt4);
        thinkOk = thinkOk && mt4 && trunc.IndexOf("截断") >= 0 && trunc.IndexOf("还没想完") < 0;
        // 只有闭合标签、其后为空 → 提示未给出答复
        bool mt5; string onlyClose = ModelText.Clean("思考</think>", out mt5);
        thinkOk = thinkOk && mt5 && onlyClose.IndexOf("没有给出正式答复") >= 0;
        if (!thinkOk) failed++;
        Console.WriteLine((thinkOk ? "[OK]  " : "[FAIL] ") + "思考段剥离（5 个用例）");

        Console.WriteLine(failed == 0 ? "ALL PASS" : (failed + " FAILED"));
        return failed == 0 ? 0 : 2;
    }

    // 规划层第二期自测：OfficeAgent.exe /plantest
    // 覆盖加权预算、产物登记、失败换路（计划显式化是 prompt 文本，不在此断言）。
    static int RunPlanTest()
    {
        int failed = 0;

        // ---------- 加权预算 ----------
        HopBudget.State b = new HopBudget.State();
        double c1 = HopBudget.Cost("list_directory", b);
        double c2 = HopBudget.Cost("read_text_file", b);
        bool ok = c1 == 0.5 && c2 == 0.5 && b.ReadCalls == 2;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "读型半价（前 2 次各 " + c1 + "/" + c2 + "）");

        double c3 = HopBudget.Cost("task_plan", b);
        ok = c3 == 0.0;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "计划工具免费（" + c3 + "）");

        double c4 = HopBudget.Cost("skill_acct_tools_bonus", b);
        double c5 = HopBudget.Cost("convert_document", b);
        ok = c4 == 1.0 && c5 == 1.0;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "技能/写型全价（" + c4 + "/" + c5 + "）");

        // 读型超出 ReadCap 后转全价（防止靠疯狂列目录规避预算）
        HopBudget.State b2 = new HopBudget.State();
        double last = 0;
        for (int i = 0; i < HopBudget.ReadCap + 1; i++) last = HopBudget.Cost("list_directory", b2);
        ok = last == 1.0 && b2.ReadCalls == HopBudget.ReadCap + 1;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "读型超额度转全价（第 " + b2.ReadCalls +
            " 次 = " + last + "）");

        // 纯读型连续调用应在等效上限处触发收尾
        HopBudget.State b3 = new HopBudget.State();
        int calls = 0;
        while (!HopBudget.ShouldWrapUp(b3) && calls < 1000)
        {
            b3.Spent += HopBudget.Cost("read_text_file", b3);
            calls++;
        }
        ok = calls > HopBudget.ReadCap && calls < 1000 && b3.Spent >= HopBudget.MaxEquivalent;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "预算耗尽触发收尾（读型 " + calls +
            " 次消耗 " + b3.Spent.ToString("0.#") + "）");

        // 加权后混合任务能容纳的真实跳数必须多于旧固定上限 12，否则这次改动没有意义
        HopBudget.State b4 = new HopBudget.State();
        int mixed = 0;
        while (!HopBudget.ShouldWrapUp(b4) && mixed < 1000)
        {
            b4.Spent += HopBudget.Cost(mixed % 2 == 0 ? "list_directory" : "skill_acct_tools_vat", b4);
            mixed++;
        }
        ok = mixed > 12;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "混合任务容量优于旧上限 12（可容纳 " +
            mixed + " 跳）");

        // ---------- 产物登记 ----------
        ArtifactRegistry.Reset();
        string tmpDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "oa_plantest");
        try
        {
            if (System.IO.Directory.Exists(tmpDir)) System.IO.Directory.Delete(tmpDir, true);
            System.IO.Directory.CreateDirectory(tmpDir);
            string f1 = System.IO.Path.Combine(tmpDir, "a.xlsx");
            string f2 = System.IO.Path.Combine(tmpDir, "b.pdf");
            System.IO.File.WriteAllText(f1, "x", new UTF8Encoding(false));
            System.IO.File.WriteAllText(f2, "y", new UTF8Encoding(false));

            string a1 = ArtifactRegistry.Add(f1);
            string a2 = ArtifactRegistry.Add(f2);
            string again = ArtifactRegistry.Add(f1);      // 重复登记应返回原别名
            ok = a1 == "产物1" && a2 == "产物2" && again == "产物1" && ArtifactRegistry.Items.Count == 2;
            if (!ok) failed++;
            Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "产物登记与去重（" + a1 + "/" + a2 +
                "/repeat=" + again + " count=" + ArtifactRegistry.Items.Count + "）");

            string ghost = ArtifactRegistry.Add(System.IO.Path.Combine(tmpDir, "nope.xlsx"));
            ok = ghost == "" && ArtifactRegistry.Items.Count == 2;
            if (!ok) failed++;
            Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "不存在的文件不登记（返回 '" + ghost + "'）");

            string desc = ArtifactRegistry.Describe();
            ok = desc.IndexOf(f1) >= 0 && desc.IndexOf(f2) >= 0 && desc.IndexOf("产物1") >= 0
                && ArtifactRegistry.AllExist();
            if (!ok) failed++;
            Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "产物清单文本");

            string ghostDesc = ArtifactRegistry.Describe();
            ArtifactRegistry.Reset();
            ok = ArtifactRegistry.Items.Count == 0 && ArtifactRegistry.Describe() == "";
            if (!ok) failed++;
            Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "回合隔离（Reset 后为空）");
        }
        finally { try { System.IO.Directory.Delete(tmpDir, true); } catch { } }

        // ---------- 失败换路 ----------
        string root = OfficeAgent.Core.EnvDetect.FindRoot();
        SkillToolBridge.InvalidateCache();
        SkillToolBridge.InvalidateAlternatives();

        List<SkillActionSpec> alts = SkillToolBridge.Alternatives(root, "acct-tools", "bank-recon", 3);
        bool foundRecon = false;
        foreach (SkillActionSpec sp in alts)
        {
            if (sp.SkillId == "bank-recon" && sp.Action == "recon") foundRecon = true;
        }
        ok = alts.Count > 0 && foundRecon;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "跨技能换路（" + alts.Count +
            " 个候选，含 bank-recon/recon=" + foundRecon + "）");

        // 跨技能替代优先于同技能兄弟动作（回归：初版给同技能固定 100 分，
        // 一个 10 动作的技能会把名额占满，真正有用的跨技能替代永远排不进来）
        List<SkillActionSpec> sameSkill = SkillToolBridge.Alternatives(root, "acct-tools", "vat", 3);
        ok = sameSkill.Count > 0 && sameSkill[0].SkillId == "bank-recon";
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "跨技能候选排在同技能之前（首个=" +
            (sameSkill.Count > 0 ? sameSkill[0].SkillId + "/" + sameSkill[0].Action : "无") + "）");

        // 名额有余时，同技能兄弟动作作为兜底补足
        List<SkillActionSpec> deep = SkillToolBridge.Alternatives(root, "acct-tools", "vat", 4);
        bool hasSame = false;
        foreach (SkillActionSpec sp in deep) { if (sp.SkillId == "acct-tools") hasSame = true; }
        ok = hasSame && deep.Count > 1;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "同技能动作作兜底补足（候选 " + deep.Count + " 个）");

        string rewritten = SkillToolBridge.WithAlternatives(root, "acct-tools", "bank-recon", "执行失败：缺少依赖", 2);
        ok = rewritten.IndexOf("执行失败：缺少依赖") >= 0 && rewritten.IndexOf("skill_") >= 0
            && rewritten.IndexOf("替代") >= 0;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "失败文本携带替代建议");

        string none = SkillToolBridge.WithAlternatives(root, "no-such-skill", "no-such-action", "失败", 2);
        ok = none == "失败";
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "无候选时不改写（'" + none + "'）");

        // ---------- acct-tools 列歧义回归（实测踩到的三个 bug） ----------
        // 造一张同时含期初/本期/期末借贷六列的科目余额表：旧实现按"第一个包含命中"取列，
        // 会读到期初列，导致金额与方向全错。
        string tbDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "oa_plantest_tb");
        try
        {
            if (System.IO.Directory.Exists(tbDir)) System.IO.Directory.Delete(tbDir, true);
            System.IO.Directory.CreateDirectory(tbDir);
            string tbFile = System.IO.Path.Combine(tbDir, "科目余额表.csv");
            // 期初 540000/540000，本期 275000/275000，期末 650000/650000 三处均平
            string[] lines = new string[] {
                "科目代码,科目名称,期初借方,期初贷方,本期借方,本期贷方,期末借方,期末贷方",
                "1001,库存现金,35000,0,20000,0,55000,0",
                "1002,银行存款,300000,0,150000,80000,370000,0",
                "1122,应收账款,120000,0,60000,40000,140000,0",
                "1405,库存商品,85000,0,25000,25000,85000,0",
                "2202,应付账款,0,40000,0,70000,0,110000",
                "2211,应付职工薪酬,0,30000,0,5000,0,35000",
                "4001,实收资本,0,470000,0,0,0,470000",
                "4103,本年利润,0,0,0,0,0,35000",
                "6001,主营业务收入,0,0,0,200000,0,0",
                "6401,主营业务成本,0,0,120000,0,0,0",
                "6602,管理费用,0,0,45000,0,0,0",
            };
            System.IO.File.WriteAllText(tbFile, string.Join("\n", lines) + "\n", new UTF8Encoding(false));

            // 试算平衡必须读【期末】列：650000/650000，而不是期初的 540000
            bool tbOk = false;
            string tbMsg = AgentTools.Dispatch("skill_acct_tools_trial_balance",
                "{\"inputs\":\"" + tbFile.Replace("\\", "\\\\") + "\"}", null, new ConvertEngine(), null, out tbOk);
            ok = tbOk && tbMsg.IndexOf("650000") >= 0 && tbMsg.IndexOf("540000") < 0;
            if (!ok) failed++;
            Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "试算平衡读期末列而非期初（" +
                tbMsg.Split('\n')[0] + "）");

            // 财务报表：必须能处理借贷分列表，且资产 = 负债+权益（三个 bug 的综合回归）。
            // 注意：AgentTools.Dispatch 返回的是技能 message 文本（data 段不外露），
            // 故这里断言 message 里的摘要与实际数字。技能自己的配平自检会写进 message。
            bool fsOk = false;
            string fsMsg = AgentTools.Dispatch("skill_acct_tools_statements",
                "{\"inputs\":\"" + tbFile.Replace("\\", "\\\\") + "\"}", null, new ConvertEngine(), null, out fsOk);
            bool selfClaimsBalanced = fsMsg.IndexOf("资产 = 负债+权益") >= 0;
            bool admitsUnbalanced = fsMsg.IndexOf("不平") >= 0 && fsMsg.IndexOf("✗") >= 0;
            ok = fsOk && selfClaimsBalanced && !admitsUnbalanced
                && fsMsg.IndexOf("650000") >= 0 && fsMsg.IndexOf("200000") >= 0
                && fsMsg.IndexOf("35000") >= 0;
            if (!ok) failed++;
            Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "财务报表借贷分列与配平（自检平衡=" +
                selfClaimsBalanced + " 自报不平=" + admitsUnbalanced + "）");

            // 技能返回的 data JSON 里 balanced 必须为 true（Dispatch 不外露 data，故直接调一次 SkillRunner）
            SkillRunResult fsr = SkillRunner.Run(root, System.IO.Path.Combine(
                System.IO.Path.Combine(root, "skills"), "acct-tools"),
                "{\"task\":\"acct-tools\",\"action\":\"statements\",\"inputs\":[\"" +
                MiniJson.Esc(tbFile) + "\"]}", 120);
            string rawOut = fsr.DataJson == null ? "" : fsr.DataJson;
            bool dataBalanced = rawOut.IndexOf("\"balanced\":true") >= 0
                || rawOut.IndexOf("\"balanced\": true") >= 0;
            ok = fsr.Ok && dataBalanced && rawOut.IndexOf("505000") >= 0;
            if (!ok) failed++;
            Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "data.balanced=true 且权益 505000（" +
                (rawOut.Length > 160 ? rawOut.Substring(0, 160) : rawOut) + "）");

            // ★ 关键回归：负债/权益不得为负（旧实现把贷方科目当借方，全变负数）
            bool noNegative = rawOut.IndexOf("-110000") < 0 && rawOut.IndexOf("-35000") < 0
                && rawOut.IndexOf("-470000") < 0 && fsMsg.IndexOf("-110000") < 0;
            if (!noNegative) failed++;
            Console.WriteLine((noNegative ? "[OK]  " : "[FAIL] ") + "负债权益符号为正（无负数漏出）");

            // ★ 净利不得重复计入权益：表内已有「本年利润」35000 时不再叠加本期净利，
            //   故权益合计 = 实收资本 470000 + 本年利润 35000 = 505000。
            //   （若不修，会变成 470000+35000+35000 = 540000 而虚增、报表自报不平）
            bool noDouble = rawOut.IndexOf("505000") >= 0 && rawOut.IndexOf("540000") < 0;
            if (!noDouble) failed++;
            Console.WriteLine((noDouble ? "[OK]  " : "[FAIL] ") + "净利不重复计入权益（权益合计 505000）");
        }
        finally { try { System.IO.Directory.Delete(tbDir, true); } catch { } }

        // ---------- 第三期：路由提示 / 目标继承 / 成功率台账 ----------

        // 1. 路由提示：命中技能时必须给出技能名与动作清单，且必须带"可以推翻"的免责声明
        ActionPlan hp = new ActionPlan();
        hp.Kind = ActionKind.Skill;
        hp.SkillId = "acct-tools";
        hp.Inputs.Add(System.IO.Path.Combine(tmpDir, "不存在.xlsx"));
        hp.Missing = new string[] { "至少一个输入文件" };
        string hint = RouteHint.Build(hp, "帮我把这个月的税算一下");
        ok = hint.IndexOf("acct-tools") >= 0 && hint.IndexOf("vat") >= 0
            && hint.IndexOf("以你自己的判断为准") >= 0;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "路由提示含技能名/动作清单/免责声明（" +
            hint.Replace("\n", " ").Trim() + "）");

        // 2. 未命中时不得产生任何提示（不能平白往每次请求里塞固定文本）
        ActionPlan missPlan = new ActionPlan();
        missPlan.Kind = ActionKind.None;
        string noHint = RouteHint.Build(missPlan, "你好");
        ok = noHint.Length == 0;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "未命中不产生提示（len=" + noHint.Length + "）");

        // 3. 规则判断与模型判断冲突时，提示必须**不**是命令语气。
        //    这是"提示"与"指令"的分界线：一旦措辞变成"你必须用 X"，规则误判就无从纠正。
        bool notCommand = hint.IndexOf("必须") < 0 && hint.IndexOf("只能") < 0
            && hint.IndexOf("可能判错") >= 0;
        if (!notCommand) failed++;
        Console.WriteLine((notCommand ? "[OK]  " : "[FAIL] ") + "提示是建议而非命令（误判可被模型推翻）");

        // 4. 成功率台账：记录 → 查询 → 样本不足时不得用于排序
        //    直接验证**扁平编码**能被 MiniJson 正确读回（嵌套写法会被静默读错，见类注释）
        SkillStats.ResetForTest();
        SkillStats.Record("acct-tools", "vat", true);
        bool unknownYet = SkillStats.Rate("acct-tools", "vat") < 0;
        if (!unknownYet) failed++;
        Console.WriteLine((unknownYet ? "[OK]  " : "[FAIL] ") + "样本不足时成功率视为未知（不用 1 次就下结论）");

        // 造"失败多于成功"：1 成功 2 失败 = 33%。
        // ★ 这里刻意**不用 2/2（正好 50%）**：BadRate 是严格小于才点名，
        //   50% 属"各占一半、还谈不上差"。初版自检正好踩在这个边界上而误判代码有 bug。
        SkillStats.Record("acct-tools", "vat", false);
        SkillStats.Record("acct-tools", "vat", false);
        SkillStats.InvalidateCache();   // 强制重新读盘，验证落盘格式确实可解析
        SkillStats.Entry ve = SkillStats.Get("acct-tools", "vat");
        ok = ve != null && ve.Ok == 1 && ve.Fail == 2;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "成功率台账落盘后可读回（ok=" +
            (ve == null ? "null" : ve.Ok + " fail=" + ve.Fail) + "，期望 1/2）");

        // 5. 差动作提示：成功率严格低于 50% 且样本足够才点名
        List<string[]> bad = SkillStats.BadActions("acct-tools", 3);
        ok = bad.Count == 1 && bad[0][0] == "vat";
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "低成功率动作被点名（" +
            (bad.Count > 0 ? bad[0][0] + " " + bad[0][1] + "/" + bad[0][2] : "无") + "）");

        // 5b. 边界：恰好 50% 不点名（钉住上面踩过的语义）
        SkillStats.Record("acct-tools", "vat", true);   // 2 成功 2 失败 = 50%
        List<string[]> half = SkillStats.BadActions("acct-tools", 3);
        bool notNamedAtHalf = true;
        foreach (string[] x in half) { if (x[0] == "vat") notNamedAtHalf = false; }
        if (!notNamedAtHalf) failed++;
        Console.WriteLine((notNamedAtHalf ? "[OK]  " : "[FAIL] ") + "恰好 50% 不点名（边界语义）");

        // 6. 好动作不得被点名（避免"用得好也被提醒"的噪声）
        SkillStats.Record("acct-tools", "statements", true);
        SkillStats.Record("acct-tools", "statements", true);
        List<string[]> bad2 = SkillStats.BadActions("acct-tools", 5);
        bool onlyBad = true;
        foreach (string[] x in bad2) { if (x[0] == "statements") onlyBad = false; }
        if (!onlyBad) failed++;
        Console.WriteLine((onlyBad ? "[OK]  " : "[FAIL] ") + "高成功率动作不被点名（共 " + bad2.Count + " 项）");

        // 7. 换路候选仍可复现（成功率台账不得改变"谁有资格当候选"）
        SkillStats.ResetForTest();
        SkillStats.InvalidateCache();
        List<SkillActionSpec> baseAlts = SkillToolBridge.Alternatives(root, "acct-tools", "bank-recon", 10);
        int total = baseAlts.Count;
        ok = total > 0 && baseAlts[0].SkillId == "bank-recon" && baseAlts[0].Action == "recon";
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "换路候选可复现（共 " + total +
            " 项，首个=" + (total > 0 ? baseAlts[0].SkillId + "/" + baseAlts[0].Action : "") + "）");

        // 7b. 成功率只作**同分**的二级排序，不得越权压过场景词重叠分。
        //     ★ 这条是实测撞出来的：最初想验"给差评就降权"，直接给得分最高的
        //       bank-recon/recon 记 3 次失败，结果它**纹丝不动**——因为 acct-tools 与
        //       bank-recon 共享 4 个场景词（表格核对/对账/勾稽/流水核对），
        //       而与其他技能共享 0 个，它是唯一的跨技能候选，根本没有"同分对手"可比。
        //       即：**成功率排序只在该生效的地方生效**，不能改变候选的相对能力判断。
        //     故这里改为直接断言排序语义本身：分高的必须压过分低的，成功率只在分相同时起作用。
        //     用一个可复现的构造：同技能兄弟动作分都是 0，此时才轮到成功率说话。
        SkillStats.ResetForTest();
        SkillStats.Record("bank-recon", "recon", false);
        SkillStats.Record("bank-recon", "recon", false);
        SkillStats.Record("bank-recon", "recon", false);
        SkillStats.InvalidateCache();
        List<SkillActionSpec> afterAlts = SkillToolBridge.Alternatives(root, "acct-tools", "bank-recon", 10);
        // 重叠分优先：即便如此，bank-recon/recon 仍然是唯一跨技能候选，位置不该被撼动
        bool stillFirst = afterAlts.Count > 0
            && afterAlts[0].SkillId == "bank-recon" && afterAlts[0].Action == "recon";
        if (!stillFirst) failed++;
        Console.WriteLine((stillFirst ? "[OK]  " : "[FAIL] ") +
            "重叠分优先于成功率（差评不撼动高分候选，" + (afterAlts.Count > 0 ? afterAlts[0].SkillId : "空") + "）");

        // 7c. 同分（重叠分都为 0）的**同技能兄弟动作**内部，成功率高的排前面。
        //     ★ 必须挑"跨技能候选之后"的位置。上一步连着两次踩了同一个坑：
        //       候选表里第 1 项始终是跨技能候选（bank-recon/recon，重叠分 4），
        //       而成功率只作同分二级排序，**不可能**把它挤下去——给它记差评当然没反应。
        //       真正由成功率决定次序的是它后面的同技能兄弟（分数前缀全为 0000）。
        //     故这里在候选表里定位第一个「非 bank-recon」的项来降权。
        SkillStats.ResetForTest();
        SkillStats.InvalidateCache();
        List<SkillActionSpec> sib = SkillToolBridge.Alternatives(root, "acct-tools", "vat", 10);
        int sibIdx = -1;
        for (int i = 0; i < sib.Count; i++)
        {
            if (sib[i].SkillId == "acct-tools") { sibIdx = i; break; }
        }
        if (sibIdx >= 0 && sib.Count - sibIdx > 1)
        {
            SkillActionSpec v = sib[sibIdx];
            for (int i = 0; i < 3; i++) SkillStats.Record(v.SkillId, v.Action, false);
            SkillStats.InvalidateCache();
            List<SkillActionSpec> sib2 = SkillToolBridge.Alternatives(root, "acct-tools", "vat", 10);
            // 降权后：该动作不应再是"同技能兄弟里的第一个"
            int newFirstSib = -1;
            for (int i = 0; i < sib2.Count; i++)
            {
                if (sib2[i].SkillId == "acct-tools") { newFirstSib = i; break; }
            }
            bool demoted = !(newFirstSib >= 0 && sib2[newFirstSib].Action == v.Action);
            if (!demoted) failed++;
            Console.WriteLine((demoted ? "[OK]  " : "[FAIL] ") + "同分兄弟动作按成功率降权（" +
                v.Action + " → 同技能首个变为 " +
                (newFirstSib >= 0 ? sib2[newFirstSib].Action : "无") + "）");
        }
        else
        {
            Console.WriteLine("[INFO] 无足够的同技能兄弟候选，跳过降权断言");
        }

        SkillStats.ResetForTest();
        SkillStats.InvalidateCache();
        double rUnknown = SkillStats.Rate("acct-tools", "vat");
        ok = rUnknown < 0;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "清空台账后成功率回到未知（不残留旧数据）");

        SkillStats.ResetForTest();
        SkillStats.InvalidateCache();

        // ---------- 第三期：目标继承 ----------

        // 8. 技能从工具日志里解析：状态符号（✓/✗/跳过）不得混进工具名
        string logSample = "🔧 list_directory ✓\n🔧 skill_acct_tools_trial_balance ✓\n" +
            "🔧 skill_acct_tools_statements ✗\n🔧 skill_acct_tools_statements ↷跳过\n" +
            "🔧 read_text_file ✓";
        List<string> sk = TurnMemory.SkillsFromLog(logSample);
        ok = sk.Count == 2
            && sk[0] == "skill_acct_tools_trial_balance"
            && sk[1] == "skill_acct_tools_statements";
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "从工具日志解析技能（" +
            string.Join(",", sk.ToArray()) + "，期望 2 项且不含状态符号）");

        // 9. 非技能工具不入表（读文件只是手段，不是"在做什么"的信号）
        bool noPlain = true;
        foreach (string s in sk) { if (s == "read_text_file" || s == "list_directory") noPlain = false; }
        if (!noPlain) failed++;
        Console.WriteLine((noPlain ? "[OK]  " : "[FAIL] ") + "普通工具不入目标继承（只记技能）");

        // 10. 产物去重 + 空值跳过
        List<string> prodIn = new List<string>();
        prodIn.Add("C:\\a.xlsx"); prodIn.Add(""); prodIn.Add("C:\\a.xlsx"); prodIn.Add(null);
        prodIn.Add("C:\\b.pdf");
        List<string> prod = TurnMemory.Products(prodIn);
        ok = prod.Count == 2;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "产物去重且跳过空值（" + prod.Count + " 项，期望 2）");

        // 11. 目标继承文本：必须含"仅指代时参考"的限定，且不得是命令语气
        string inherit = TurnMemory.Build(sk, prod, 4);
        ok = inherit.IndexOf("skill_acct_tools_trial_balance") >= 0
            && inherit.IndexOf("a.xlsx") >= 0
            && inherit.IndexOf("指代上一轮") >= 0
            && inherit.IndexOf("若用户开启了新话题") >= 0;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "目标继承含技能/产物/限定语（len=" + inherit.Length + "）");

        // 12. 无内容时不产生文本（否则每轮都塞固定文本，白烧 token）
        string emptyInherit = TurnMemory.Build(new List<string>(), new List<string>(), 4);
        ok = emptyInherit.Length == 0;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "无上一轮内容时不产生继承段（len=" + emptyInherit.Length + "）");

        // 12b. 递增不得覆盖别的进程写入（初版用进程内快照整表覆盖 → 会静默丢数据）。
        //      构造：A 记账后**丢弃缓存**（模拟另一个进程刚写过盘），B 再记账时
        //      必须重读盘、把 A 的那笔一并保留，而不是拿旧快照盖掉。
        SkillStats.ResetForTest();
        SkillStats.Record("acct-tools", "vat", true);
        SkillStats.InvalidateCache();              // 丢弃缓存；盘上有 1 笔
        SkillStats.Record("acct-tools", "aging", true);   // 再记另一动作
        SkillStats.InvalidateCache();
        SkillStats.Entry vatAfter = SkillStats.Get("acct-tools", "vat");
        SkillStats.Entry agingAfter = SkillStats.Get("acct-tools", "aging");
        ok = vatAfter != null && vatAfter.Ok == 1 && agingAfter != null && agingAfter.Ok == 1;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "记账不覆盖既有记录（vat=" +
            (vatAfter == null ? "null" : vatAfter.Ok.ToString()) + " aging=" +
            (agingAfter == null ? "null" : agingAfter.Ok.ToString()) + "，期望各 1）");

        // 12c. 落盘后不得残留临时文件（初版固定 .tmp 名，且 Delete+Move 有丢文件窗口）
        bool noTmpLeft = true;
        try
        {
            string dir = System.IO.Path.GetDirectoryName(SkillStats.FilePath());
            foreach (string f2 in System.IO.Directory.GetFiles(dir, "skill-stats.json.*"))
            {
                if (f2.EndsWith(".tmp")) noTmpLeft = false;
            }
        }
        catch { }
        if (!noTmpLeft) failed++;
        Console.WriteLine((noTmpLeft ? "[OK]  " : "[FAIL] ") + "落盘后无残留临时文件");
        SkillStats.ResetForTest();
        SkillStats.InvalidateCache();

        // 13. 产物超量时只列 maxProducts 个，其余报数量（防提示膨胀）
        List<string> many = new List<string>();
        for (int i = 0; i < 7; i++) many.Add("C:\\f" + i + ".xlsx");
        string manyText = TurnMemory.Build(null, many, 4);
        ok = manyText.IndexOf("f3.xlsx") >= 0 && manyText.IndexOf("f4.xlsx") < 0
            && manyText.IndexOf("还有 3 个") >= 0;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "产物超量只列前 4 个并报余量");

        Console.WriteLine(failed == 0 ? "ALL PASS" : (failed + " FAILED"));
        return failed == 0 ? 0 : 2;
    }

    // 低配优化自测：OfficeAgent.exe /perftest
    //
    // 覆盖两件事，都必须**断言到数字**而不是"跑通了"：
    //   ① 正确性：新的有界预览 XlsxPreview 与旧的 LoadGrid 必须逐单元格一致。
    //      这是本类最重要的一条——性能优化若改变了数据，用户看到的就是错的表。
    //      实测中就靠它抓到"合并单元格没回填"（跨列标题只剩第一列）。
    //   ② 内存上界：预览保留的行数必须有界，不能随表的大小增长。
    //      这是"低配能跑"的核心保证：内存占用从 O(总行数) 变成 O(保留行数)。
    static int RunPerfTest()
    {
        int failed = 0;
        string dir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "oa_perftest");
        try
        {
            if (System.IO.Directory.Exists(dir)) System.IO.Directory.Delete(dir, true);
            System.IO.Directory.CreateDirectory(dir);

            // ---------- ① 数据一致性：有界预览 vs 全量网格 ----------
            // 造一张"宽 + 长 + 带合并单元格"的表，一次覆盖三类风险。
            string wide = System.IO.Path.Combine(dir, "wide.csv");
            StringBuilder csv = new StringBuilder();
            for (int c = 0; c < 40; c++) { if (c > 0) csv.Append(','); csv.Append("列" + c); }
            csv.Append('\n');
            for (int r = 0; r < 5000; r++)
            {
                for (int c = 0; c < 40; c++)
                {
                    if (c > 0) csv.Append(',');
                    csv.Append("v" + r + "_" + c);
                }
                csv.Append('\n');
            }
            System.IO.File.WriteAllText(wide, csv.ToString(), new UTF8Encoding(false));
            string wideXlsx = System.IO.Path.Combine(dir, "wide.xlsx");
            bool madeWide = false;
            try
            {
                // 用项目自己的写入器造 xlsx（不依赖 LibreOffice，CI/无 LO 环境也能跑）
                ReportSheet sh = new ReportSheet("S1");
                for (int r = 0; r <= 5000; r++)
                {
                    string[] cells = new string[40];
                    for (int c = 0; c < 40; c++) cells[c] = (r == 0) ? ("列" + c) : ("v" + (r - 1) + "_" + c);
                    sh.AddTextRow(cells);
                }
                List<ReportSheet> sheets = new List<ReportSheet>();
                sheets.Add(sh);
                string saveErr = MiniXlsxWrite.Save(wideXlsx, sheets);
                madeWide = saveErr == null && System.IO.File.Exists(wideXlsx);
                if (!madeWide) Console.WriteLine("[INFO] 造表失败: " + saveErr);
            }
            catch { }

            if (madeWide)
            {
                XlsxBook bk = XlsxBook.Open(wideXlsx);
                int tr1, uc1;
                string[,] full = bk.LoadGrid(0, 6000, 64, out tr1, out uc1);
                bk.Dispose();

                XlsxPreview pv = XlsxPreview.Load(wideXlsx, 0, XlsxPreview.DefaultPreviewRows, XlsxPreview.MaxPreviewCols);
                int diff = 0;
                string firstDiff = "";
                int cmpRows = Math.Min(full.GetLength(0), pv.Rows.Length);
                int cmpCols = Math.Min(full.GetLength(1), pv.UsedCols);
                for (int r = 0; r < cmpRows; r++)
                {
                    for (int c = 0; c < cmpCols; c++)
                    {
                        string a = full[r, c] == null ? "" : full[r, c];
                        string b = (c < pv.Rows[r].Length && pv.Rows[r][c] != null) ? pv.Rows[r][c] : "";
                        if (a != b)
                        {
                            if (firstDiff.Length == 0)
                                firstDiff = "r" + r + "c" + c + " 全量='" + a + "' 预览='" + b + "'";
                            diff++;
                        }
                    }
                }
                bool same = diff == 0 && cmpRows == XlsxPreview.DefaultPreviewRows;
                if (!same) failed++;
                Console.WriteLine((same ? "[OK]  " : "[FAIL] ") + "有界预览与全量网格逐单元格一致（比对 " +
                    cmpRows + "x" + cmpCols + "，差异 " + diff + (firstDiff.Length > 0 ? "，首个 " + firstDiff : "") + "）");

                // 总行数必须如实统计（预览只留前 N 行，但**总行数不能是 N**）
                bool totalOk = pv.TotalRows == 5001;
                if (!totalOk) failed++;
                Console.WriteLine((totalOk ? "[OK]  " : "[FAIL] ") + "预览保留行有界但总行数如实（保留=" +
                    pv.Rows.Length + " 总数=" + pv.TotalRows + "，期望 5001）");

                // ② 内存上界：保留行数不得超过配置上限，且与表大小无关
                bool bounded = pv.Rows.Length <= XlsxPreview.DefaultPreviewRows;
                if (!bounded) failed++;
                Console.WriteLine((bounded ? "[OK]  " : "[FAIL] ") + "保留行数有界（" + pv.Rows.Length +
                    " <= " + XlsxPreview.DefaultPreviewRows + "）");

                // 列数也必须有界（宽表不能靠列数把内存撑爆）
                XlsxPreview widePv = XlsxPreview.Load(wideXlsx, 0, 10, XlsxPreview.MaxPreviewCols);
                bool colBounded = widePv.UsedCols <= XlsxPreview.MaxPreviewCols;
                if (!colBounded) failed++;
                Console.WriteLine((colBounded ? "[OK]  " : "[FAIL] ") + "列数有界（" + widePv.UsedCols +
                    " <= " + XlsxPreview.MaxPreviewCols + "）");

                // 预览信息必须如实标注"仅预览前 N 行 / 共 M 行"
                string desc = pv.Describe(1, "S1");
                bool descOk = desc.IndexOf("仅预览前 " + XlsxPreview.DefaultPreviewRows + " 行") >= 0
                    && desc.IndexOf("共 5001 行") >= 0;
                if (!descOk) failed++;
                Console.WriteLine((descOk ? "[OK]  " : "[FAIL] ") + "预览提示如实标注截断（" + desc + "）");
            }
            else
            {
                Console.WriteLine("[INFO] 跳过 xlsx 一致性断言（MiniXlsxWrite 不可用）");
            }

            // ---------- ③ 流式读取条目：必须与全量解压结果一致 ----------
            // MiniZip.OpenEntryStream 是本次低配优化的核心（省掉工作表 XML 的整份解压）。
            // 它放弃了 CRC32，故必须证明"读出来的字节与全量解压完全相同"。
            if (madeWide)
            {
                MiniZipFile z = MiniZipFile.OpenRead(wideXlsx);
                MiniZipEntryInfo target = null;
                foreach (MiniZipEntryInfo e in z.Entries)
                {
                    if (e.FullName != null && e.FullName.StartsWith("xl/worksheets/")) { target = e; break; }
                }
                bool streamOk = false;
                string streamMsg = "未找到工作表条目";
                if (target != null)
                {
                    // 全量解压
                    byte[] fullBytes;
                    using (Stream fs1 = z.OpenEntry(target.FullName))
                    {
                        fullBytes = new byte[fs1.Length];
                        int off = 0;
                        while (off < fullBytes.Length)
                        {
                            int n = fs1.Read(fullBytes, off, fullBytes.Length - off);
                            if (n <= 0) break;
                            off += n;
                        }
                    }
                    // 流式解压
                    byte[] streamBytes;
                    using (Stream fs2 = z.OpenEntryStream(target.FullName))
                    {
                        MemoryStream acc = new MemoryStream();
                        byte[] buf = new byte[8192];
                        int n;
                        while ((n = fs2.Read(buf, 0, buf.Length)) > 0) acc.Write(buf, 0, n);
                        streamBytes = acc.ToArray();
                    }
                    streamOk = fullBytes.Length == streamBytes.Length;
                    if (streamOk)
                    {
                        for (int i = 0; i < fullBytes.Length; i++)
                        {
                            if (fullBytes[i] != streamBytes[i]) { streamOk = false; streamMsg = "第 " + i + " 字节不同"; break; }
                        }
                    }
                    else streamMsg = "长度不同 全量=" + fullBytes.Length + " 流式=" + streamBytes.Length;
                    if (streamOk) streamMsg = fullBytes.Length + " 字节完全一致";
                }
                if (!streamOk) failed++;
                Console.WriteLine((streamOk ? "[OK]  " : "[FAIL] ") + "流式解压与全量解压字节一致（" + streamMsg + "）");
                z.Dispose();
            }

            // ---------- ④ 流式读取必须能发现截断 ----------
            // 放弃了 CRC32，就必须保证"数据坏了会报错"而不是静默少给几行。
            string broken = System.IO.Path.Combine(dir, "broken.xlsx");
            bool truncatedDetected = false;
            try
            {
                // 复制一个 xlsx，把中央目录声明的解压长度改大（模拟截断/损坏）
                if (madeWide)
                {
                    byte[] raw = System.IO.File.ReadAllBytes(wideXlsx);
                    System.IO.File.WriteAllBytes(broken, raw);
                    // 直接把文件截掉尾部一段 → 解压必然不足
                    using (FileStream fs = new FileStream(broken, FileMode.Open, FileAccess.Write))
                    {
                        fs.SetLength(fs.Length - 200);
                    }
                    try
                    {
                        MiniZipFile zb = MiniZipFile.OpenRead(broken);
                        foreach (MiniZipEntryInfo e in zb.Entries)
                        {
                            if (e.FullName != null && e.FullName.StartsWith("xl/worksheets/"))
                            {
                                using (Stream st = zb.OpenEntryStream(e.FullName))
                                {
                                    byte[] buf = new byte[8192];
                                    while (st.Read(buf, 0, buf.Length) > 0) { }
                                }
                                break;
                            }
                        }
                        zb.Dispose();
                    }
                    catch { truncatedDetected = true; }
                }
            }
            catch { }
            if (!truncatedDetected) failed++;
            Console.WriteLine((truncatedDetected ? "[OK]  " : "[FAIL] ") +
                "流式读取能发现截断（不静默少给数据）");
        }
        finally
        {
            try { System.IO.Directory.Delete(dir, true); } catch { }
        }

        Console.WriteLine(failed == 0 ? "ALL PASS" : (failed + " FAILED"));
        return failed == 0 ? 0 : 2;
    }

    // 技能→工具投影自测（规划层第一期）：OfficeAgent.exe /bridgetest
    // 验证 skill.json 的 actions/actionParams 能被正确投影成 function-calling schema，
    // 并且模型给出的工具名能反查回技能、参数能正确落成 request.json。
    static int RunBridgeTest()
    {
        int failed = 0;
        string root = EnvDetect.FindRoot();
        SkillToolBridge.InvalidateCache();

        // 1. 投影：五个技能的动作总数应为 19
        List<SkillActionSpec> all = SkillToolBridge.Collect(root);
        bool ok = all.Count == 19;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "动作投影（" + all.Count + " 个动作，期望 19）");

        // 2. 每个动作都能定位到技能且名字唯一
        Dictionary<string, int> nameCount = new Dictionary<string, int>();
        bool uniq = true;
        foreach (SkillActionSpec sp in all)
        {
            string n = sp.SkillId + "/" + sp.Action;
            if (nameCount.ContainsKey(n)) uniq = false;
            else nameCount[n] = 1;
        }
        if (!uniq) failed++;
        Console.WriteLine((uniq ? "[OK]  " : "[FAIL] ") + "动作名唯一性");

        // 3. schema JSON 结构合法：括号配平、逗号数量正确
        string schemas = SkillToolBridge.BuildSchemas(root);
        int depth = 0; bool balanced = true;
        foreach (char c in schemas)
        {
            if (c == '{' || c == '[') depth++;
            else if (c == '}' || c == ']') { depth--; if (depth < 0) { balanced = false; break; } }
        }
        balanced = balanced && depth == 0 && schemas.Length > 0;
        // 不能出现空的 properties 后紧跟空 required 造成 "[," 之类的非法拼接
        bool noDangling = schemas.IndexOf(",]") < 0 && schemas.IndexOf("[,]") < 0 && schemas.IndexOf("{,") < 0;
        ok = balanced && noDangling;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "schema JSON 配平（len=" + schemas.Length +
            " balanced=" + balanced + " dangling=" + !noDangling + "）");

        // 4. 拼接进完整工具表后仍是合法 JSON 片段（回归：空技能表时不得产生 "[...,]"）
        string baseTools = AgentTools.SchemasJson(true);
        string joined = baseTools.Substring(0, baseTools.Length - 1) + "," + schemas + "]";
        depth = 0; balanced = true;
        foreach (char c in joined)
        {
            if (c == '{' || c == '[') depth++;
            else if (c == '}' || c == ']') { depth--; if (depth < 0) { balanced = false; break; } }
        }
        balanced = balanced && depth == 0;
        // 顶层应是 8 个基础工具（6 通用 + task_plan + repair_environment 等）+ 19 个技能动作 = 27 个 "name"
        // 注意：此处的 "name" 计数也包含每个工具 parameters 里名为 name 的参数（如 xlsx-ops 的 add-sheet.name），
        // 所以这里比对的是"拼接后总数必须 = 基础数 + 技能数"，基础数由同一函数单独取一次保证一致。
        string baseOnly = AgentTools.SchemasJson(true);
        int baseNames = 0;
        int bi = 0;
        while (true)
        {
            int at = baseOnly.IndexOf("\"name\":\"", bi, StringComparison.Ordinal);
            if (at < 0) break;
            baseNames++;
            bi = at + 8;
        }
        int skillNames = 0;
        int si = 0;
        while (true)
        {
            int at = schemas.IndexOf("\"name\":\"", si, StringComparison.Ordinal);
            if (at < 0) break;
            skillNames++;
            si = at + 8;
        }
        int toolCount = 0;
        int idx = 0;
        while (true)
        {
            int at = joined.IndexOf("\"name\":\"", idx, StringComparison.Ordinal);
            if (at < 0) break;
            toolCount++;
            idx = at + 8;
        }
        ok = balanced && toolCount == baseNames + skillNames && toolCount > baseNames;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "完整工具表拼接（基础=" + baseNames +
            " 技能=" + skillNames + " 合计=" + toolCount + "，balanced=" + balanced + "）");

        // 5. 反查：用生成的名字能找回原技能动作（含 '-' 被转 '_' 的情形）
        SkillActionSpec r1 = SkillToolBridge.Resolve(root, "skill_acct_tools_trial_balance");
        SkillActionSpec r2 = SkillToolBridge.Resolve(root, "skill_xlsx_ops_add_sheet");
        SkillActionSpec r3 = SkillToolBridge.Resolve(root, "skill_acct_tools_nonexistent");
        ok = r1 != null && r1.SkillId == "acct-tools" && r1.Action == "trial-balance"
            && r2 != null && r2.SkillId == "xlsx-ops" && r2.Action == "add-sheet"
            && r3 == null;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "工具名反查（trial-balance=" +
            (r1 == null ? "?" : r1.Action) + " add-sheet=" + (r2 == null ? "?" : r2.Action) +
            " 未知名=" + (r3 == null ? "null(正确)" : "误命中") + "）");

        // 6. 参数解析：acct-tools/vat 应有两个 number 必填参数
        SkillActionSpec vat = SkillToolBridge.Resolve(root, "skill_acct_tools_vat");
        ok = vat != null && vat.Params.Count == 2;
        if (ok)
        {
            foreach (string[] p in vat.Params)
            {
                if (p[1] != "number" || p[2] != "required") ok = false;
            }
        }
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "参数解析（vat 参数数=" +
            (vat == null ? "?" : vat.Params.Count.ToString()) + "）");

        // 7. 必填缺失 → 结构化报错，而不是让 Python 端报晦涩错误
        SkillActionSpec dv = SkillToolBridge.Resolve(root, "skill_acct_tools_depreciation");
        bool missOk = false;
        string missMsg = SkillToolBridge.Invoke(root, dv, new Dictionary<string, string>(), null, out missOk);
        ok = !missOk && missMsg.IndexOf("cost") >= 0 && missMsg.IndexOf("years") >= 0;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "必填缺失拦截（" + missMsg + "）");

        // 8. 端到端：走 AgentTools.Dispatch 真实执行 vat（纯计算，无输入文件）
        Dictionary<string, string> vatArgs = new Dictionary<string, string>();
        vatArgs["sales"] = "100000";
        vatArgs["purchases"] = "60000";
        List<string> prods = new List<string>();
        bool vatOk = false;
        string vatMsg = AgentTools.Dispatch("skill_acct_tools_vat", "{\"sales\":\"100000\",\"purchases\":\"60000\"}",
            null, new ConvertEngine(), prods, out vatOk);
        // 一般计税 100000*13% - 60000*13% = 5200
        ok = vatOk && vatMsg.IndexOf("5200") >= 0;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "Dispatch 执行技能（" + vatMsg.Replace("\n", " ") + "）");

        // 9. 未知技能工具名 → 明确报错（不得静默当成基础工具）
        bool unkOk = false;
        string unkMsg = AgentTools.Dispatch("skill_does_not_exist", "{}", null, new ConvertEngine(), null, out unkOk);
        ok = !unkOk && unkMsg.IndexOf("未知技能工具") >= 0;
        if (!ok) failed++;
        Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "未知技能工具拦截（" + unkMsg + "）");

        // 10. 向后兼容：无 actions 段的老技能不投影，且不影响其余技能
        string tmpUserSkills = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "oa-bridgetest-user");
        try
        {
            if (System.IO.Directory.Exists(tmpUserSkills)) System.IO.Directory.Delete(tmpUserSkills, true);
            string legacy = System.IO.Path.Combine(tmpUserSkills, "legacy-skill");
            System.IO.Directory.CreateDirectory(legacy);
            System.IO.File.WriteAllText(System.IO.Path.Combine(legacy, "skill.json"),
                "{\"id\":\"legacy-skill\",\"name\":\"老技能\",\"version\":\"1.0.0\",\"host\":\"1.0.0\"," +
                "\"runtime\":\"builtin\",\"entry\":\"recon\",\"scenarios\":\"老\"}",
                new UTF8Encoding(false));
            List<SkillRegistryEntry> reg = SkillSystem.Scan(root, null, tmpUserSkills, false);
            int legacyActions = 0;
            foreach (SkillRegistryEntry e in reg)
            {
                if (e.Id == "legacy-skill") legacyActions += SkillToolBridge.ParseActions(e).Count;
            }
            ok = legacyActions == 0;
            if (!ok) failed++;
            Console.WriteLine((ok ? "[OK]  " : "[FAIL] ") + "无 actions 的老技能不投影（动作数=" + legacyActions + "）");
        }
        finally { try { System.IO.Directory.Delete(tmpUserSkills, true); } catch { } }

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
