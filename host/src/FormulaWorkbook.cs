// FormulaWorkbook —— create_formula_workbook 工具：生成「带公式、用户填数即自动计算」的 xlsx。
// 三种用法：
//   ① 模板模式 template=payroll|vat|ledger：列结构与公式由本文件内置定义保证正确，模型只给数据行；
//   ② 自由模式 sheets=JSON 规格：任意表头/数据/公式列（{r} 代表当前行号）；
//   ③ summary 参数（自由模式）：按分组列 SUMIF 自动生成汇总 sheet——改明细、汇总自动变。
// 公式一律无缓存值写入（ReportCell.F），配合 workbook.xml 的 calcPr fullCalcOnLoad 打开即重算。
// 设计要点：模板模式的公式列在数据行之外还预置 N 行「空白公式行」（IF 守卫显示为空），
//           用户在 Excel 里继续填数即自动计算；合计/SUMIF 区间覆盖数据+空白行。
// 红线：C# 3.0 语法；参数是模型给的数据，绝不参与命令行/shell 构造。
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using OfficeAgent.Core;

namespace OfficeAgent.Host
{
    public static class FormulaWorkbook
    {
        const int DefaultBlankRows = 50;
        const int LedgerBlankRows = 100;

        public static string Create(string argsJson, out bool ok)
        {
            ok = false;
            Dictionary<string, object> a;
            try { a = JsonVal.ParseObject(argsJson == null ? "{}" : argsJson); }
            catch (Exception ex) { return "参数不是合法 JSON: " + ex.Message; }

            string fenceErr;
            string p = AgentTools.SafeOutputPath(JsonVal.Str(a, "path"), ".xlsx", out fenceErr);
            if (p == null) return fenceErr;

            string template = JsonVal.Str(a, "template");
            List<ReportSheet> sheets;
            if (template != null && template.Trim().Length > 0)
            {
                sheets = BuildTemplate(template.Trim().ToLowerInvariant(),
                    JsonVal.Str(a, "rows"), JsonVal.Str(a, "params"));
                if (sheets == null) return "未知模板: " + template + "（可用：payroll=工资表, vat=增值税台账, ledger=流水账）";
            }
            else if (JsonVal.List(a, "sheets") != null)
            {
                string buildErr;
                sheets = BuildFree(JsonVal.List(a, "sheets"), JsonVal.Str(a, "summary"), out buildErr);
                if (sheets == null) return buildErr;
            }
            else
            {
                return "缺少内容：请给 template（payroll|vat|ledger）或 sheets（自由规格 JSON）。";
            }

            string saveErr = MiniXlsxWrite.Save(p, sheets);
            if (saveErr != null) return "生成失败: " + saveErr;
            ok = true;
            AgentTools.LastProduct = p;
            AuditLog.Record("file_write", "agent_make_formula_xlsx " + p);
            StringBuilder names = new StringBuilder();
            for (int i = 0; i < sheets.Count; i++)
            {
                if (i > 0) names.Append("/");
                names.Append(sheets[i].Name);
            }
            return "已生成带公式的工作簿: " + p + "\n包含工作表: " + names.ToString() +
                "\n公式已预置，用户直接填数即可自动计算（汇总页随明细联动）；行不够时在合计行上方插入行。";
        }

        // ---------- 基础换算 ----------

        // 模型给的文本 → ReportCell："="开头=公式（<f> 内不带等号）；可解析数值=数值；其余=文本。
        // 前导 0 的纯数字串（单号/编码）按文本保留。
        static ReportCell Cell(object raw, string style)
        {
            string t = raw as string;
            if (t == null && raw is double) return ReportCell.N((double)raw, style);
            if (t == null) return null;
            t = t.Trim();
            if (t.Length == 0) return null;
            if (t[0] == '=' && t.Length > 1) return ReportCell.F(t.Substring(1), style);
            double num;
            bool leadingZero = t.Length > 1 && t[0] == '0';
            if (!leadingZero && double.TryParse(t, NumberStyles.Any, CultureInfo.InvariantCulture, out num))
                return ReportCell.N(num, style);
            ReportCell c = new ReportCell(t);
            c.Style = style;
            return c;
        }

        static string ColName(int idx)
        {
            StringBuilder sb = new StringBuilder();
            idx++;
            while (idx > 0)
            {
                int m = (idx - 1) % 26;
                sb.Insert(0, (char)('A' + m));
                idx = (idx - 1) / 26;
            }
            return sb.ToString();
        }

        static int ColIndex(string col)
        {
            if (col == null || col.Length == 0 || col.Length > 3) return -1;
            int v = 0;
            string up = col.ToUpperInvariant();
            for (int i = 0; i < up.Length; i++)
            {
                char ch = up[i];
                if (ch < 'A' || ch > 'Z') return -1;
                v = v * 26 + (ch - 'A' + 1);
            }
            return v - 1;
        }

        // "key=value;key2=value2" → 字典（模板参数覆盖）
        static Dictionary<string, string> ParseParams(string s)
        {
            Dictionary<string, string> d = new Dictionary<string, string>();
            if (s == null) return d;
            foreach (string pair in s.Split(';', '；'))
            {
                string seg = pair.Trim();
                if (seg.Length == 0) continue;
                int eq = seg.IndexOf('=');
                if (eq <= 0) continue;
                d[seg.Substring(0, eq).Trim().ToLowerInvariant()] = seg.Substring(eq + 1).Trim();
            }
            return d;
        }

        static double ParamNum(Dictionary<string, string> ps, string key, double def)
        {
            if (!ps.ContainsKey(key)) return def;
            double v;
            if (double.TryParse(ps[key], NumberStyles.Any, CultureInfo.InvariantCulture, out v)) return v;
            return def;
        }

        // 模板数据行 CSV：允许为空（只给结构）；容错剥掉模型顺手带的表头行
        static List<string[]> ParseCsvRows(string csv)
        {
            if (csv == null || csv.Trim().Length == 0) return new List<string[]>();
            List<string[]> rows = MiniCsv.Parse(csv);
            if (rows.Count > 0)
            {
                string first = (rows[0][0] == null ? "" : rows[0][0]).Trim();
                if (first == "姓名" || first == "日期") rows.RemoveAt(0);
            }
            return rows;
        }

        static ReportSheet HeaderSheet(string name, string[] header, int[] widths)
        {
            ReportSheet sh = new ReportSheet(name);
            List<ReportCell> row = new List<ReportCell>();
            foreach (string h in header)
            {
                ReportCell c = new ReportCell(h);
                c.Style = "h";
                row.Add(c);
            }
            sh.Rows.Add(row);
            sh.ColWidths = widths;
            sh.FreezeRows = 1;
            return sh;
        }

        static ReportSheet TextSheet(string name, string[] lines)
        {
            ReportSheet sh = new ReportSheet(name);
            foreach (string line in lines)
            {
                List<ReportCell> row = new List<ReportCell>();
                ReportCell c = new ReportCell(line);
                c.Style = "p";
                row.Add(c);
                sh.Rows.Add(row);
            }
            return sh;
        }

        static List<ReportCell> TriRow(string k, string note)
        {
            List<ReportCell> row = new List<ReportCell>();
            ReportCell c = new ReportCell(k);
            row.Add(c);
            row.Add(null);
            row.Add(note == null ? null : new ReportCell(note));
            return row;
        }

        static string CellAt(string[] arr, int idx)
        {
            if (arr == null || idx >= arr.Length) return null;
            return arr[idx];
        }

        // ============================================================
        // 模板模式
        // ============================================================
        static List<ReportSheet> BuildTemplate(string template, string rowsCsv, string paramsStr)
        {
            if (template == "payroll") return BuildPayroll(rowsCsv, paramsStr);
            if (template == "vat") return BuildVat(rowsCsv);
            if (template == "ledger") return BuildLedger(rowsCsv, paramsStr);
            return null;
        }

        // ---- 工资表（标准套账）：参数 / 工资表 / 说明 ----
        static List<ReportSheet> BuildPayroll(string rowsCsv, string paramsStr)
        {
            Dictionary<string, string> ps = ParseParams(paramsStr);
            double pension = ParamNum(ps, "pensionrate", 0.08);
            double medical = ParamNum(ps, "medicalrate", 0.02);
            double unemployment = ParamNum(ps, "unemploymentrate", 0.005);
            double fund = ParamNum(ps, "housingfundrate", 0.12);
            double threshold = ParamNum(ps, "taxthreshold", 5000);
            int blank = (int)ParamNum(ps, "blankrows", DefaultBlankRows);
            if (blank < 0) blank = 0;

            List<string[]> data = ParseCsvRows(rowsCsv);
            int n = data.Count;
            int last = 1 + n + blank;   // 表头 1 行 + 数据 + 空白公式行

            ReportSheet cfg = HeaderSheet("参数", new string[] { "项目", "数值", "说明" }, new int[] { 24, 12, 56 });
            cfg.Rows.Add(ParamRow("养老个人比例", pension, "占缴费基数；模板按基本工资作基数，可让 agent 调整"));
            cfg.Rows.Add(ParamRow("医疗个人比例", medical, null));
            cfg.Rows.Add(ParamRow("失业个人比例", unemployment, null));
            // 合计比例是公式：用户改分项，全表自动联动
            List<ReportCell> sumRow = TriRow("社保个人合计比例", "公式=分项相加，改分项自动更新");
            sumRow[1] = ReportCell.F("B2+B3+B4", "m");
            cfg.Rows.Add(sumRow);
            cfg.Rows.Add(ParamRow("公积金个人比例", fund, null));
            cfg.Rows.Add(ParamRow("个税起征点（月）", threshold, "全国统一减除费用"));

            ReportSheet sh = HeaderSheet("工资表",
                new string[] { "序号", "姓名", "部门", "基本工资", "岗位津贴", "加班费",
                               "应发合计", "社保个人", "公积金个人", "应纳税所得额", "个税", "实发合计" },
                new int[] { 6, 10, 12, 12, 12, 12, 12, 12, 13, 14, 12, 12 });
            for (int i = 0; i < n; i++)
            {
                string[] d = data[i];
                sh.Rows.Add(PayrollRow(i + 2, CellAt(d, 0), CellAt(d, 1), CellAt(d, 2), CellAt(d, 3), CellAt(d, 4)));
            }
            for (int b = 0; b < blank; b++)
            {
                sh.Rows.Add(PayrollRow(n + b + 2, null, null, null, null, null));
            }
            // 合计行：数值列 SUM（区间含空白公式行，SUM 忽略其中的空串）
            sh.Rows.Add(PayrollTotalRow(last));

            ReportSheet help = TextSheet("说明", new string[] {
                "工资表（标准套账）使用说明",
                "1. 直接在「工资表」页填：姓名、部门、基本工资、岗位津贴、加班费——其余列自动计算。",
                "2. 已预置空白公式行，继续往下填即可；行不够时在「合计」行上方插入行，合计自动扩展。",
                "3. 社保/公积金比例在「参数」页修改，全表自动联动。",
                "4. 个税按月度税率表计算（应纳税所得额=应发-起征点-社保-公积金）；如需累计预扣预缴，可让 agent 加列。",
                "5. 打开文件即自动重算；若个别单元格显示异常，按 F9 重算。" });

            List<ReportSheet> sheets = new List<ReportSheet>();
            sheets.Add(cfg); sheets.Add(sh); sheets.Add(help);
            return sheets;
        }

        static List<ReportCell> ParamRow(string k, double v, string note)
        {
            List<ReportCell> row = TriRow(k, note);
            row[1] = ReportCell.N(v, "m");
            return row;
        }

        // 工资表一行：姓名/部门/基本/岗位/加班 + 公式列；空行公式带 IF 守卫，显示为空不显 0
        static List<ReportCell> PayrollRow(int r, string name, string dept, string basePay, string post, string overtime)
        {
            List<ReportCell> row = new List<ReportCell>();
            row.Add(ReportCell.F("IF(B" + r + "=\"\",\"\"," + (r - 1) + ")", "n"));
            row.Add(Cell(name, "n"));
            row.Add(Cell(dept, "n"));
            row.Add(Cell(basePay, "m"));
            row.Add(Cell(post, "m"));
            row.Add(Cell(overtime, "m"));
            row.Add(ReportCell.F("IF(D" + r + "=\"\",\"\",ROUND(D" + r + "+E" + r + "+F" + r + ",2))", "m"));
            row.Add(ReportCell.F("IF(D" + r + "=\"\",\"\",ROUND(D" + r + "*参数!B5,2))", "m"));
            row.Add(ReportCell.F("IF(D" + r + "=\"\",\"\",ROUND(D" + r + "*参数!B6,2))", "m"));
            row.Add(ReportCell.F("IF(G" + r + "=\"\",\"\",ROUND(MAX(0,G" + r + "-参数!B7-H" + r + "-I" + r + "),2))", "m"));
            row.Add(ReportCell.F(
                "IF(J" + r + "=\"\",\"\",ROUND(IF(J" + r + "<=0,0," +
                "IF(J" + r + "<=3000,J" + r + "*0.03," +
                "IF(J" + r + "<=12000,J" + r + "*0.1-210," +
                "IF(J" + r + "<=25000,J" + r + "*0.2-1410," +
                "IF(J" + r + "<=35000,J" + r + "*0.25-2660," +
                "IF(J" + r + "<=55000,J" + r + "*0.3-4410," +
                "IF(J" + r + "<=80000,J" + r + "*0.35-7160," +
                "J" + r + "*0.45-15160))))))),2))", "m"));
            row.Add(ReportCell.F("IF(G" + r + "=\"\",\"\",ROUND(G" + r + "-H" + r + "-I" + r + "-K" + r + ",2))", "m"));
            return row;
        }

        static List<ReportCell> PayrollTotalRow(int last)
        {
            List<ReportCell> total = new List<ReportCell>();
            total.Add(new ReportCell("合计"));
            for (int c = 1; c < 12; c++)
            {
                if (c >= 3)
                {
                    string col = ColName(c);
                    total.Add(ReportCell.F("SUM(" + col + "2:" + col + last + ")", "b"));
                }
                else total.Add(null);
            }
            return total;
        }

        // ---- 增值税台账：台账 / 汇总 / 说明 ----
        static List<ReportSheet> BuildVat(string rowsCsv)
        {
            List<string[]> data = ParseCsvRows(rowsCsv);
            int n = data.Count;
            int last = 1 + n + LedgerBlankRows;

            ReportSheet sh = HeaderSheet("台账",
                new string[] { "日期", "摘要", "类型（销项/进项）", "金额（不含税）", "税率", "税额", "价税合计" },
                new int[] { 12, 24, 16, 14, 8, 12, 14 });
            for (int i = 0; i < n; i++)
            {
                string[] d = data[i];
                sh.Rows.Add(VatRow(i + 2, CellAt(d, 0), CellAt(d, 1), CellAt(d, 2), CellAt(d, 3), CellAt(d, 4)));
            }
            for (int b = 0; b < LedgerBlankRows; b++)
            {
                sh.Rows.Add(VatRow(n + b + 2, null, null, null, null, null));
            }

            ReportSheet sum = HeaderSheet("汇总", new string[] { "项目", "金额（元）" }, new int[] { 34, 16 });
            sum.Rows.Add(SumRow("销项金额合计", "SUMIF(台账!C2:C" + last + ",\"销项\",台账!D2:D" + last + ")"));
            sum.Rows.Add(SumRow("销项税额合计", "SUMIF(台账!C2:C" + last + ",\"销项\",台账!F2:F" + last + ")"));
            sum.Rows.Add(SumRow("进项金额合计", "SUMIF(台账!C2:C" + last + ",\"进项\",台账!D2:D" + last + ")"));
            sum.Rows.Add(SumRow("进项税额合计", "SUMIF(台账!C2:C" + last + ",\"进项\",台账!F2:F" + last + ")"));
            sum.Rows.Add(SumRow("应纳增值税（销项-进项，负数=留抵）", "ROUND(B3-B5,2)"));

            ReportSheet help = TextSheet("说明", new string[] {
                "增值税台账使用说明",
                "1. 在「台账」页逐笔登记：日期、摘要、类型只填「销项」或「进项」、不含税金额、税率（如 0.13）。",
                "2. 税额=金额×税率、价税合计=金额+税额，自动计算；已预置 100 行空白公式。",
                "3. 「汇总」页自动统计销项/进项与应纳税额，随台账实时联动。" });

            List<ReportSheet> sheets = new List<ReportSheet>();
            sheets.Add(sh); sheets.Add(sum); sheets.Add(help);
            return sheets;
        }

        static List<ReportCell> VatRow(int r, string date, string memo, string type, string amount, string rate)
        {
            List<ReportCell> row = new List<ReportCell>();
            row.Add(Cell(date, "n"));
            row.Add(Cell(memo, "n"));
            row.Add(Cell(type, "n"));
            row.Add(Cell(amount, "m"));
            row.Add(Cell(rate, "n"));
            row.Add(ReportCell.F("IF(D" + r + "=\"\",\"\",ROUND(D" + r + "*E" + r + ",2))", "m"));
            row.Add(ReportCell.F("IF(D" + r + "=\"\",\"\",ROUND(D" + r + "+F" + r + ",2))", "m"));
            return row;
        }

        static List<ReportCell> SumRow(string label, string formula)
        {
            List<ReportCell> row = new List<ReportCell>();
            row.Add(new ReportCell(label));
            row.Add(ReportCell.F(formula, "m"));
            return row;
        }

        // ---- 流水账：参数 / 流水 / 汇总 / 说明 ----
        static List<ReportSheet> BuildLedger(string rowsCsv, string paramsStr)
        {
            Dictionary<string, string> ps = ParseParams(paramsStr);
            double opening = ParamNum(ps, "openingbalance", 0);

            List<string[]> data = ParseCsvRows(rowsCsv);
            int n = data.Count;
            int last = 1 + n + LedgerBlankRows;

            ReportSheet cfg = HeaderSheet("参数", new string[] { "项目", "数值" }, new int[] { 16, 14 });
            cfg.Rows.Add(ParamRow("期初余额", opening, null));

            ReportSheet sh = HeaderSheet("流水",
                new string[] { "日期", "摘要", "类别", "收入", "支出", "余额" },
                new int[] { 12, 28, 12, 12, 12, 14 });
            for (int i = 0; i < n + LedgerBlankRows; i++)
            {
                string[] d = i < n ? data[i] : null;
                int r = i + 2;
                List<ReportCell> row = new List<ReportCell>();
                row.Add(Cell(d == null ? null : CellAt(d, 0), "n"));
                row.Add(Cell(d == null ? null : CellAt(d, 1), "n"));
                row.Add(Cell(d == null ? null : CellAt(d, 2), "n"));
                row.Add(Cell(d == null ? null : CellAt(d, 3), "m"));
                row.Add(Cell(d == null ? null : CellAt(d, 4), "m"));
                // 余额=期初+累计收入-累计支出（扩张区间 SUM：对空行、跳行填写都稳健）
                row.Add(ReportCell.F("IF(AND(D" + r + "=\"\",E" + r + "=\"\"),\"\",ROUND(参数!B2+SUM($D$2:D" + r + ")-SUM($E$2:E" + r + "),2))", "m"));
                sh.Rows.Add(row);
            }

            // 分类汇总：类别取自已给数据（之后新增的类别让 agent 补公式）
            List<string> cats = new List<string>();
            for (int i = 0; i < n; i++)
            {
                string c = (CellAt(data[i], 2) == null ? "" : (data[i][2] == null ? "" : data[i][2].Trim()));
                if (c.Length > 0 && !cats.Contains(c)) cats.Add(c);
            }
            ReportSheet sum = HeaderSheet("汇总", new string[] { "项目", "金额（元）" }, new int[] { 20, 14 });
            sum.Rows.Add(SumRow("收入合计", "SUM(流水!D2:D" + last + ")"));
            sum.Rows.Add(SumRow("支出合计", "SUM(流水!E2:E" + last + ")"));
            sum.Rows.Add(SumRow("期末结余", "ROUND(B2-B3+参数!B2,2)"));
            if (cats.Count > 0)
            {
                sum.AddSpacer();
                List<ReportCell> catHead = new List<ReportCell>();
                string[] heads = new string[] { "类别", "收入", "支出" };
                for (int i = 0; i < 3; i++)
                {
                    ReportCell c = new ReportCell(heads[i]);
                    c.Style = "h";
                    catHead.Add(c);
                }
                sum.Rows.Add(catHead);
                for (int gi = 0; gi < cats.Count; gi++)
                {
                    List<ReportCell> row = new List<ReportCell>();
                    row.Add(new ReportCell(cats[gi]));
                    row.Add(ReportCell.F("SUMIF(流水!C2:C" + last + ",A" + (sum.Rows.Count + 1) + ",流水!D2:D" + last + ")", "m"));
                    row.Add(ReportCell.F("SUMIF(流水!C2:C" + last + ",A" + (sum.Rows.Count + 1) + ",流水!E2:E" + last + ")", "m"));
                    sum.Rows.Add(row);
                }
            }

            ReportSheet help = TextSheet("说明", new string[] {
                "流水账使用说明",
                "1. 期初余额在「参数」页修改；「流水」页逐笔登记日期、摘要、类别、收入、支出。",
                "2. 余额=期初+累计收入-累计支出，自动计算（支持中间空行）；已预置 100 行空白公式。",
                "3. 「汇总」页自动统计收入/支出/期末结余与分类小计；登记时用了新类别，让 agent 补一行分类公式。" });

            List<ReportSheet> sheets = new List<ReportSheet>();
            sheets.Add(cfg); sheets.Add(sh); sheets.Add(sum); sheets.Add(help);
            return sheets;
        }

        // ============================================================
        // 自由模式：sheets JSON 规格 + 可选 summary
        // ============================================================
        static List<ReportSheet> BuildFree(List<object> sheetsSpec, string summaryJson, out string err)
        {
            err = null;
            if (sheetsSpec.Count == 0) { err = "sheets 为空"; return null; }
            List<ReportSheet> sheets = new List<ReportSheet>();
            for (int si = 0; si < sheetsSpec.Count; si++)
            {
                Dictionary<string, object> spec = sheetsSpec[si] as Dictionary<string, object>;
                if (spec == null) { err = "sheets[" + si + "] 应为对象"; return null; }
                string name = JsonVal.Str(spec, "name");
                if (name == null || name.Trim().Length == 0) { err = "sheets[" + si + "].name 缺失"; return null; }
                name = name.Trim();
                if (name.IndexOf('\'') >= 0 || name.Length > 31)
                {
                    err = "sheet 名不合法（不得含单引号，长度≤31）: " + name;
                    return null;
                }

                List<object> header = JsonVal.List(spec, "header");
                List<object> rows = JsonVal.List(spec, "rows");
                List<object> formulaCols = JsonVal.List(spec, "formulaCols");
                int blank = (int)JsonVal.Num(spec, "blankRows", 0);
                if (blank < 0) blank = 0;

                ReportSheet sh = new ReportSheet(name);
                int width = 0;
                if (header != null && header.Count > 0)
                {
                    List<ReportCell> row = new List<ReportCell>();
                    for (int c = 0; c < header.Count; c++)
                    {
                        ReportCell hc = Cell(header[c], "n");
                        if (hc == null) hc = new ReportCell("");
                        hc.Style = "h";
                        row.Add(hc);
                    }
                    sh.Rows.Add(row);
                    sh.FreezeRows = 1;
                    width = header.Count;
                }
                int n = rows != null ? rows.Count : 0;
                for (int ri = 0; ri < n; ri++)
                {
                    List<object> src = rows[ri] as List<object>;
                    if (src == null) { err = "sheets[" + si + "].rows[" + ri + "] 应为数组"; return null; }
                    List<ReportCell> row = new List<ReportCell>();
                    for (int c = 0; c < src.Count; c++) row.Add(Cell(src[c], "n"));
                    ApplyFormulaCols(row, formulaCols, ri + 2);
                    sh.Rows.Add(row);
                }
                for (int b = 0; b < blank; b++)
                {
                    List<ReportCell> row = new List<ReportCell>();
                    ApplyFormulaCols(row, formulaCols, n + b + 2);
                    sh.Rows.Add(row);
                }
                if (JsonVal.Bool(spec, "totalRow", false) && n + blank > 0)
                {
                    int last = 1 + n + blank;
                    List<ReportCell> total = new List<ReportCell>();
                    total.Add(new ReportCell("合计"));
                    for (int c = 1; c < Math.Max(width, 1); c++)
                    {
                        string col = ColName(c);
                        total.Add(ReportCell.F("SUM(" + col + "2:" + col + last + ")", "b"));
                    }
                    sh.Rows.Add(total);
                }
                ApplyWidths(sh, spec);
                sheets.Add(sh);
            }

            // 汇总 sheet（可选）：按分组列 SUMIF，改明细自动变
            if (summaryJson != null && summaryJson.Trim().Length > 0)
            {
                string sumErr;
                ReportSheet sm2 = BuildSummarySheet(sheets, summaryJson, out sumErr);
                if (sm2 == null) { err = sumErr; return null; }
                sheets.Add(sm2);
            }
            return sheets;
        }

        static ReportSheet BuildSummarySheet(List<ReportSheet> sheets, string summaryJson, out string err)
        {
            err = null;
            Dictionary<string, object> sm;
            try { sm = JsonVal.ParseObject(summaryJson); }
            catch (Exception ex) { err = "summary 不是合法 JSON: " + ex.Message; return null; }
            string srcName = JsonVal.Str(sm, "source");
            string groupCol = JsonVal.Str(sm, "groupCol");
            List<object> sumCols = JsonVal.List(sm, "sumCols");
            if (srcName == null || groupCol == null || sumCols == null || sumCols.Count == 0)
            {
                err = "summary 需要 source、groupCol、sumCols";
                return null;
            }
            srcName = srcName.Trim();
            ReportSheet srcSheet = null;
            for (int i = 0; i < sheets.Count; i++)
                if (sheets[i].Name == srcName) srcSheet = sheets[i];
            if (srcSheet == null) { err = "summary.source 指向的 sheet 不存在: " + srcName; return null; }
            int gc = ColIndex(groupCol);
            if (gc < 0) { err = "summary.groupCol 列号不合法: " + groupCol; return null; }
            // 分组值取自源 sheet 数据行（按出现顺序去重；分组列需是文本/数值字面量）
            List<string> groups = new List<string>();
            for (int ri = 1; ri < srcSheet.Rows.Count; ri++)
            {
                List<ReportCell> row = srcSheet.Rows[ri];
                if (gc >= row.Count || row[gc] == null) continue;
                string v = row[gc].IsFormula ? "" : (row[gc].Text == null ? "" : row[gc].Text.Trim());
                if (v.Length > 0 && !groups.Contains(v)) groups.Add(v);
            }
            int last = srcSheet.Rows.Count;

            ReportSheet sm2 = new ReportSheet("汇总");
            string label = JsonVal.Str(sm, "labelHeader");
            List<ReportCell> head = new List<ReportCell>();
            ReportCell h0 = new ReportCell(label == null ? "分类" : label);
            h0.Style = "h";
            head.Add(h0);
            for (int c = 0; c < sumCols.Count; c++)
            {
                Dictionary<string, object> sc = sumCols[c] as Dictionary<string, object>;
                if (sc == null || ColIndex(JsonVal.Str(sc, "col")) < 0)
                {
                    err = "summary.sumCols[" + c + "] 需要 col（字母列号）";
                    return null;
                }
                string lh = JsonVal.Str(sc, "header");
                ReportCell hc = new ReportCell(lh == null ? JsonVal.Str(sc, "col") : lh);
                hc.Style = "h";
                head.Add(hc);
            }
            sm2.Rows.Add(head);
            string srcRef = "'" + srcName + "'!";
            string gRef = srcRef + groupCol.ToUpperInvariant() + "2:" + groupCol.ToUpperInvariant() + last;
            for (int gi = 0; gi < groups.Count; gi++)
            {
                List<ReportCell> row = new List<ReportCell>();
                row.Add(new ReportCell(groups[gi]));
                for (int c = 0; c < sumCols.Count; c++)
                {
                    Dictionary<string, object> sc = (Dictionary<string, object>)sumCols[c];
                    string colUp = JsonVal.Str(sc, "col").ToUpperInvariant();
                    row.Add(ReportCell.F("SUMIF(" + gRef + ",A" + (gi + 2) + "," +
                        srcRef + colUp + "2:" + colUp + last + ")", "m"));
                }
                sm2.Rows.Add(row);
            }
            List<ReportCell> tot = new List<ReportCell>();
            tot.Add(new ReportCell("合计"));
            for (int c = 0; c < sumCols.Count; c++)
            {
                string col = ColName(c + 1);
                tot.Add(ReportCell.F("SUM(" + col + "2:" + col + (groups.Count + 1) + ")", "b"));
            }
            sm2.Rows.Add(tot);
            return sm2;
        }

        // 公式列规格 [{col:"F", formula:"=D{r}-E{r}"}]：{r} 替换为实际行号后落到指定列
        static void ApplyFormulaCols(List<ReportCell> row, List<object> formulaCols, int excelRow)
        {
            if (formulaCols == null) return;
            for (int fi = 0; fi < formulaCols.Count; fi++)
            {
                Dictionary<string, object> fc = formulaCols[fi] as Dictionary<string, object>;
                if (fc == null) continue;
                int idx = ColIndex(JsonVal.Str(fc, "col"));
                string formula = JsonVal.Str(fc, "formula");
                if (idx < 0 || formula == null) continue;
                while (row.Count <= idx) row.Add(null);
                string f = formula.Replace("{r}", excelRow.ToString(CultureInfo.InvariantCulture));
                if (f.Length > 0 && f[0] == '=') f = f.Substring(1);
                row[idx] = ReportCell.F(f, "m");
            }
        }

        static void ApplyWidths(ReportSheet sh, Dictionary<string, object> spec)
        {
            List<object> w = JsonVal.List(spec, "widths");
            if (w == null || w.Count == 0) return;
            int[] arr = new int[w.Count];
            for (int i = 0; i < w.Count; i++) arr[i] = w[i] is double ? (int)(double)w[i] : 8;
            sh.ColWidths = arr;
        }
    }
}
