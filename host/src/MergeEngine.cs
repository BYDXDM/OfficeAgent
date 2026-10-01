// MergeEngine —— 报表汇总（设计方案 §6.2）：多工作簿按表头名映射抽取 → 合并底稿 + 勾稽校验页
// 场景：月结 20 家子公司报表归集——各文件列序可不同，按表头名匹配进统一底稿；
// 金额列自动合计，校验页逐文件小计 vs 底稿合计勾稽。原始文件只读；.xls 经转换总线预转。
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace OfficeAgent.Host
{
    public class MergeFileResult
    {
        public string File = "";
        public string Title = "";
        public int Rows;
        public bool Ok;
        public string Err = "";
        public Dictionary<string, decimal> Subtotals = new Dictionary<string, decimal>();
    }

    public class MergeResult
    {
        public string Err = null;
        public int MergedRows;
        public List<MergeFileResult> Files = new List<MergeFileResult>();
        public List<ReconCheck> Checks = new List<ReconCheck>();
        public string OutputPath = null;
        public long ElapsedMs;
        public int FailedChecks()
        {
            int n = 0;
            foreach (ReconCheck c in Checks) { if (!c.Pass && !c.Warn) n++; }
            return n;
        }
    }

    public class MergeEngine
    {
        public MergeResult Run(MergeTemplate tpl, List<string> files, string outPath, ConvertEngine conv)
        {
            MergeResult r = new MergeResult();
            System.Diagnostics.Stopwatch sw = System.Diagnostics.Stopwatch.StartNew();
            try
            {
                if (tpl.TargetCols.Length == 0 || tpl.SrcCols.Length != tpl.TargetCols.Length)
                {
                    r.Err = "模板列映射不完整（目标列与源列一一对应）";
                    return r;
                }
                List<string[]> merged = new List<string[]>();
                List<bool[]> numeric = new List<bool[]>();    // 与 merged 平行：每格是否数值

                foreach (string file in files)
                {
                    MergeFileResult fr = LoadAndExtract(tpl, file, conv, merged, numeric, r);
                    r.Files.Add(fr);
                }

                // 金额列合计
                Dictionary<string, decimal> totals = new Dictionary<string, decimal>();
                foreach (string col in tpl.AmountCols) totals[col] = 0;
                for (int i = 0; i < merged.Count; i++)
                {
                    for (int j = 0; j < tpl.TargetCols.Length; j++)
                    {
                        if (!IsAmountCol(tpl, tpl.TargetCols[j])) continue;
                        decimal v;
                        if (numeric[i][j] && decimal.TryParse(merged[i][j], NumberStyles.Number,
                            CultureInfo.InvariantCulture, out v))
                            totals[tpl.TargetCols[j]] += v;
                    }
                }
                r.MergedRows = merged.Count;

                // 校验 1：行数守恒
                int fileRows = 0;
                foreach (MergeFileResult fr in r.Files) { if (fr.Ok) fileRows += fr.Rows; }
                ReconCheck c1 = new ReconCheck();
                c1.Name = "行数守恒";
                c1.Pass = fileRows == merged.Count;
                c1.Detail = "各文件合计 " + fileRows + " 行；底稿 " + merged.Count + " 行";
                r.Checks.Add(c1);

                // 校验 2：金额勾稽（逐文件小计之和 = 底稿合计）
                foreach (string col in tpl.AmountCols)
                {
                    decimal sumFiles = 0;
                    foreach (MergeFileResult fr in r.Files) { if (fr.Ok && fr.Subtotals.ContainsKey(col)) sumFiles += fr.Subtotals[col]; }
                    ReconCheck c = new ReconCheck();
                    c.Name = "合计勾稽：「" + col + "」";
                    decimal tot = totals.ContainsKey(col) ? totals[col] : 0;
                    c.Pass = sumFiles == tot;
                    c.Detail = "各文件小计和 " + sumFiles.ToString("N2", CultureInfo.InvariantCulture) +
                        "；底稿合计 " + tot.ToString("N2", CultureInfo.InvariantCulture);
                    r.Checks.Add(c);
                }

                // 校验 3：文件装载错误提示（警告）
                ReconCheck c3 = new ReconCheck();
                c3.Name = "文件装载";
                c3.Warn = true;
                c3.Pass = true;
                int bad = 0;
                foreach (MergeFileResult fr in r.Files) { if (!fr.Ok) bad++; }
                c3.Detail = bad == 0 ? ("全部 " + r.Files.Count + " 个文件装载成功")
                    : (bad + " 个文件装载失败并跳过（明细见校验页）");
                r.Checks.Add(c3);

                string werr = WriteWorkbook(r, tpl, merged, numeric, totals, outPath);
                if (werr != null) { r.Err = "写出汇总底稿失败: " + werr; return r; }
                r.OutputPath = outPath;
                return r;
            }
            catch (Exception ex)
            {
                r.Err = "汇总异常: " + ex.Message;
                return r;
            }
            finally { sw.Stop(); r.ElapsedMs = sw.ElapsedMilliseconds; }
        }

        static bool IsAmountCol(MergeTemplate tpl, string targetCol)
        {
            foreach (string a in tpl.AmountCols) { if (a == targetCol) return true; }
            return false;
        }

        // ---------- 装载与抽取 ----------

        MergeFileResult LoadAndExtract(MergeTemplate tpl, string file, ConvertEngine conv,
            List<string[]> merged, List<bool[]> numeric, MergeResult r)
        {
            MergeFileResult fr = new MergeFileResult();
            fr.File = file;
            try
            {
                string workPath = file;
                string ext = (Path.GetExtension(file) ?? "").ToLowerInvariant();
                if (ext == ".xls")
                {
                    if (conv == null) { fr.Err = ".xls 需要转换引擎（COM/LibreOffice）预转 xlsx"; return fr; }
                    string outPath;
                    string err = conv.Convert(file, ConvTarget.Xlsx, out outPath);
                    if (err != null) { fr.Err = ".xls 预转失败: " + err; return fr; }
                    workPath = outPath;
                    ext = ".xlsx";
                }
                if (ext != ".xlsx" && ext != ".csv") { fr.Err = "不支持的格式: " + ext; return fr; }

                string[] header = null;
                List<string[]> rows = new List<string[]>();
                if (ext == ".xlsx")
                {
                    int sheetIdx;
                    using (XlsxBook book = XlsxBook.Open(workPath))
                    {
                        if (!ResolveSheet(book, tpl.SheetSpec, out sheetIdx)) { fr.Err = "找不到工作表: " + tpl.SheetSpec; return fr; }
                        fr.Title = book.Sheets[sheetIdx].Name;
                        List<string[]> all = new List<string[]>();
                        List<int> rowNums = new List<int>();
                        string err = book.StreamRows(sheetIdx, 64, delegate(string[] cells, int excelRow)
                        {
                            rowNums.Add(excelRow);
                            all.Add(Trim(cells));
                        });
                        if (err != null) { fr.Err = "读取失败: " + err; return fr; }
                        XlsxBook.FillMerged(all.ToArray(), book.Sheets[sheetIdx].Merges);
                        int hr = Math.Max(1, tpl.HeaderRow);
                        bool gotHeader = false;
                        for (int i = 0; i < rowNums.Count; i++)
                        {
                            if (!gotHeader && rowNums[i] == hr) { header = all[i]; gotHeader = true; continue; }
                            if (gotHeader && rowNums[i] > hr) rows.Add(all[i]);
                        }
                        if (!gotHeader) { fr.Err = "表头行 " + tpl.HeaderRow + " 不存在"; return fr; }
                    }
                }
                else
                {
                    Encoding used;
                    List<string[]> all = MiniCsv.Parse(MiniCsv.DetectRead(workPath, out used));
                    if (all.Count < tpl.HeaderRow) { fr.Err = "表头行 " + tpl.HeaderRow + " 不存在"; return fr; }
                    fr.Title = Path.GetFileNameWithoutExtension(file);
                    header = all[tpl.HeaderRow - 1];
                    for (int i = tpl.HeaderRow; i < all.Count; i++) rows.Add(all[i]);
                }

                // 按表头名解析源列 → 目标列位置
                int[] srcIdx = new int[tpl.TargetCols.Length];
                if (header == null) { fr.Err = "表头缺失"; return fr; }
                for (int j = 0; j < tpl.TargetCols.Length; j++)
                {
                    int c; string err;
                    if (!ResolveCol(tpl.SrcCols[j], header, out c, out err)) { fr.Err = err; return fr; }
                    srcIdx[j] = c;
                }

                foreach (string[] row in rows)
                {
                    string[] outRow = new string[tpl.TargetCols.Length];
                    bool[] num = new bool[tpl.TargetCols.Length];
                    for (int j = 0; j < tpl.TargetCols.Length; j++)
                    {
                        int c = srcIdx[j];
                        string v = c < row.Length && row[c] != null ? row[c] : "";
                        if (v.Length > 0 && v[0] == '=') v = "";      // 公式无缓存值：空处理
                        outRow[j] = v;
                        decimal d;
                        if (IsAmountCol(tpl, tpl.TargetCols[j]) && v.Length > 0 &&
                            ReconEngine.TryAmount(v, out d))
                        {
                            outRow[j] = d.ToString("0.############################", CultureInfo.InvariantCulture);
                            num[j] = true;
                        }
                    }
                    merged.Add(outRow);
                    numeric.Add(num);
                    fr.Rows++;
                }
                fr.Ok = true;

                // 逐文件小计（金额列）
                foreach (string col in tpl.AmountCols) fr.Subtotals[col] = 0;
                int baseIdx = merged.Count - fr.Rows;
                for (int i = baseIdx; i < merged.Count; i++)
                {
                    for (int j = 0; j < tpl.TargetCols.Length; j++)
                    {
                        if (!IsAmountCol(tpl, tpl.TargetCols[j]) || !numeric[i][j]) continue;
                        decimal v;
                        if (decimal.TryParse(merged[i][j], NumberStyles.Number, CultureInfo.InvariantCulture, out v))
                            fr.Subtotals[tpl.TargetCols[j]] += v;
                    }
                }
                return fr;
            }
            catch (Exception ex)
            {
                fr.Err = "装载异常: " + ex.Message;
                return fr;
            }
        }

        static bool ResolveSheet(XlsxBook book, string spec, out int idx)
        {
            idx = 0;
            string s = (spec ?? "").Trim();
            if (s.Length == 0) return book.Sheets.Count > 0;
            int n;
            if (int.TryParse(s, out n))
            {
                if (n >= 1 && n <= book.Sheets.Count) { idx = n - 1; return true; }
                return false;
            }
            for (int i = 0; i < book.Sheets.Count; i++)
            {
                if (string.Equals(book.Sheets[i].Name.Trim(), s, StringComparison.Ordinal)) { idx = i; return true; }
            }
            return false;
        }

        static bool ResolveCol(string token, string[] header, out int col, out string err)
        {
            col = -1; err = null;
            string t = (token ?? "").Trim();
            if (t.Length == 0) { err = "模板源列为空"; return false; }
            int n;
            if (int.TryParse(t, out n))
            {
                if (n >= 1 && n <= 64) { col = n - 1; return true; }
                err = "列序号越界: " + t; return false;
            }
            string want = ReconEngine.NormHeader(t);
            for (int i = 0; i < header.Length; i++)
            {
                if (ReconEngine.NormHeader(header[i]) == want) { col = i; return true; }
            }
            err = "找不到列「" + t + "」";
            return false;
        }

        static string[] Trim(string[] cells)
        {
            int last = -1;
            for (int i = cells.Length - 1; i >= 0; i--) { if (cells[i] != null) { last = i; break; } }
            if (last < 0) return new string[0];
            string[] copy = new string[last + 1];
            Array.Copy(cells, copy, last + 1);
            return copy;
        }

        // ---------- 输出 ----------

        string WriteWorkbook(MergeResult r, MergeTemplate tpl, List<string[]> merged, List<bool[]> numeric,
            Dictionary<string, decimal> totals, string outPath)
        {
            ReportSheet draft = new ReportSheet("合并底稿");
            draft.FreezeRows = 1;
            int[] widths = new int[tpl.TargetCols.Length];
            for (int j = 0; j < widths.Length; j++) widths[j] = IsAmountCol(tpl, tpl.TargetCols[j]) ? 14 : 18;
            draft.ColWidths = widths;
            ReportCell[] head = new ReportCell[tpl.TargetCols.Length];
            for (int j = 0; j < head.Length; j++) head[j] = new ReportCell(tpl.TargetCols[j]) { Style = "h" };
            draft.AddRow(head);
            for (int i = 0; i < merged.Count; i++)
            {
                ReportCell[] row = new ReportCell[tpl.TargetCols.Length];
                for (int j = 0; j < row.Length; j++)
                {
                    bool amt = IsAmountCol(tpl, tpl.TargetCols[j]);
                    if (amt && numeric[i][j])
                        row[j] = ReportCell.N(double.Parse(merged[i][j], CultureInfo.InvariantCulture), "m");
                    else
                        row[j] = new ReportCell(merged[i][j]);
                }
                draft.AddRow(row);
            }
            if (merged.Count > 0 && tpl.AmountCols.Length > 0)
            {
                ReportCell[] total = new ReportCell[tpl.TargetCols.Length];
                total[0] = new ReportCell("合计") { Style = "b" };
                for (int j = 1; j < total.Length; j++) total[j] = new ReportCell("");
                for (int j = 0; j < tpl.TargetCols.Length; j++)
                {
                    if (!IsAmountCol(tpl, tpl.TargetCols[j])) continue;
                    string colL = ColLetters(j);
                    total[j] = ReportCell.F("SUM(" + colL + "2:" + colL + (merged.Count + 1) + ")", "b");
                }
                draft.AddRow(total);
            }

            ReportSheet check = new ReportSheet("校验页");
            check.ColWidths = new int[] { 40, 16, 40 };
            check.AddRow(new ReportCell[] { new ReportCell("报表汇总校验") { Style = "t" } });
            check.AddRow(new ReportCell[] { new ReportCell("生成时间") { Style = "b" },
                new ReportCell(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")) { Style = "p" }, null });
            check.AddRow(new ReportCell[] { new ReportCell("源文件") { Style = "h" }, new ReportCell("数据行") { Style = "h" },
                new ReportCell("金额列小计") { Style = "h" } });
            foreach (MergeFileResult fr in r.Files)
            {
                if (!fr.Ok)
                {
                    check.AddRow(new ReportCell[] { new ReportCell(fr.File) { Style = "r" },
                        new ReportCell("-") { Style = "p" }, new ReportCell(fr.Err) { Style = "r" } });
                    continue;
                }
                StringBuilder st = new StringBuilder();
                foreach (KeyValuePair<string, decimal> kv in fr.Subtotals)
                {
                    if (st.Length > 0) st.Append("；");
                    st.Append(kv.Key).Append("=").Append(kv.Value.ToString("N2", CultureInfo.InvariantCulture));
                }
                check.AddRow(new ReportCell[] { new ReportCell(fr.File), ReportCell.N(fr.Rows, "m"),
                    new ReportCell(st.ToString()) { Style = "p" } });
            }
            check.AddSpacer();
            check.AddRow(new ReportCell[] { new ReportCell("勾稽校验") { Style = "h" }, new ReportCell("结果") { Style = "h" },
                new ReportCell("明细") { Style = "h" } });
            foreach (ReconCheck c in r.Checks)
            {
                string verdict = c.Warn ? "提示" : (c.Pass ? "✓ 通过" : "✗ 失败");
                check.AddRow(new ReportCell[] { new ReportCell(c.Name),
                    new ReportCell(verdict) { Style = c.Warn ? "o" : (c.Pass ? "g" : "r") },
                    new ReportCell(c.Detail) });
            }

            List<ReportSheet> sheets = new List<ReportSheet>();
            sheets.Add(draft);
            sheets.Add(check);
            return MiniXlsxWrite.Save(outPath, sheets);
        }

        static string ColLetters(int idx)
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
    }
}
