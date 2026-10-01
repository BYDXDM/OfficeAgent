// InvoiceEngine —— 发票提取（设计方案 §6.4）：PDF 文本层直提 → 标准表 + 勾稽
//   提取字段：发票号码 / 开票日期 / 购买方·销售方名称与税号 / 金额 / 税额 / 价税合计
//   正则按行+全文匹配，覆盖数电票（20 位号码）与增值税电子票（8 位）；扫描件无文本层 → 进失败清单（OCR 后续版本）。
//   Verifier：金额 + 税额 = 价税合计（逐票勾稽，容差 0.01）。
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using OfficeAgent.Core;

namespace OfficeAgent.Host
{
    public class InvoiceRow
    {
        public string File = "";
        public int Page = 1;
        public string InvoiceNo = "";
        public string Date = "";
        public string Buyer = "";
        public string BuyerTaxNo = "";
        public string Seller = "";
        public string SellerTaxNo = "";
        public decimal Amount, Tax, Total;      // 金额 / 税额 / 价税合计
        public bool HasAmount, HasTax, HasTotal;
        public decimal Balance() { return Amount + Tax - Total; }
    }

    public class InvoiceFailure
    {
        public string File = "";
        public string Reason = "";
    }

    public class InvoiceResult
    {
        public string Err = null;
        public List<InvoiceRow> Rows = new List<InvoiceRow>();
        public List<InvoiceFailure> Failures = new List<InvoiceFailure>();
        public List<ReconCheck> Checks = new List<ReconCheck>();
        public string OutputPath = null;
        public long ElapsedMs;
    }

    public class InvoiceEngine
    {
        static readonly Regex ReNo = new Regex("发\\s*票\\s*号\\s*码\\s*[:：]?\\s*(\\d{8,20})");
        static readonly Regex ReDate = new Regex("开\\s*票\\s*日\\s*期\\s*[:：]?\\s*(\\d{4})\\s*年\\s*(\\d{1,2})\\s*月\\s*(\\d{1,2})\\s*日");
        static readonly Regex ReTotal = new Regex("[（(]\\s*小写\\s*[)）]\\s*[¥￥]?\\s*([0-9,]+\\.[0-9]{2})");
        static readonly Regex ReSum = new Regex("合\\s*计\\s*[¥￥]\\s*([0-9,]+\\.[0-9]{2})\\s*[¥￥]\\s*([0-9,]+\\.[0-9]{2})");
        static readonly Regex ReTaxNo = new Regex("(?:统一\\s*社\\s*会\\s*信\\s*用\\s*代\\s*码/纳\\s*税\\s*人\\s*识\\s*别\\s*号|纳\\s*税\\s*人\\s*识\\s*别\\s*号|统\\s*一\\s*社\\s*会\\s*信\\s*用\\s*代\\s*码)\\s*[:：]?\\s*([0-9A-Z]{15,20})");
        static readonly Regex ReBuyerName = new Regex("购\\s*买\\s*方[^名]{0,30}?名\\s*称\\s*[:：]?\\s*([^\\n:：]{4,60}?)(?:\\s{2,}|统一|纳税人|$)");
        static readonly Regex ReSellerName = new Regex("销\\s*售\\s*方[^名]{0,30}?名\\s*称\\s*[:：]?\\s*([^\\n:：]{4,60}?)(?:\\s{2,}|统一|纳税人|$)");

        public InvoiceResult Run(List<string> pdfFiles, string outPath)
        {
            InvoiceResult r = new InvoiceResult();
            System.Diagnostics.Stopwatch sw = System.Diagnostics.Stopwatch.StartNew();
            try
            {
                string root = EnvDetect.FindRoot();
                Pdfium.EnsureInit(root);

                foreach (string file in pdfFiles)
                {
                    InvoiceRow row;
                    string err;
                    if (!ExtractOne(file, out row, out err))
                    {
                        InvoiceFailure f = new InvoiceFailure();
                        f.File = file;
                        f.Reason = err;
                        r.Failures.Add(f);
                        continue;
                    }
                    r.Rows.Add(row);
                }

                // 勾稽：金额 + 税额 = 价税合计
                int checkedN = 0, badN = 0;
                foreach (InvoiceRow row in r.Rows)
                {
                    if (!(row.HasAmount && row.HasTax && row.HasTotal)) continue;
                    checkedN++;
                    if (Math.Abs(row.Balance()) > 0.01m) badN++;
                }
                ReconCheck c1 = new ReconCheck();
                c1.Name = "单票勾稽（金额 + 税额 = 价税合计）";
                c1.Pass = badN == 0;
                c1.Detail = checkedN == 0 ? "无同时含三个金额的票可勾稽"
                    : ("可勾稽 " + checkedN + " 票，不平衡 " + badN + " 票");
                r.Checks.Add(c1);

                ReconCheck c2 = new ReconCheck();
                c2.Name = "提取成功率";
                c2.Warn = true;
                c2.Pass = true;
                c2.Detail = "成功 " + r.Rows.Count + " 票；失败 " + r.Failures.Count + " 票（失败原因见「未识别清单」页）";
                r.Checks.Add(c2);

                string werr = WriteWorkbook(r, outPath);
                if (werr != null) { r.Err = "写出发票清单失败: " + werr; return r; }
                r.OutputPath = outPath;
                return r;
            }
            catch (Exception ex)
            {
                r.Err = "发票提取异常: " + ex.Message;
                return r;
            }
            finally { sw.Stop(); r.ElapsedMs = sw.ElapsedMilliseconds; }
        }

        bool ExtractOne(string file, out InvoiceRow row, out string err)
        {
            row = null;
            err = null;
            try
            {
                string werr;
                using (Pdfium.PdfDoc doc = Pdfium.PdfDoc.Open(file, out werr))
                {
                    if (doc == null) { err = "打开失败: " + werr; return false; }
                    StringBuilder text = new StringBuilder();
                    int foundPage = -1;
                    for (int p = 0; p < doc.PageCount; p++)
                    {
                        List<string> lines = doc.ExtractLines(p, out werr);
                        if (lines == null) { err = werr; return false; }
                        foreach (string ln in lines) text.Append(ln).Append('\n');
                        foundPage = 0;
                        if (p == 0 && text.Length > 100) break;   // 发票信息在第一页（批量拼页的票逐页处理时由调用方拆分）
                    }
                    string full = text.ToString();
                    if (full.Trim().Length < 20)
                    {
                        err = "无文本层（可能为扫描件）；当前版本不做 OCR";
                        return false;
                    }
                    row = Parse(file, foundPage + 1, full);
                    if (row.InvoiceNo.Length == 0 && !row.HasTotal)
                    {
                        err = "未识别出发票要素（号码与价税合计均缺失；可能不是增值税/数电发票）";
                        row = null;
                        return false;
                    }
                    return true;
                }
            }
            catch (Exception ex)
            {
                err = "提取异常: " + ex.Message;
                return false;
            }
        }

        InvoiceRow Parse(string file, int page, string text)
        {
            InvoiceRow row = new InvoiceRow();
            row.File = file;
            row.Page = page;
            text = MergeSpacedAscii(text);   // 行内数字/字母伪空格合并（老字宽/字体伪影鲁棒性）
            Match m = ReNo.Match(text);
            if (m.Success) row.InvoiceNo = m.Groups[1].Value;
            m = ReDate.Match(text);
            if (m.Success)
                row.Date = m.Groups[1].Value + "-" + Pad2(m.Groups[2].Value) + "-" + Pad2(m.Groups[3].Value);
            m = ReTotal.Match(text);
            if (m.Success) { row.Total = ParseMoney(m.Groups[1].Value); row.HasTotal = true; }
            m = ReSum.Match(text);
            if (m.Success)
            {
                row.Amount = ParseMoney(m.Groups[1].Value);
                row.Tax = ParseMoney(m.Groups[2].Value);
                row.HasAmount = true;
                row.HasTax = true;
            }
            // 税号：第一处=购买方，第二处=销售方（数电票/电子票版式约定）
            MatchCollection taxes = ReTaxNo.Matches(text);
            if (taxes.Count > 0) row.BuyerTaxNo = taxes[0].Groups[1].Value;
            if (taxes.Count > 1) row.SellerTaxNo = taxes[1].Groups[1].Value;
            m = ReBuyerName.Match(text);
            if (m.Success) row.Buyer = CleanName(m.Groups[1].Value);
            if (row.Buyer.Length == 0 && taxes.Count > 0)
            {
                // 版式差异兜底：税号前的一段长文本通常为公司名
                string guess = GuessNameNear(text, row.BuyerTaxNo);
                if (guess != null) row.Buyer = guess;
            }
            m = ReSellerName.Match(text);
            if (m.Success) row.Seller = CleanName(m.Groups[1].Value);
            if (row.Seller.Length == 0 && taxes.Count > 1)
            {
                string guess = GuessNameNear(text, row.SellerTaxNo);
                if (guess != null) row.Seller = guess;
            }
            return row;
        }

        // 行内归一化：把「2 4 5 1」这类被伪空格拆开的数字/字母串接回连续 token。
        // 只在行内做（不跨行），只合并两侧均为 ASCII 数字/字母/小数点/百分号的空格。
        static string MergeSpacedAscii(string text)
        {
            string[] lines = text.Split('\n');
            StringBuilder sb = new StringBuilder(text.Length);
            foreach (string ln in lines)
            {
                sb.Append(Regex.Replace(ln, "(?<=[0-9A-Za-z.%])\\s+(?=[0-9A-Za-z.%])", "")).Append('\n');
            }
            return sb.ToString();
        }

        string GuessNameNear(string text, string taxNo)
        {
            if (taxNo == null || taxNo.Length == 0) return null;
            int idx = text.IndexOf(taxNo, StringComparison.Ordinal);
            if (idx < 0) return null;
            int start = Math.Max(0, idx - 100);
            string seg = text.Substring(start, idx - start);
            // 候选：段内最后一个「名 称：X」（版式常见：名称在前，税号标签紧随其后）
            MatchCollection mc = Regex.Matches(seg, "名\\s*称\\s*[:：]\\s*([^\\n:：]{4,60})");
            for (int i = mc.Count - 1; i >= 0; i--)
            {
                string name = mc[i].Groups[1].Value;
                int cut = name.IndexOf("统一社会信用", StringComparison.Ordinal);
                int cut2 = name.IndexOf("纳税人识别号", StringComparison.Ordinal);
                if (cut < 0 || (cut2 >= 0 && cut2 < cut)) cut = cut2;
                if (cut > 0) name = name.Substring(0, cut);
                name = CleanName(name);
                if (name.Length >= 4) return name;
            }
            return null;
        }

        static string CleanName(string s)
        {
            string t = s.Trim();
            t = t.TrimEnd('：', ':', '，', ',', ' ');
            return t;
        }

        static string Pad2(string s)
        {
            return s.Length == 1 ? "0" + s : s;
        }

        static decimal ParseMoney(string s)
        {
            decimal v;
            string t = s.Replace(",", "");
            return decimal.TryParse(t, NumberStyles.Number, CultureInfo.InvariantCulture, out v) ? v : 0;
        }

        string WriteWorkbook(InvoiceResult r, string outPath)
        {
            ReportSheet list = new ReportSheet("发票清单");
            list.FreezeRows = 1;
            list.ColWidths = new int[] { 30, 22, 12, 24, 22, 20, 20, 12, 12, 12, 10 };
            list.AddRow(new ReportCell[] {
                new ReportCell("文件") { Style = "h" }, new ReportCell("发票号码") { Style = "h" },
                new ReportCell("开票日期") { Style = "h" }, new ReportCell("购买方") { Style = "h" },
                new ReportCell("销售方") { Style = "h" }, new ReportCell("购买方税号") { Style = "h" },
                new ReportCell("销售方税号") { Style = "h" }, new ReportCell("金额") { Style = "h" },
                new ReportCell("税额") { Style = "h" }, new ReportCell("价税合计") { Style = "h" },
                new ReportCell("页") { Style = "h" } });
            decimal sumAmount = 0, sumTax = 0, sumTotal = 0;
            foreach (InvoiceRow w in r.Rows)
            {
                list.AddRow(new ReportCell[] {
                    new ReportCell(Path.GetFileName(w.File)),
                    new ReportCell(w.InvoiceNo),
                    new ReportCell(w.Date),
                    new ReportCell(w.Buyer),
                    new ReportCell(w.Seller),
                    new ReportCell(w.BuyerTaxNo),
                    new ReportCell(w.SellerTaxNo),
                    w.HasAmount ? (ReportCell)ReportCell.N((double)w.Amount, "m") : (ReportCell)new ReportCell("-") { Style = "p" },
                    w.HasTax ? (ReportCell)ReportCell.N((double)w.Tax, "m") : (ReportCell)new ReportCell("-") { Style = "p" },
                    w.HasTotal ? (ReportCell)ReportCell.N((double)w.Total, "m") : (ReportCell)new ReportCell("-") { Style = "p" },
                    ReportCell.N(w.Page, "m") });
                if (w.HasAmount) sumAmount += w.Amount;
                if (w.HasTax) sumTax += w.Tax;
                if (w.HasTotal) sumTotal += w.Total;
            }
            if (r.Rows.Count > 0)
            {
                list.AddRow(new ReportCell[] {
                    new ReportCell("合计") { Style = "b" }, null, null, null, null, null, null,
                    (ReportCell)ReportCell.N((double)sumAmount, "b"),
                    (ReportCell)ReportCell.N((double)sumTax, "b"),
                    (ReportCell)ReportCell.N((double)sumTotal, "b"), null });
            }

            ReportSheet check = new ReportSheet("校验与说明");
            check.ColWidths = new int[] { 42, 14, 60 };
            check.AddRow(new ReportCell[] { new ReportCell("发票提取校验") { Style = "t" } });
            check.AddRow(new ReportCell[] { new ReportCell("生成时间") { Style = "b" },
                new ReportCell(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")) { Style = "p" }, null });
            check.AddRow(new ReportCell[] { new ReportCell("清单合计") { Style = "b" },
                new ReportCell("金额 " + sumAmount.ToString("N2", CultureInfo.InvariantCulture) +
                    "；税额 " + sumTax.ToString("N2", CultureInfo.InvariantCulture) +
                    "；价税合计 " + sumTotal.ToString("N2", CultureInfo.InvariantCulture)) { Style = "p" }, null });
            foreach (ReconCheck c in r.Checks)
            {
                string verdict = c.Warn ? "提示" : (c.Pass ? "✓ 通过" : "✗ 失败");
                check.AddRow(new ReportCell[] { new ReportCell(c.Name),
                    new ReportCell(verdict) { Style = c.Warn ? "o" : (c.Pass ? "g" : "r") },
                    new ReportCell(c.Detail) });
            }

            ReportSheet fails = new ReportSheet("未识别清单");
            fails.FreezeRows = 1;
            fails.ColWidths = new int[] { 50, 70 };
            fails.AddRow(new ReportCell[] { new ReportCell("文件") { Style = "h" }, new ReportCell("原因") { Style = "h" } });
            foreach (InvoiceFailure f in r.Failures)
                fails.AddRow(new ReportCell[] { new ReportCell(f.File) { Style = "o" }, new ReportCell(f.Reason) { Style = "o" } });
            if (r.Failures.Count == 0)
                fails.AddRow(new ReportCell[] { new ReportCell("无") { Style = "g" } });

            List<ReportSheet> sheets = new List<ReportSheet>();
            sheets.Add(list);
            sheets.Add(check);
            sheets.Add(fails);
            return MiniXlsxWrite.Save(outPath, sheets);
        }
    }
}
