// ReconEngine —— 两表核对（设计方案 §6.1 旗舰场景）
//   归一化键匹配（去空白/全半角/大写）→ 同键 FIFO 配对 → 容差内判匹配 → 差异分类
//   （仅A有 / 仅B有 / 金额不等）→ Verifier 勾稽自证 → 差异表 Excel（MiniXlsxWrite 标红）。
// 规则：原始文件只读；金额一律 decimal；解析失败的金额行剔除并列错误清单，绝不静默当 0。
// 模板记忆：映射配置存 %LOCALAPPDATA%\OfficeAgent\recon-templates.json，"把上月对账再来一次"直接复用。
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace OfficeAgent.Host
{
    public class ReconMapping
    {
        public string SheetA = "0";          // sheet 序号（1 起）或名称
        public string SheetB = "0";
        public int HeaderRowA = 1;           // 1-based
        public int HeaderRowB = 1;
        public string[] KeyColsA = new string[0];   // 列名或 1-based 序号（字符串，模板友好）
        public string[] KeyColsB = new string[0];
        public string DebitA = "";           // 借方/收入列；"" = 无
        public string CreditA = "";          // 贷方/支出列；单列签名口径时留空
        public string DebitB = "";
        public string CreditB = "";
        public double Tolerance = 0.01;
        public bool SqueezeKey = true;       // 键归一化：去空白+全半角+大写
        public bool FillMergedCells = true;  // 合并单元格回填
        public bool IncludeMatched = false;  // 差异表附"全部匹配明细"页
    }

    public class ReconRow
    {
        public int ExcelRow;          // 1-based 原始行号（引用单元格用）
        public string KeyRaw = "";
        public string KeyNorm = "";
        public decimal Amount;
        public bool HasAmount;        // 金额列全空 → 0（HasAmount=false 仍按 0 参与）
    }

    public class ReconParseError
    {
        public string Side = "";      // "A" / "B"
        public int ExcelRow;
        public string ColName = "";
        public string Raw = "";
        public string Reason = "";
    }

    public class ReconDiffRow
    {
        public string Type = "";      // "金额不等" / "仅A有" / "仅B有"
        public string Key = "";       // 归一化键
        public decimal AmountA, AmountB, Delta;
        public int RowA = -1, RowB = -1;
        public string RefA = "-", RefB = "-";
    }

    public class ReconCheck
    {
        public string Name = "";
        public bool Pass;
        public bool Warn;             // 警告级（不计入失败）
        public string Detail = "";
    }

    public class ReconResult
    {
        public string Err = null;
        public int TotalA, TotalB;
        public int MatchedCount, PairDiffCount, OnlyACount, OnlyBCount;
        public decimal SumA, SumB;
        public List<ReconDiffRow> Diffs = new List<ReconDiffRow>();
        public List<ReconDiffRow> Matched = new List<ReconDiffRow>();   // 匹配明细（IncludeMatched 时输出）
        public List<ReconCheck> Checks = new List<ReconCheck>();
        public List<ReconParseError> ParseErrors = new List<ReconParseError>();
        public string OutputPath = null;
        public long ElapsedMs;
        public bool UsedXlsConversion;     // .xls 经转换总线预转 xlsx
        public int FailedChecks()
        {
            int n = 0;
            foreach (ReconCheck c in Checks) { if (!c.Pass && !c.Warn) n++; }
            return n;
        }
    }

    public class ReconEngine
    {
        // ================= 对外入口 =================

        public ReconResult Run(ReconMapping map, string pathA, string pathB, string outPath, ConvertEngine conv)
        {
            ReconResult r = new ReconResult();
            System.Diagnostics.Stopwatch sw = System.Diagnostics.Stopwatch.StartNew();
            try
            {
                Table ta = LoadSide(pathA, map.SheetA, map.HeaderRowA, map.FillMergedCells, "A", conv, r);
                if (ta == null) { r.Err = LastErr; return r; }
                Table tb = LoadSide(pathB, map.SheetB, map.HeaderRowB, map.FillMergedCells, "B", conv, r);
                if (tb == null) { r.Err = LastErr; return r; }

                Side sa;
                if (!ResolveSide(map.KeyColsA, map.DebitA, map.CreditA, ta.Header, out sa)) { r.Err = LastErr; return r; }
                Side sb;
                if (!ResolveSide(map.KeyColsB, map.DebitB, map.CreditB, tb.Header, out sb)) { r.Err = LastErr; return r; }
                if (sa.Debit < 0 && sa.Credit < 0) { r.Err = "A 侧未指定金额列"; return r; }
                if (sb.Debit < 0 && sb.Credit < 0) { r.Err = "B 侧未指定金额列"; return r; }
                if (sa.KeyCols.Length == 0 || sb.KeyCols.Length == 0) { r.Err = "未指定键列"; return r; }

                List<ReconRow> rowsA = ExtractRows(ta, sa, "A", map.SqueezeKey, r);
                List<ReconRow> rowsB = ExtractRows(tb, sb, "B", map.SqueezeKey, r);
                r.TotalA = rowsA.Count;
                r.TotalB = rowsB.Count;

                Match(rowsA, rowsB, map, sa, sb, ta, tb, r);
                Verify(rowsA, rowsB, map, sa, sb, ta, tb, r);
                string werr = WriteReport(r, map, pathA, pathB, ta, tb, outPath);
                if (werr != null) { r.Err = "写出差异表失败: " + werr; return r; }
                r.OutputPath = outPath;
                return r;
            }
            catch (Exception ex)
            {
                r.Err = "核对异常: " + ex.Message;
                return r;
            }
            finally { sw.Stop(); r.ElapsedMs = sw.ElapsedMilliseconds; }
        }

        string LastErr = "";

        // ================= 表格装载 =================

        class Table
        {
            public string FilePath = "";
            public string Title = "";        // 引用前缀：xlsx 用 sheet 名，csv 用文件名
            public string[] Header = new string[0];
            public int HeaderRow = 1;
            public Dictionary<int, string[]> ByRow = new Dictionary<int, string[]>();  // 1-based 行号 → 单元格
            public List<int> DataRows = new List<int>();                               // 数据行号（>表头行）
        }

        // .xls → 转换总线预转 xlsx（原始文件只读，输出 <名>_conv.xlsx）
        string PreConvertIfXls(string path, ConvertEngine conv, ReconResult r)
        {
            string ext = (Path.GetExtension(path) ?? "").ToLowerInvariant();
            if (ext != ".xls") return path;
            if (conv == null) { LastErr = ".xls 需要转换引擎（COM/LibreOffice）预转 xlsx"; return null; }
            string outPath;
            string err = conv.Convert(path, ConvTarget.Xlsx, out outPath);
            if (err != null) { LastErr = ".xls 预转 xlsx 失败: " + err; return null; }
            r.UsedXlsConversion = true;
            return outPath;
        }

        Table LoadSide(string path, string sheetSpec, int headerRow, bool fillMerged, string side,
            ConvertEngine conv, ReconResult r)
        {
            LastErr = "";
            if (!File.Exists(path)) { LastErr = side + " 侧文件不存在: " + path; return null; }
            string workPath = PreConvertIfXls(path, conv, r);
            if (workPath == null) return null;
            if (headerRow < 1) headerRow = 1;
            string ext = (Path.GetExtension(workPath) ?? "").ToLowerInvariant();

            Table t = new Table();
            t.FilePath = path;
            t.HeaderRow = headerRow;

            if (ext == ".xlsx")
            {
                int sheetIdx;
                using (XlsxBook book = XlsxBook.Open(workPath))
                {
                    if (!ResolveSheet(book, sheetSpec, out sheetIdx)) { LastErr = LastErr.Length == 0 ? side + " 侧找不到工作表: " + sheetSpec : LastErr; return null; }
                    t.Title = book.Sheets[sheetIdx].Name;
                    List<string[]> all = new List<string[]>();
                    List<int> rowNums = new List<int>();
                    string err = book.StreamRows(sheetIdx, 64, delegate(string[] cells, int excelRow)
                    {
                        rowNums.Add(excelRow);
                        all.Add(TrimRow(cells));
                    });
                    if (err != null) { LastErr = side + " 侧读取 xlsx 失败: " + err; return null; }
                    if (fillMerged) XlsxBook.FillMerged(all.ToArray(), book.Sheets[sheetIdx].Merges);
                    for (int i = 0; i < rowNums.Count; i++) t.ByRow[rowNums[i]] = all[i];
                }
            }
            else if (ext == ".csv")
            {
                Encoding used;
                List<string[]> rows = MiniCsv.Parse(MiniCsv.DetectRead(workPath, out used));
                for (int i = 0; i < rows.Count; i++) t.ByRow[i + 1] = rows[i];
                t.Title = Path.GetFileNameWithoutExtension(path);
            }
            else
            {
                LastErr = side + " 侧不支持的格式: " + ext + "（支持 xlsx/csv；xls 将自动预转）";
                return null;
            }

            if (!t.ByRow.TryGetValue(headerRow, out t.Header)) t.Header = new string[0];
            foreach (int rn in t.ByRow.Keys) { if (rn > headerRow) t.DataRows.Add(rn); }
            t.DataRows.Sort();
            return t;
        }

        static void ResolveSheetErr(string msg) { }

        bool ResolveSheet(XlsxBook book, string spec, out int idx)
        {
            idx = 0;
            if (spec == null || spec.Trim().Length == 0) { LastErr = ""; return book.Sheets.Count > 0; }
            string s = spec.Trim();
            int n;
            if (int.TryParse(s, out n))
            {
                if (n >= 1 && n <= book.Sheets.Count) { idx = n - 1; return true; }
                LastErr = "sheet 序号越界: " + s;
                return false;
            }
            for (int i = 0; i < book.Sheets.Count; i++)
            {
                if (string.Equals(book.Sheets[i].Name.Trim(), s, StringComparison.Ordinal)) { idx = i; return true; }
            }
            return false;
        }

        static string[] TrimRow(string[] cells)
        {
            int last = -1;
            for (int i = cells.Length - 1; i >= 0; i--) { if (cells[i] != null) { last = i; break; } }
            if (last < 0) return new string[0];
            string[] copy = new string[last + 1];
            Array.Copy(cells, copy, last + 1);
            return copy;
        }

        // ================= 列解析 =================

        class Side
        {
            public int[] KeyCols = new int[0];
            public int Debit = -1, Credit = -1;
            public Table T;
        }

        public static string NormHeader(string s)
        {
            string h = Full2Half(s == null ? "" : s);
            StringBuilder sb = new StringBuilder();
            foreach (char ch in h)
            {
                if (ch == ' ' || ch == '\t' || ch == '\r' || ch == '\n' || ch == '\u00A0') continue;
                sb.Append(ch);
            }
            return sb.ToString().ToUpperInvariant();
        }

        // token：纯数字 → 1-based 列序号；否则按表头名（归一化）匹配
        static bool ResolveCol(string token, string[] header, out int col, out string err)
        {
            col = -1; err = null;
            string t = (token ?? "").Trim();
            if (t.Length == 0) return false;
            int n;
            if (int.TryParse(t, out n))
            {
                if (n >= 1 && n <= 64) { col = n - 1; return true; }
                err = "列序号越界: " + t; return false;
            }
            string want = NormHeader(t);
            for (int i = 0; i < header.Length; i++)
            {
                if (NormHeader(header[i]) == want) { col = i; return true; }
            }
            StringBuilder avail = new StringBuilder();
            for (int i = 0; i < header.Length; i++)
            {
                if (header[i] == null || header[i].Length == 0) continue;
                if (avail.Length > 0) avail.Append("、");
                avail.Append(header[i]);
            }
            err = "找不到列「" + t + "」（表头有：" + avail + "）";
            return false;
        }

        bool ResolveSide(string[] keyTokens, string debitTok, string creditTok, string[] header, out Side side)
        {
            side = new Side();
            LastErr = "";
            List<int> keys = new List<int>();
            foreach (string tok in keyTokens)
            {
                int c; string err;
                if (!ResolveCol(tok, header, out c, out err)) { LastErr = err; return false; }
                keys.Add(c);
            }
            side.KeyCols = keys.ToArray();
            if (debitTok != null && debitTok.Trim().Length > 0)
            {
                int c; string err;
                if (!ResolveCol(debitTok, header, out c, out err)) { LastErr = err; return false; }
                side.Debit = c;
            }
            if (creditTok != null && creditTok.Trim().Length > 0)
            {
                int c; string err;
                if (!ResolveCol(creditTok, header, out c, out err)) { LastErr = err; return false; }
                side.Credit = c;
            }
            return true;
        }

        // ================= 行抽取 =================

        List<ReconRow> ExtractRows(Table t, Side s, string sideTag, bool squeeze, ReconResult r)
        {
            List<ReconRow> rows = new List<ReconRow>();
            int maxKey = 0;
            foreach (int c in s.KeyCols) { if (c > maxKey) maxKey = c; }
            foreach (int rn in t.DataRows)
            {
                string[] cells = t.ByRow[rn];
                StringBuilder keyRaw = new StringBuilder();
                for (int i = 0; i < s.KeyCols.Length; i++)
                {
                    int c = s.KeyCols[i];
                    if (i > 0) keyRaw.Append('\x1f');
                    keyRaw.Append(c < cells.Length && cells[c] != null ? cells[c] : "");
                }
                ReconRow row = new ReconRow();
                row.ExcelRow = rn;
                row.KeyRaw = keyRaw.ToString();
                row.KeyNorm = squeeze ? NormKey(row.KeyRaw) : row.KeyRaw.Trim().ToUpperInvariant();

                decimal amount = 0;
                bool hasAmount = false, bad = false;
                if (s.Debit >= 0)
                {
                    string raw = s.Debit < cells.Length ? cells[s.Debit] : null;
                    decimal v;
                    if (TryAmount(raw, out v)) { amount += v; if (raw != null && raw.Trim().Length > 0) hasAmount = true; }
                    else { AddParseErr(r, sideTag, rn, t, s.Debit, raw); bad = true; }
                }
                if (!bad && s.Credit >= 0)
                {
                    string raw = s.Credit < cells.Length ? cells[s.Credit] : null;
                    decimal v;
                    if (TryAmount(raw, out v)) { amount -= v; if (raw != null && raw.Trim().Length > 0) hasAmount = true; }
                    else { AddParseErr(r, sideTag, rn, t, s.Credit, raw); bad = true; }
                }
                if (bad) continue;    // 解析失败行剔除（错误清单里），绝不静默当 0
                row.Amount = amount;
                row.HasAmount = hasAmount;
                rows.Add(row);
            }
            return rows;
        }

        static void AddParseErr(ReconResult r, string side, int rn, Table t, int col, string raw)
        {
            ReconParseError e = new ReconParseError();
            e.Side = side; e.ExcelRow = rn;
            e.ColName = col < t.Header.Length ? t.Header[col] : ("列" + (col + 1));
            e.Raw = raw == null ? "" : raw;
            e.Reason = "金额解析失败";
            r.ParseErrors.Add(e);
        }

        // ================= 归一化 / 金额解析（独立路径，供抽样重算复用） =================

        public static string Full2Half(string s)
        {
            if (s == null) return "";
            StringBuilder sb = new StringBuilder(s.Length);
            foreach (char c in s)
            {
                if (c == '\u3000') sb.Append(' ');
                else if (c >= '\uFF01' && c <= '\uFF5E') sb.Append((char)(c - 0xFEE0));
                else sb.Append(c);
            }
            return sb.ToString();
        }

        public static string NormKey(string raw)
        {
            string s = Full2Half(raw);
            StringBuilder sb = new StringBuilder(s.Length);
            foreach (char ch in s)
            {
                if (ch == ' ' || ch == '\t' || ch == '\r' || ch == '\n' || ch == '\u00A0') continue;
                sb.Append(ch);
            }
            return sb.ToString().ToUpperInvariant();
        }

        public static bool TryAmount(string raw, out decimal v)
        {
            v = 0;
            if (raw == null) return true;                    // 空单元格 = 0
            string s = Full2Half(raw).Trim();
            if (s.Length == 0) return true;
            bool neg = false;
            if ((s.StartsWith("(") && s.EndsWith(")")) || (s.StartsWith("（") && s.EndsWith("）")))
            {
                neg = true;
                s = s.Substring(1, s.Length - 2);
            }
            StringBuilder sb = new StringBuilder();
            foreach (char ch in s)
            {
                if ((ch >= '0' && ch <= '9') || ch == '.' || ch == '-') sb.Append(ch);
                else if (ch == ',' || ch == ' ' || ch == '\'' || ch == '¥' || ch == '￥'
                     || ch == '$' || ch == '€' || ch == '£' || ch == '元') { /* 千分位/币符 */ }
                else return false;                            // 未知字符 → 解析失败
            }
            string t = sb.ToString();
            if (t.Length == 0 || t == "-" || t == ".") return true;   // "-" 是银行流水 0 惯例
            if (!decimal.TryParse(t, NumberStyles.Number, CultureInfo.InvariantCulture, out v)) return false;
            if (neg) v = -v;
            return true;
        }

        // ================= 匹配与差异分类 =================

        void Match(List<ReconRow> rowsA, List<ReconRow> rowsB, ReconMapping map, Side sa, Side sb,
            Table ta, Table tb, ReconResult r)
        {
            decimal tol = Convert.ToDecimal(map.Tolerance);
            Dictionary<string, List<ReconRow>> byKey = new Dictionary<string, List<ReconRow>>();
            Dictionary<string, int> cursor = new Dictionary<string, int>();
            foreach (ReconRow b in rowsB)
            {
                List<ReconRow> list;
                if (!byKey.TryGetValue(b.KeyNorm, out list))
                {
                    list = new List<ReconRow>();
                    byKey[b.KeyNorm] = list;
                    cursor[b.KeyNorm] = 0;
                }
                list.Add(b);
            }
            foreach (ReconRow a in rowsA)
            {
                List<ReconRow> list;
                if (!byKey.TryGetValue(a.KeyNorm, out list))
                {
                    AddDiff(r, "仅A有", a.KeyNorm, a.Amount, 0, a.ExcelRow, -1, RefOf(ta, sa.Debit, sa.Credit, a.ExcelRow), null);
                    r.SumA += a.Amount;
                    continue;
                }
                int idx = cursor[a.KeyNorm];
                if (idx >= list.Count)
                {
                    AddDiff(r, "仅A有", a.KeyNorm, a.Amount, 0, a.ExcelRow, -1, RefOf(ta, sa.Debit, sa.Credit, a.ExcelRow), null);
                    r.SumA += a.Amount;
                    continue;
                }
                ReconRow b = list[idx];
                cursor[a.KeyNorm] = idx + 1;
                decimal delta = a.Amount - b.Amount;
                r.SumA += a.Amount; r.SumB += b.Amount;
                if (Math.Abs(delta) <= tol)
                {
                    r.MatchedCount++;
                    ReconDiffRow m = new ReconDiffRow();
                    m.Type = "匹配"; m.Key = a.KeyNorm; m.AmountA = a.Amount; m.AmountB = b.Amount;
                    m.Delta = delta; m.RowA = a.ExcelRow; m.RowB = b.ExcelRow;
                    m.RefA = RefOf(ta, sa.Debit, sa.Credit, a.ExcelRow);
                    m.RefB = RefOf(tb, sb.Debit, sb.Credit, b.ExcelRow);
                    r.Matched.Add(m);
                }
                else AddDiff(r, "金额不等", a.KeyNorm, a.Amount, b.Amount, a.ExcelRow, b.ExcelRow,
                    RefOf(ta, sa.Debit, sa.Credit, a.ExcelRow), RefOf(tb, sb.Debit, sb.Credit, b.ExcelRow));
            }
            // 未被 A 消费的 B 行（行内位置 >= 消费游标）→ 仅B有
            Dictionary<string, int> pos = new Dictionary<string, int>();
            foreach (ReconRow b in rowsB)
            {
                int p;
                pos.TryGetValue(b.KeyNorm, out p);
                pos[b.KeyNorm] = p + 1;
                if (p < cursor[b.KeyNorm]) continue;    // 已配对
                AddDiff(r, "仅B有", b.KeyNorm, 0, b.Amount, -1, b.ExcelRow, null, RefOf(tb, sb.Debit, sb.Credit, b.ExcelRow));
                r.SumB += b.Amount;
            }
            r.PairDiffCount = CountType(r.Diffs, "金额不等");
            r.OnlyACount = CountType(r.Diffs, "仅A有");
            r.OnlyBCount = CountType(r.Diffs, "仅B有");
        }

        static int CountType(List<ReconDiffRow> diffs, string type)
        {
            int n = 0;
            foreach (ReconDiffRow d in diffs) { if (d.Type == type) n++; }
            return n;
        }

        static void AddDiff(ReconResult r, string type, string key, decimal a, decimal b,
            int rowA, int rowB, string refA, string refB)
        {
            ReconDiffRow d = new ReconDiffRow();
            d.Type = type; d.Key = key; d.AmountA = a; d.AmountB = b; d.Delta = a - b;
            d.RowA = rowA; d.RowB = rowB;
            d.RefA = refA == null ? "-" : refA;
            d.RefB = refB == null ? "-" : refB;
            r.Diffs.Add(d);
        }

        static string RefOf(Table t, int debit, int credit, int excelRow)
        {
            int col = debit >= 0 ? debit : credit;
            if (col < 0) return "-";
            string colName = ColLetters(col);
            return t.Title + "!" + colName + excelRow;
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

        // ================= Verifier（勾稽自证） =================

        void Verify(List<ReconRow> rowsA, List<ReconRow> rowsB, ReconMapping map, Side sa, Side sb,
            Table ta, Table tb, ReconResult r)
        {
            decimal sumA = 0, sumB = 0;
            foreach (ReconRow x in rowsA) sumA += x.Amount;
            foreach (ReconRow x in rowsB) sumB += x.Amount;

            // 恒等式：两侧合计差 = 容差内匹配残差 + 差异合计
            //   全部配对贡献 Σ(a-b) = 容差内配对残差 + 超容差配对 delta（后者已并入差异合计）
            decimal tolerance = Convert.ToDecimal(map.Tolerance);
            decimal matchedTolResidual = 0;
            Dictionary<string, List<ReconRow>> bk = new Dictionary<string, List<ReconRow>>();
            Dictionary<string, int> cur = new Dictionary<string, int>();
            foreach (ReconRow b in rowsB)
            {
                List<ReconRow> list;
                if (!bk.TryGetValue(b.KeyNorm, out list)) { list = new List<ReconRow>(); bk[b.KeyNorm] = list; cur[b.KeyNorm] = 0; }
                list.Add(b);
            }
            foreach (ReconRow a in rowsA)
            {
                List<ReconRow> list;
                if (!bk.TryGetValue(a.KeyNorm, out list)) continue;
                int idx = cur[a.KeyNorm];
                if (idx >= list.Count) continue;
                ReconRow b = list[idx];
                cur[a.KeyNorm] = idx + 1;
                decimal delta = a.Amount - b.Amount;
                if (Math.Abs(delta) <= tolerance) matchedTolResidual += delta;
            }
            // AddDiff 统一 Delta = A - B（仅A有 Delta=金额A；仅B有 Delta=-金额B），直接求和
            decimal diffTotal = 0;
            foreach (ReconDiffRow d in r.Diffs) diffTotal += d.Delta;
            decimal left = sumA - sumB;
            decimal expected = matchedTolResidual + diffTotal;
            ReconCheck c1 = new ReconCheck();
            c1.Name = "合计勾稽（两侧合计差 = 差异合计 + 容差内匹配残差）";
            c1.Detail = "两侧合计差 " + MoneyText(left) + "；差异合计 " + MoneyText(diffTotal) +
                "；容差内匹配残差 " + MoneyText(matchedTolResidual);
            c1.Pass = Math.Abs(left - expected) <= tolerance * 10;
            if (!c1.Pass) c1.Detail += "；不匹配（差 " + MoneyText(left - expected) + "）";
            r.Checks.Add(c1);

            // 2) 行数守恒
            ReconCheck c2 = new ReconCheck();
            c2.Name = "行数守恒";
            int pairedA = r.MatchedCount + r.PairDiffCount;
            int pairedB = pairedA;
            bool cons = (pairedA + r.OnlyACount == rowsA.Count) && (pairedB + r.OnlyBCount == rowsB.Count);
            c2.Pass = cons;
            c2.Detail = "A 侧 " + pairedA + "(配对) + " + r.OnlyACount + "(仅A) = " + rowsA.Count +
                "；B 侧 " + pairedB + "(配对) + " + r.OnlyBCount + "(仅B) = " + rowsB.Count;
            r.Checks.Add(c2);

            // 3) 抽样重算：差异行原始单元格文本二次解析比对
            ReconCheck c3 = new ReconCheck();
            c3.Name = "差异行抽样重算（20 行）";
            int sampled = 0, bad = 0;
            foreach (ReconDiffRow d in r.Diffs)
            {
                if (sampled >= 20) break;
                if (d.Type == "仅A有" || d.Type == "仅B有") continue;
                string rawA = RawCell(ta, sa, d.RowA);
                string rawB = RawCell(tb, sb, d.RowB);
                decimal va, vb;
                bool ok = TryAmount(rawA, out va) && TryAmount(rawB, out vb) &&
                          Math.Abs(d.Delta - (va - vb)) <= tolerance;
                if (!ok) bad++;
                sampled++;
            }
            c3.Pass = bad == 0;
            c3.Detail = sampled == 0 ? "无配对差异行可抽样" : ("抽样 " + sampled + " 行，重算不一致 " + bad + " 行");
            r.Checks.Add(c3);

            // 4) 重复键提示（警告级）
            ReconCheck c4 = new ReconCheck();
            c4.Name = "重复键提示";
            c4.Warn = true;
            c4.Pass = true;
            Dictionary<string, int> cnt = new Dictionary<string, int>();
            foreach (ReconRow a in rowsA) { int n; cnt.TryGetValue(a.KeyNorm, out n); cnt[a.KeyNorm] = n + 1; }
            int dup = 0;
            foreach (KeyValuePair<string, int> kv in cnt) { if (kv.Value > 1) dup++; }
            c4.Detail = "A 侧重复键 " + dup + " 个（重复键按 FIFO 顺序配对，请人工复核）";
            r.Checks.Add(c4);

            // 5) 金额解析失败提示（警告级）
            ReconCheck c5 = new ReconCheck();
            c5.Name = "金额解析错误";
            c5.Warn = true;
            c5.Pass = true;
            c5.Detail = r.ParseErrors.Count == 0
                ? "全部金额单元格解析成功"
                : r.ParseErrors.Count + " 个金额单元格解析失败（已剔除并列入错误清单，见差异表「解析错误」页）";
            r.Checks.Add(c5);
        }

        static string RawCell(Table t, Side s, int excelRow)
        {
            if (excelRow < 0) return null;
            string[] cells;
            if (!t.ByRow.TryGetValue(excelRow, out cells)) return null;
            int col = s.Debit >= 0 ? s.Debit : s.Credit;
            if (col < 0 || col >= cells.Length) return null;
            return cells[col];
        }

        static string MoneyText(decimal v)
        {
            return v.ToString("N2", CultureInfo.InvariantCulture);
        }

        // ================= 差异表输出 =================

        string WriteReport(ReconResult r, ReconMapping map, string pathA, string pathB,
            Table ta, Table tb, string outPath)
        {
            ReportSheet sum = new ReportSheet("汇总与勾稽");
            sum.FreezeRows = 0;
            sum.ColWidths = new int[] { 42, 30, 30 };
            sum.AddRow(new ReportCell[] { new ReportCell("两表核对报告") { Style = "t" } });
            sum.AddRow(new ReportCell[] { new ReportCell("生成时间") { Style = "b" },
                new ReportCell(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")) { Style = "p" }, null });
            sum.AddRow(new ReportCell[] { new ReportCell("A 侧文件") { Style = "b" }, new ReportCell(pathA) { Style = "p" }, null });
            sum.AddRow(new ReportCell[] { new ReportCell("B 侧文件") { Style = "b" }, new ReportCell(pathB) { Style = "p" }, null });
            sum.AddRow(new ReportCell[] { new ReportCell("核对口径") { Style = "b" },
                new ReportCell("键列 A[" + JoinCols(map.KeyColsA) + "] B[" + JoinCols(map.KeyColsB) + "]；金额 A[" +
                    map.DebitA + (map.CreditA.Length > 0 ? "-" + map.CreditA : "") + "] B[" +
                    map.DebitB + (map.CreditB.Length > 0 ? "-" + map.CreditB : "") + "]；容差 ±" +
                    map.Tolerance.ToString("0.####", CultureInfo.InvariantCulture)) { Style = "p" }, null });
            sum.AddSpacer();
            sum.AddRow(new ReportCell[] { new ReportCell("指标") { Style = "h" }, new ReportCell("数值") { Style = "h" },
                new ReportCell("说明") { Style = "h" } });
            sum.AddRow(new ReportCell[] { new ReportCell("A 侧数据行"), ReportCell.N(r.TotalA, "m"),
                new ReportCell(ta.Title + "（表头第 " + ta.HeaderRow + " 行）") });
            sum.AddRow(new ReportCell[] { new ReportCell("B 侧数据行"), ReportCell.N(r.TotalB, "m"),
                new ReportCell(tb.Title + "（表头第 " + tb.HeaderRow + " 行）") });
            sum.AddRow(new ReportCell[] { new ReportCell("匹配成功"), ReportCell.N(r.MatchedCount, "mg"), new ReportCell("键相同且金额差在容差内") });
            sum.AddRow(new ReportCell[] { new ReportCell("金额不等"), ReportCell.N(r.PairDiffCount, "mr"), new ReportCell("键相同但金额差超容差") });
            sum.AddRow(new ReportCell[] { new ReportCell("仅 A 有"), ReportCell.N(r.OnlyACount, "o"), new ReportCell("A 侧存在、B 侧无同键") });
            sum.AddRow(new ReportCell[] { new ReportCell("仅 B 有"), ReportCell.N(r.OnlyBCount, "o"), new ReportCell("B 侧存在、A 侧无同键") });
            sum.AddRow(new ReportCell[] { new ReportCell("A 侧合计"), ReportCell.N((double)r.SumA, "m"), new ReportCell("解析成功的金额（借-贷口径）") });
            sum.AddRow(new ReportCell[] { new ReportCell("B 侧合计"), ReportCell.N((double)r.SumB, "m"), new ReportCell("解析成功的金额（借-贷口径）") });
            sum.AddSpacer();
            sum.AddRow(new ReportCell[] { new ReportCell("Verifier 勾稽校验") { Style = "h" }, new ReportCell("结果") { Style = "h" },
                new ReportCell("明细") { Style = "h" } });
            foreach (ReconCheck c in r.Checks)
            {
                string verdict = c.Warn ? "提示" : (c.Pass ? "✓ 通过" : "✗ 失败");
                sum.AddRow(new ReportCell[] { new ReportCell(c.Name),
                    new ReportCell(verdict) { Style = c.Warn ? "o" : (c.Pass ? "g" : "r") },
                    new ReportCell(c.Detail) });
            }
            if (r.ParseErrors.Count > 0)
            {
                sum.AddSpacer();
                sum.AddRow(new ReportCell[] { new ReportCell("解析错误清单") { Style = "h" }, new ReportCell("原始文本") { Style = "h" },
                    new ReportCell("位置") { Style = "h" } });
                foreach (ReconParseError e in r.ParseErrors)
                {
                    sum.AddRow(new ReportCell[] { new ReportCell(e.Side + " 侧 " + e.ExcelRow + " 行 " + e.ColName),
                        new ReportCell(e.Raw) { Style = "o" }, new ReportCell(e.Reason) });
                }
            }

            ReportSheet diff = new ReportSheet("差异明细");
            diff.FreezeRows = 1;
            diff.ColWidths = new int[] { 10, 22, 14, 14, 14, 18, 18, 10 };
            diff.AddRow(new ReportCell[] { new ReportCell("差异类型") { Style = "h" }, new ReportCell("键值") { Style = "h" },
                new ReportCell("A侧金额") { Style = "h" }, new ReportCell("B侧金额") { Style = "h" },
                new ReportCell("差额") { Style = "h" }, new ReportCell("A侧引用") { Style = "h" },
                new ReportCell("B侧引用") { Style = "h" }, new ReportCell("原始行") { Style = "h" } });
            foreach (ReconDiffRow d in r.Diffs)
            {
                string style = d.Type == "金额不等" ? "mr" : "om";
                string textStyle = d.Type == "金额不等" ? "r" : "o";
                diff.AddRow(new ReportCell[] {
                    new ReportCell(d.Type) { Style = textStyle },
                    new ReportCell(d.Key),
                    d.Type == "仅B有" ? new ReportCell("-") { Style = "p" } : (ReportCell)ReportCell.N((double)d.AmountA, style),
                    d.Type == "仅A有" ? new ReportCell("-") { Style = "p" } : (ReportCell)ReportCell.N((double)d.AmountB, style),
                    (ReportCell)ReportCell.N((double)d.Delta, style),
                    new ReportCell(d.RefA) { Style = "p" },
                    new ReportCell(d.RefB) { Style = "p" },
                    new ReportCell((d.RowA > 0 ? "A:" + d.RowA + " " : "") + (d.RowB > 0 ? "B:" + d.RowB : "")) { Style = "p" } });
            }
            if (r.Diffs.Count == 0)
                diff.AddRow(new ReportCell[] { new ReportCell("无差异") { Style = "g" } });

            List<ReportSheet> sheets = new List<ReportSheet>();
            sheets.Add(sum);
            sheets.Add(diff);

            if (map.IncludeMatched && r.Matched.Count > 0)
            {
                ReportSheet ms = new ReportSheet("匹配明细");
                ms.FreezeRows = 1;
                ms.ColWidths = new int[] { 22, 14, 14, 14, 18, 18 };
                ms.AddRow(new ReportCell[] { new ReportCell("键值") { Style = "h" }, new ReportCell("A侧金额") { Style = "h" },
                    new ReportCell("B侧金额") { Style = "h" }, new ReportCell("差额") { Style = "h" },
                    new ReportCell("A侧引用") { Style = "h" }, new ReportCell("B侧引用") { Style = "h" } });
                foreach (ReconDiffRow d in r.Matched)
                {
                    ms.AddRow(new ReportCell[] {
                        new ReportCell(d.Key),
                        (ReportCell)ReportCell.N((double)d.AmountA, "m"),
                        (ReportCell)ReportCell.N((double)d.AmountB, "m"),
                        (ReportCell)ReportCell.N((double)d.Delta, Math.Abs(d.Delta) > 0 ? "om" : "mg"),
                        new ReportCell(d.RefA) { Style = "p" },
                        new ReportCell(d.RefB) { Style = "p" } });
                }
                sheets.Add(ms);
            }
            return MiniXlsxWrite.Save(outPath, sheets);
        }

        static string JoinCols(string[] cols)
        {
            StringBuilder sb = new StringBuilder();
            foreach (string c in cols) { if (sb.Length > 0) sb.Append("+"); sb.Append(c); }
            return sb.ToString();
        }
    }
}
