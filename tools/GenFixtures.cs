// GenFixtures —— 生成 M2 测试 fixtures（纯 BCL：ZipArchive 手写 OOXML，无 openpyxl 依赖）
// 构建：powershell -NoProfile -ExecutionPolicy Bypass -File tools\genfixtures.build.ps1
// 运行：build\test\GenFixtures.exe   （输出到 <repo>\tests\fixtures，路径取 exe 所在仓库根）
// 产物：recon\flow.csv + recon\ledger.csv（核对场景已知差异集）、dirty.xlsx（脏表）、
//       gb18030.csv、utf8n.csv（编码探测）、big.xlsx（10 万行性能冒烟）
//
// 核对场景差异集（flow=A 流水, ledger=B 账面，键=对方账号，金额=借-贷口径）：
//   正常匹配 10 组；脏键归一化匹配 3 组；容差内匹配 1 组(差 0.005 ≤ 0.01)；
//   重复键 1 组各 2 行全部配对；金额不等 2 组；仅 A 有 3 行；仅 B 有 2 行。
//   期望：matched=16, amount_diff=2, only_a=3, only_b=2
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace OfficeAgent.Tools
{
    internal static class GenFixtures
    {
        const int FmtDate = 2;      // cellXfs: yyyy-mm-dd
        const int FmtMoney = 3;     // cellXfs: #,##0.00
        const int FmtPct = 4;       // cellXfs: 0.00%

        static string repoRoot;

        static int Main(string[] args)
        {
            // exe 位于 <repo>\build\test\ → 仓库根为上两级；输出锚定仓库根，杜绝任意路径
            repoRoot = AppDomain.CurrentDomain.BaseDirectory;
            repoRoot = Path.GetFullPath(Path.Combine(repoRoot, "..", ".."));
            string fixDir = Path.Combine(repoRoot, "tests");
            fixDir = Path.Combine(fixDir, "fixtures");
            string reconDir = Path.Combine(fixDir, "recon");
            if (!Directory.Exists(reconDir)) Directory.CreateDirectory(reconDir);

            int flowRows, ledgerRows;
            GenRecon(Path.Combine(reconDir, "flow.csv"), Path.Combine(reconDir, "ledger.csv"),
                out flowRows, out ledgerRows);
            GenDirty(Path.Combine(fixDir, "dirty.xlsx"));
            WriteCsv(Path.Combine(fixDir, "gb18030.csv"),
                new string[][] {
                    new string[] { "科目", "余额" },
                    new string[] { "库存现金", "1234.50" },
                    new string[] { "银行存款", "9876.00" } }, 936, false);   // GB18030 → cp936，无 BOM
            WriteCsv(Path.Combine(fixDir, "utf8n.csv"),
                new string[][] {
                    new string[] { "科目", "余额" },
                    new string[] { "库存现金", "1234.50" } }, 65001, false); // UTF-8 无 BOM（编码探测用）
            GenBig(Path.Combine(fixDir, "big.xlsx"), 100000);
            GenInvoicePdf(Path.Combine(fixDir, "invoice-sample.pdf"));

            Console.WriteLine("fixtures done: flow_rows=" + flowRows + " ledger_rows=" + ledgerRows);
            Console.WriteLine("expect recon: matched=16 amount_diff=2 only_a=3 only_b=2");
            DumpSize(Path.Combine(reconDir, "flow.csv"));
            DumpSize(Path.Combine(reconDir, "ledger.csv"));
            DumpSize(Path.Combine(fixDir, "dirty.xlsx"));
            DumpSize(Path.Combine(fixDir, "gb18030.csv"));
            DumpSize(Path.Combine(fixDir, "utf8n.csv"));
            DumpSize(Path.Combine(fixDir, "big.xlsx"));
            DumpSize(Path.Combine(fixDir, "invoice-sample.pdf"));
            return 0;
        }

        static void DumpSize(string path)
        {
            FileInfo fi = new FileInfo(path);
            Console.WriteLine(string.Format(CultureInfo.InvariantCulture,
                "  {0,-40} {1,9} bytes", path.Substring(repoRoot.Length + 1),
                fi.Exists ? fi.Length : 0));
        }

        // ================= CSV =================

        static void WriteCsv(string path, string[][] rows, int codepage, bool bom)
        {
            StringBuilder sb = new StringBuilder();
            foreach (string[] row in rows)
            {
                for (int i = 0; i < row.Length; i++)
                {
                    if (i > 0) sb.Append(',');
                    sb.Append(row[i]);
                }
                sb.Append("\r\n");
            }
            Encoding enc = codepage == 65001 ? (Encoding)new UTF8Encoding(bom) : Encoding.GetEncoding(codepage);
            File.WriteAllText(path, sb.ToString(), enc);
        }

        // ================= 核对场景（A 流水 / B 账面） =================

        static string DayFor(int i)
        {
            return new DateTime(2026, 8, 1).AddDays(i).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        }

        static void GenRecon(string flowPath, string ledgerPath, out int flowRows, out int ledgerRows)
        {
            List<string[]> flow = new List<string[]>();
            List<string[]> ledger = new List<string[]>();
            flow.Add(new string[] { "账号", "户名", "交易日期", "摘要", "收入", "支出" });
            ledger.Add(new string[] { "凭证日期", "对方账号", "摘要", "借方", "贷方" });

            string[][] normal = new string[][] {
                new string[] { "1100101001", "北京宏图贸易", "5000.00", "货款" },
                new string[] { "1100101002", "上海蓝天科技", "3200.50", "服务费" },
                new string[] { "1100101003", "广州恒信电子", "880.00", "材料款" },
                new string[] { "1100101004", "深圳迅捷物流", "15000.00", "运费" },
                new string[] { "1100101005", "杭州云帆软件", "24000.00", "年费" },
                new string[] { "1100101006", "成都锦程咨询", "6000.75", "咨询费" },
                new string[] { "1100101007", "南京金陵印刷", "1200.00", "印刷费" },
                new string[] { "1100101008", "武汉长江建材", "4300.10", "建材款" },
                new string[] { "1100101009", "西安秦风机械", "21000.00", "设备款" },
                new string[] { "1100101010", "重庆山城餐饮", "350.00", "餐费" }
            };
            for (int i = 0; i < normal.Length; i++)
            {
                string dt = DayFor(i);
                string acct = normal[i][0], name = normal[i][1], amt = normal[i][2], memo = normal[i][3];
                flow.Add(new string[] { acct, name, dt, memo, amt, "0" });
                ledger.Add(new string[] { dt, acct, memo, amt, "0" });
            }

            // 3 组脏键：空格 / 全角 / 大小写差异 → 归一化后应匹配
            flow.Add(new string[] { " 1100202001 ", "天津海河商贸", "2026-08-20", "货款", "950.00", "0" });
            ledger.Add(new string[] { "2026-08-20", "1100202001", "货款", "950.00", "0" });
            flow.Add(new string[] { "１１００２０２００２", "苏州园林景观", "2026-08-20", "绿化款", "1800.00", "0" });
            ledger.Add(new string[] { "2026-08-20", "1100202002", "绿化款", "1800.00", "0" });
            flow.Add(new string[] { "abc003corp", "青岛黄海实业", "2026-08-20", "加工款", "777.77", "0" });
            ledger.Add(new string[] { "2026-08-20", "AbC003Corp", "加工款", "777.77", "0" });

            // 1 组容差内（差 0.005 ≤ 0.01）
            flow.Add(new string[] { "1100303001", "无锡太湖酒店", "2026-08-21", "住宿费", "1000.005", "0" });
            ledger.Add(new string[] { "2026-08-21", "1100303001", "住宿费", "1000.01", "0" });

            // 1 组重复键各 2 行，全部配对
            flow.Add(new string[] { "DUP0001", "重复键供应商A", "2026-08-22", "分期款1", "200.00", "0" });
            flow.Add(new string[] { "DUP0001", "重复键供应商A", "2026-08-22", "分期款2", "300.00", "0" });
            ledger.Add(new string[] { "2026-08-22", "DUP0001", "分期款1", "200.00", "0" });
            ledger.Add(new string[] { "2026-08-22", "DUP0001", "分期款2", "300.00", "0" });

            // 2 组金额不等（键两侧都有，差异超容差）
            flow.Add(new string[] { "1100404001", "佛山顺德五金", "2026-08-23", "五金款", "1500.00", "0" });
            ledger.Add(new string[] { "2026-08-23", "1100404001", "五金款", "1450.00", "0" });
            flow.Add(new string[] { "1100404002", "珠海拱北文具", "2026-08-24", "文具款", "88.80", "0" });
            ledger.Add(new string[] { "2026-08-24", "1100404002", "文具款", "8.88", "0" });

            // 3 行仅 A 有
            flow.Add(new string[] { "1100505001", "郑州中原广告", "2026-08-25", "广告费", "660.00", "0" });
            flow.Add(new string[] { "1100505002", "长沙湘江食品", "2026-08-25", "食品款", "420.00", "0" });
            flow.Add(new string[] { "1100505003", "合肥庐州仪器", "2026-08-25", "仪器款", "9999.99", "0" });

            // 2 行仅 B 有
            ledger.Add(new string[] { "2026-08-26", "1100606001", "化工款", "5000.00", "0" });
            ledger.Add(new string[] { "2026-08-26", "1100606002", "电费", "1234.00", "0" });

            WriteCsv(flowPath, flow.ToArray(), 65001, true);      // UTF-8 BOM（Excel 双击不乱码）
            WriteCsv(ledgerPath, ledger.ToArray(), 65001, true);
            flowRows = flow.Count - 1;
            ledgerRows = ledger.Count - 1;
        }

        // ================= 脏表 dirty.xlsx =================

        static void GenDirty(string path)
        {
            Sheet merge = new Sheet();
            merge.Name = "合并单元格";
            merge.Cols = new int[] { 18, 18, 14, 14, 12, 12 };
            merge.Merges = new string[] { "A1:B1", "B2:B3", "C2:D3", "C4:D4" };
            merge.Rows.Add(new Cell[] { Cell.S("季度费用汇总（标题跨两列）"), null, Cell.S("口径：含税"),
                Cell.S("备注"), Cell.S("金额"), Cell.S("占比") });
            merge.Rows.Add(new Cell[] { Cell.S("差旅费"), Cell.S("会计一部"), Cell.S("元"), null,
                Cell.N(12345.67, FmtMoney), Cell.N(0.256, FmtPct) });
            merge.Rows.Add(null);
            merge.Rows.Add(new Cell[] { Cell.S("办公费"), Cell.S("会计二部"), Cell.S("元"), null,
                Cell.N(9876.54, FmtMoney), Cell.N(0.431, FmtPct) });

            Sheet formula = new Sheet();
            formula.Name = "公式缓存";
            formula.Rows.Add(new Cell[] { Cell.S("项目"), Cell.S("数值"), Cell.S("日期"), Cell.S("占比"), Cell.S("全角") });
            formula.Rows.Add(new Cell[] { Cell.S("甲"), Cell.N(1, 0),
                Cell.N(ExcelSerial(new DateTime(2026, 9, 19)), FmtDate), Cell.N(0.156, FmtPct),
                Cell.S("１２３ ＡＢＣ　测试") });
            formula.Rows.Add(new Cell[] { Cell.S("乙"), Cell.N(2, 0),
                Cell.N(ExcelSerial(new DateTime(2026, 12, 31)), FmtDate), Cell.N(0.044, FmtPct),
                Cell.S("ａｂｃ") });
            formula.Rows.Add(new Cell[] { Cell.S("合计(有缓存)"), Cell.F("SUM(B2:B3)", 3), null, null, null });
            formula.Rows.Add(new Cell[] { Cell.S("引用(无缓存)"), Cell.F("B4*2", null), null, null, null });

            Sheet text = new Sheet();
            text.Name = "文本金额";
            text.Rows.Add(new Cell[] { Cell.S("科目"), Cell.S("余额") });
            text.Rows.Add(new Cell[] { Cell.S("库存现金"), Cell.S("1,234.56") });
            text.Rows.Add(new Cell[] { Cell.S("应收账款"), Cell.S("(123.45)") });
            text.Rows.Add(new Cell[] { Cell.S("银行存款"), Cell.N(1234567.89, FmtMoney) });
            text.Rows.Add(new Cell[] { Cell.S("坏账准备"), Cell.N(-987.65, FmtMoney) });
            text.Rows.Add(new Cell[] { Cell.S("备用金"), Cell.S("￥500元") });

            WriteXlsx(path, new Sheet[] { merge, formula, text });
        }

        // ================= 大表 big.xlsx =================

        static void GenBig(string path, int rows)
        {
            Sheet s = new Sheet();
            s.Name = "大表";
            s.Rows.Add(new Cell[] { Cell.S("序号"), Cell.S("科目"), Cell.S("借方"), Cell.S("贷方"),
                Cell.S("余额"), Cell.S("序号平方") });
            for (int i = 1; i <= rows; i++)
            {
                s.Rows.Add(new Cell[] { Cell.N(i, 0), Cell.S("科目" + (i % 1000).ToString("0000")),
                    Cell.N(i + 0.25, 0), Cell.N(i * 2 + 0.5, 0), Cell.N(i * 3 + 0.75, 0), Cell.N((double)i * i, 0) });
            }
            WriteXlsx(path, new Sheet[] { s });
        }

        // ================= 发票样例 PDF（Type0/UniGB-UCS2-H 非嵌入字体，pdfium 可提取中文） =================

        static void GenInvoicePdf(string path)
        {
            string[] lines = new string[] {
                "电子发票（普通发票）",
                "发票号码：24512000000123456789",
                "开票日期：2026年08月15日",
                "购买方信息",
                "名 称：北京宏图贸易有限公司  统一社会信用代码/纳税人识别号：91110108MA01ABC45F",
                "销售方信息",
                "名 称：上海蓝天科技有限公司  统一社会信用代码/纳税人识别号：91310115MA02DEF78G",
                "项目名称  数量  单价  金额  税率  税额",
                "*信息技术服务*软件服务费  1  2000.00  2000.00  6%  120.00",
                "合 计 ￥2000.00 ￥120.00",
                "价税合计（大写）贰仟壹佰贰拾元整 （小写）￥2120.00"
            };
            StringBuilder content = new StringBuilder();
            content.Append("BT /F1 11 Tf 14 TL\n");
            double y = 780;
            foreach (string ln in lines)
            {
                content.Append("1 0 0 1 40 ").Append(y.ToString("0", CultureInfo.InvariantCulture)).Append(" Td <");
                byte[] utf16 = Encoding.BigEndianUnicode.GetBytes(ln);
                StringBuilder hex = new StringBuilder(utf16.Length * 2);
                foreach (byte b in utf16) hex.Append(b.ToString("X2", CultureInfo.InvariantCulture));
                content.Append(hex).Append("> Tj\n");
                y -= 22;
            }
            content.Append("ET\n");

            // 组装 PDF 对象（xref 偏移手工计算）
            List<string> objs = new List<string>();
            objs.Add("<< /Type /Catalog /Pages 2 0 R >>");
            objs.Add("<< /Type /Pages /Kids [3 0 R] /Count 1 >>");
            objs.Add("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>");
            objs.Add("<< /Type /Font /Subtype /Type0 /BaseFont /SimSun /Encoding /UniGB-UCS2-H /DescendantFonts [6 0 R] >>");
            objs.Add("<< /Length " + content.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n" + content.ToString() + "endstream");
            objs.Add("<< /Type /Font /Subtype /CIDFontType2 /BaseFont /SimSun /CIDSystemInfo << /Registry (Adobe) /Ordering (GB1) /Supplement 5 >> /DW 1000 >>");

            StringBuilder pdf = new StringBuilder();
            pdf.Append("%PDF-1.7\n");
            List<int> offsets = new List<int>();
            for (int i = 0; i < objs.Count; i++)
            {
                offsets.Add(pdf.Length);
                pdf.Append((i + 1)).Append(" 0 obj\n").Append(objs[i]).Append("\nendobj\n");
            }
            long xref = pdf.Length;
            pdf.Append("xref\n0 ").Append(objs.Count + 1).Append("\n");
            pdf.Append("0000000000 65535 f \n");
            foreach (int off in offsets)
                pdf.Append(off.ToString("0000000000", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
            pdf.Append("trailer\n<< /Size ").Append(objs.Count + 1).Append(" /Root 1 0 R >>\nstartxref\n")
               .Append(xref).Append("\n%%EOF\n");
            File.WriteAllText(path, pdf.ToString(), new UTF8Encoding(false));
        }

        class Cell
        {
            public string Text;        // "s"
            public double Num;         // "n"/"f" 缓存
            public string Formula;     // "f"
            public int Xf;             // 样式索引
            public static Cell S(string t) { Cell c = new Cell(); c.Text = t; return c; }
            public static Cell N(double v, int xf) { Cell c = new Cell(); c.Num = v; c.Xf = xf; return c; }
            public static Cell F(string f, double? cached)
            {
                Cell c = new Cell(); c.Formula = f; c.Xf = 0;
                if (cached.HasValue) c.Num = cached.Value; else c.Num = double.NaN;
                return c;
            }
            public bool IsFormula { get { return Formula != null; } }
            public bool HasCached { get { return !double.IsNaN(Num); } }
        }

        class Sheet
        {
            public string Name = "";
            public List<Cell[]> Rows = new List<Cell[]>();
            public string[] Merges = null;
            public int[] Cols = null;
        }

        static double ExcelSerial(DateTime d)
        {
            return d.ToOADate();   // 与 Excel 序列一致（1899-12-30 基准）
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

        static string Esc(string s)
        {
            return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
                    .Replace("\"", "&quot;");
        }

        static string NumText(double v)
        {
            string s = v.ToString("0.##########", CultureInfo.InvariantCulture);
            if (s == "" || s == "-") s = "0";
            return s;
        }

        static string SheetXml(Sheet sh)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            if (sh.Cols != null)
            {
                sb.Append("<cols>");
                for (int j = 0; j < sh.Cols.Length; j++)
                    sb.Append("<col min=\"").Append(j + 1).Append("\" max=\"").Append(j + 1)
                      .Append("\" width=\"").Append(sh.Cols[j]).Append("\" customWidth=\"1\"/>");
                sb.Append("</cols>");
            }
            sb.Append("<sheetData>");
            for (int r = 0; r < sh.Rows.Count; r++)
            {
                sb.Append("<row r=\"").Append(r + 1).Append("\">");
                Cell[] row = sh.Rows[r];
                if (row != null)
                {
                    for (int c = 0; c < row.Length; c++)
                    {
                        Cell cell = row[c];
                        if (cell == null) continue;
                        string rf = ColName(c) + (r + 1);
                        string xf = cell.Xf > 0 ? " s=\"" + cell.Xf + "\"" : "";
                        if (cell.IsFormula)
                        {
                            sb.Append("<c r=\"").Append(rf).Append("\"><f>").Append(Esc(cell.Formula)).Append("</f>");
                            if (cell.HasCached) sb.Append("<v>").Append(NumText(cell.Num)).Append("</v>");
                            sb.Append("</c>");
                        }
                        else if (cell.Text != null)
                        {
                            sb.Append("<c r=\"").Append(rf).Append("\" t=\"inlineStr\"><is><t xml:space=\"preserve\">")
                              .Append(Esc(cell.Text)).Append("</t></is></c>");
                        }
                        else
                        {
                            sb.Append("<c r=\"").Append(rf).Append("\"").Append(xf)
                              .Append("><v>").Append(NumText(cell.Num)).Append("</v></c>");
                        }
                    }
                }
                sb.Append("</row>");
            }
            sb.Append("</sheetData>");
            if (sh.Merges != null && sh.Merges.Length > 0)
            {
                sb.Append("<mergeCells count=\"").Append(sh.Merges.Length).Append("\">");
                foreach (string m in sh.Merges) sb.Append("<mergeCell ref=\"").Append(m).Append("\"/>");
                sb.Append("</mergeCells>");
            }
            sb.Append("</worksheet>");
            return sb.ToString();
        }

        const string CtHead = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
            "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
            "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
            "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>";
        const string CtSheet = "<Override PartName=\"/xl/worksheets/sheet{0}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>";
        const string RelsRoot = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
            "</Relationships>";

        static readonly string StylesFull =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<numFmts count=\"3\">" +
            "<numFmt numFmtId=\"164\" formatCode=\"yyyy\\-mm\\-dd\"/>" +
            "<numFmt numFmtId=\"165\" formatCode=\"#,##0.00\"/>" +
            "<numFmt numFmtId=\"166\" formatCode=\"0.00%\"/>" +
            "</numFmts>" +
            "<fonts count=\"2\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font>" +
            "<font><b/><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>" +
            "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill>" +
            "<fill><patternFill patternType=\"gray125\"/></fill></fills>" +
            "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>" +
            "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>" +
            "<cellXfs count=\"5\">" +
            "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>" +
            "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>" +
            "<xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>" +
            "<xf numFmtId=\"165\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>" +
            "<xf numFmtId=\"166\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>" +
            "</cellXfs>" +
            "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>" +
            "</styleSheet>";

        static void WriteXlsx(string path, Sheet[] sheets)
        {
            if (File.Exists(path)) File.Delete(path);
            using (FileStream fs = new FileStream(path, FileMode.Create))
            using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Create))
            {
                StringBuilder ct = new StringBuilder(CtHead);
                StringBuilder wbSheets = new StringBuilder();
                StringBuilder wbRels = new StringBuilder();
                for (int i = 0; i < sheets.Length; i++)
                {
                    ct.Append(CtSheet.Replace("{0}", (i + 1).ToString(CultureInfo.InvariantCulture)));
                    wbSheets.Append("<sheet name=\"").Append(Esc(sheets[i].Name))
                      .Append("\" sheetId=\"").Append(i + 1)
                      .Append("\" r:id=\"rId").Append(i + 1).Append("\"/>");
                    wbRels.Append("<Relationship Id=\"rId").Append(i + 1)
                      .Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet")
                      .Append(i + 1).Append(".xml\"/>");
                }
                ct.Append("</Types>");
                AddEntry(zip, "[Content_Types].xml", ct.ToString());
                AddEntry(zip, "_rels/.rels", RelsRoot);
                AddEntry(zip, "xl/workbook.xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
                    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                    "<sheets>" + wbSheets.ToString() + "</sheets></workbook>");
                AddEntry(zip, "xl/_rels/workbook.xml.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    wbRels.ToString() + "</Relationships>");
                AddEntry(zip, "xl/styles.xml", StylesFull);
                for (int i = 0; i < sheets.Length; i++)
                    AddEntry(zip, "xl/worksheets/sheet" + (i + 1) + ".xml", SheetXml(sheets[i]));
            }
        }

        static void AddEntry(ZipArchive zip, string name, string content)
        {
            ZipArchiveEntry e = zip.CreateEntry(name, CompressionLevel.Optimal);
            using (StreamWriter w = new StreamWriter(e.Open(), new UTF8Encoding(false)))
            {
                w.Write(content);
            }
        }
    }
}
