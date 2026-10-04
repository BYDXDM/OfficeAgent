// MiniXlsxWrite —— 带样式的极简 xlsx 报表写出器（核对差异表 / 汇总底稿 / 发票清单共用）
// 纯 BCL（MiniZip），与 ConvertEngine.WriteMinimalXlsx 同源思路；红线：C# 3.0 语法、UTF-8 输出。
// 样式集：表头(灰底加粗居中) / 金额(#,##0.00) / 红(差异) / 绿(通过) / 橙(警示) / 标题 / 首行冻结。
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using OfficeAgent.Core;

namespace OfficeAgent.Host
{
    public class ReportCell
    {
        public string Text = null;       // 文本单元格（与 Number 二选一）
        public double? Number = null;    // 数值单元格
        public string Style = "n";       // n h m mr mg r g o om b t
        public bool IsFormula = false;   // true 时 Text 为公式（无缓存值，Excel 打开时计算）

        public ReportCell() { }
        public ReportCell(string text) { Text = text; }

        public static ReportCell S(string t) { return new ReportCell(t); }
        public static ReportCell N(double v, string style)
        {
            ReportCell c = new ReportCell();
            c.Number = v; c.Style = style;
            return c;
        }
        public static ReportCell F(string formula, string style)
        {
            ReportCell c = new ReportCell();
            c.Text = formula; c.IsFormula = true; c.Style = style;
            return c;
        }
    }

    public class ReportSheet
    {
        public string Name;
        public List<List<ReportCell>> Rows = new List<List<ReportCell>>();
        public int[] ColWidths = null;   // 字符宽度；null = 默认
        public int FreezeRows = 0;       // 冻结前 N 行（表头 1）
        // 数据区结束行（0 基，含表头所在行；-1 = 未指定，按 Rows.Count 处理）。
        // 用于"明细区"与"追加的合计行"区分：汇总 sheet 的分组扫描与 SUMIF 区间都必须止于此处，
        // 否则会把追加的「合计」行当成一个分组值并重复计入（自由模式 totalRow+summary 双计缺陷）。
        public int DataEndRow = -1;

        public ReportSheet(string name) { Name = name; }

        public void AddRow(params ReportCell[] cells)
        {
            List<ReportCell> row = new List<ReportCell>();
            foreach (ReportCell c in cells) row.Add(c);
            Rows.Add(row);
        }

        public void AddTextRow(params string[] texts)
        {
            List<ReportCell> row = new List<ReportCell>();
            foreach (string t in texts) row.Add(new ReportCell(t));
            Rows.Add(row);
        }

        public void AddSpacer()
        {
            Rows.Add(new List<ReportCell>());
        }
    }

    public static class MiniXlsxWrite
    {
        // 样式名 → cellXfs 索引（与下方 StylesXml 顺序严格一致）
        // n 普通带边框 / h 表头 / m 金额 / mr 红金额 / mg 绿金额 / om 橙金额
        // r 红文本 / g 绿文本 / o 橙文本 / b 加粗 / t 标题 / p 纯文本无边框
        static int StyleIndex(string style)
        {
            switch (style)
            {
                case "h": return 1;
                case "m": return 2;
                case "mr": return 3;
                case "mg": return 4;
                case "r": return 5;
                case "g": return 6;
                case "o": return 7;
                case "om": return 8;
                case "b": return 9;
                case "t": return 10;
                case "p": return 11;
                default: return 0;
            }
        }

        static string StylesXml()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            sb.Append("<numFmts count=\"1\"><numFmt numFmtId=\"165\" formatCode=\"#,##0.00\"/></numFmts>");
            // 字体：0 普通 1 加粗 2 加粗14(标题) 3 红 4 绿 5 橙
            sb.Append("<fonts count=\"6\">");
            sb.Append("<font><sz val=\"11\"/><name val=\"Calibri\"/></font>");
            sb.Append("<font><b/><sz val=\"11\"/><name val=\"Calibri\"/></font>");
            sb.Append("<font><b/><sz val=\"14\"/><name val=\"Calibri\"/></font>");
            sb.Append("<font><b/><sz val=\"11\"/><color rgb=\"FF9C0006\"/><name val=\"Calibri\"/></font>");
            sb.Append("<font><b/><sz val=\"11\"/><color rgb=\"FF006100\"/><name val=\"Calibri\"/></font>");
            sb.Append("<font><b/><sz val=\"11\"/><color rgb=\"FF9C6500\"/><name val=\"Calibri\"/></font>");
            sb.Append("</fonts>");
            // 填充：0 none 1 gray125 2 表头浅灰蓝 3 红 4 绿 5 橙
            sb.Append("<fills count=\"6\">");
            sb.Append("<fill><patternFill patternType=\"none\"/></fill>");
            sb.Append("<fill><patternFill patternType=\"gray125\"/></fill>");
            sb.Append("<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFD9E1F2\"/><bgColor indexed=\"64\"/></patternFill></fill>");
            sb.Append("<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFFFC7CE\"/><bgColor indexed=\"64\"/></patternFill></fill>");
            sb.Append("<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFC6EFCE\"/><bgColor indexed=\"64\"/></patternFill></fill>");
            sb.Append("<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFFFEB9C\"/><bgColor indexed=\"64\"/></patternFill></fill>");
            sb.Append("</fills>");
            sb.Append("<borders count=\"2\">");
            sb.Append("<border><left/><right/><top/><bottom/><diagonal/></border>");
            sb.Append("<border><left style=\"thin\"><color rgb=\"FFB0B7C3\"/></left><right style=\"thin\"><color rgb=\"FFB0B7C3\"/></right>" +
                      "<top style=\"thin\"><color rgb=\"FFB0B7C3\"/></top><bottom style=\"thin\"><color rgb=\"FFB0B7C3\"/></bottom><diagonal/></border>");
            sb.Append("</borders>");
            sb.Append("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");
            sb.Append("<cellXfs count=\"12\">");
            // 0 n
            sb.Append("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyBorder=\"1\"/>");
            // 1 h 表头
            sb.Append("<xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\">" +
                      "<alignment horizontal=\"center\" vertical=\"center\" wrapText=\"1\"/></xf>");
            // 2 m 金额
            sb.Append("<xf numFmtId=\"165\" fontId=\"0\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyNumberFormat=\"1\" applyBorder=\"1\"/>");
            // 3 mr 红金额
            sb.Append("<xf numFmtId=\"165\" fontId=\"3\" fillId=\"3\" borderId=\"1\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\"/>");
            // 4 mg 绿金额
            sb.Append("<xf numFmtId=\"165\" fontId=\"4\" fillId=\"4\" borderId=\"1\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\"/>");
            // 5 r 红文本
            sb.Append("<xf numFmtId=\"0\" fontId=\"3\" fillId=\"3\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\"/>");
            // 6 g 绿文本
            sb.Append("<xf numFmtId=\"0\" fontId=\"4\" fillId=\"4\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\"/>");
            // 7 o 橙文本
            sb.Append("<xf numFmtId=\"0\" fontId=\"5\" fillId=\"5\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\"/>");
            // 8 om 橙金额
            sb.Append("<xf numFmtId=\"165\" fontId=\"5\" fillId=\"5\" borderId=\"1\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\"/>");
            // 9 b 加粗
            sb.Append("<xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyBorder=\"1\"/>");
            // 10 t 标题
            sb.Append("<xf numFmtId=\"0\" fontId=\"2\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>");
            // 11 p 纯文本无边框
            sb.Append("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>");
            sb.Append("</cellXfs>");
            sb.Append("<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>");
            sb.Append("</styleSheet>");
            return sb.ToString();
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
            if (s == null) return "";
            return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
                    .Replace("\"", "&quot;");
        }

        static string SheetXml(ReportSheet sh)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            if (sh.FreezeRows > 0)
            {
                sb.Append("<sheetViews><sheetView workbookViewId=\"0\"><pane ySplit=\"").Append(sh.FreezeRows)
                  .Append("\" topLeftCell=\"A").Append(sh.FreezeRows + 1)
                  .Append("\" activePane=\"bottomLeft\" state=\"frozen\"/></sheetView></sheetViews>");
            }
            if (sh.ColWidths != null && sh.ColWidths.Length > 0)
            {
                sb.Append("<cols>");
                for (int j = 0; j < sh.ColWidths.Length; j++)
                    sb.Append("<col min=\"").Append(j + 1).Append("\" max=\"").Append(j + 1)
                      .Append("\" width=\"").Append(sh.ColWidths[j]).Append("\" customWidth=\"1\"/>");
                sb.Append("</cols>");
            }
            sb.Append("<sheetData>");
            for (int r = 0; r < sh.Rows.Count; r++)
            {
                List<ReportCell> row = sh.Rows[r];
                sb.Append("<row r=\"").Append(r + 1).Append("\">");
                for (int c = 0; c < row.Count; c++)
                {
                    ReportCell cell = row[c];
                    if (cell == null) continue;
                    string rf = ColName(c) + (r + 1);
                    string sattr = " s=\"" + StyleIndex(cell.Style) + "\"";
                    if (cell.IsFormula)
                    {
                        sb.Append("<c r=\"").Append(rf).Append("\"").Append(sattr).Append("><f>")
                          .Append(Esc(cell.Text)).Append("</f></c>");
                    }
                    else if (cell.Number.HasValue)
                    {
                        sb.Append("<c r=\"").Append(rf).Append("\"").Append(sattr).Append("><v>")
                          .Append(cell.Number.Value.ToString("0.##########", CultureInfo.InvariantCulture))
                          .Append("</v></c>");
                    }
                    else if (cell.Text != null && cell.Text.Length > 0)
                    {
                        sb.Append("<c r=\"").Append(rf).Append("\"").Append(sattr)
                          .Append(" t=\"inlineStr\"><is><t xml:space=\"preserve\">")
                          .Append(Esc(cell.Text)).Append("</t></is></c>");
                    }
                    else
                    {
                        sb.Append("<c r=\"").Append(rf).Append("\"").Append(sattr).Append("/>");
                    }
                }
                sb.Append("</row>");
            }
            sb.Append("</sheetData></worksheet>");
            return sb.ToString();
        }

        // 保存报表；返回 null=成功，否则错误信息
        public static string Save(string path, List<ReportSheet> sheets)
        {
            return Save(path, sheets, 0);
        }

        // activeTab：打开文件时落在第几张表（0 起）——模板把工作表排第一时用
        public static string Save(string path, List<ReportSheet> sheets, int activeTab)
        {
            try
            {
                if (File.Exists(path)) File.Delete(path);
                using (FileStream fs = new FileStream(path, FileMode.Create))
                {
                    List<MiniZipEntryOut> ents = new List<MiniZipEntryOut>();
                    StringBuilder wbSheets = new StringBuilder();
                    for (int i = 0; i < sheets.Count; i++)
                    {
                        wbSheets.Append("<sheet name=\"").Append(Esc(sheets[i].Name))
                          .Append("\" sheetId=\"").Append(i + 1)
                          .Append("\" r:id=\"rId").Append(i + 1).Append("\"/>");
                    }
                    XlsxSkeleton.Add(ents, "[Content_Types].xml", XlsxSkeleton.ContentTypes(sheets.Count, true));
                    XlsxSkeleton.Add(ents, "_rels/.rels", XlsxSkeleton.RootRels());
                    XlsxSkeleton.Add(ents, "xl/workbook.xml", XlsxSkeleton.WorkbookXml(wbSheets.ToString(), activeTab));
                    XlsxSkeleton.Add(ents, "xl/_rels/workbook.xml.rels", XlsxSkeleton.SheetsRels(sheets.Count));
                    XlsxSkeleton.Add(ents, "xl/styles.xml", StylesXml());
                    for (int i = 0; i < sheets.Count; i++)
                        XlsxSkeleton.Add(ents, "xl/worksheets/sheet" + (i + 1) + ".xml", SheetXml(sheets[i]));
                    MiniZipWriter.Write(fs, ents);
                }
                return null;
            }
            catch (Exception ex) { return ex.Message; }
        }
    }
}
