// MiniXlsx —— 纯 BCL 的 xlsx 读取器（MiniZip + XmlReader 流式；net35 降级壳同源）
// M1 定位：预览与 CSV 导出足够；M2 加固：合并单元格、公式(无缓存值标记 "=")、
// 共享字符串 StringBuilder 累加、zip 解压体积防护、原始行号暴露（核对场景引用单元格）。
// 支持：多 sheet、共享字符串、inlineStr、日期/时间格式还原（styles.xml numFmt）、大文件流式。
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;
using OfficeAgent.Core;

namespace OfficeAgent.Host
{
    public class XlsxSheetInfo
    {
        public string Name = "";
        public string Target = "";   // zip 内路径
        public List<XlsxMerge> Merges = new List<XlsxMerge>();   // 流式读取时填充
    }

    // 合并区域（0-based 闭区间，含端点）
    public class XlsxMerge
    {
        public int R1, C1, R2, C2;
    }

    public class XlsxBook : IDisposable
    {
        MiniZipFile zip;
        List<string> sharedList = new List<string>();
        List<int> xfNumFmtId = new List<int>();
        Dictionary<int, bool> customFmtIsDate = new Dictionary<int, bool>();

        public List<XlsxSheetInfo> Sheets = new List<XlsxSheetInfo>();

        public static XlsxBook Open(string path)
        {
            XlsxBook b = new XlsxBook();
            b.zip = MiniZipFile.OpenRead(path);
            b.LoadSharedStrings();
            b.LoadStyles();
            b.LoadSheetList();
            return b;
        }

        public void Dispose()
        {
            if (zip != null) { zip.Dispose(); zip = null; }
        }

        MiniZipEntryInfo FindEntry(string name)
        {
            return zip.FindEntry(name);
        }

        // zip 炸弹防护：解压后体积上限（xlsx 常规远小于此；超出直接拒绝）
        const long MaxEntryBytes = 512L * 1024 * 1024;

        MiniZipEntryInfo FindEntryGuarded(string name)
        {
            MiniZipEntryInfo e = FindEntry(name);
            if (e == null) return null;
            if (e.UncompressedSize > MaxEntryBytes)
                throw new IOException("xlsx 条目过大（疑似 zip 炸弹）: " + name);
            return e;
        }

        void LoadSharedStrings()
        {
            MiniZipEntryInfo ss = FindEntryGuarded("xl/sharedStrings.xml");
            if (ss == null) return;
            using (XmlReader r = XmlReader.Create(zip.OpenEntry(ss.FullName)))
            {
                StringBuilder cur = null;
                while (r.Read())
                {
                    if (r.NodeType == XmlNodeType.Element && r.LocalName == "si") cur = new StringBuilder();
                    else if (r.NodeType == XmlNodeType.Text && cur != null) cur.Append(r.Value);
                    else if (r.NodeType == XmlNodeType.EndElement && r.LocalName == "si")
                    {
                        sharedList.Add(cur == null ? "" : cur.ToString());
                        cur = null;
                    }
                }
            }
        }

        void LoadStyles()
        {
            MiniZipEntryInfo st = FindEntry("xl/styles.xml");
            if (st == null) return;
            using (XmlReader r = XmlReader.Create(zip.OpenEntry(st.FullName)))
            {
                int section = 0; // 0 其他 1 numFmts 2 cellXfs
                while (r.Read())
                {
                    if (r.NodeType == XmlNodeType.Element)
                    {
                        if (r.LocalName == "numFmts") section = 1;
                        else if (r.LocalName == "cellXfs") section = 2;
                        else if (r.LocalName == "numFmt" && section == 1)
                        {
                            string id = r.GetAttribute("numFmtId");
                            string code = r.GetAttribute("formatCode");
                            int nid;
                            if (id != null && int.TryParse(id, out nid) && code != null)
                            {
                                string c = code.ToLowerInvariant();
                                customFmtIsDate[nid] = c.Contains("y") || c.Contains("d") || c.Contains("h\"");
                            }
                        }
                        else if (r.LocalName == "xf" && section == 2)
                        {
                            string fid = r.GetAttribute("numFmtId");
                            int n;
                            xfNumFmtId.Add(fid != null && int.TryParse(fid, out n) ? n : 0);
                        }
                    }
                    else if (r.NodeType == XmlNodeType.EndElement)
                    {
                        if (r.LocalName == "numFmts" || r.LocalName == "cellXfs") section = 0;
                    }
                }
            }
        }

        void LoadSheetList()
        {
            const string rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            Dictionary<string, string> rels = new Dictionary<string, string>();
            MiniZipEntryInfo re = FindEntry("xl/_rels/workbook.xml.rels");
            if (re != null)
            {
                using (XmlReader r = XmlReader.Create(zip.OpenEntry(re.FullName)))
                {
                    while (r.Read())
                    {
                        if (r.NodeType == XmlNodeType.Element && r.LocalName == "Relationship")
                        {
                            string id = r.GetAttribute("Id"), tg = r.GetAttribute("Target");
                            if (id != null && tg != null) rels[id] = tg;
                        }
                    }
                }
            }
            MiniZipEntryInfo wb = FindEntry("xl/workbook.xml");
            if (wb == null) return;
            using (XmlReader r = XmlReader.Create(zip.OpenEntry(wb.FullName)))
            {
                while (r.Read())
                {
                    if (r.NodeType == XmlNodeType.Element && r.LocalName == "sheet")
                    {
                        XlsxSheetInfo si = new XlsxSheetInfo();
                        si.Name = r.GetAttribute("name") ?? "";
                        string rid = r.GetAttribute("id", rns);
                        if (rid == null) rid = r.GetAttribute("r:id");
                        string target;
                        if (rid != null && rels.TryGetValue(rid, out target)) si.Target = target;
                        if (si.Target.StartsWith("/")) si.Target = si.Target.Substring(1);
                        else if (si.Target.Length > 0 && !si.Target.StartsWith("xl/", StringComparison.OrdinalIgnoreCase)) si.Target = "xl/" + si.Target;
                        Sheets.Add(si);
                    }
                }
            }
        }

        bool IsDateFmt(int xfIndex)
        {
            if (xfIndex < 0 || xfIndex >= xfNumFmtId.Count) return false;
            int id = xfNumFmtId[xfIndex];
            if (id >= 14 && id <= 22) return true;
            if (id >= 27 && id <= 36) return true;
            if (id >= 45 && id <= 47) return true;
            if (id >= 50 && id <= 58) return true;
            bool d;
            return customFmtIsDate.TryGetValue(id, out d) && d;
        }

        string ResolveCell(string val, string type, int styleIdx)
        {
            if (val == null) return null;
            if (type == "s")
            {
                int idx;
                if (int.TryParse(val, out idx) && idx >= 0 && idx < sharedList.Count) return sharedList[idx];
                return val;
            }
            if (type == "b") return val == "1" ? "TRUE" : "FALSE";
            if (type == "str" || type == "inlineStr" || type == "e") return val;
            // 数字：应用日期格式还原（会计表格刚需）
            double d;
            if (double.TryParse(val, NumberStyles.Float, CultureInfo.InvariantCulture, out d) && IsDateFmt(styleIdx)
                && d >= 0 && d <= 2958465)
            {
                try
                {
                    DateTime dt = DateTime.FromOADate(d);
                    return dt.TimeOfDay == TimeSpan.Zero
                        ? dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                        : dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                }
                catch { }
            }
            return val;
        }

        static int ColIndex(string addr, int fallback)
        {
            if (addr == null) return fallback;
            int n = 0, letters = 0;
            for (int i = 0; i < addr.Length; i++)
            {
                char c = addr[i];
                if (c >= 'A' && c <= 'Z') { n = n * 26 + (c - 'A' + 1); letters++; }
                else if (c >= 'a' && c <= 'z') { n = n * 26 + (c - 'a' + 1); letters++; }
                else break;
            }
            return letters > 0 ? n - 1 : fallback;
        }

        // 流式读行（CSV 导出与核对引擎用；rowHandler 收到列数组，可含 null）
        public string StreamRows(int sheetIndex, int maxCols, Action<string[]> rowHandler)
        {
            return StreamRows(sheetIndex, maxCols, delegate(string[] row, int excelRow)
            {
                rowHandler(row);
            });
        }

        // 带原始 Excel 行号（1-based）的重载：核对场景输出「文件!单元格引用」必需。
        // 顺带把工作表的 mergeCells 解析进 Sheets[sheetIndex].Merges（替换式写入）。
        public string StreamRows(int sheetIndex, int maxCols, Action<string[], int> rowHandler)
        {
            if (sheetIndex < 0 || sheetIndex >= Sheets.Count) return "sheet 序号越界";
            MiniZipEntryInfo se;
            try { se = FindEntryGuarded(Sheets[sheetIndex].Target); }
            catch (IOException ex) { return ex.Message; }
            if (se == null) return "找不到工作表 XML";
            List<XlsxMerge> merges = new List<XlsxMerge>();
            try
            {
                using (XmlReader r = XmlReader.Create(zip.OpenEntry(se.FullName)))
                {
                    int rowNumber = -1;
                    int curCol = 0, styleIdx = 0;
                    string cellType = null, val = null, pendingFormula = null;
                    bool haveCell = false, hadFormula = false;
                    string[] cells = new string[maxCols];
                    while (r.Read())
                    {
                        if (r.NodeType == XmlNodeType.Element && r.LocalName == "row")
                        {
                            if (rowNumber >= 0) { rowHandler(cells, rowNumber + 1); }
                            string rr = r.GetAttribute("r");
                            int n;
                            rowNumber = rr != null && int.TryParse(rr, out n) ? n - 1 : rowNumber + 1;
                            Array.Clear(cells, 0, cells.Length);
                            haveCell = false; hadFormula = false; pendingFormula = null;
                            if (r.IsEmptyElement)
                            {
                                // 自闭合空行 <row/>：保持行占位
                                rowHandler(cells, rowNumber + 1);
                                rowNumber = -1;
                            }
                        }
                        else if (r.NodeType == XmlNodeType.Element && r.LocalName == "c")
                        {
                            cellType = r.GetAttribute("t");
                            val = null;
                            haveCell = true; hadFormula = false; pendingFormula = null;
                            string s = r.GetAttribute("s");
                            int n;
                            styleIdx = s != null && int.TryParse(s, out n) ? n : 0;
                            curCol = ColIndex(r.GetAttribute("r"), curCol);
                        }
                        else if (r.NodeType == XmlNodeType.Element && r.LocalName == "f" && haveCell)
                        {
                            hadFormula = true;
                            string ftxt = r.ReadElementContentAsString();
                            if (ftxt == null || ftxt.Length == 0) ftxt = "(共享公式)";
                            if (ftxt.Length > 128) ftxt = ftxt.Substring(0, 128) + "…";
                            pendingFormula = ftxt;
                            // ReadElementContentAsString 已推进越过 </f>，随后的 <v> 照常解析
                        }
                        else if (r.NodeType == XmlNodeType.Element && r.LocalName == "v" && haveCell)
                        {
                            val = r.ReadElementContentAsString();
                            // ReadElementContentAsString 已把位置推进越过 </v>，
                            // EndElement("c") 分支不会再触发，必须就地提交
                            if (curCol >= 0 && curCol < maxCols) cells[curCol] = ResolveCell(val, cellType, styleIdx);
                            haveCell = false;
                        }
                        else if (r.NodeType == XmlNodeType.Element && r.LocalName == "t" && haveCell && cellType == "inlineStr")
                        {
                            val = (val == null ? "" : val) + r.ReadElementContentAsString();
                        }
                        else if (r.NodeType == XmlNodeType.EndElement && r.LocalName == "c" && haveCell)
                        {
                            if (curCol >= 0 && curCol < maxCols)
                            {
                                string sv = val == null && hadFormula
                                    ? "=" + pendingFormula      // 公式无缓存值：标记为公式文本，核对场景按非数值处理
                                    : ResolveCell(val, cellType, styleIdx);
                                cells[curCol] = sv;
                            }
                            haveCell = false;
                        }
                        else if (r.NodeType == XmlNodeType.Element && r.LocalName == "mergeCell")
                        {
                            string rf = r.GetAttribute("ref");
                            XlsxMerge m;
                            if (rf != null && (m = ParseMergeRef(rf)) != null) merges.Add(m);
                        }
                        else if (r.NodeType == XmlNodeType.EndElement && r.LocalName == "row" && rowNumber >= 0)
                        {
                            rowHandler(cells, rowNumber + 1);
                            rowNumber = -1; // 行已提交
                        }
                    }
                    // 收尾：文档结束时若仍有未提交的行（无 </row> 的畸形 XML）补交一次。
                    // 注意必须放在 while 之外——放在循环内会每个节点都提交当前行，
                    // 导致同一行被反复输出、后行覆盖前行（表现为表头丢失、数据行重复）。
                    if (rowNumber >= 0) rowHandler(cells, rowNumber + 1);
                }
                Sheets[sheetIndex].Merges = merges;
                return null;
            }
            catch (Exception ex) { return ex.Message; }
        }

        static XlsxMerge ParseMergeRef(string rf)
        {
            int ci = rf.IndexOf(':');
            if (ci < 0) return null;
            int c1, r1, c2, r2;
            if (!ParseAddr(rf.Substring(0, ci), out c1, out r1)) return null;
            if (!ParseAddr(rf.Substring(ci + 1), out c2, out r2)) return null;
            XlsxMerge m = new XlsxMerge();
            m.C1 = Math.Min(c1, c2); m.C2 = Math.Max(c1, c2);
            m.R1 = Math.Min(r1, r2); m.R2 = Math.Max(r1, r2);
            return m;
        }

        static bool ParseAddr(string addr, out int col, out int row)
        {
            col = -1; row = 0;
            if (addr == null || addr.Length == 0) return false;
            int i = 0, cn = 0, letters = 0;
            for (; i < addr.Length; i++)
            {
                char c = addr[i];
                if (c >= 'A' && c <= 'Z') { cn = cn * 26 + (c - 'A' + 1); letters++; }
                else if (c >= 'a' && c <= 'z') { cn = cn * 26 + (c - 'a' + 1); letters++; }
                else break;
            }
            int rn = 0, digits = 0;
            for (; i < addr.Length; i++)
            {
                char c = addr[i];
                if (c >= '0' && c <= '9') { rn = rn * 10 + (c - '0'); digits++; }
                else return false;
            }
            if (letters == 0 || digits == 0 || rn <= 0) return false;
            col = cn - 1; row = rn - 1;
            return true;
        }

        // 把合并区域锚点值回填到区域内其余单元格（预览/核对口径统一）
        public static void FillMerged(string[][] rows, List<XlsxMerge> merges)
        {
            if (merges == null || merges.Count == 0) return;
            foreach (XlsxMerge m in merges)
            {
                string anchor = null;
                if (m.R1 < rows.Length && m.C1 < rows[m.R1].Length) anchor = rows[m.R1][m.C1];
                if (anchor == null) continue;
                for (int r = m.R1; r <= m.R2; r++)
                {
                    if (r >= rows.Length) break;
                    for (int c = m.C1; c <= m.C2; c++)
                    {
                        if (c >= rows[r].Length) break;
                        if (r == m.R1 && c == m.C1) continue;
                        rows[r][c] = anchor;
                    }
                }
            }
        }

        // 预览加载：返回网格；totalRows=实际行数；usedCols=实际使用列数
        public string[,] LoadGrid(int sheetIndex, int maxRows, int maxCols, out int totalRows, out int usedCols)
        {
            totalRows = 0;
            usedCols = 0;
            if (sheetIndex < 0 || sheetIndex >= Sheets.Count) return new string[0, 0];
            MiniZipEntryInfo se = FindEntry(Sheets[sheetIndex].Target);
            if (se == null) return new string[0, 0];

            string dim = null;
            using (XmlReader r = XmlReader.Create(zip.OpenEntry(se.FullName)))
            {
                while (r.Read())
                {
                    if (r.NodeType == XmlNodeType.Element && r.LocalName == "dimension")
                    {
                        dim = r.GetAttribute("ref");
                        break;
                    }
                    if (r.NodeType == XmlNodeType.Element && r.LocalName == "row") break;
                }
            }
            if (dim != null)
            {
                int ci = dim.IndexOf(':');
                if (ci >= 0)
                {
                    int dr = ParseRowNumber(dim.Substring(ci + 1));
                    if (dr > 0) totalRows = dr;
                }
            }

            List<string[]> rows = new List<string[]>();
            int totalRowsLocal = totalRows;
            string err = StreamRows(sheetIndex, maxCols, delegate(string[] row)
            {
                // 必须拷贝：StreamRows 复用同一个 cells 缓冲并在行间 Clear，
                // 直接存引用会让所有行都指向最后一次的内容（表现为表头丢失、各行数据相同）
                if (rows.Count < maxRows)
                {
                    string[] copy = new string[row.Length];
                    Array.Copy(row, copy, row.Length);
                    rows.Add(copy);
                }
            });
            if (err != null && rows.Count == 0) return new string[0, 0];
            totalRows = rows.Count > 0 && totalRowsLocal < rows.Count ? rows.Count : Math.Max(totalRowsLocal, rows.Count);
            XlsxBook.FillMerged(rows.ToArray(), Sheets[sheetIndex].Merges);   // 合并单元格回填（预览口径）
            foreach (string[] row in rows)
            {
                for (int c = row.Length - 1; c >= 0; c--)
                {
                    if (row[c] != null) { if (c + 1 > usedCols) usedCols = c + 1; break; }
                }
            }
            string[,] grid = new string[rows.Count, Math.Max(usedCols, 1)];
            for (int i = 0; i < rows.Count; i++)
            {
                for (int j = 0; j < usedCols && j < rows[i].Length; j++) grid[i, j] = rows[i][j];
            }
            return grid;
        }

        static int ParseRowNumber(string addr)
        {
            if (addr == null) return 0;
            int start = 0, n = 0;
            for (int i = 0; i < addr.Length; i++)
            {
                char c = addr[i];
                if (c >= 'A' && c <= 'Z') { start = 1; continue; }
                if (c >= '0' && c <= '9') { n = n * 10 + (c - '0'); start = 1; }
                else if (start == 1) break;
            }
            return n;
        }
    }
}
