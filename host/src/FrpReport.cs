// FrpReport —— 解析 .frp 表单打印模板（Delphi/FastReport 系二进制格式，用户实样逆向）。
//
// 文件结构（2026-10-05 由用户提供的 1.frp 实样破解，1563 对象零校验失败）：
//   顺序数据流，字符串 = [1 字节类型 0|1][u16 长度(LE)][内容]；GBK/ANSI 文本。
//   对象记录 = [字符串 "MemoNN"/"LineNN"/…][00][02 00][u32 左][u32 上][u32 宽][u32 高]
//   随后若干属性串（字体名 宋体/Arial、@ 等杂项）与对象的文本串混排——
//   取该对象名到下一对象名之间"最后一个非字体名字符串"为对象文本（实样验证：序号"1"、
//   标题、时间 "07:47-     " 全部命中；末位规则同时甩掉几何后的 1 字符属性杂值）。
//   版面为坐标定位（员工区块左右两栏并排）：行=按上坐标分带，列=按左坐标分带。
// 用途：查看（read_text_file/预览）与输出（转 xlsx，交给既有表格链路）。
// 红线：C# 3.0 语法；只读解析，不执行任何来自文件的内容。
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OfficeAgent.Host
{
    public static class FrpReport
    {
        class FrpObj
        {
            public string Name = "";
            public int L, T, W, H;
            public string Text = "";
        }

        static readonly string[] NamePrefixes = new string[] {
            "Memo", "Line", "Text", "Shape", "Picture", "CheckBox", "Page", "Band", "Sub", "R" };

        static readonly string[] FontNames = new string[] {
            "宋体", "Arial", "黑体", "楷体", "楷体_GB2312", "仿宋", "仿宋_GB2312", "微软雅黑",
            "Times New Roman", "Courier New", "Wingdings", "Webdings", "MS Sans Serif", "Tahoma" };

        static bool IsObjName(string s)
        {
            if (s == null || s.Length < 2 || s.Length > 32) return false;
            int d = 0;
            for (int p = 0; p < NamePrefixes.Length; p++)
            {
                string pre = NamePrefixes[p];
                if (s.StartsWith(pre, StringComparison.Ordinal) && s.Length > pre.Length)
                {
                    d = pre.Length;
                    break;
                }
            }
            if (d == 0) return false;
            for (int i = d; i < s.Length; i++)
            {
                if (s[i] < '0' || s[i] > '9') return false;
            }
            return true;
        }

        static bool IsFontName(string s)
        {
            string t = (s ?? "").Trim();
            for (int i = 0; i < FontNames.Length; i++)
            {
                if (string.Equals(FontNames[i], t, StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        }

        // 串是否"足够可打印"（GBK/ANSI；二进制乱码段不放行）。
        // 规则：NUL 一票否决（二进制特征，且 XML 不容忍）；其余控制字符按不可打印计数，
        // 可打印数 ≥ max(1, 长度×0.9)。★ 下限 1 是关键——纯控制短串（如 "\x02"）在
        // len=1 时 (len*9)/10 取整为 0，旧写法 0>=0 恒过 → 钻进单元格写坏 XML。
        // ★ 不能"见控制字符就整串否决"——.NET 936 对串尾杂字节做 best-fit 映射会产生
        //   U+0001 类尾随控制符（如真文本 "1" 的原始串是 "1\x01"），整串否决会丢掉
        //   98% 的对象文本（实测 withText 从 1218 掉到 27）；残余控制符由 CleanCell 清掉。
        static bool Printable(string s, int rawLen)
        {
            if (s == null || s.Length == 0) return false;
            int ok = 0;
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c == 0) return false;
                if (c == '\r' || c == '\n' || c == '\t' || !char.IsControl(c)) ok++;
            }
            return ok >= Math.Max(1, (s.Length * 9) / 10);
        }

        // 单元格文本清洗：去掉 XML 不容忍的控制符、压平换行
        static string CleanCell(string s)
        {
            if (s == null) return "";
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c == '\r' || c == '\n') { sb.Append(' '); continue; }
                if (c >= 0x20) sb.Append(c);
            }
            return sb.ToString().Trim();
        }

        // 顺序解析：返回对象列表（含几何与文本）。不抛异常，坏段跳过。
        static List<FrpObj> ParseObjects(byte[] b)
        {
            List<FrpObj> objs = new List<FrpObj>();
            FrpObj cur = null;
            List<string> group = new List<string>();   // 当前对象名之后收集到的字符串
            int n = b.Length;
            int i = 0;
            while (i < n - 3)
            {
                int flag = b[i];
                if (flag != 0 && flag != 1) { i++; continue; }
                int ln = b[i + 1] | (b[i + 2] << 8);
                if (ln < 1 || ln > 600 || i + 3 + ln > n) { i++; continue; }
                string t;
                {
                    byte[] seg = new byte[ln];
                    Array.Copy(b, i + 3, seg, 0, ln);
                    bool ascii = true;
                    for (int k = 0; k < ln; k++)
                    {
                        byte c = seg[k];
                        if (!((c >= 0x20 && c < 0x7f) || c == 9 || c == 10 || c == 13)) { ascii = false; break; }
                    }
                    Encoding enc = ascii ? Encoding.ASCII : Encoding.GetEncoding(936);
                    t = enc.GetString(seg);
                    if (!Printable(t, ln)) { i++; continue; }
                }
                int after = i + 3 + ln;
                if (IsObjName(t))
                {
                    if (after + 19 <= n && b[after] == 0 && b[after + 1] == 2 && b[after + 2] == 0)
                    {
                        int o = after + 3;
                        int L = BitConverter.ToInt32(b, o);
                        int T = BitConverter.ToInt32(b, o + 4);
                        int W = BitConverter.ToInt32(b, o + 8);
                        int H = BitConverter.ToInt32(b, o + 12);
                        if (L >= 0 && L < 5000 && T >= 0 && T < 5000 && W > 0 && W < 5000 && H >= 0 && H < 1500)
                        {
                            if (cur != null) cur.Text = PickText(group);
                            group = new List<string>();
                            cur = new FrpObj();
                            cur.Name = t; cur.L = L; cur.T = T; cur.W = W; cur.H = H;
                            objs.Add(cur);
                            i = o + 16;
                            continue;
                        }
                    }
                }
                group.Add(t);
                i = after;
            }
            if (cur != null) cur.Text = PickText(group);
            return objs;
        }

        // 对象文本 = 名字到下一名字之间"最后一个非字体名"字符串（见文件头注释）
        static string PickText(List<string> group)
        {
            string last = "";
            for (int i = 0; i < group.Count; i++)
            {
                if (!IsFontName(group[i])) last = group[i];
            }
            return CleanCell(last);
        }

        // 解析为行列网格：行=上坐标分带（锚点容差 5），列=左坐标分带（全局锚点，容差 8）。
        public static List<string[]> ParseToGrid(string path, out string err)
        {
            err = null;
            List<string[]> empty = new List<string[]>();
            byte[] b;
            try { b = File.ReadAllBytes(path); }
            catch (Exception ex) { err = "读取失败: " + ex.Message; return empty; }
            if (b.Length < 16) { err = "文件太小，不是有效的 frp 模板"; return empty; }
            List<FrpObj> objs = ParseObjects(b);
            List<FrpObj> withText = new List<FrpObj>();
            for (int i = 0; i < objs.Count; i++)
            {
                if (objs[i].Text != null && objs[i].Text.Trim().Length > 0) withText.Add(objs[i]);
            }
            if (withText.Count == 0) { err = "未在模板中找到文本对象"; return empty; }

            withText.Sort(delegate(FrpObj a, FrpObj c)
            {
                if (a.T != c.T) return a.T - c.T;
                return a.L - c.L;
            });

            // 全局列锚点
            List<int> colAnchor = new List<int>();
            foreach (FrpObj o in withText)
            {
                int hit = -1;
                for (int k = 0; k < colAnchor.Count; k++)
                {
                    if (Math.Abs(colAnchor[k] - o.L) <= 8) { hit = k; break; }
                }
                if (hit < 0)
                {
                    colAnchor.Add(o.L);
                    colAnchor.Sort();
                }
            }

            List<string[]> grid = new List<string[]>();
            int rowAnchor = -1000;
            List<FrpObj> rowItems = new List<FrpObj>();
            List<int> colOfRow = new List<int>();
            Action flush = delegate
            {
                if (rowItems.Count == 0) return;
                string[] r = new string[colAnchor.Count];
                for (int k = 0; k < r.Length; k++) r[k] = "";
                for (int k = 0; k < rowItems.Count; k++)
                {
                    r[colOfRow[k]] = rowItems[k].Text.Replace("\r", " ").Replace("\n", " ").Trim();
                }
                grid.Add(r);
                rowItems = new List<FrpObj>();
                colOfRow = new List<int>();
            };
            for (int i = 0; i < withText.Count; i++)
            {
                FrpObj o = withText[i];
                if (rowAnchor < -500 || o.T - rowAnchor > 5)
                {
                    flush();
                    rowAnchor = o.T;
                }
                int ci = -1;
                for (int k = 0; k < colAnchor.Count; k++)
                {
                    if (Math.Abs(colAnchor[k] - o.L) <= 8) { ci = k; break; }
                }
                if (ci < 0) ci = 0;
                rowItems.Add(o);
                colOfRow.Add(ci);
            }
            flush();
            return grid;
        }

        // 纯文本视图（read_text_file / 查看用）
        public static string ToText(string path, out string err)
        {
            List<string[]> grid = ParseToGrid(path, out err);
            if (err != null) return "";
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("（frp 打印模板解析结果；每行=版面同一横行，| 分隔单元格）");
            for (int r = 0; r < grid.Count; r++)
            {
                StringBuilder line = new StringBuilder();
                for (int c = 0; c < grid[r].Length; c++)
                {
                    if (c > 0) line.Append(" | ");
                    line.Append(grid[r][c]);
                }
                string s = line.ToString().TrimEnd(' ', '|');
                if (s.Trim().Length > 0) sb.AppendLine(s);
            }
            return sb.ToString();
        }

        // 输出：模板 → xlsx（落盘路径由调用方决定；返回 null=成功）
        public static string ConvertToXlsx(string frpPath, string outXlsx)
        {
            string err;
            List<string[]> grid = ParseToGrid(frpPath, out err);
            if (err != null) return err;
            ReportSheet sheet = new ReportSheet("Sheet1");
            for (int r = 0; r < grid.Count; r++)
            {
                List<ReportCell> row = new List<ReportCell>();
                for (int c = 0; c < grid[r].Length; c++)
                {
                    ReportCell cell;
                    if (r == 0) { cell = ReportCell.S(grid[r][c]); cell.Style = "h"; }
                    else
                    {
                        double num;
                        string t = grid[r][c];
                        if (t.Length > 1 && t[0] == '=') cell = ReportCell.F(t.Substring(1), "n");
                        else if (t.Length > 0 && double.TryParse(t, out num)) cell = ReportCell.N(num, "n");
                        else cell = ReportCell.S(t);
                    }
                    row.Add(cell);
                }
                sheet.Rows.Add(row);
            }
            sheet.FreezeRows = 0;
            List<ReportSheet> sheets = new List<ReportSheet>();
            sheets.Add(sheet);
            return MiniXlsxWrite.Save(outXlsx, sheets);
        }
    }
}
