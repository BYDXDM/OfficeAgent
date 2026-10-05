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

        // ============================================================
        // 考勤分析（0.8.5）：从模板对象结构还原「员工 → 日期 → 打卡段」。
        //
        // 归属规则（实样验证）：
        //   名字单元格 = L==376 且纯中文 2-4 字（同位重叠时取流序靠后者=打印在上层）；
        //   日期行    = 一行内 ≥3 个「DD 星期」形如 "01 二" 的单元格，归属其上方最近的名字；
        //   打卡段    = 同一列 L、位于该日期行之下、下一日期行之前的对象文本，
        //               形如 "07:33-11:40"（完整）/"07:47-"（缺下班）/"-07:27"（缺上班，
        //               晚班跨零点下班）/ "-"（无打卡）；"晚转白/白转晚/常白班" 是班次标记。
        // 工时口径（确定性部分）：完整段 (end-start)，end<=start 视为跨零点 +24h；
        //   半段不计工时、列入"打卡不完整"；工作日（一~五）全天无打卡 = 缺勤/请假候选——
        //   两者无法从打卡数据区分，须由 agent 向用户确认口径（反问铁律）。
        // ============================================================
        public class AttDay
        {
            public int Day = 0;
            public string Weekday = "";
            public List<string> Segments = new List<string>();   // 原始段文本
            public double Hours = 0;                              // 完整段工时
            public bool HasIncomplete = false;                    // 有半段打卡
        }

        public class AttEmp
        {
            public string Name = "";
            public string Shift = "";
            public List<AttDay> Days = new List<AttDay>();
            public double TotalHours = 0;
            public int DaysWorked = 0;
            public List<string> NoPunchWorkdays = new List<string>();   // 工作日无任何打卡（缺勤/请假候选）
            public List<string> Incomplete = new List<string>();         // 打卡不完整明细
        }

        static bool IsDateLabel(string t)
        {
            // "01 二" / "9"（个别日期行单元格只剩数字）
            string s = (t ?? "").Trim();
            if (s.Length == 0) return false;
            int sp = s.IndexOf(' ');
            string num = sp > 0 ? s.Substring(0, sp) : s;
            if (num.Length == 0 || num.Length > 2) return false;
            for (int i = 0; i < num.Length; i++) { if (num[i] < '0' || num[i] > '9') return false; }
            int v = int.Parse(num);
            if (sp > 0)
            {
                string wd = s.Substring(sp + 1).Trim();
                if (wd.Length == 1 && "一二三四五六日".IndexOf(wd) >= 0) return v >= 1 && v <= 31;
                return false;
            }
            return v >= 1 && v <= 31;
        }

        static bool IsTimeSeg(string t)
        {
            // "HH:MM-HH:MM" / "HH:MM-" / "-HH:MM" / "-"（两侧可带空格）
            string s = (t ?? "").Trim();
            if (s.Length == 0) return false;
            if (s == "-") return true;
            int dash = s.IndexOf('-');
            if (dash < 0) return false;
            string a = s.Substring(0, dash).Trim();
            string b = s.Substring(dash + 1).Trim();
            if (a.Length == 0 && b.Length == 0) return true;
            if (a.Length > 0 && !IsHHMM(a)) return false;
            if (b.Length > 0 && !IsHHMM(b)) return false;
            return true;
        }

        static bool IsHHMM(string s)
        {
            if (s == null || s.Length != 5 || s[2] != ':') return false;
            for (int i = 0; i < 5; i++)
            {
                if (i == 2) continue;
                if (s[i] < '0' || s[i] > '9') return false;
            }
            int h = int.Parse(s.Substring(0, 2));
            int m = int.Parse(s.Substring(3, 2));
            return h <= 23 && m <= 59;
        }

        // 段时长（小时）；-1 = 无法计算（半段）
        static double SegHours(string seg)
        {
            string s = (seg ?? "").Trim();
            if (s == "-") return -1;
            int dash = s.IndexOf('-');
            if (dash < 0) return -1;
            string a = s.Substring(0, dash).Trim();
            string b = s.Substring(dash + 1).Trim();
            if (a.Length == 0 || b.Length == 0) return -1;
            if (!IsHHMM(a) || !IsHHMM(b)) return -1;
            double ta = int.Parse(a.Substring(0, 2)) + int.Parse(a.Substring(3, 2)) / 60.0;
            double tb = int.Parse(b.Substring(0, 2)) + int.Parse(b.Substring(3, 2)) / 60.0;
            if (tb <= ta) tb += 24;   // 晚班跨零点
            double d = tb - ta;
            if (d > 16) return -1;    // 超过 16 小时的"段"视为解析异常，不计
            return d;
        }

        static bool IsShiftMark(string t)
        {
            string s = (t ?? "").Trim();
            return s == "常白班" || s == "白转晚" || s == "晚转白";
        }

        static List<AttEmp> ParseAttendance(List<FrpObj> objs, out string period)
        {
            period = "";
            List<FrpObj> withText = new List<FrpObj>();
            for (int i = 0; i < objs.Count; i++)
            {
                string ct = CleanCell(objs[i].Text);
                if (ct.Length > 0) { objs[i].Text = ct; withText.Add(objs[i]); }
            }
            // 考期
            for (int i = 0; i < withText.Count; i++)
            {
                string t = withText[i].Text;
                int p = t.IndexOf("--");
                if (p > 4 && t.Length > p + 2)
                {
                    string a = t.Substring(0, p).Trim();
                    string b2 = t.Substring(p + 2).Trim();
                    if (a.Length >= 8 && b2.Length >= 8 && a[4] == '-' && b2[4] == '-') { period = a + " ~ " + b2; break; }
                }
            }
            // 名字单元格（L=376 纯中文 2-4 字；同位重叠取流序靠后者=打印在上层）。
            // ★ 模板存在同位叠名的遗留对象（实样：卢忠华/杨春胜、何加纯/凌炜彬、周少宁/莫达华、杨开荣/樊燕芳），
            //   两名都真实存在于名册时无法从几何判断谁可见——取靠后者并在报告标注，由用户核对。
            List<FrpObj> names = new List<FrpObj>();
            for (int i = 0; i < withText.Count; i++)
            {
                FrpObj o = withText[i];
                if (o.L != 376) continue;
                string t = o.Text.Trim();
                if (t.Length < 2 || t.Length > 4) continue;
                bool cjk = true;
                for (int k = 0; k < t.Length; k++) { if (t[k] < 0x4e00 || t[k] > 0x9fff) { cjk = false; break; } }
                if (!cjk) continue;
                bool dup = false;
                for (int k = 0; k < names.Count; k++)
                {
                    if (names[k].T == o.T) { names[k] = o; dup = true; break; }   // 同位重叠：后者在上层
                }
                if (!dup) names.Add(o);
            }
            StringBuilder overlapped = new StringBuilder();
            {
                Dictionary<int, string> seenT = new Dictionary<int, string>();
                for (int i = 0; i < withText.Count; i++)
                {
                    FrpObj o = withText[i];
                    if (o.L != 376 || o.Text.Trim().Length < 2 || o.Text.Trim().Length > 4) continue;
                    bool cjk2 = true;
                    string t2 = o.Text.Trim();
                    for (int k = 0; k < t2.Length; k++) { if (t2[k] < 0x4e00 || t2[k] > 0x9fff) { cjk2 = false; break; } }
                    if (!cjk2) continue;
                    if (seenT.ContainsKey(o.T) && seenT[o.T] != t2)
                        overlapped.Append(seenT[o.T]).Append("/").Append(t2).Append(" ");
                    else if (!seenT.ContainsKey(o.T)) seenT[o.T] = t2;
                }
            }

            // 员工登记（T → 员工）
            Dictionary<int, AttEmp> empByT = new Dictionary<int, AttEmp>();
            List<AttEmp> emps = new List<AttEmp>();
            foreach (FrpObj nm in names)
            {
                AttEmp e = new AttEmp();
                e.Name = nm.Text.Trim();
                empByT[nm.T] = e;
                emps.Add(e);
            }

            // 全局按 (T,L) 走一遍，模拟读表人：
            //   日期行（≥3 个「DD 星期」）→ 归属当前员工（其上方最近的名字），并记下列→天号；
            //   打卡段 → 归属其上方最近的日期行（而不是最近的名字——名字行会与上一员工的时间行重叠）；
            //   班次标记（常白班/白转晚/晚转白）→ 当前员工。
            withText.Sort(delegate(FrpObj a, FrpObj c)
            {
                if (a.T != c.T) return a.T - c.T;
                return a.L - c.L;
            });

            AttEmp curEmp = null;
            AttEmp rowEmp = null;                 // 当前日期行所属员工
            Dictionary<int, int> colDay = null;   // 当前日期行：列 L → 天号
            Dictionary<int, string> colWd = null;
            Dictionary<string, AttDay> acc = new Dictionary<string, AttDay>();   // (员工,天号) → 日
            int layerT = -1;
            List<FrpObj> layer = new List<FrpObj>();
            Action flushLayer = delegate
            {
                if (layer.Count == 0) return;
                int cnt = 0;
                foreach (FrpObj o in layer) { if (IsDateLabel(o.Text)) cnt++; }
                if (cnt >= 3)
                {
                    // 新日期行
                    rowEmp = curEmp;
                    colDay = new Dictionary<int, int>();
                    colWd = new Dictionary<int, string>();
                    foreach (FrpObj o in layer)
                    {
                        if (!IsDateLabel(o.Text)) continue;
                        string s2 = o.Text.Trim();
                        int sp2 = s2.IndexOf(' ');
                        int day = int.Parse(sp2 > 0 ? s2.Substring(0, sp2) : s2);
                        if (!colDay.ContainsKey(o.L)) { colDay[o.L] = day; colWd[o.L] = sp2 > 0 ? s2.Substring(sp2 + 1).Trim() : ""; }
                    }
                }
                else
                {
                    // 打卡段层：归当前日期行
                    if (rowEmp != null && colDay != null)
                    {
                        foreach (FrpObj o in layer)
                        {
                            if (!colDay.ContainsKey(o.L)) continue;
                            if (IsDateLabel(o.Text)) continue;
                            string[] lines = o.Text.Split(new char[] { '\r', '\n' });
                            for (int li = 0; li < lines.Length; li++)
                            {
                                string seg = lines[li].Trim();
                                if (seg.Length == 0) continue;
                                if (IsShiftMark(seg)) { if (rowEmp.Shift.Length == 0) rowEmp.Shift = seg; continue; }
                                if (!IsTimeSeg(seg)) continue;
                                int day = colDay[o.L];
                                string key = rowEmp.Name + "#" + day;
                                if (!acc.ContainsKey(key))
                                {
                                    AttDay ad = new AttDay();
                                    ad.Day = day; ad.Weekday = colWd[o.L];
                                    acc[key] = ad;
                                }
                                if (acc[key].Segments.Count < 4) acc[key].Segments.Add(seg);
                            }
                        }
                    }
                }
                layer = new List<FrpObj>();
            };

            for (int i = 0; i < withText.Count; i++)
            {
                FrpObj o = withText[i];
                if (o.T != layerT) { flushLayer(); layerT = o.T; }
                // 名字/班次在进入层处理前先结算（名字切换当前员工）
                if (o.L == 376)
                {
                    flushLayer();
                    if (empByT.ContainsKey(o.T)) curEmp = empByT[o.T];
                    continue;
                }
                if (IsShiftMark(o.Text)) { flushLayer(); if (curEmp != null && curEmp.Shift.Length == 0) curEmp.Shift = o.Text; continue; }
                layer.Add(o);
            }
            flushLayer();

            // 归集：acc → 各员工 Days（按天号排序；同员工同天号来自多个日期行的已在 key 里合并）
            foreach (AttEmp e in emps)
            {
                List<int> ds = new List<int>();
                Dictionary<int, AttDay> byDay = new Dictionary<int, AttDay>();
                foreach (KeyValuePair<string, AttDay> kv in acc)
                {
                    int hz = kv.Key.IndexOf('#');
                    if (kv.Key.Substring(0, hz) != e.Name) continue;
                    int d = int.Parse(kv.Key.Substring(hz + 1));
                    if (!byDay.ContainsKey(d)) { byDay[d] = kv.Value; ds.Add(d); }
                }
                ds.Sort();
                foreach (int d in ds) e.Days.Add(byDay[d]);
            }
            return emps;
        }

        // 考勤分析报告（agent 转述；请假/缺勤口径由 agent 向用户确认后自行统计）
        public static string AnalyzeAttendance(string path, out string err)
        {
            byte[] b;
            err = null;
            try { b = File.ReadAllBytes(path); }
            catch (Exception ex) { err = "读取失败: " + ex.Message; return ""; }
            List<FrpObj> objs = ParseObjects(b);
            if (objs.Count == 0) { err = "不是有效的 frp 模板"; return ""; }
            string period;
            List<AttEmp> emps = ParseAttendance(objs, out period);
            if (emps.Count == 0) { err = "未识别到员工考勤区块"; return ""; }

            StringBuilder sb = new StringBuilder();
            sb.Append("考勤解析结果（frp 打印模板）");
            if (period.Length > 0) sb.Append("　考期：" + period);
            sb.Append("\n口径说明：工时=完整打卡段之和（晚班跨零点自动 +24h）；\"HH:MM-\"或\"-HH:MM\"=打卡不完整，不计工时；");
            sb.Append("工作日（一~五）全天无打卡列为「无打卡工作日」——缺勤还是请假无法从打卡数据区分，需向用户确认口径后统计。\n");
            int ti2 = 0;
            foreach (AttEmp e in emps)
            {
                e.TotalHours = 0; e.DaysWorked = 0; e.NoPunchWorkdays = new List<string>(); e.Incomplete = new List<string>();
                foreach (AttDay d in e.Days)
                {
                    double dayH = 0;
                    foreach (string s in d.Segments)
                    {
                        double h = SegHours(s);
                        if (h >= 0) dayH += h;
                        else if (s != "-") d.HasIncomplete = true;
                    }
                    d.Hours = Math.Round(dayH, 2);
                    if (dayH > 0) e.DaysWorked++;
                    string tag = d.Day + "日(" + (d.Weekday.Length == 0 ? "?" : d.Weekday) + ")";
                    if (d.Segments.Count == 0 || (d.Segments.Count == 1 && d.Segments[0] == "-"))
                    {
                        if (d.Weekday != "六" && d.Weekday != "日") e.NoPunchWorkdays.Add(tag);
                    }
                    if (d.HasIncomplete) e.Incomplete.Add(tag + "[" + string.Join("/", d.Segments.ToArray()) + "]");
                    e.TotalHours += dayH;
                }
                ti2++;
                sb.Append("\n【" + e.Name + "】" + (e.Shift.Length > 0 ? "（" + e.Shift + "）" : "")
                    + " 出勤 " + e.DaysWorked + " 天，工时 " + Math.Round(e.TotalHours, 2) + " 小时");
                List<string> np = new List<string>();
                for (int k = 0; k < e.NoPunchWorkdays.Count; k++) { if (!np.Contains(e.NoPunchWorkdays[k])) np.Add(e.NoPunchWorkdays[k]); }
                List<string> ic = new List<string>();
                for (int k = 0; k < e.Incomplete.Count; k++) { if (!ic.Contains(e.Incomplete[k])) ic.Add(e.Incomplete[k]); }
                if (np.Count > 0)
                    sb.Append("；无打卡工作日 " + np.Count + " 天（" + string.Join("、", np.ToArray()) + "）");
                if (ic.Count > 0)
                    sb.Append("；打卡不完整 " + ic.Count + " 处（" + string.Join("、", ic.ToArray()) + "）");
                // 每日明细（供 agent 按用户口径统计请假/缺勤）
                StringBuilder det = new StringBuilder();
                foreach (AttDay d in e.Days)
                {
                    det.Append(d.Day + "日" + (d.Weekday.Length == 0 ? "" : d.Weekday) + ":"
                        + (d.Segments.Count == 0 ? "-" : string.Join("/", d.Segments.ToArray())) + " ");
                }
                if (det.Length > 0) sb.Append("\n  明细：" + det.ToString().TrimEnd());
            }
            sb.Append("\n（共 " + ti2 + " 名员工。请假次数/缺勤扣时需用户口径：无打卡工作日算缺勤还是请假？请假标记是什么？标准日工时多少？）");
            return sb.ToString();
        }
    }
}
