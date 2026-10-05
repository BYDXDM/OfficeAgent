// TextLayout —— 文本折行与字符命中（0.9.6）
//
// 用途：让对话气泡里的文字**可被鼠标选中**（用户要求）。
// 为什么单独成文件：这段是**纯逻辑**（只依赖 TextRenderer 测量），可以无头自测——
// 「选中」这种交互一旦几何算错，表现就是高亮和文字对不上（错位），
// 所以几何必须能被断言，而不是靠肉眼。
//
// 口径说明：绘制用的是 TextRenderer + TextFormatFlags.WordBreak。
//   本实现按"能放下的最长前缀"逐行折行；若有空格，优先退到空格后（近似 WordBreak
//   的按词折行）。中文无空格，逐字折行与 WordBreak 一致。
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace OfficeAgent.Host
{
    public static class TextLayout
    {
        public class Line
        {
            public int Start;   // 在原文中的起始下标
            public int Len;     // 该行字符数（不含换行符）
            public int End { get { return Start + Len; } }
        }

        // 把 text 按 width 折行；返回每行的字符区间（行数 ≥ 1）
        public static List<Line> ComputeLines(string text, Font font, int width)
        {
            List<Line> lines = new List<Line>();
            if (text == null) text = "";
            if (width < 10) width = 10;
            if (text.Length == 0) { lines.Add(new Line()); return lines; }

            int pos = 0, guard = 0;
            while (guard++ < 5000)
            {
                if (pos >= text.Length) break;
                int nl = text.IndexOf('\n', pos);
                int hardEnd = nl < 0 ? text.Length : nl;

                // 在这一"硬行"内二分：还能放下的最长前缀
                int lo = pos, hi = hardEnd, fit = -1;
                while (lo <= hi)
                {
                    int mid = (lo + hi) / 2;
                    int len = mid - pos;
                    bool ok;
                    if (len <= 0) ok = true;
                    else ok = Width(text.Substring(pos, len), font) <= width;
                    if (ok) { fit = mid; lo = mid + 1; }
                    else hi = mid - 1;
                }
                if (fit < pos) fit = pos;
                if (fit == pos) fit = Math.Min(pos + 1, hardEnd);   // 单字都放不下也要前进，防死循环

                // 近似 WordBreak：没到行尾且不是空格断点，就退到最近的空格之后
                if (fit < hardEnd && fit > pos)
                {
                    char at = text[fit];
                    if (at != ' ')
                    {
                        int sp = text.LastIndexOf(' ', fit - 1, fit - pos);
                        if (sp >= pos) fit = sp + 1;
                    }
                }

                lines.Add(new Line { Start = pos, Len = fit - pos });
                pos = fit;
                if (nl >= 0 && pos == nl) pos = nl + 1;   // 跳过换行符本身
            }
            if (lines.Count == 0) lines.Add(new Line());
            return lines;
        }

        // 行内命中：给定行内 x 偏移（相对行左），返回该行内应插入的字符下标（0..len）
        // 采用"取最近边界"：超过半个字符宽就算下一个字符，拖选手感更自然。
        public static int CharInLine(string text, int lineStart, int lineLen, Font font, int x)
        {
            if (text == null || lineLen <= 0) return 0;
            if (x <= 0) return 0;
            int prevW = 0;
            for (int i = 1; i <= lineLen; i++)
            {
                int w = Width(text.Substring(lineStart, i), font);
                if (x < w)
                {
                    // 落在第 i 个字符内：靠左还是靠右
                    return (x - prevW) * 2 < (w - prevW) ? i - 1 : i;
                }
                prevW = w;
            }
            return lineLen;
        }

        public static int Width(string s, Font font)
        {
            if (s == null || s.Length == 0) return 0;
            try
            {
                return TextRenderer.MeasureText(s, font,
                    new Size(int.MaxValue, int.MaxValue), TextFormatFlags.NoPadding).Width;
            }
            catch { return 0; }
        }

        // 文本总行高（与绘制一致：每行 font.Height）
        public static int LineHeight(Font font)
        {
            try { return TextRenderer.MeasureText("测Mg", font, new Size(int.MaxValue, int.MaxValue), TextFormatFlags.NoPadding).Height; }
            catch { return 18; }
        }
    }
}
