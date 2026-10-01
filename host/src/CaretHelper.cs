using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OfficeAgent.Host
{
    // 输入框自定义光标（2px 粗、与字号等高）与定位。
    //
    // 为什么不能直接用 EM_POSFROMCHAR 的 POINT* 形式：
    // 本项目输入框（多行 TextBox）实测该形式恒返回 -1，缓冲区保持 (0,0)；
    // 旧代码把失败当坐标 → SetCaretPos(0,-1) → 光标每次都闪回第一个字符之前。
    // 可用的是 wParam = 字符索引、返回值 LOWORD=X / HIWORD=Y 的打包形式，
    // 但索引 == 文本长度（行尾）时同样返回 -1，必须由调用方补偿最后一个字符的字宽。
    static class CaretHelper
    {
        [DllImport("user32.dll")]
        static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wp, IntPtr lp);
        [DllImport("user32.dll")]
        static extern bool CreateCaret(IntPtr hWnd, IntPtr hBitmap, int width, int height);
        [DllImport("user32.dll")]
        static extern bool ShowCaret(IntPtr hWnd);
        [DllImport("user32.dll")]
        static extern bool SetCaretPos(int x, int y);
        [DllImport("user32.dll")]
        static extern bool GetCaretPos(ref NativePoint p);

        [StructLayout(LayoutKind.Sequential)]
        public struct NativePoint { public int X; public int Y; }

        public const int EM_POSFROMCHAR = 0x00D6;

        // 取字符索引处的左缘坐标（输入框客户区坐标）。成功返回 true；
        // 索引越界 / 行尾 / 控件未就绪都返回 false，绝不输出 (-1,-1) 这类垃圾坐标。
        public static bool TryPosFromChar(IntPtr hWnd, int index, out int x, out int y)
        {
            x = 0; y = 0;
            if (hWnd == IntPtr.Zero || index < 0) return false;
            IntPtr r = SendMessage(hWnd, EM_POSFROMCHAR, (IntPtr)index, IntPtr.Zero);
            long v = r.ToInt64();
            if (v < 0) return false;
            int px = (short)(v & 0xFFFF);
            int py = (short)((v >> 16) & 0xFFFF);
            if (px < 0 || py < 0) return false;
            x = px; y = py;
            return true;
        }

        // 计算插入点应处的坐标：行尾取前一字符左缘 + 该字符字宽。
        // 返回 false 表示拿不到可信坐标，调用方应保持光标原位而不是写垃圾值。
        public static bool TryInsertionPoint(TextBox box, Font font, out int x, out int y)
        {
            x = 0; y = 0;
            if (box == null || !box.IsHandleCreated) return false;

            int idx = box.SelectionStart;
            if (idx < 0) idx = 0;
            if (idx > box.TextLength) idx = box.TextLength;

            if (idx > 0 && TryPosFromChar(box.Handle, idx - 1, out x, out y))
            {
                if (idx == box.TextLength)
                {
                    string last = box.Text.Substring(idx - 1, 1);
                    x += TextRenderer.MeasureText(last, font).Width;
                }
                else
                {
                    int nx, ny;
                    if (TryPosFromChar(box.Handle, idx, out nx, out ny)) { x = nx; y = ny; }
                }
                return true;
            }

            // 位置 0：取索引 0 的左缘；空文本时该调用也失败 → 退回控件原点，语义正确
            if (idx == 0)
            {
                if (TryPosFromChar(box.Handle, 0, out x, out y)) return true;
                x = 0; y = 0;
                return true;
            }
            return false;
        }

        // 重建加粗光标（仅在获得焦点/句柄创建时调用一次）。
        // 关键设计：编辑控件在 WM_SETFOCUS 里已把系统光标放到准确位置——先拍下它，
        // 重建后原位恢复，之后**完全交给控件自己管理**（打字/点击/IME/滚动都是控件原生维护）。
        // 不要在按键/文本变化时用 EM_POSFROMCHAR 自算位置再 SetCaretPos：真机上
        // （DPI 感知 + 中文 IME）自算坐标有偏差，光标看起来乱跳；无头测试测不出该差异。
        // 尺寸 3px 宽、比行高再高一点（上提 3px 居中）：用户反馈 2px 行高太细太小。
        public static void Build(TextBox box, Font font)
        {
            try
            {
                if (box == null || !box.IsHandleCreated) return;
                NativePoint cur = new NativePoint();
                bool have = GetCaretPos(ref cur);
                CreateCaret(box.Handle, IntPtr.Zero, 3, font.Height + 6);
                if (have && cur.X >= 0 && cur.Y >= 0)
                    SetCaretPos(cur.X, Math.Max(0, cur.Y - 3));
                else
                {
                    int x, y;
                    if (TryInsertionPoint(box, font, out x, out y)) SetCaretPos(x + 1, Math.Max(0, y - 3));
                }
                ShowCaret(box.Handle);
            }
            catch { }
        }

        // 移动到当前插入点；坐标不可信时原地不动
        public static void Move(TextBox box, Font font)
        {
            try
            {
                int x, y;
                if (!TryInsertionPoint(box, font, out x, out y)) return;
                SetCaretPos(x + 1, y);
            }
            catch { }
        }
    }
}
