// pdfium 封装 —— PDF 打开/渲染为位图（Chromium 109 世代构建，Win7 可用）
// dll 定位：优先主程序同目录，其次安装根 bin\。找不到时降级为"无 PDF 预览"。
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace OfficeAgent.Host
{
    public static class Pdfium
    {
        const int FPDF_ANNOT = 1;

        [DllImport("pdfium.dll")] static extern void FPDF_InitLibrary();
        [DllImport("pdfium.dll")] static extern void FPDF_DestroyLibrary();
        // 不用 FPDF_LoadDocument（Ansi 路径在中文系统=GBK，pdfium 要求 UTF-8，中文路径必失败）；
        // 改用内存加载，彻底绕开路径编码问题
        [DllImport("pdfium.dll")] static extern IntPtr FPDF_LoadMemDocument(byte[] data, int size, string password);
        [DllImport("pdfium.dll")] static extern void FPDF_CloseDocument(IntPtr doc);
        [DllImport("pdfium.dll")] static extern int FPDF_GetPageCount(IntPtr doc);
        [DllImport("pdfium.dll")] static extern IntPtr FPDF_LoadPage(IntPtr doc, int index);
        [DllImport("pdfium.dll")] static extern void FPDF_ClosePage(IntPtr page);
        [DllImport("pdfium.dll")] static extern float FPDF_GetPageWidthF(IntPtr page);
        [DllImport("pdfium.dll")] static extern float FPDF_GetPageHeightF(IntPtr page);
        [DllImport("pdfium.dll")] static extern IntPtr FPDFBitmap_Create(int w, int h, int alpha);
        [DllImport("pdfium.dll")] static extern void FPDFBitmap_FillRect(IntPtr bmp, int left, int top, int w, int h, uint color);
        [DllImport("pdfium.dll")] static extern IntPtr FPDFBitmap_GetBuffer(IntPtr bmp);
        [DllImport("pdfium.dll")] static extern void FPDFBitmap_Destroy(IntPtr bmp);
        [DllImport("pdfium.dll")] static extern void FPDF_RenderPageBitmap(IntPtr bmp, IntPtr page, int sx, int sy, int sizeX, int sizeY, int rotate, int flags);

        // ---------- 文本层（发票提取用，M2）----------
        // 本构建（Chromium109 世代 pinned dll）中加载入口名为 FPDFText_LoadPage（同义于新版 FPDFText_LoadTextPage）
        [DllImport("pdfium.dll", EntryPoint = "FPDFText_LoadPage")] static extern IntPtr FPDFText_LoadTextPage(IntPtr page);
        [DllImport("pdfium.dll")] static extern void FPDFText_ClosePage(IntPtr textPage);
        [DllImport("pdfium.dll")] static extern int FPDFText_CountChars(IntPtr textPage);
        [DllImport("pdfium.dll")] static extern uint FPDFText_GetUnicode(IntPtr textPage, int index);
        [DllImport("pdfium.dll")] static extern int FPDFText_GetCharBox(IntPtr textPage, int index,
            out double left, out double right, out double bottom, out double top);

        static bool inited = false;

        public static string FindDll(string root)
        {
            string[] candidates = new string[] {
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "pdfium.dll"),
                Path.Combine(Path.Combine(root, "bin"), "pdfium.dll"),
                Path.Combine(Path.Combine(root, "host"), "pdfium.dll")
            };
            foreach (string p in candidates)
            {
                try { if (File.Exists(p)) return p; } catch { }
            }
            return null;
        }

        public static void EnsureInit(string root)
        {
            if (inited) return;
            string dll = FindDll(root);
            if (dll == null) throw new DllNotFoundException("未找到 pdfium.dll（" + AppDomain.CurrentDomain.BaseDirectory + " 或 bin\\）");
            // 指定绝对目录加载（避免 PATH 问题）
            IntPtr h = LoadLibrary(dll);
            if (h == IntPtr.Zero) throw new DllNotFoundException("pdfium.dll 加载失败: " + dll);
            FPDF_InitLibrary();
            inited = true;
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        static extern IntPtr LoadLibrary(string path);

        public class PdfDoc : IDisposable
        {
            IntPtr doc;
            public int PageCount = 0;

            PdfDoc(IntPtr doc) { this.doc = doc; PageCount = FPDF_GetPageCount(doc); }

            public static PdfDoc Open(string path, out string err)
            {
                err = null;
                byte[] data;
                try { data = File.ReadAllBytes(path); }
                catch (Exception ex) { err = "读取失败: " + ex.Message; return null; }
                IntPtr d = FPDF_LoadMemDocument(data, data.Length, null);
                if (d == IntPtr.Zero)
                {
                    err = "PDF 打开失败（可能损坏或加密）";
                    return null;
                }
                return new PdfDoc(d);
            }

            public Bitmap RenderPage(int index, double dpi)
            {
                IntPtr page = FPDF_LoadPage(doc, index);
                if (page == IntPtr.Zero) throw new InvalidOperationException("第 " + (index + 1) + " 页加载失败");
                try
                {
                    double wpt = FPDF_GetPageWidthF(page);
                    double hpt = FPDF_GetPageHeightF(page);
                    if (wpt <= 0) wpt = 595;   // A4 兜底
                    if (hpt <= 0) hpt = 842;
                    int w = Math.Max(1, (int)(wpt / 72.0 * dpi));
                    int h = Math.Max(1, (int)(hpt / 72.0 * dpi));
                    IntPtr bmp = FPDFBitmap_Create(w, h, 0);
                    if (bmp == IntPtr.Zero) throw new InvalidOperationException("位图创建失败 " + w + "x" + h);
                    try
                    {
                        FPDFBitmap_FillRect(bmp, 0, 0, w, h, 0xFFFFFFFFu);
                        FPDF_RenderPageBitmap(bmp, page, 0, 0, w, h, 0, FPDF_ANNOT);
                        IntPtr buf = FPDFBitmap_GetBuffer(bmp);
                        Bitmap bm = new Bitmap(w, h, PixelFormat.Format32bppArgb);
                        BitmapData bd = bm.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb);
                        try
                        {
                            int rowBytes = w * 4;
                            byte[] row = new byte[rowBytes];
                            for (int y = 0; y < h; y++)
                            {
                                Marshal.Copy(new IntPtr(buf.ToInt64() + (long)y * rowBytes), row, 0, rowBytes);
                                Marshal.Copy(row, 0, new IntPtr(bd.Scan0.ToInt64() + (long)y * bd.Stride), rowBytes);
                            }
                        }
                        finally { bm.UnlockBits(bd); }
                        return bm;
                    }
                    finally { FPDFBitmap_Destroy(bmp); }
                }
                finally { FPDF_ClosePage(page); }
            }

            public void Dispose()
            {
                if (doc != IntPtr.Zero) { FPDF_CloseDocument(doc); doc = IntPtr.Zero; }
            }

            // 提取第 index 页的文本行：按基线聚类成视觉行、行内按 x 排序，词间加空格。
            // 返回 null = 失败（err 给出原因）。空页返回空列表。
            public List<string> ExtractLines(int index, out string err)
            {
                err = null;
                IntPtr page = FPDF_LoadPage(doc, index);
                if (page == IntPtr.Zero) { err = "第 " + (index + 1) + " 页加载失败"; return null; }
                try
                {
                    IntPtr tp = FPDFText_LoadTextPage(page);
                    if (tp == IntPtr.Zero) { err = "文本层加载失败（可能为扫描件）"; return null; }
                    try
                    {
                        int count = FPDFText_CountChars(tp);
                        List<string> lines = new List<string>();
                        if (count <= 0) return lines;
                        // 收集可见字符（box 有效）
                        List<double> lefts = new List<double>(), rights = new List<double>();
                        List<double> tops = new List<double>(), bottoms = new List<double>();
                        List<int> codes = new List<int>();
                        for (int i = 0; i < count; i++)
                        {
                            double l, rt, b, t;
                            int ok = FPDFText_GetCharBox(tp, i, out l, out rt, out b, out t);
                            uint u = FPDFText_GetUnicode(tp, i);
                            if (ok == 0 || u == 0 || u == 0xFFFD)
                            {
                                // 无框字符（如生成器未给 box）：仍保留文本
                                if (u != 0 && u != 0xFFFD)
                                {
                                    codes.Add((int)u);
                                    lefts.Add(-1); rights.Add(-1); tops.Add(-1); bottoms.Add(-1);
                                }
                                continue;
                            }
                            codes.Add((int)u);
                            lefts.Add(l); rights.Add(rt); tops.Add(t); bottoms.Add(b);
                        }
                        // 基线聚类
                        int start = 0;
                        while (start < codes.Count)
                        {
                            double baseTop = tops[start], baseBottom = bottoms[start];
                            int end = start + 1;
                            while (end < codes.Count)
                            {
                                double ct = tops[end], cb = bottoms[end];
                                if (ct < 0 || cb < 0) break;
                                double mid = (baseTop + baseBottom) / 2.0, cm = (ct + cb) / 2.0;
                                if (Math.Abs(mid - cm) > Math.Max(3.0, Math.Abs(baseTop - baseBottom) / 2.0)) break;
                                end++;
                            }
                            // 行内拼接：x 升序（字符序通常即阅读序），间隙 > 平均字宽 0.45 插空格
                            StringBuilder sb = new StringBuilder();
                            double prevRight = double.MinValue;
                            double widthSum = 0;
                            int widthN = 0;
                            for (int i = start; i < end; i++)
                            {
                                if (lefts[i] >= 0) { widthSum += rights[i] - lefts[i]; widthN++; }
                            }
                            double avgW = widthN > 0 ? widthSum / widthN : 0;
                            for (int i = start; i < end; i++)
                            {
                                char c = (char)codes[i];
                                if (codes[i] > 0xFFFF) c = '?';
                                if (lefts[i] >= 0 && prevRight > double.MinValue &&
                                    lefts[i] - prevRight > avgW * 0.45 && c != ' ') sb.Append(' ');
                                sb.Append(c);
                                if (rights[i] >= 0) prevRight = rights[i];
                            }
                            string line = sb.ToString().Trim();
                            if (line.Length > 0) lines.Add(line);
                            start = end;
                        }
                        return lines;
                    }
                    finally { FPDFText_ClosePage(tp); }
                }
                finally { FPDF_ClosePage(page); }
            }
        }
    }
}
