// XlsxPreview —— 大表预览的**单遍扫描 + 有界保留**（低配优化）
//
// 问题（实测，开发机 12 核/16GB；目标画像 4GB Win7 老机只会更差）：
//   MainForm.ShowXlsx 原先走 XlsxBook.LoadGrid(0, 20000, 128)，一次性物化全部单元格：
//     20000 行 x 6 列   -> 906ms，托管堆 +41MB
//     20000 行 x 128 列 -> 3230ms，托管堆 +167MB
//   这只是 string[,] 数组本身；DataGridView 还要为每个单元格建 DataGridViewCell
//   （20000x128 = 256 万个对象，数百 MB 量级），且逐行 Rows.Add 全程在 UI 线程、未挂起布局。
//   在 4GB 老机上打开一张宽表 = OOM 或长时间假死。
//
// ★ 走过的弯路（留档，避免后来者重犯）：
//   第一版想做成"虚拟滚动 + 滑动窗口按需重扫"（DataGridView.VirtualMode +
//   滚动到未缓存块时重读文件）。实测**比原实现更糟**：
//     宽表取 120 行耗时 9375ms、堆 +158MB（原实现 3168ms / +167MB）。
//   根因是 xlsx 的本质：zip 压缩的变长 XML，**无法随机跳到第 N 行**，
//   要读第 N 行只能从头解压。于是"按需"退化成每滚动一次 O(N) 重扫，
//   且每行都要拷贝 maxCols(128) 长度的数组（哪怕该行只有 6 列有值）。
//   结论：**xlsx 加载大表做虚拟滚动是个伪需求**——除非先自建行偏移索引，
//   而那需要解压整个表，成本与直接读一遍相同。
//
// 最终方案：单遍流式扫描，**只保留前 previewRows 行**，同时如实报出总行数与实际列数。
//   内存从 O(总行数 x 请求列数) 降到 O(previewRows x 实际列数)，
//   且列数取"实际用到的"而非写死的 128——后者对 6 列的表是 20 倍浪费。
//   预览本来就是"看一眼开头"，如实标注"仅预览前 N 行 / 共 M 行"即可。
using System;
using System.Collections.Generic;

namespace OfficeAgent.Host
{
    public class XlsxPreview
    {
        // 保留给预览的最大行数。2000 行足够判断表结构，内存可控：
        // 2000 x 64 列 x 8B 引用 ≈ 1MB（外加字符串本身）。
        public const int DefaultPreviewRows = 2000;

        // 列数硬上限：再宽的表现在也没有预览价值（128 列已远超屏幕），
        // 且它是内存占用的乘数。需要看更多列的极端场景应走"转换后阅读"而非预览。
        public const int MaxPreviewCols = 64;

        public string[][] Rows = new string[0][];
        public int TotalRows;      // 实际总行数（如实告知用户，不假装只有预览这些）
        public int UsedCols;       // 实际使用到的列数
        public string Error = "";

        // 单遍扫描：保留前 maxRows 行，但**继续读完全表**以得到真实总行数。
        // 不能"读够 maxRows 就停"：那样滚动条/提示里的总行数会是错的，
        // 而"这个表到底多大"恰恰是会计判断要不要继续处理的关键信息。
        //
        // maxCols 传该表实际需要的列数上限即可；此处默认 MaxPreviewCols。
        public static XlsxPreview Load(string path, int sheetIndex, int maxRows, int maxCols)
        {
            XlsxPreview p = new XlsxPreview();
            if (maxRows < 1) maxRows = 1;
            if (maxCols < 1) maxCols = 1;

            XlsxBook book = null;
            try
            {
                book = XlsxBook.Open(path);
                if (book.Sheets.Count == 0) { p.Error = "工作簿里没有工作表"; return p; }
                if (sheetIndex < 0 || sheetIndex >= book.Sheets.Count) sheetIndex = 0;
                // 无缓存值的公式格：预览用短标记而不是公式原文。
                // 系统/ERP 生成的 xlsx 常写了公式却不写缓存值，照搬原文会把
                // "=IF($N6=...,_xlfn.TEXTJOIN(...))" 这种上百字符的串铺进网格，看起来就是乱码。
                book.UncachedFormulaMarker = "〔公式〕";

                List<string[]> kept = new List<string[]>();
                int count = 0;
                int used = 0;
                string err = book.StreamRows(sheetIndex, maxCols, delegate(string[] row)
                {
                    count++;
                    // 只按"最后一个非空单元格"推进 usedCols，与 LoadGrid 口径一致
                    for (int c = row.Length - 1; c >= 0; c--)
                    {
                        if (row[c] != null) { if (c + 1 > used) used = c + 1; break; }
                    }
                    if (kept.Count >= maxRows) return;   // 只保留前 maxRows 行，其余仅计数
                    // 必须拷贝：StreamRows 复用同一个 cells 缓冲并在行间 Clear
                    string[] copy = new string[row.Length];
                    Array.Copy(row, copy, row.Length);
                    kept.Add(copy);
                });

                if (err != null && err.Length > 0 && kept.Count == 0) p.Error = err;
                p.Rows = kept.ToArray();
                p.TotalRows = count;
                p.UsedCols = used;

                // ★ 合并单元格回填，必须与 LoadGrid 口径一致。
                //   漏掉这步的后果是实测抓出来的：dirty.xlsx 有跨列标题，
                //   LoadGrid 会把锚点值填满整个合并区域，而本方法不回填 →
                //   同一张表"旧预览"与"新预览"内容不同（差异 5 个单元格），
                //   表现为跨列表头只在第一列显示、其余列空白。
                //   这不是解析错误，是**功能遗漏**；一致性验证（逐单元格对比）才暴露出来——
                //   只看性能数字的话，这个 bug 会直接进生产。
                //   注：Merges 由 StreamRows 在扫描时填充到 Sheets[sheetIndex].Merges，
                //   而**回填只需要前 N 行**：合并区域若整体在保留范围之外，回填自然跳过
                //   （FillMerged 内有 r >= rows.Length 的边界处理）。
                XlsxBook.FillMerged(p.Rows, book.Sheets[sheetIndex].Merges);
                for (int i = 0; i < p.Rows.Length; i++)
                {
                    // 回填后列数可能变宽（合并区域跨越了原本"未使用"的列），重新核算
                    string[] row = p.Rows[i];
                    for (int c = row.Length - 1; c >= 0; c--)
                    {
                        if (row[c] != null) { if (c + 1 > p.UsedCols) p.UsedCols = c + 1; break; }
                    }
                }
                if (p.UsedCols < 1) p.UsedCols = 1;
            }
            catch (Exception ex)
            {
                p.Error = ex.Message == null ? ex.GetType().Name : ex.Message;
                p.Error = "预览失败: " + p.Error;
            }
            finally
            {
                if (book != null) { try { book.Dispose(); } catch { } }
            }
            return p;
        }

        // 预览提示文本（含"仅预览前 N 行 / 共 M 行"的如实告知）。
        // 单独成方法以便自检直测，不必启动 UI。
        public string Describe(int sheetCount, string sheetNames)
        {
            if (Error.Length > 0) return Error;
            string extra = TotalRows > Rows.Length
                ? "（仅预览前 " + Rows.Length + " 行 / 共 " + TotalRows + " 行）"
                : "";
            string cols = "";
            if (UsedCols > MaxPreviewCols)
                cols = "（列数超出预览上限 " + MaxPreviewCols + "，仅显示前 " + MaxPreviewCols + " 列）";
            return "工作表: " + sheetNames + "  " + extra + cols;
        }
    }
}
