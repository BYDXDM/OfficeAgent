using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OfficeAgent.Host
{
    // PDF 文本读取（供 agent read_text_file 工具直接使用）。
    // 复用 pdfium 逐字符框基线聚类的行提取，不依赖任何外部阅读器、不改动原文件。
    static class PdfTextReader
    {
        public const int MaxChars = 200000;   // 单次返回上限，超出截断（避免撑爆模型上下文）
        public const int MaxPages = 200;      // 单次读取页数上限

        // 返回可读文本；失败时 err 非空、返回值不含部分结果
        public static string Read(string path, string root, out string err)
        {
            err = null;
            try
            {
                Pdfium.EnsureInit(root);
            }
            catch (Exception ex)
            {
                err = "pdfium 初始化失败: " + ex.Message;
                return null;
            }

            try
            {
                using (Pdfium.PdfDoc doc = Pdfium.PdfDoc.Open(path, out err))
                {
                    if (doc == null) return null;

                    int pages = doc.PageCount;
                    int limit = Math.Min(pages, MaxPages);
                    StringBuilder sb = new StringBuilder();
                    int totalChars = 0;
                    bool truncated = false;

                    for (int p = 0; p < limit; p++)
                    {
                        string perr;
                        List<string> lines = doc.ExtractLines(p, out perr);
                        if (lines == null)
                        {
                            sb.Append("\n[第 ").Append(p + 1).Append(" 页文本提取失败: ")
                              .Append(perr == null ? "未知原因" : perr).Append("]\n");
                            continue;
                        }
                        sb.Append("\n---- 第 ").Append(p + 1).Append(" 页 ----\n");
                        foreach (string ln in lines)
                        {
                            sb.Append(ln).Append('\n');
                            totalChars += ln.Length + 1;
                        }
                        if (totalChars >= MaxChars) { truncated = true; break; }
                    }

                    string text = sb.ToString();
                    if (text.Length > MaxChars) { text = text.Substring(0, MaxChars); truncated = true; }

                    // 图片型（扫描件）PDF 没有文本层：明确告知，避免模型以为文件是空的
                    bool noTextLayer = text.Trim().Length == 0;
                    StringBuilder head = new StringBuilder();
                    head.Append("PDF ").Append(path)
                        .Append("（共 ").Append(pages).Append(" 页");
                    if (limit < pages) head.Append("，本次读取前 ").Append(limit).Append(" 页");
                    head.Append("）");
                    if (noTextLayer)
                    {
                        head.Append("\n未提取到任何文本：该 PDF 可能是扫描件/纯图片，没有文字层。")
                            .Append("可改用 convert_document 转成图片后另行处理，或在预览页查看。");
                        return head.ToString();
                    }
                    head.Append("内容：\n").Append(text);
                    if (truncated)
                        head.Append("\n…[内容过长，已截断；如需其余部分请分段读取]…");
                    return head.ToString();
                }
            }
            catch (Exception ex)
            {
                err = "读取 PDF 失败: " + ex.Message;
                return null;
            }
        }
    }
}
