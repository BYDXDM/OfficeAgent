// ConvertEngine —— 三级降级转换总线（设计方案 §5.3）
//   ① MS Office COM（保真最高）  ② LibreOffice 7.6 headless（无 Office 兜底）  ③ 原生库直写（xlsx↔csv）
// 规则：原始文件只读；输出写 *_conv.<ext>；外部进程经 ToolRunner 白名单；COM 走 STA 线程 + 看门狗。
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using OfficeAgent.Core;

namespace OfficeAgent.Host
{
    public enum ConvTarget { Pdf, Csv, Xlsx }

    public class ConvertEngine
    {
        public string SofficePath = null;   // 探测到的 soffice.exe；null = 无 LO
        // 最近一次转换的**附加提示**（null=无）。用于"转换成功但需用户知晓"的情况——
        // 例如源表公式没有缓存值、CSV 里这些格被置空（否则会写成一墙公式文本，被当成乱码）。
        public string LastWarning = null;

        public static string FindSoffice(string root)
        {
            // 64 位系统上 32 位进程的 ProgramFiles 指向 (x86)，需补 ProgramW6432 覆盖 64 位 LO 安装
            string pf64 = Environment.GetEnvironmentVariable("ProgramW6432");
            List<string> candidates = new List<string>();
            candidates.Add(Path.Combine(Path.Combine(root, "lo76"), Path.Combine(Path.Combine("LibreOffice", "program"), "soffice.exe")));
            candidates.Add(Path.Combine(Path.Combine(root, "lo76"), Path.Combine("program", "soffice.exe")));
            if (pf64 != null && pf64.Length > 0)
                candidates.Add(Path.Combine(Path.Combine(pf64, "LibreOffice"), Path.Combine("program", "soffice.exe")));
            candidates.Add(Path.Combine(Path.Combine(Environment.GetEnvironmentVariable("ProgramFiles(x86)") ?? "", "LibreOffice"), Path.Combine("program", "soffice.exe")));
            candidates.Add(Path.Combine(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "LibreOffice"), Path.Combine("program", "soffice.exe")));
            foreach (string p in candidates)
            {
                try { if (File.Exists(p)) return p; } catch { }
            }
            return null;
        }

        public string OutputPathOf(string input, ConvTarget target)
        {
            string ext = target == ConvTarget.Pdf ? ".pdf" : (target == ConvTarget.Csv ? ".csv" : ".xlsx");
            string dir = Path.GetDirectoryName(input);
            string name = Path.GetFileNameWithoutExtension(input) + "_conv" + ext;
            return Path.Combine(dir, name);
        }

        // 返回 null=成功（outputPath 有效），否则错误信息
        public string Convert(string input, ConvTarget target, out string outputPath)
        {
            outputPath = null;
            LastWarning = null;   // 每次转换重置，避免把上一次的提示带给本次调用方
            if (!File.Exists(input)) return "输入文件不存在: " + input;
            input = Path.GetFullPath(input);   // COM/LO 均要求绝对路径
            string ext = (Path.GetExtension(input) ?? "").ToLowerInvariant();
            try
            {
                if (target == ConvTarget.Csv)
                {
                    if (ext == ".xlsx") return XlsxToCsv(input, out outputPath);
                    if (ext == ".csv") return "输入已是 CSV";
                    if (ext == ".xls")
                    {
                        string comErr = ConvertCom(input, "Excel", ConvTarget.Csv, out outputPath);
                        if (comErr == null) return null;
                        string loErr = ConvertLo(input, ConvTarget.Csv, out outputPath);
                        if (loErr == null) return null;
                        return "COM 失败（" + comErr + "），LibreOffice " + (SofficePath == null ? "未安装" : "也失败（" + loErr + "）");
                    }
                    return "暂不支持 " + ext + " → CSV";
                }
                if (target == ConvTarget.Xlsx)
                {
                    if (ext == ".csv") return CsvToXlsx(input, out outputPath);
                    if (ext == ".xls")
                    {
                        string comErr = ConvertCom(input, "Excel", ConvTarget.Xlsx, out outputPath);
                        if (comErr == null) return null;
                        string loErr = ConvertLo(input, ConvTarget.Xlsx, out outputPath);
                        if (loErr == null) return null;
                        return "COM 失败（" + comErr + "），LibreOffice " + (SofficePath == null ? "未安装" : "也失败（" + loErr + "）");
                    }
                    return "暂不支持 " + ext + " → XLSX";
                }
                if (target == ConvTarget.Pdf)
                {
                    switch (ext)
                    {
                        case ".xlsx":
                        case ".xls":
                        case ".docx":
                        case ".doc":
                        case ".pptx":
                        case ".ppt":
                            string kind = KindOf(ext);
                            // 跳过 COM 时给哨兵错误（不能是 null——null 被视为成功而 outputPath 未赋值）
                            string comErr = (kind == "Excel" && (excelPdfBroken || ExcelPdfFlagSet()))
                                ? "Excel 导出此前挂死，本进程直接走 LibreOffice" : null;
                            if (comErr == null) comErr = ConvertCom(input, kind, ConvTarget.Pdf, out outputPath);
                            if (comErr == null) return null;
                            if (comErr.IndexOf("COM 转换超时", StringComparison.Ordinal) >= 0 && kind == "Excel") MarkExcelPdfBroken();
                            string loErr = ConvertLo(input, ConvTarget.Pdf, out outputPath);
                            if (loErr == null) return null;
                            return "COM 失败（" + comErr + "），LibreOffice " + (SofficePath == null ? "未安装" : "也失败（" + loErr + "）");
                        default:
                            return "暂不支持 " + ext + " → PDF（M1 支持 xlsx/xls/doc/docx/ppt/pptx）";
                    }
                }
                return "未知转换目标";
            }
            catch (Exception ex) { return "转换异常: " + ex.Message; }
        }

        static string KindOf(string ext)
        {
            switch (ext)
            {
                case ".xlsx":
                case ".xls": return "Excel";
                case ".docx":
                case ".doc": return "Word";
                default: return "PowerPoint";
            }
        }

        // ---------- ① Office COM（反射调用，不依赖 PIA；STA 线程 + 看门狗）----------

        const int ComTimeoutSeconds = 300;
        // Excel→PDF 实测坑：隐藏会话下 ExportAsFixedFormat/SaveAs(57) 可能挂死（Word/PP 同会话正常，
        // PrintCommunication/IgnorePrintAreas/工作表级/全参数六种变体均挂）。对策：看门狗 90s 快败，
        // 且一次超时后本进程内 Excel→PDF 直接走 LO，不再反复烧 300s。
        static bool excelPdfBroken = false;

        // Excel→PDF 挂死标记跨进程持久化（CLI 每次新进程）：7 天后自动失效（打印机环境修复后恢复 COM 快路径）
        static string ExcelPdfFlagFile()
        {
            return Path.Combine(Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeAgent"),
                "excel-pdf-broken.txt");
        }

        static bool ExcelPdfFlagSet()
        {
            try
            {
                string f = ExcelPdfFlagFile();
                if (!File.Exists(f)) return false;
                DateTime t;
                if (!DateTime.TryParse(File.ReadAllText(f).Trim(), out t)) return false;
                if ((DateTime.Now - t).TotalDays > 7) { try { File.Delete(f); } catch { } return false; }
                return true;
            }
            catch { return false; }
        }

        static void MarkExcelPdfBroken()
        {
            try
            {
                string dir = Path.GetDirectoryName(ExcelPdfFlagFile());
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                File.WriteAllText(ExcelPdfFlagFile(), DateTime.Now.ToString("o"));
            }
            catch { }
        }

        int ComTimeout(string kind, ConvTarget target)
        {
            if (kind == "Excel" && target == ConvTarget.Pdf) return 90;
            return ComTimeoutSeconds;
        }

        string ConvertCom(string input, string kind, ConvTarget target, out string outputPath)
        {
            outputPath = OutputPathOf(input, target);
            Type appType = null;
            if (kind == "Excel") appType = Type.GetTypeFromProgID("Excel.Application");
            else if (kind == "Word") appType = Type.GetTypeFromProgID("Word.Application");
            else appType = Type.GetTypeFromProgID("PowerPoint.Application");
            if (appType == null) return kind + " COM 不可用（未装 Office）";
            if (File.Exists(outputPath)) { try { File.Delete(outputPath); } catch { } }

            string result = null, step = "创建 " + kind + " 应用";
            string outPathLocal = outputPath;   // out 参数不能进匿名方法，用局部变量桥接
            Thread worker = new Thread(new ThreadStart(delegate
            {
                object app = null, doc = null;
                try
                {
                    app = Activator.CreateInstance(appType);
                    Type at = app.GetType();
                    step = "DisplayAlerts";
                    if (kind != "PowerPoint") at.InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, app, new object[] { false });
                    step = "获取文档集合";
                    object docs = at.InvokeMember(kind == "Excel" ? "Workbooks" : (kind == "Word" ? "Documents" : "Presentations"),
                        BindingFlags.GetProperty, null, app, null);
                    step = "打开文件";
                    doc = docs.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, docs, new object[] { input });
                    Type dt = doc.GetType();
                    if (target == ConvTarget.Pdf)
                    {
                        step = "导出 PDF";
                        // 注意两家参数顺序相反：Excel=(Type, Filename)；Word=(OutputFileName, ExportFormat)
                        if (kind == "PowerPoint")
                            dt.InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, doc, new object[] { outPathLocal, 32 /*ppSaveAsPDF*/ });
                        else if (kind == "Word")
                            dt.InvokeMember("ExportAsFixedFormat", BindingFlags.InvokeMethod, null, doc,
                                new object[] { outPathLocal, 17 /*wdExportFormatPDF*/ });
                        else
                            dt.InvokeMember("ExportAsFixedFormat", BindingFlags.InvokeMethod, null, doc,
                                new object[] { 0 /*xlTypePDF*/, outPathLocal });
                    }
                    else
                    {
                        step = "SaveAs";
                        // Excel: 51 = xlOpenXMLWorkbook；62 = xlCSVUTF8
                        dt.InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, doc,
                            new object[] { outPathLocal, target == ConvTarget.Xlsx ? 51 : 62, 6 /*Local*/ });
                    }
                }
                catch (Exception ex)
                {
                    Exception inner = ex is TargetInvocationException && ex.InnerException != null ? ex.InnerException : ex;
                    result = "[" + step + "] " + inner.Message;
                }
                finally
                {
                    if (doc != null)
                    {
                        try { doc.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, doc, new object[] { false }); } catch { }
                        try { Marshal.FinalReleaseComObject(doc); } catch { }
                    }
                    if (app != null)
                    {
                        try { app.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, app, null); } catch { }
                        try { Marshal.FinalReleaseComObject(app); } catch { }
                    }
                }
            }));
            worker.SetApartmentState(ApartmentState.STA);   // Office COM 必须 STA
            worker.IsBackground = true;
            worker.Start();
            if (!worker.Join(new TimeSpan(0, 0, ComTimeout(kind, target))))
            {
                // 看门狗：可能是 Office 在后台弹了不可见对话框（首次运行/激活）。
                // 不越权杀进程（用户可能开着 Excel），返回明确错误并提示。
                if (kind == "Excel" && target == ConvTarget.Pdf) excelPdfBroken = true;
                return "COM 转换超时（" + ComTimeout(kind, target) + " 秒，" + step + " 阶段）。若反复出现，请在桌面手动打开一次 " + kind + " 完成首次运行向导。";
            }
            outputPath = outPathLocal;
            if (result != null) outputPath = null;
            return result;
        }

        // ---------- ② LibreOffice headless ----------

        // 超时分级：PDF 导出（大表上千页）远慢于格式互转；上限放宽到 900s，Job Object 仍在
        // 内存/整树上兜底。超时一律视为失败——绝不把被杀进程残留的半成品文件当成功。
        int LoTimeoutSeconds(ConvTarget target)
        {
            return target == ConvTarget.Pdf ? 900 : 300;
        }

        string ConvertLo(string input, ConvTarget target, out string outputPath)
        {
            outputPath = null;
            if (SofficePath == null) return "LibreOffice 未安装";
            string fmt = target == ConvTarget.Pdf ? "pdf" : (target == ConvTarget.Csv ? "csv" : "xlsx");
            string outdir = Path.GetDirectoryName(Path.GetFullPath(input));
            string expected = Path.Combine(outdir, Path.GetFileNameWithoutExtension(input) + "." + fmt);
            if (File.Exists(expected)) { try { File.Delete(expected); } catch { } }
            // soffice 为常量探测路径；ToolRunner 白名单校验；参数模板固定，用户路径仅做引号转义；UseShellExecute=false
            string args = "--headless --norestore --convert-to " + fmt + " --outdir " + Q(outdir) + " " + Q(input);
            int code = ToolRunner.ConvertViaSoffice(SofficePath, args, LoTimeoutSeconds(target));
            if (code == ProcRunner.ExitTimeout) return "LibreOffice 转换超时（进程已整树终止，无输出）";
            if (File.Exists(expected))
            {
                string dst = OutputPathOf(input, target);
                if (!string.Equals(expected, dst, StringComparison.OrdinalIgnoreCase)) { try { File.Move(expected, dst); } catch { } }
                outputPath = dst;
                return null;
            }
            return "soffice exit=" + code;
        }

        static string Q(string s)
        {
            return "\"" + (s == null ? "" : s.Replace("\"", "\"\"")) + "\"";
        }

        // ---------- ③ 原生直写 ----------

        string XlsxToCsv(string input, out string outputPath)
        {
            outputPath = OutputPathOf(input, ConvTarget.Csv);
            using (XlsxBook book = XlsxBook.Open(input))
            {
                // 导出场景：无缓存值的公式格**置空**。
                // 若照搬 "=公式" 文本，系统/ERP 生成的表（写了公式但不写缓存值）会导出
                // 一墙 "=IF(...,_xlfn.TEXTJOIN(...))"，用户看到的就是"乱码"。
                book.BlankUncachedFormula = true;
                List<string[]> rows = new List<string[]>();
                string err = book.StreamRows(0, 256, delegate(string[] row)
                {
                    List<string> clean = new List<string>();
                    int last = -1;
                    for (int i = row.Length - 1; i >= 0; i--) { if (row[i] != null) { last = i; break; } }
                    for (int i = 0; i <= last; i++) clean.Add(row[i] ?? "");
                    rows.Add(clean.ToArray());
                });
                if (err != null) { outputPath = null; return "读取 xlsx 失败: " + err; }
                if (book.UncachedFormulaCount > 0)
                {
                    LastWarning = "源表有 " + book.UncachedFormulaCount + " 个公式单元格没有缓存值" +
                        "（常见于系统/ERP 导出的 xlsx），CSV 中这些格已留空。" +
                        "如需它们的计算值：先用 Excel/WPS 打开该表并另存一次（让公式算出结果），再转换。";
                }
                return MiniCsv.Write(outputPath, rows) ?? null;
            }
        }

        string CsvToXlsx(string input, out string outputPath)
        {
            outputPath = OutputPathOf(input, ConvTarget.Xlsx);
            try
            {
                Encoding used;
                string text = MiniCsv.DetectRead(input, out used);
                List<string[]> rows = MiniCsv.Parse(text);
                return WriteMinimalXlsx(outputPath, rows);
            }
            catch (Exception ex) { outputPath = null; return ex.Message; }
        }

        // 最小可用 xlsx 写出（inlineStr，Excel/WPS 均可打开）
        static string WriteMinimalXlsx(string path, List<string[]> rows)
        {
            try
            {
                if (File.Exists(path)) File.Delete(path);
                StringBuilder sheet = new StringBuilder();
                sheet.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                sheet.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>");
                for (int i = 0; i < rows.Count; i++)
                {
                    sheet.Append("<row r=\"").Append(i + 1).Append("\">");
                    string[] row = rows[i];
                    for (int j = 0; j < row.Length; j++)
                    {
                        string v = row[j];
                        if (string.IsNullOrEmpty(v)) continue;
                        string cellRef = ColName(j) + (i + 1);
                        double num;
                        if (double.TryParse(v, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out num))
                        {
                            sheet.Append("<c r=\"").Append(cellRef).Append("\"><v>").Append(v).Append("</v></c>");
                        }
                        else
                        {
                            sheet.Append("<c r=\"").Append(cellRef).Append("\" t=\"inlineStr\"><is><t xml:space=\"preserve\">")
                                 .Append(EscapeXml(v)).Append("</t></is></c>");
                        }
                    }
                    sheet.Append("</row>");
                }
                sheet.Append("</sheetData></worksheet>");

                using (FileStream fs = new FileStream(path, FileMode.Create))
                {
                    List<MiniZipEntryOut> ents = new List<MiniZipEntryOut>();
                    XlsxSkeleton.Add(ents, "[Content_Types].xml", XlsxSkeleton.ContentTypes(1));
                    XlsxSkeleton.Add(ents, "_rels/.rels", XlsxSkeleton.RootRels());
                    XlsxSkeleton.Add(ents, "xl/workbook.xml",
                        XlsxSkeleton.WorkbookXml("<sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>"));
                    XlsxSkeleton.Add(ents, "xl/_rels/workbook.xml.rels",
                        XlsxSkeleton.WorkbookRels("<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"));
                    XlsxSkeleton.Add(ents, "xl/worksheets/sheet1.xml", sheet.ToString());
                    MiniZipWriter.Write(fs, ents);
                }
                return null;
            }
            catch (Exception ex) { return ex.Message; }
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

        static string EscapeXml(string s)
        {
            return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
                    .Replace("\"", "&quot;").Replace("\r", "").Replace("\n", "_x000a_");
        }
    }
}
