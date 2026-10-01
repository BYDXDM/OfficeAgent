// PptWriter —— 最小 PPT(.pptx) 生成器：模板替换式
// 为什么用模板：OOXML 演示文稿的母版/版式/主题手写极易产出 PowerPoint 拒开的文件；
// 模板取自真实 PowerPoint 实产的 test.pptx（母版/12 版式/主题齐全），母版部分原样保留，必然合法。
// 生成策略：保留模板全部部件，仅重写 幻灯片列表（presentation.xml + rels + Content_Types + slideN.xml）。
// 每页两个独立文本框（标题 + 要点），不引用版式占位符 —— 与母版解耦，任何模板都兼容。
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeAgent.Core;

namespace OfficeAgent.Host
{
    public static class PptWriter
    {
        // slides: 每项 [标题, 要点（\n 分段）]
        public static string Save(string templatePath, string outPath, List<string[]> slides)
        {
            try
            {
                if (slides == null || slides.Count == 0) return "没有幻灯片内容";
                if (!File.Exists(templatePath)) return "缺少 PPT 模板: " + templatePath;

                StringBuilder presRels = new StringBuilder();
                StringBuilder sldIdLst = new StringBuilder();
                StringBuilder ctOverrides = new StringBuilder();
                List<MiniZipEntryOut> outs = new List<MiniZipEntryOut>();

                using (MiniZipFile tpl = MiniZipFile.OpenRead(templatePath))
                {
                    // 1) 原样保留：母版/版式/主题/docProps 等全部非幻灯片部件
                    foreach (MiniZipEntryInfo e in tpl.Entries)
                    {
                        string n = e.FullName;
                        if (n.StartsWith("ppt/slides/", StringComparison.OrdinalIgnoreCase)) continue;      // 旧幻灯片与 rels 全部丢弃
                        if (n == "[Content_Types].xml" || n == "ppt/presentation.xml" ||
                            n == "ppt/_rels/presentation.xml.rels") continue;                               // 这三个重写
                        using (Stream s = tpl.OpenEntry(n))
                            outs.Add(new MiniZipEntryOut(n, ReadAll(s)));
                    }

                    // 2) 逐页生成 slideN.xml + rels
                    for (int i = 0; i < slides.Count; i++)
                    {
                        string slideXml = SlideXml(slides[i][0], slides[i][1]);
                        outs.Add(new MiniZipEntryOut("ppt/slides/slide" + (i + 1) + ".xml",
                            new UTF8Encoding(false).GetBytes(slideXml)));
                        outs.Add(new MiniZipEntryOut("ppt/slides/_rels/slide" + (i + 1) + ".xml.rels",
                            new UTF8Encoding(false).GetBytes(
                                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/>" +
                                "</Relationships>")));
                        int rid = i + 2;   // rId1 留给母版
                        sldIdLst.Append("<p:sldId id=\"").Append(256 + i).Append("\" r:id=\"rId").Append(rid).Append("\"/>");
                        presRels.Append("<Relationship Id=\"rId").Append(rid)
                          .Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide").Append(i + 1).Append(".xml\"/>");
                        ctOverrides.Append("<Override PartName=\"/ppt/slides/slide").Append(i + 1)
                          .Append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>");
                    }
                }

                // 3) presentation.xml + rels
                outs.Add(new MiniZipEntryOut("ppt/presentation.xml", new UTF8Encoding(false).GetBytes(
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<p:presentation xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                    "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" " +
                    "saveSubsetFonts=\"1\">" +
                    "<p:sldMasterIdLst><p:sldMasterId id=\"2147483648\" r:id=\"rId1\"/></p:sldMasterIdLst>" +
                    "<p:sldIdLst>" + sldIdLst.ToString() + "</p:sldIdLst>" +
                    "<p:sldSz cx=\"12192000\" cy=\"6858000\"/><p:notesSz cx=\"6858000\" cy=\"9144000\"/></p:presentation>")));
                outs.Add(new MiniZipEntryOut("ppt/_rels/presentation.xml.rels", new UTF8Encoding(false).GetBytes(
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\" Target=\"slideMasters/slideMaster1.xml\"/>" +
                    presRels.ToString() +
                    "<Relationship Id=\"rId" + (slides.Count + 2) + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>" +
                    "</Relationships>")));

                // 4) Content_Types：以模板为底，去掉旧 slide 覆盖项，追加新 slide 覆盖项
                string ct;
                using (Stream s = MiniZipFile.OpenRead(templatePath).OpenEntry("[Content_Types].xml"))
                    ct = new UTF8Encoding(false).GetString(ReadAll(s));
                int oldSlide = ct.IndexOf("<Override PartName=\"/ppt/slides/", StringComparison.Ordinal);
                if (oldSlide >= 0)
                {
                    int oldEnd = ct.IndexOf("/>", oldSlide, StringComparison.Ordinal);
                    if (oldEnd >= 0) ct = ct.Remove(oldSlide, oldEnd - oldSlide + 2);
                }
                ct = ct.Replace("</Types>", ctOverrides.ToString() + "</Types>");
                outs.Add(new MiniZipEntryOut("[Content_Types].xml", new UTF8Encoding(false).GetBytes(ct)));

                using (FileStream fs = new FileStream(outPath, FileMode.Create))
                    MiniZipWriter.Write(fs, outs);
                return null;
            }
            catch (Exception ex) { return ex.Message; }
        }

        // 单页：标题框（顶部，32pt 粗体）+ 要点框（18pt，逐段）
        static string SlideXml(string title, string body)
        {
            StringBuilder sp = new StringBuilder();
            sp.Append(TextBox(1, "标题", 600000, 400000, 10992000, 1100000, Esc(title), "3200", true));
            StringBuilder paras = new StringBuilder();
            string[] lines = (body == null ? "" : body).Split('\n');
            bool first = true;
            foreach (string raw in lines)
            {
                string t = raw.Trim();
                if (t.Length == 0) continue;
                if (!first) paras.Append("</a:p>");
                paras.Append("<a:p><a:r><a:rPr lang=\"zh-CN\" sz=\"1800\"/><a:t>• ").Append(Esc(t)).Append("</a:t></a:r>");
                first = false;
            }
            if (first) paras.Append("<a:p><a:r><a:rPr lang=\"zh-CN\" sz=\"1800\"/><a:t></a:t></a:r>");
            paras.Append("</a:p>");
            sp.Append(TextBoxRaw(2, "要点", 600000, 1700000, 10992000, 4600000, paras.ToString()));
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">" +
                "<p:cSld><p:spTree>" +
                "<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>" +
                sp.ToString() +
                "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>";
        }

        // 纯文本框（单段）
        static string TextBox(int id, string name, long x, long y, long cx, long cy, string text, string sz, bool bold)
        {
            return "<p:sp><p:nvSpPr><p:cNvPr id=\"" + id + "\" name=\"" + Esc(name) + "\"/>" +
                "<p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr>" +
                "<p:spPr><a:xfrm><a:off x=\"" + x + "\" y=\"" + y + "\"/><a:ext cx=\"" + cx + "\" cy=\"" + cy + "\"/></a:xfrm>" +
                "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr>" +
                "<p:txBody><a:bodyPr wrap=\"square\"/><a:lstStyle/>" +
                "<a:p><a:r><a:rPr lang=\"zh-CN\" sz=\"" + sz + "\"" + (bold ? " b=\"1\"" : "") + "/>" +
                "<a:t>" + text + "</a:t></a:r></a:p>" +
                "</p:txBody></p:sp>";
        }

        // 预组好段落的文本框
        static string TextBoxRaw(int id, string name, long x, long y, long cx, long cy, string paragraphsXml)
        {
            return "<p:sp><p:nvSpPr><p:cNvPr id=\"" + id + "\" name=\"" + Esc(name) + "\"/>" +
                "<p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr>" +
                "<p:spPr><a:xfrm><a:off x=\"" + x + "\" y=\"" + y + "\"/><a:ext cx=\"" + cx + "\" cy=\"" + cy + "\"/></a:xfrm>" +
                "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr>" +
                "<p:txBody><a:bodyPr wrap=\"square\"/><a:lstStyle/>" + paragraphsXml +
                "</p:txBody></p:sp>";
        }

        static string Esc(string s)
        {
            if (s == null) return "";
            return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
        }

        static byte[] ReadAll(Stream s)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                byte[] buf = new byte[65536];
                int n;
                while ((n = s.Read(buf, 0, buf.Length)) > 0) ms.Write(buf, 0, n);
                return ms.ToArray();
            }
        }
    }
}
