// MainForm.Recon.cs —— partial：表格核对页 + 汇总·发票页（M2 P0 场景 UI）
// 布局沿用工程惯例：绝对定位 + 线程回 Invoke；引擎调用在后台线程，UI 只做参数收集与结果展示。
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using OfficeAgent.Core;

namespace OfficeAgent.Host
{
    public partial class MainForm
    {
        // ---------- 表格核对页控件 ----------
        TextBox reconA, reconB;
        TextBox reconKeyA, reconKeyB;
        TextBox reconDebitA, reconCreditA, reconDebitB, reconCreditB;
        TextBox reconTol, reconHdrA, reconHdrB, reconSheetA, reconSheetB;
        CheckBox reconNorm, reconAll;
        ComboBox reconTpl;
        TextBox reconTplName;
        Label reconStatus;
        string reconOutput = null;

        Panel BuildReconPage()
        {
            Panel p = new Panel();
            p.Dock = DockStyle.Fill;
            p.BackColor = Color.White;
            Font f = new Font("Microsoft YaHei UI", 9F);

            int y = 16;
            AddField(p, f, "文件 A（如银行流水）", 24, ref y, 320, out reconA);
            AddBrowse(p, f, reconA, y - 30);
            AddField(p, f, "文件 B（如账面记录）", 424, ref y, 320, out reconB);
            AddBrowse(p, f, reconB, y - 30);

            AddTip(p, f, "键列（可多列，逗号分隔；填列名或第几列，如「账号」或「1」）", 24, y); y += 22;
            AddField2(p, f, "A 键列", 24, y, 150, out reconKeyA);
            AddField2(p, f, "B 键列", 224, y, 150, out reconKeyB);
            y += 52;

            AddTip(p, f, "金额列（借/贷两列口径，如「收入,支出」；单列签名口径只填借方）", 24, y); y += 22;
            AddField2(p, f, "A 借方列", 24, y, 110, out reconDebitA);
            AddField2(p, f, "A 贷方列", 154, y, 110, out reconCreditA);
            AddField2(p, f, "B 借方列", 294, y, 110, out reconDebitB);
            AddField2(p, f, "B 贷方列", 424, y, 110, out reconCreditB);
            y += 52;

            AddTip(p, f, "容差（如 0.01）/ 表头行 A、B / 工作表 A、B（填序号或名称；CSV 忽略）", 24, y); y += 22;
            AddField2(p, f, "容差", 24, y, 70, out reconTol); reconTol.Text = "0.01";
            AddField2(p, f, "表头行A", 114, y, 50, out reconHdrA); reconHdrA.Text = "1";
            AddField2(p, f, "表头行B", 184, y, 50, out reconHdrB); reconHdrB.Text = "1";
            AddField2(p, f, "工作表A", 254, y, 70, out reconSheetA); reconSheetA.Text = "1";
            AddField2(p, f, "工作表B", 344, y, 70, out reconSheetB); reconSheetB.Text = "1";
            y += 52;

            reconNorm = new CheckBox();
            reconNorm.Text = "键归一化（去空格/全半角/大小写）";
            reconNorm.Location = new Point(24, y);
            reconNorm.AutoSize = true;
            reconNorm.Checked = true;
            p.Controls.Add(reconNorm);
            reconAll = new CheckBox();
            reconAll.Text = "差异表附匹配明细页";
            reconAll.Location = new Point(280, y);
            reconAll.AutoSize = true;
            p.Controls.Add(reconAll);
            y += 32;

            AddTip(p, f, "模板（映射配置存本机，「上月对账再来一次」直接载入）", 24, y); y += 22;
            reconTpl = new ComboBox();
            reconTpl.Location = new Point(24, y);
            reconTpl.Size = new Size(180, 26);
            reconTpl.DropDownStyle = ComboBoxStyle.DropDownList;
            p.Controls.Add(reconTpl);
            Button btnLoadTpl = SmallBtn("载入模板", 214, y, 84, delegate { LoadReconTemplate(); });
            p.Controls.Add(btnLoadTpl);
            reconTplName = new TextBox();
            reconTplName.Location = new Point(308, y - 1);
            reconTplName.Size = new Size(120, 26);
            p.Controls.Add(reconTplName);
            Button btnSaveTpl = SmallBtn("保存模板", 436, y, 84, delegate { SaveReconTemplate(); });
            p.Controls.Add(btnSaveTpl);
            Button btnGuess = SmallBtn("自动猜列", 530, y, 84, delegate { GuessReconColumns(); });
            p.Controls.Add(btnGuess);
            y += 40;

            Button go = new Button();
            go.Text = "开始核对";
            go.Location = new Point(24, y);
            go.Size = new Size(120, 34);
            go.FlatStyle = FlatStyle.Flat;
            go.FlatAppearance.BorderSize = 0;
            go.BackColor = Color.FromArgb(46, 76, 196);
            go.ForeColor = Color.White;
            go.Cursor = Cursors.Hand;
            go.Click += new EventHandler(delegate(object s, EventArgs ev) { RunReconUi(); });
            reconGoBtn = go;
            p.Controls.Add(go);
            Button openOut = SmallBtn("打开差异表", 154, y + 3, 100, delegate
            {
                if (reconOutput != null && File.Exists(reconOutput)) OpenInExplorer(reconOutput);
                else MessageBox.Show("还没有核对结果。");
            });
            p.Controls.Add(openOut);
            y += 50;

            reconStatus = new Label();
            reconStatus.Location = new Point(24, y);
            reconStatus.Size = new Size(900, 56);
            reconStatus.ForeColor = Color.FromArgb(90, 93, 105);
            reconStatus.Text = "两表核对：键相同且金额差在容差内 → 匹配；其余进入差异表（金额不等 / 仅A有 / 仅B有）。";
            p.Controls.Add(reconStatus);
            y += 60;

            // 检查结果列表：大白话 + "❗" 悬停显示通俗解释（鼠标停 1 秒即可看到）
            lvChecks = new ListView();
            lvChecks.View = View.Details;
            lvChecks.FullRowSelect = true;
            lvChecks.ShowItemToolTips = true;
            lvChecks.Location = new Point(24, y);
            lvChecks.Size = new Size(900, 116);
            lvChecks.Columns.Add("程序自动检查了什么", 300);
            lvChecks.Columns.Add("结果", 70);
            lvChecks.Columns.Add("数字说明", 520);
            p.Controls.Add(lvChecks);
            return p;
        }

        ListView lvChecks;

        delegate void FillChecksD(ReconResult r);

        void FillReconChecks(ReconResult r)
        {
            if (InvokeRequired) { Invoke(new FillChecksD(FillReconChecks), r); return; }
            if (lvChecks == null) return;
            lvChecks.BeginUpdate();
            lvChecks.Items.Clear();
            foreach (ReconCheck c in r.Checks)
            {
                string verdict = c.Warn ? "提醒" : (c.Pass ? "通过" : "未通过");
                ListViewItem it = new ListViewItem(PlainTips.PlainCheck(c.Name) + " ❗");
                it.SubItems.Add(verdict);
                it.SubItems.Add(c.Detail);
                it.ToolTipText = "这条检查的意思：\n" + PlainTips.CheckHover(c.Name);
                it.ForeColor = c.Warn ? Color.FromArgb(156, 101, 0) : (c.Pass ? Color.FromArgb(0, 97, 0) : Color.FromArgb(156, 0, 6));
                lvChecks.Items.Add(it);
            }
            lvChecks.EndUpdate();
        }

        void AddField(Panel p, Font f, string label, int x, ref int y, int w, out TextBox tb)
        {
            Label l = new Label();
            l.Text = label;
            l.ForeColor = Color.FromArgb(90, 93, 105);
            l.Location = new Point(x, y);
            l.AutoSize = true;
            p.Controls.Add(l);
            y += 22;
            tb = new TextBox();
            tb.Location = new Point(x, y);
            tb.Size = new Size(w, 26);
            p.Controls.Add(tb);
            y += 34;
        }

        void AddBrowse(Panel p, Font f, TextBox target, int y)
        {
            Button b = SmallBtn("浏览…", 0, y, 60, delegate
            {
                using (OpenFileDialog dlg = new OpenFileDialog())
                {
                    dlg.Title = "选择文件";
                    dlg.Filter = "表格|*.xlsx;*.csv;*.xls|全部文件|*.*";
                    if (dlg.ShowDialog(this) == DialogResult.OK)
                    {
                        target.Text = dlg.FileName;
                        UpdateChatContext();
                    }
                }
            });
            b.Location = new Point(target.Right + 8, target.Top);
            p.Controls.Add(b);
        }

        // 把核对页当前选择的文件同步进会话上下文（自然语言路由的候选输入）
        void UpdateChatContext()
        {
            List<string> files = new List<string>();
            if (reconA != null && File.Exists(reconA.Text.Trim())) files.Add(reconA.Text.Trim());
            if (reconB != null && File.Exists(reconB.Text.Trim())) files.Add(reconB.Text.Trim());
            chat.SetContextFiles(files.ToArray());
        }

        void AddTip(Panel p, Font f, string text, int x, int y)
        {
            Label l = new Label();
            l.Text = text;
            l.ForeColor = Color.FromArgb(140, 143, 156);
            l.Location = new Point(x, y);
            l.AutoSize = true;
            p.Controls.Add(l);
        }

        void AddField2(Panel p, Font f, string label, int x, int y, int w, out TextBox tb)
        {
            Label l = new Label();
            l.Text = label;
            l.ForeColor = Color.FromArgb(90, 93, 105);
            l.Location = new Point(x, y);
            l.AutoSize = true;
            p.Controls.Add(l);
            tb = new TextBox();
            tb.Location = new Point(x, y + 20);
            tb.Size = new Size(w, 26);
            p.Controls.Add(tb);
        }

        Button SmallBtn(string text, int x, int y, int w, EventHandler onClick)
        {
            Button b = new Button();
            b.Text = text;
            b.Location = new Point(x, y);
            b.Size = new Size(w, 28);
            b.Cursor = Cursors.Hand;
            b.Click += onClick;
            return b;
        }

        void OpenInExplorer(string path)
        {
            try { System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + path + "\""); }
            catch (Exception ex) { MessageBox.Show("打开失败: " + ex.Message); }
        }

        ReconMapping ReconMappingFromUi()
        {
            ReconMapping m = new ReconMapping();
            m.KeyColsA = SplitTokensLocal(reconKeyA.Text);
            m.KeyColsB = SplitTokensLocal(reconKeyB.Text);
            m.DebitA = reconDebitA.Text.Trim();
            m.CreditA = reconCreditA.Text.Trim();
            m.DebitB = reconDebitB.Text.Trim();
            m.CreditB = reconCreditB.Text.Trim();
            m.SheetA = reconSheetA.Text.Trim();
            m.SheetB = reconSheetB.Text.Trim();
            int n;
            if (int.TryParse(reconHdrA.Text.Trim(), out n) && n >= 1) m.HeaderRowA = n;
            if (int.TryParse(reconHdrB.Text.Trim(), out n) && n >= 1) m.HeaderRowB = n;
            double d;
            if (double.TryParse(reconTol.Text.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out d) && d >= 0)
                m.Tolerance = d;
            m.SqueezeKey = reconNorm.Checked;
            m.IncludeMatched = reconAll.Checked;
            return m;
        }

        void FillReconUi(ReconMapping m)
        {
            reconKeyA.Text = string.Join(",", Array.ConvertAll(m.KeyColsA, delegate(string s) { return s; }));
            reconKeyB.Text = string.Join(",", Array.ConvertAll(m.KeyColsB, delegate(string s) { return s; }));
            reconDebitA.Text = m.DebitA; reconCreditA.Text = m.CreditA;
            reconDebitB.Text = m.DebitB; reconCreditB.Text = m.CreditB;
            reconSheetA.Text = m.SheetA; reconSheetB.Text = m.SheetB;
            reconHdrA.Text = m.HeaderRowA.ToString(CultureInfo.InvariantCulture);
            reconHdrB.Text = m.HeaderRowB.ToString(CultureInfo.InvariantCulture);
            reconTol.Text = m.Tolerance.ToString("0.####", CultureInfo.InvariantCulture);
            reconNorm.Checked = m.SqueezeKey;
            reconAll.Checked = m.IncludeMatched;
        }

        void RefreshReconTplCombo()
        {
            reconTpl.Items.Clear();
            foreach (ReconTemplate t in ReconTemplateStore.LoadAll())
            {
                if (t.Name == ReconTemplateStore.LastName) continue;
                reconTpl.Items.Add(t.Name);
            }
        }

        void LoadReconTemplate()
        {
            if (reconTpl.SelectedItem == null) { MessageBox.Show("先在列表中选择一个模板。"); return; }
            ReconTemplate t = ReconTemplateStore.Get(reconTpl.SelectedItem.ToString());
            if (t == null) { MessageBox.Show("模板不存在。"); return; }
            FillReconUi(t.ToMapping());
            if (File.Exists(t.FileA)) reconA.Text = t.FileA;
            if (File.Exists(t.FileB)) reconB.Text = t.FileB;
            SetReconStatus("已载入模板「" + t.Name + "」。", null);
        }

        void SaveReconTemplate()
        {
            string name = reconTplName.Text.Trim();
            if (name.Length == 0) { MessageBox.Show("请先填写模板名称。"); return; }
            ReconTemplate t = ReconTemplate.FromMapping(name, ReconMappingFromUi(), reconA.Text.Trim(), reconB.Text.Trim());
            ReconTemplateStore.Upsert(t);
            RefreshReconTplCombo();
            SetReconStatus("模板已保存：「" + name + "」。", null);
        }

        // 从 A/B 文件表头自动猜键列与金额列（命中常见会计表头关键词）
        void GuessReconColumns()
        {
            string[] keyWords = new string[] { "账号", "流水号", "凭证号", "单号", "票号", "编码", "摘要", "户名", "对方" };
            string[] debitWords = new string[] { "借方", "收入", "收方", "入金", "发生额" };
            string[] creditWords = new string[] { "贷方", "支出", "付方", "出金" };
            GuessSide(reconA.Text.Trim(), reconHdrA.Text.Trim(), keyWords, debitWords, creditWords, true);
            GuessSide(reconB.Text.Trim(), reconHdrB.Text.Trim(), keyWords, debitWords, creditWords, false);
            SetReconStatus("已按表头关键词自动填列，请核对后开始。", null);
        }

        void GuessSide(string file, string hdrText, string[] keyWords, string[] debitWords, string[] creditWords, bool sideA)
        {
            if (file.Length == 0 || !File.Exists(file)) return;
            string[] header = null;
            int hr = 1;
            int.TryParse(hdrText, out hr);
            try
            {
                string ext = (Path.GetExtension(file) ?? "").ToLowerInvariant();
                if (ext == ".csv")
                {
                    Encoding used;
                    List<string[]> rows = MiniCsv.Parse(MiniCsv.DetectRead(file, out used));
                    if (rows.Count >= hr) header = rows[hr - 1];
                }
                else if (ext == ".xlsx")
                {
                    using (XlsxBook book = XlsxBook.Open(file))
                    {
                        string err = book.StreamRows(0, 64, delegate(string[] cells, int excelRow)
                        {
                            if (header == null && excelRow == hr) header = (string[])cells.Clone();
                            return;
                        });
                    }
                }
                if (header == null) return;
                StringBuilder keys = new StringBuilder();
                string debit = "", credit = "";
                for (int i = 0; i < header.Length; i++)
                {
                    string h = header[i] == null ? "" : header[i].Trim();
                    if (h.Length == 0) continue;
                    foreach (string kw in keyWords)
                    {
                        if (h.IndexOf(kw, StringComparison.Ordinal) >= 0)
                        {
                            if (keys.Length > 0) keys.Append(",");
                            keys.Append(h);
                            break;
                        }
                    }
                    if (debit.Length == 0)
                    {
                        foreach (string kw in debitWords) { if (h.IndexOf(kw, StringComparison.Ordinal) >= 0) { debit = h; break; } }
                    }
                    if (credit.Length == 0)
                    {
                        foreach (string kw in creditWords) { if (h.IndexOf(kw, StringComparison.Ordinal) >= 0) { credit = h; break; } }
                    }
                }
                if (sideA)
                {
                    if (keys.Length > 0) reconKeyA.Text = keys.ToString();
                    reconDebitA.Text = debit; reconCreditA.Text = credit;
                }
                else
                {
                    if (keys.Length > 0) reconKeyB.Text = keys.ToString();
                    reconDebitB.Text = debit; reconCreditB.Text = credit;
                }
            }
            catch { }
        }

        void SetReconStatus(string s, Color? color)
        {
            if (InvokeRequired) { Invoke((MethodInvoker)delegate { SetReconStatus(s, color); }); return; }
            reconStatus.Text = s;
            reconStatus.ForeColor = color.HasValue ? color.Value : Color.FromArgb(90, 93, 105);
        }

        void RunReconUi()
        {
            string fileA = reconA.Text.Trim(), fileB = reconB.Text.Trim();
            if (fileA.Length == 0 || fileB.Length == 0 || !File.Exists(fileA) || !File.Exists(fileB))
            {
                SetReconStatus("请先选择 A、B 两个文件。", Color.FromArgb(178, 58, 58));
                return;
            }
            ReconMapping map = ReconMappingFromUi();
            if (map.KeyColsA.Length == 0 || map.KeyColsB.Length == 0)
            {
                SetReconStatus("请先指定两侧键列（可点「自动猜列」）。", Color.FromArgb(178, 58, 58));
                return;
            }
            string outPath = Path.Combine(Path.GetDirectoryName(Path.GetFullPath(fileA)),
                "核对_" + TrimBaseLocal(fileA, 24) + "_vs_" + TrimBaseLocal(fileB, 24) + ".xlsx");
            SetReconStatus("核对中…", null);
            if (reconGoBtn != null) reconGoBtn.Enabled = false;
            Thread t = new Thread(new ThreadStart(delegate
            {
                ConvertEngine conv = new ConvertEngine();
                conv.SofficePath = ConvertEngine.FindSoffice(EnvDetect.FindRoot());
                ReconEngine engine = new ReconEngine();
                ReconResult r = engine.Run(map, fileA, fileB, outPath, conv);
                AuditLog.Record("recon_run", "A=" + fileA + "; B=" + fileB +
                    "; matched=" + r.MatchedCount + "; diffs=" + r.Diffs.Count + "; err=" + (r.Err ?? "none"));
                if (r.Err != null)
                {
                    SetReconStatus("✗ " + r.Err, Color.FromArgb(178, 58, 58));
                }
                else
                {
                    if (File.Exists(outPath)) AuditLog.Record("file_write", outPath);
                    reconOutput = outPath;
                    StringBuilder sb = new StringBuilder();
                    sb.Append("完成（").Append(r.ElapsedMs).Append(" ms）：匹配 ").Append(r.MatchedCount)
                      .Append("；金额不等 ").Append(r.PairDiffCount)
                      .Append("；仅A ").Append(r.OnlyACount)
                      .Append("；仅B ").Append(r.OnlyBCount).Append("\n");
                    sb.Append("合计：A=").Append(r.SumA.ToString("N2", CultureInfo.InvariantCulture))
                      .Append("  B=").Append(r.SumB.ToString("N2", CultureInfo.InvariantCulture))
                      .Append("  差=").Append((r.SumA - r.SumB).ToString("N2", CultureInfo.InvariantCulture)).Append("\n");
                    foreach (ReconCheck c in r.Checks)
                        sb.Append(c.Warn ? "[提醒] " : (c.Pass ? "[✓] " : "[✗] ")).Append(PlainTips.PlainCheck(c.Name)).Append("  ").Append(c.Detail).Append("\n");
                    sb.Append("差异表：").Append(outPath);
                    bool hasFail = r.FailedChecks() > 0;
                    SetReconStatus(sb.ToString(), hasFail ? Color.FromArgb(178, 58, 58) : Color.FromArgb(22, 130, 93));
                    FillReconChecks(r);
                    ReconTemplateStore.Upsert(ReconTemplate.FromMapping(ReconTemplateStore.LastName, map, fileA, fileB));
                }
                EnableReconGo();
            }));
            t.IsBackground = true;
            t.Start();
        }

        Button reconGoBtn = null;

        void EnableReconGo()
        {
            if (InvokeRequired) { Invoke((MethodInvoker)EnableReconGo); return; }
            if (reconGoBtn != null) reconGoBtn.Enabled = true;
        }

        // 本地帮助：与 Program.SplitTokens/TrimBase 同义（partial 类不共享彼此私有成员）
        static string[] SplitTokensLocal(string s)
        {
            List<string> parts = new List<string>();
            foreach (string p in s.Split(','))
            {
                string t = p.Trim();
                if (t.Length > 0) parts.Add(t);
            }
            return parts.ToArray();
        }

        static string TrimBaseLocal(string path, int max)
        {
            string b = Path.GetFileNameWithoutExtension(path);
            if (b.Length > max) b = b.Substring(0, max);
            foreach (char c in Path.GetInvalidFileNameChars()) b = b.Replace(c, '_');
            return b;
        }

        // ================= 汇总·发票页 =================

        ListBox mergeFiles;
        TextBox mergeTplName, mergeTargets, mergeSrcs, mergeAmounts, mergeHdr, mergeSheet;
        ComboBox mergeTpl;
        Label mergeStatus, invoiceStatus;
        string mergeOutput = null, invoiceOutput = null;
        ListBox invoiceFiles;

        Panel BuildExtractPage()
        {
            Panel p = new Panel();
            p.Dock = DockStyle.Fill;
            p.BackColor = Color.White;
            Font f = new Font("Microsoft YaHei UI", 9F);

            // ---------- 上半：报表汇总 ----------
            Label cap1 = new Label();
            cap1.Text = "报表汇总（多簿按表头名归集，月结多子公司场景）";
            cap1.Font = new Font("Microsoft YaHei UI", 11F, FontStyle.Bold);
            cap1.Location = new Point(24, 14);
            cap1.AutoSize = true;
            p.Controls.Add(cap1);

            mergeFiles = new ListBox();
            mergeFiles.Location = new Point(24, 44);
            mergeFiles.Size = new Size(420, 110);
            p.Controls.Add(mergeFiles);
            p.Controls.Add(SmallBtn("添加文件", 454, 44, 90, delegate
            {
                using (OpenFileDialog dlg = new OpenFileDialog())
                {
                    dlg.Title = "选择要汇总的表格（可多选）";
                    dlg.Multiselect = true;
                    dlg.Filter = "表格|*.xlsx;*.csv;*.xls|全部文件|*.*";
                    if (dlg.ShowDialog(this) == DialogResult.OK)
                        foreach (string s in dlg.FileNames) mergeFiles.Items.Add(s);
                }
            }));
            p.Controls.Add(SmallBtn("移除所选", 454, 78, 90, delegate
            {
                if (mergeFiles.SelectedIndex >= 0) mergeFiles.Items.RemoveAt(mergeFiles.SelectedIndex);
            }));
            p.Controls.Add(SmallBtn("清空", 454, 112, 90, delegate { mergeFiles.Items.Clear(); }));

            Label l1 = new Label();
            l1.Text = "模板：目标列,源列（一一对应，逗号分隔）；金额列参与合计勾稽";
            l1.ForeColor = Color.FromArgb(140, 143, 156);
            l1.Location = new Point(560, 44);
            l1.AutoSize = true;
            p.Controls.Add(l1);
            Label lt = new Label();
            lt.Text = "目标列";
            lt.Location = new Point(560, 70);
            lt.AutoSize = true;
            p.Controls.Add(lt);
            mergeTargets = new TextBox();
            mergeTargets.Location = new Point(614, 66);
            mergeTargets.Size = new Size(280, 26);
            p.Controls.Add(mergeTargets);
            Label ls = new Label();
            ls.Text = "源  列";
            ls.Location = new Point(560, 100);
            ls.AutoSize = true;
            p.Controls.Add(ls);
            mergeSrcs = new TextBox();
            mergeSrcs.Location = new Point(614, 96);
            mergeSrcs.Size = new Size(280, 26);
            p.Controls.Add(mergeSrcs);
            Label la = new Label();
            la.Text = "金额列";
            la.Location = new Point(560, 130);
            la.AutoSize = true;
            p.Controls.Add(la);
            mergeAmounts = new TextBox();
            mergeAmounts.Location = new Point(614, 126);
            mergeAmounts.Size = new Size(280, 26);
            p.Controls.Add(mergeAmounts);

            Label lh = new Label();
            lh.Text = "表头行";
            lh.Location = new Point(24, 164);
            lh.AutoSize = true;
            p.Controls.Add(lh);
            mergeHdr = new TextBox();
            mergeHdr.Text = "1";
            mergeHdr.Location = new Point(78, 160);
            mergeHdr.Size = new Size(40, 26);
            p.Controls.Add(mergeHdr);
            Label lsh = new Label();
            lsh.Text = "工作表";
            lsh.Location = new Point(130, 164);
            lsh.AutoSize = true;
            p.Controls.Add(lsh);
            mergeSheet = new TextBox();
            mergeSheet.Text = "1";
            mergeSheet.Location = new Point(184, 160);
            mergeSheet.Size = new Size(60, 26);
            p.Controls.Add(mergeSheet);

            mergeTpl = new ComboBox();
            mergeTpl.Location = new Point(260, 160);
            mergeTpl.Size = new Size(150, 26);
            mergeTpl.DropDownStyle = ComboBoxStyle.DropDownList;
            p.Controls.Add(mergeTpl);
            p.Controls.Add(SmallBtn("载入", 418, 160, 56, delegate
            {
                if (mergeTpl.SelectedItem == null) return;
                MergeTemplate t = MergeTemplateStore.Get(mergeTpl.SelectedItem.ToString());
                if (t == null) return;
                mergeTargets.Text = string.Join(",", t.TargetCols);
                mergeSrcs.Text = string.Join(",", t.SrcCols);
                mergeAmounts.Text = string.Join(",", t.AmountCols);
                mergeHdr.Text = t.HeaderRow.ToString(CultureInfo.InvariantCulture);
                mergeSheet.Text = t.SheetSpec;
            }));
            mergeTplName = new TextBox();
            mergeTplName.Location = new Point(484, 161);
            mergeTplName.Size = new Size(110, 26);
            p.Controls.Add(mergeTplName);
            p.Controls.Add(SmallBtn("存模板", 600, 160, 70, delegate
            {
                string name = mergeTplName.Text.Trim();
                if (name.Length == 0) { MessageBox.Show("请填模板名。"); return; }
                MergeTemplateStore.Upsert(MergeTemplate.FromForm(name, mergeSheet.Text.Trim(),
                    ParseInt(mergeHdr.Text, 1), SplitTokensLocal(mergeTargets.Text), SplitTokensLocal(mergeSrcs.Text),
                    SplitTokensLocal(mergeAmounts.Text)));
                RefreshMergeTplCombo();
            }));
            p.Controls.Add(SmallBtn("删模板", 676, 160, 70, delegate
            {
                if (mergeTpl.SelectedItem != null)
                {
                    MergeTemplateStore.Delete(mergeTpl.SelectedItem.ToString());
                    RefreshMergeTplCombo();
                }
            }));
            Button mergeGo = new Button();
            mergeGo.Text = "开始汇总";
            mergeGo.Location = new Point(760, 156);
            mergeGo.Size = new Size(110, 34);
            mergeGo.FlatStyle = FlatStyle.Flat;
            mergeGo.FlatAppearance.BorderSize = 0;
            mergeGo.BackColor = Color.FromArgb(46, 76, 196);
            mergeGo.ForeColor = Color.White;
            mergeGo.Cursor = Cursors.Hand;
            mergeGo.Click += new EventHandler(delegate(object s, EventArgs ev) { RunMergeUi(); });
            p.Controls.Add(mergeGo);
            p.Controls.Add(SmallBtn("打开底稿", 880, 162, 90, delegate
            {
                if (mergeOutput != null && File.Exists(mergeOutput)) OpenInExplorer(mergeOutput);
            }));

            mergeStatus = new Label();
            mergeStatus.Location = new Point(24, 196);
            mergeStatus.Size = new Size(940, 66);
            mergeStatus.ForeColor = Color.FromArgb(90, 93, 105);
            mergeStatus.Text = "各文件列序可不同，按表头名匹配进统一底稿；校验页自动勾稽（各文件小计和 = 底稿合计）。";
            p.Controls.Add(mergeStatus);

            // ---------- 下半：发票提取 ----------
            Label cap2 = new Label();
            cap2.Text = "发票批量提取（PDF 文本层直提 → 标准表 + 单票勾稽；扫描件需 OCR，另行处理）";
            cap2.Font = new Font("Microsoft YaHei UI", 11F, FontStyle.Bold);
            cap2.Location = new Point(24, 276);
            cap2.AutoSize = true;
            p.Controls.Add(cap2);

            invoiceFiles = new ListBox();
            invoiceFiles.Location = new Point(24, 306);
            invoiceFiles.Size = new Size(420, 110);
            p.Controls.Add(invoiceFiles);
            p.Controls.Add(SmallBtn("添加 PDF", 454, 306, 90, delegate
            {
                using (OpenFileDialog dlg = new OpenFileDialog())
                {
                    dlg.Title = "选择发票 PDF（可多选）";
                    dlg.Multiselect = true;
                    dlg.Filter = "PDF|*.pdf|全部文件|*.*";
                    if (dlg.ShowDialog(this) == DialogResult.OK)
                        foreach (string s in dlg.FileNames) invoiceFiles.Items.Add(s);
                }
            }));
            p.Controls.Add(SmallBtn("添加文件夹", 454, 340, 90, delegate
            {
                using (FolderBrowserDialog dlg = new FolderBrowserDialog())
                {
                    dlg.Description = "选择含发票 PDF 的文件夹";
                    if (dlg.ShowDialog(this) == DialogResult.OK)
                    {
                        foreach (string s in Directory.GetFiles(dlg.SelectedPath, "*.pdf"))
                            invoiceFiles.Items.Add(s);
                    }
                }
            }));
            p.Controls.Add(SmallBtn("移除所选", 454, 374, 90, delegate
            {
                if (invoiceFiles.SelectedIndex >= 0) invoiceFiles.Items.RemoveAt(invoiceFiles.SelectedIndex);
            }));
            p.Controls.Add(SmallBtn("清空", 454, 408, 90, delegate { invoiceFiles.Items.Clear(); }));

            Button invGo = new Button();
            invGo.Text = "开始提取";
            invGo.Location = new Point(560, 320);
            invGo.Size = new Size(110, 34);
            invGo.FlatStyle = FlatStyle.Flat;
            invGo.FlatAppearance.BorderSize = 0;
            invGo.BackColor = Color.FromArgb(46, 76, 196);
            invGo.ForeColor = Color.White;
            invGo.Cursor = Cursors.Hand;
            invGo.Click += new EventHandler(delegate(object s, EventArgs ev) { RunInvoiceUi(); });
            p.Controls.Add(invGo);
            p.Controls.Add(SmallBtn("打开清单", 680, 324, 90, delegate
            {
                if (invoiceOutput != null && File.Exists(invoiceOutput)) OpenInExplorer(invoiceOutput);
            }));

            invoiceStatus = new Label();
            invoiceStatus.Location = new Point(560, 366);
            invoiceStatus.Size = new Size(460, 120);
            invoiceStatus.ForeColor = Color.FromArgb(90, 93, 105);
            invoiceStatus.Text = "提取字段：发票号码 / 日期 / 购销双方及税号 / 金额 / 税额 / 价税合计。\n自动检查：金额 + 税额 是否等于 价税合计；失败票进「未识别清单」。";
            p.Controls.Add(invoiceStatus);
            return p;
        }

        void RefreshMergeTplCombo()
        {
            mergeTpl.Items.Clear();
            foreach (MergeTemplate t in MergeTemplateStore.LoadAll()) mergeTpl.Items.Add(t.Name);
        }

        static int ParseInt(string s, int def)
        {
            int n;
            return int.TryParse((s ?? "").Trim(), out n) && n >= 1 ? n : def;
        }

        void RunMergeUi()
        {
            if (mergeFiles.Items.Count == 0) { mergeStatus.Text = "请先添加要汇总的文件。"; return; }
            string[] targets = SplitTokensLocal(mergeTargets.Text);
            string[] srcs = SplitTokensLocal(mergeSrcs.Text);
            if (targets.Length == 0 || srcs.Length != targets.Length)
            {
                mergeStatus.Text = "目标列与源列需一一对应（逗号分隔）。";
                return;
            }
            List<string> files = new List<string>();
            foreach (object o in mergeFiles.Items) files.Add(o.ToString());
            MergeTemplate tpl = MergeTemplate.FromForm("__ui", mergeSheet.Text.Trim(),
                ParseInt(mergeHdr.Text, 1), targets, srcs, SplitTokensLocal(mergeAmounts.Text));
            string outPath = Path.Combine(Path.GetDirectoryName(Path.GetFullPath(files[0])),
                "汇总底稿_" + DateTime.Now.ToString("yyyyMMdd_HHmm", CultureInfo.InvariantCulture) + ".xlsx");
            mergeStatus.Text = "汇总中…";
            Thread t = new Thread(new ThreadStart(delegate
            {
                ConvertEngine conv = new ConvertEngine();
                conv.SofficePath = ConvertEngine.FindSoffice(EnvDetect.FindRoot());
                MergeEngine engine = new MergeEngine();
                MergeResult r = engine.Run(tpl, files, outPath, conv);
                AuditLog.Record("merge_run", "files=" + files.Count + "; rows=" + r.MergedRows + "; err=" + (r.Err ?? "none"));
                if (r.Err != null) { SetLabel(mergeStatus, "✗ " + r.Err); return; }
                if (File.Exists(outPath)) AuditLog.Record("file_write", outPath);
                mergeOutput = outPath;
                StringBuilder sb = new StringBuilder();
                sb.Append("完成（").Append(r.ElapsedMs).Append(" ms）：底稿 ").Append(r.MergedRows).Append(" 行；来自 ").Append(r.Files.Count).Append(" 个文件\n");
                foreach (ReconCheck c in r.Checks)
                    sb.Append(c.Warn ? "[提醒] " : (c.Pass ? "[✓] " : "[✗] ")).Append(PlainTips.PlainCheck(c.Name)).Append("  ").Append(c.Detail).Append("\n");
                sb.Append("底稿：").Append(outPath);
                SetLabel(mergeStatus, sb.ToString());
            }));
            t.IsBackground = true;
            t.Start();
        }

        void RunInvoiceUi()
        {
            if (invoiceFiles.Items.Count == 0) { invoiceStatus.Text = "请先添加发票 PDF。"; return; }
            List<string> files = new List<string>();
            foreach (object o in invoiceFiles.Items) files.Add(o.ToString());
            string outPath = Path.Combine(Path.GetDirectoryName(Path.GetFullPath(files[0])),
                "发票清单_" + DateTime.Now.ToString("yyyyMMdd_HHmm", CultureInfo.InvariantCulture) + ".xlsx");
            SetLabel(invoiceStatus, "提取中…");
            Thread t = new Thread(new ThreadStart(delegate
            {
                InvoiceEngine engine = new InvoiceEngine();
                InvoiceResult r = engine.Run(files, outPath);
                AuditLog.Record("invoice_run", "files=" + files.Count + "; ok=" + r.Rows.Count +
                    "; fail=" + r.Failures.Count + "; err=" + (r.Err ?? "none"));
                if (r.Err != null) { SetLabel(invoiceStatus, "✗ " + r.Err); return; }
                if (File.Exists(outPath)) AuditLog.Record("file_write", outPath);
                invoiceOutput = outPath;
                SetLabel(invoiceStatus, "完成（" + r.ElapsedMs + " ms）：成功 " + r.Rows.Count +
                    " 票；失败 " + r.Failures.Count + " 票。\n清单：" + outPath);
            }));
            t.IsBackground = true;
            t.Start();
        }

        delegate void SetLabelD(Label l, string s);

        void SetLabel(Label l, string s)
        {
            if (InvokeRequired) { Invoke(new SetLabelD(SetLabel), l, s); return; }
            l.Text = s;
        }
    }
}
