// WorkspacePicker —— 工作区位置选择器（带「磁盘」维度）
//
// 为什么需要它：既有的 FolderBrowserDialog 只能"在目录树里翻"，用户要换盘得自己找；
// 本选择器把「磁盘」提到最显眼的位置——选盘即自动落到 <盘>:\OfficeAgentFiles，
// 一键把工作区从 C 盘（系统盘）挪到 D 盘等数据盘。
//
// 附带收益：工作区默认落在 C 盘（文档\OfficeAgentFiles）正是"写 C 盘要确认"的高频来源；
// 引导用户选非系统盘可从源头减少安全确认弹窗。
//
// 红线：C# 3.0 语法；DriveInfo/DriveType 均为 .NET 2.0 API。
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace OfficeAgent.Host
{
    static class WorkspacePicker
    {
        // 默认子目录名（与 AppConfig.EffectiveWorkspace 的默认一致）
        public const string DefaultSubDir = "OfficeAgentFiles";

        // 枚举可用磁盘（固定盘 + 可移动盘，且已就绪），写入 cmb 并返回盘根列表（与 cmb 索引一一对应）
        public static List<string> FillDrives(ComboBox cmb)
        {
            List<string> roots = new List<string>();
            try { cmb.Items.Clear(); } catch { return roots; }
            string sysRoot = "";
            try { sysRoot = Path.GetPathRoot(Environment.SystemDirectory); } catch { }
            try
            {
                foreach (DriveInfo d in DriveInfo.GetDrives())
                {
                    if (d.DriveType != DriveType.Fixed && d.DriveType != DriveType.Removable) continue;
                    try { if (!d.IsReady) continue; } catch { continue; }
                    string root = "";
                    try { root = d.RootDirectory.FullName; } catch { continue; }
                    if (root == null || root.Length == 0) continue;

                    string label = root;
                    try { if (d.VolumeLabel != null && d.VolumeLabel.Length > 0) label += " " + d.VolumeLabel; } catch { }
                    try { label += "（可用 " + FormatSize(d.AvailableFreeSpace) + "）"; } catch { }
                    if (sysRoot != null && sysRoot.Length > 0 &&
                        string.Equals(root.TrimEnd('\\'), sysRoot.TrimEnd('\\'), StringComparison.OrdinalIgnoreCase))
                        label += "  [系统盘]";

                    roots.Add(root);
                    cmb.Items.Add(label);
                }
            }
            catch { }
            return roots;
        }

        public static string FormatSize(long bytes)
        {
            try
            {
                double gb = bytes / 1073741824.0;
                if (gb >= 1) return gb.ToString("0.0") + " GB";
                double mb = bytes / 1048576.0;
                return mb.ToString("0") + " MB";
            }
            catch { return "?"; }
        }

        // 在 roots 中找到 path 所在盘的索引；找不到返回 0
        public static int IndexOfDrive(List<string> roots, string path)
        {
            try
            {
                string root = Path.GetPathRoot(path);
                if (root == null) return 0;
                for (int i = 0; i < roots.Count; i++)
                    if (string.Equals(roots[i].TrimEnd('\\'), root.TrimEnd('\\'), StringComparison.OrdinalIgnoreCase))
                        return i;
            }
            catch { }
            return 0;
        }

        // 模态选择器：返回选定目录；取消返回 null
        public static string Pick(IWin32Window owner, string initialDir, string title)
        {
            string chosen = null;
            using (Form f = new Form())
            {
                f.Text = title == null || title.Length == 0 ? "选择工作区位置" : title;
                f.FormBorderStyle = FormBorderStyle.FixedDialog;
                f.MinimizeBox = false;
                f.MaximizeBox = false;
                f.ShowInTaskbar = false;
                f.StartPosition = FormStartPosition.CenterParent;
                f.ClientSize = new Size(540, 196);
                try { f.Font = new Font("Microsoft YaHei UI", 9F); } catch { }

                Label l1 = new Label();
                l1.Text = "磁盘：";
                l1.Location = new Point(22, 26);
                l1.AutoSize = true;
                f.Controls.Add(l1);

                ComboBox cmb = new ComboBox();
                cmb.DropDownStyle = ComboBoxStyle.DropDownList;
                cmb.Location = new Point(78, 22);
                cmb.Size = new Size(440, 26);
                f.Controls.Add(cmb);
                List<string> roots = FillDrives(cmb);

                Label l2 = new Label();
                l2.Text = "文件夹：";
                l2.Location = new Point(22, 68);
                l2.AutoSize = true;
                f.Controls.Add(l2);

                TextBox tb = new TextBox();
                tb.Location = new Point(78, 64);
                tb.Size = new Size(344, 26);
                f.Controls.Add(tb);

                Button bBrowse = new Button();
                bBrowse.Text = "浏览…";
                bBrowse.Location = new Point(430, 63);
                bBrowse.Size = new Size(88, 28);
                bBrowse.Cursor = Cursors.Hand;
                bBrowse.Click += delegate
                {
                    using (FolderBrowserDialog fb = new FolderBrowserDialog())
                    {
                        fb.Description = "选择工作区文件夹";
                        try { fb.SelectedPath = tb.Text; } catch { }
                        if (fb.ShowDialog(f) == DialogResult.OK) tb.Text = fb.SelectedPath;
                    }
                };
                f.Controls.Add(bBrowse);

                Label l3 = new Label();
                l3.Text = "建议选非系统盘（如 D 盘），可减少「写入 C 盘」的安全确认提示。";
                l3.ForeColor = Color.FromArgb(110, 113, 126);
                l3.Location = new Point(24, 100);
                l3.AutoSize = true;
                f.Controls.Add(l3);

                // 选盘 → 自动落到 <盘>\OfficeAgentFiles（仅当当前文本不在该盘时改写，避免覆盖用户已选目录）
                cmb.SelectedIndexChanged += delegate
                {
                    if (cmb.SelectedIndex < 0 || cmb.SelectedIndex >= roots.Count) return;
                    string root = roots[cmb.SelectedIndex];
                    string cur = tb.Text == null ? "" : tb.Text.Trim();
                    string curRoot = "";
                    try { curRoot = Path.GetPathRoot(cur); } catch { }
                    bool sameDrive = curRoot != null && curRoot.Length > 0 &&
                        string.Equals(curRoot.TrimEnd('\\'), root.TrimEnd('\\'), StringComparison.OrdinalIgnoreCase);
                    if (!sameDrive) tb.Text = Path.Combine(root, DefaultSubDir);
                };

                Button bCancel = new Button();
                bCancel.Text = "取消";
                bCancel.Location = new Point(22, 140);
                bCancel.Size = new Size(110, 32);
                bCancel.Cursor = Cursors.Hand;
                bCancel.Click += delegate { chosen = null; f.Close(); };
                f.Controls.Add(bCancel);

                Button bOk = new Button();
                bOk.Text = "使用此位置";
                bOk.Location = new Point(330, 140);
                bOk.Size = new Size(188, 32);
                bOk.Cursor = Cursors.Hand;
                bOk.Click += delegate
                {
                    string v = tb.Text == null ? "" : tb.Text.Trim();
                    if (v.Length == 0) { MessageBox.Show(f, "请先选择一个文件夹。", "提示"); return; }
                    chosen = v;
                    f.Close();
                };
                f.Controls.Add(bOk);

                // 初始值
                string init = initialDir;
                if (init == null || init.Trim().Length == 0)
                {
                    try { init = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), DefaultSubDir); } catch { init = ""; }
                }
                tb.Text = init;
                int idx = IndexOfDrive(roots, init);
                if (idx >= 0 && idx < cmb.Items.Count) cmb.SelectedIndex = idx;

                f.CancelButton = bCancel;
                f.AcceptButton = bOk;
                if (owner != null) f.ShowDialog(owner); else f.ShowDialog();
            }
            return chosen;
        }
    }
}
