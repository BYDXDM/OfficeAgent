// ModelMenu —— 模型选择弹层（0.9.0）
//
// 为什么要自建弹层而不是用 ContextMenu：ContextMenu 只能一列平铺，
// 既做不了"按密钥槽分组"，也做不了多列。用户诉求（2026-10-05）：
//   ① 按 API Key 分组显示模型，同组内的模型排在一起；
//   ② 同一供应商配多个 Key 时，用户能**自己给每个 Key 起名**（不再固定显示"号1/号2"）；
//   ③ 支持多列显示。
//
// 结构：每个 Group = 一个密钥槽（host + tag）。组标题 = 自定义名 / "号<tag>" / 主机名。
// 组内模型按固定列数网格排布；组与组纵向堆叠。
//
// 红线：C# 3.0 语法；本文件只在 UI 线程使用。
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace OfficeAgent.Host
{
    static class ModelMenu
    {
        public class Item
        {
            public string Id = "";       // 原始模型 id（含 #tag）
            public string Text = "";     // 显示文本
            public bool KeepUrl = true;  // true=留在当前服务（中转/自定义）；false=按预置映射切端点
            public bool Checked = false; // 当前选中项
        }

        public class Group
        {
            public string Host = "";
            public string Tag = "";      // "" = 主账号
            public string Title = "";    // 组标题（已含自定义名或"号N"）
            public string Sub = "";      // 副标题（主机名）
            public bool CanRename = false;
            public List<Item> Items = new List<Item>();
        }

        const int ColW = 208;      // 单元格宽
        const int RowH = 26;       // 单元格高
        const int HeadH = 34;      // 组标题高
        const int GapH = 8;        // 组间距
        const int PadX = 12;
        const int Cols = 3;        // 固定 3 列（用户要求"多列显示"）

        public static void Show(IWin32Window owner, Control anchor, List<Group> groups,
            Action<Item> onPick, Action<Group> onRename, Action onRefresh, Action onSettings)
        {
            if (groups == null) groups = new List<Group>();

            Form pop = new Form();
            pop.FormBorderStyle = FormBorderStyle.None;
            pop.ShowInTaskbar = false;
            pop.StartPosition = FormStartPosition.Manual;
            pop.TopMost = true;
            pop.BackColor = Color.White;
            pop.KeyPreview = true;
            try { pop.Font = new Font("Microsoft YaHei UI", 9F); } catch { }

            int width = PadX * 2 + ColW * Cols;

            // 纵向布局：每组 = 标题 + ceil(n/Cols) 行
            int y = 8;
            Panel body = new Panel();
            body.AutoScroll = true;
            body.BackColor = Color.White;

            foreach (Group g in groups)
            {
                // ---- 组标题 ----
                Label hd = new Label();
                hd.AutoSize = false;
                hd.Location = new Point(PadX, y);
                hd.Size = new Size(width - PadX * 2 - 60, 20);
                hd.Text = g.Title;
                try { hd.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold); } catch { }
                hd.ForeColor = Color.FromArgb(40, 44, 56);
                body.Controls.Add(hd);

                if (g.Sub != null && g.Sub.Length > 0)
                {
                    Label sb = new Label();
                    sb.AutoSize = false;
                    sb.Location = new Point(PadX + hd.Width + 6, y + 2);
                    sb.Size = new Size(200, 16);
                    sb.Text = g.Sub;
                    sb.ForeColor = Color.FromArgb(150, 153, 168);
                    try { sb.Font = new Font("Microsoft YaHei UI", 7.5F); } catch { }
                    body.Controls.Add(sb);
                }

                if (g.CanRename && onRename != null)
                {
                    LinkLabel rn = new LinkLabel();
                    rn.Text = "改名";
                    rn.AutoSize = true;
                    rn.Location = new Point(width - PadX - 34, y + 1);
                    rn.LinkColor = Color.FromArgb(62, 99, 221);
                    rn.LinkBehavior = LinkBehavior.HoverUnderline;
                    Group captured = g;
                    rn.Click += delegate { try { onRename(captured); } catch { } };
                    body.Controls.Add(rn);
                }

                // 分隔线
                Panel line = new Panel();
                line.Location = new Point(PadX, y + 22);
                line.Size = new Size(width - PadX * 2, 1);
                line.BackColor = Color.FromArgb(232, 234, 240);
                body.Controls.Add(line);

                y += HeadH;

                // ---- 组内模型（多列网格） ----
                for (int i = 0; i < g.Items.Count; i++)
                {
                    int r = i / Cols, c = i % Cols;
                    Item it = g.Items[i];

                    Label cell = new Label();
                    cell.AutoSize = false;
                    cell.Location = new Point(PadX + c * ColW, y + r * RowH);
                    cell.Size = new Size(ColW - 6, RowH - 2);
                    cell.Text = (it.Checked ? "✓ " : "   ") + it.Text;
                    cell.TextAlign = ContentAlignment.MiddleLeft;
                    cell.Cursor = Cursors.Hand;
                    cell.ForeColor = it.Checked ? Color.FromArgb(62, 99, 221) : Color.FromArgb(38, 41, 51);
                    cell.BackColor = Color.White;
                    Item captured = it;
                    Form ownerForm = pop;
                    cell.Click += delegate
                    {
                        try { ownerForm.Close(); } catch { }
                        if (onPick != null) onPick(captured);
                    };
                    // 悬停高亮（WinForms 无 :hover，用 Enter/Leave 模拟）
                    cell.MouseEnter += delegate { try { cell.BackColor = Color.FromArgb(242, 245, 252); } catch { } };
                    cell.MouseLeave += delegate { try { cell.BackColor = Color.White; } catch { } };
                    body.Controls.Add(cell);
                }

                int rows = (g.Items.Count + Cols - 1) / Cols;
                if (rows < 1) rows = 1;
                y += rows * RowH + GapH;
            }

            // ---- 底部操作 ----
            Panel foot = new Panel();
            foot.Height = 34;
            foot.BackColor = Color.FromArgb(250, 250, 252);
            LinkLabel lkRefresh = new LinkLabel();
            lkRefresh.Text = "从服务刷新模型列表…";
            lkRefresh.AutoSize = true;
            lkRefresh.Location = new Point(PadX, 9);
            lkRefresh.LinkColor = Color.FromArgb(62, 99, 221);
            lkRefresh.Click += delegate { try { pop.Close(); } catch { } if (onRefresh != null) onRefresh(); };
            foot.Controls.Add(lkRefresh);

            LinkLabel lkSet = new LinkLabel();
            lkSet.Text = "打开模型设置…";
            lkSet.AutoSize = true;
            lkSet.Location = new Point(width - PadX - 104, 9);
            lkSet.LinkColor = Color.FromArgb(62, 99, 221);
            lkSet.Click += delegate { try { pop.Close(); } catch { } if (onSettings != null) onSettings(); };
            foot.Controls.Add(lkSet);

            int bodyH = y + 6;
            int maxH = Screen.FromControl(anchor).WorkingArea.Height - 80;
            if (bodyH > maxH) bodyH = maxH;
            body.Size = new Size(width, bodyH);
            body.Location = new Point(0, 0);

            pop.ClientSize = new Size(width, bodyH + foot.Height);
            foot.Dock = DockStyle.Bottom;
            pop.Controls.Add(body);
            pop.Controls.Add(foot);

            // 定位：优先在锚点上方（输入框上沿），放不下则下方
            Point sp = anchor.PointToScreen(new Point(0, 0));
            int px = sp.X;
            int pyAbove = sp.Y - pop.Height - 2;
            int py = pyAbove >= 0 ? pyAbove : (sp.Y + anchor.Height + 2);
            Rectangle wa = Screen.FromControl(anchor).WorkingArea;
            if (px + pop.Width > wa.Right) px = wa.Right - pop.Width;
            if (px < wa.Left) px = wa.Left;
            if (py + pop.Height > wa.Bottom) py = Math.Max(wa.Top, wa.Bottom - pop.Height);
            pop.Location = new Point(px, py);

            // 点击别处 / Esc 关闭
            pop.Deactivate += delegate { try { pop.Close(); } catch { } };
            pop.KeyDown += delegate(object s, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Escape) { try { pop.Close(); } catch { } }
            };

            if (owner != null) pop.Show(owner); else pop.Show();
            pop.Activate();
        }

        // 简易输入框：返回新名称；取消返回 null
        public static string Prompt(IWin32Window owner, string title, string initial, string hint)
        {
            string result = null;
            using (Form f = new Form())
            {
                f.Text = title == null ? "重命名" : title;
                f.FormBorderStyle = FormBorderStyle.FixedDialog;
                f.MinimizeBox = false;
                f.MaximizeBox = false;
                f.ShowInTaskbar = false;
                f.StartPosition = FormStartPosition.CenterParent;
                f.ClientSize = new Size(400, 150);
                try { f.Font = new Font("Microsoft YaHei UI", 9F); } catch { }

                Label lb = new Label();
                lb.Text = hint == null ? "给它起个名字（留空则恢复默认）：" : hint;
                lb.Location = new Point(18, 18);
                lb.AutoSize = true;
                lb.ForeColor = Color.FromArgb(90, 93, 105);
                f.Controls.Add(lb);

                TextBox tb = new TextBox();
                tb.Location = new Point(20, 46);
                tb.Size = new Size(356, 26);
                tb.Text = initial == null ? "" : initial;
                tb.SelectAll();
                f.Controls.Add(tb);

                Button ok = new Button();
                ok.Text = "确定";
                ok.Location = new Point(196, 96);
                ok.Size = new Size(86, 30);
                ok.Cursor = Cursors.Hand;
                ok.Click += delegate { result = tb.Text == null ? "" : tb.Text.Trim(); f.Close(); };
                f.Controls.Add(ok);

                Button cancel = new Button();
                cancel.Text = "取消";
                cancel.Location = new Point(290, 96);
                cancel.Size = new Size(86, 30);
                cancel.Cursor = Cursors.Hand;
                cancel.Click += delegate { result = null; f.Close(); };
                f.Controls.Add(cancel);

                f.AcceptButton = ok;
                f.CancelButton = cancel;
                if (owner != null) f.ShowDialog(owner); else f.ShowDialog();
            }
            return result;
        }
    }
}
