// ModelMenu —— 模型选择弹层（0.9.0）
//
// 为什么自建而不是用 ContextMenu：ContextMenu 做不了"按密钥分组 + 组标题层级"。
//
// 版式按用户给的"理想效果"图（2026-10-05）：
//   * 深色底、浅色字；
//   * **组标题**=该 API Key 的组名，小号、弱化颜色、左对齐；
//   * 组内模型**单列纵向**排列，字号更大、颜色更亮；
//   * 组与组之间一条细分隔线；
//   * 条目可带一行弱化的说明文字（右对齐截断）。
//
// 红线：C# 3.0 语法；只在 UI 线程使用。
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
            public string Desc = "";     // 可选说明（弱化显示在右侧；空则不占位）
            public bool KeepUrl = true;  // true=留在当前服务（中转/自定义）；false=按预置映射切端点
            public bool Checked = false; // 当前选中项
        }

        public class Group
        {
            public string Host = "";
            public string Tag = "";      // "" = 主账号
            public string Title = "";    // 组标题 = 该 Key 的组名（用户自定义 / "号N" / 主机名）
            public string Sub = "";      // 副标题（主机名）；与 Title 相同则不显示
            public bool CanRename = false;
            public List<Item> Items = new List<Item>();
        }

        // 浅色配色（与应用整体一致；版式层级照"理想效果"图）
        static readonly Color Bg = Color.White;
        static readonly Color BgHover = Color.FromArgb(242, 245, 252);
        static readonly Color Border = Color.FromArgb(214, 218, 228);
        static readonly Color HeadFg = Color.FromArgb(146, 150, 168);   // 组标题：弱化
        static readonly Color ItemFg = Color.FromArgb(38, 41, 51);      // 模型名：正文色
        static readonly Color ItemOn = Color.FromArgb(62, 99, 221);     // 选中
        static readonly Color DescFg = Color.FromArgb(168, 172, 188);   // 说明：更弱
        static readonly Color Line = Color.FromArgb(234, 236, 242);
        static readonly Color LinkFg = Color.FromArgb(62, 99, 221);
        static readonly Color FootBg = Color.FromArgb(250, 250, 252);

        const int PadX = 14;
        const int Width_ = 336;    // 弹层宽（与图接近）
        const int HeadH = 24;      // 组标题高
        const int ItemH = 32;      // 模型条目高
        const int SepH = 13;       // 分隔区高

        public static void Show(IWin32Window owner, Control anchor, List<Group> groups,
            Action<Item> onPick, Action<Group> onRename, Action onRefresh, Action onSettings)
        {
            if (groups == null) groups = new List<Group>();

            Form pop = new Form();
            pop.FormBorderStyle = FormBorderStyle.None;
            pop.ShowInTaskbar = false;
            pop.StartPosition = FormStartPosition.Manual;
            pop.TopMost = true;
            pop.BackColor = Bg;
            pop.KeyPreview = true;
            try { pop.Font = new Font("Microsoft YaHei UI", 9F); } catch { }
            // 细边框：无边框窗体在浅色背景上容易"糊"掉
            pop.Padding = new Padding(1);
            pop.Paint += delegate(object s, PaintEventArgs e)
            {
                try
                {
                    using (Pen p = new Pen(Border))
                        e.Graphics.DrawRectangle(p, 0, 0, pop.ClientSize.Width - 1, pop.ClientSize.Height - 1);
                }
                catch { }
            };

            Panel body = new Panel();
            body.BackColor = Bg;
            body.AutoScroll = true;

            // ★ 弹层必须**模态**（ShowDialog）：调用方在 Show 之后会立刻 FocusInput() 把焦点
            //   还给输入框——若用非模态 Show()，弹层刚显示就被抢焦点 → Deactivate → 自关，
            //   表现为"模型菜单/模型设置打不开"（0.9.2 回归，已修）。
            //   因此所有回调都**先记录、关闭后再执行**，避免在模态循环里嵌套开新窗体。
            Item picked = null;      // 选中的模型
            Group renameTarget = null;   // 要改名的组
            int act = 0;             // 0=无 1=刷新 2=打开设置

            int y = 8;
            for (int gi = 0; gi < groups.Count; gi++)
            {
                Group g = groups[gi];

                // ---- 组标题（该 Key 的组名）----
                Label hd = new Label();
                hd.AutoSize = false;
                hd.Location = new Point(PadX, y);
                hd.Size = new Size(Width_ - PadX * 2 - 40, HeadH - 6);
                hd.Text = g.Title;
                hd.ForeColor = HeadFg;
                try { hd.Font = new Font("Microsoft YaHei UI", 7.5F); } catch { }
                hd.TextAlign = ContentAlignment.MiddleLeft;
                body.Controls.Add(hd);

                // 组名旁边弱化显示主机名（组名已等于主机名时不重复）
                if (g.Sub != null && g.Sub.Length > 0 && g.Sub != g.Title)
                {
                    Label sb = new Label();
                    sb.AutoSize = false;
                    sb.Location = new Point(PadX + hd.Width, y);
                    sb.Size = new Size(Width_ - PadX * 2 - hd.Width, HeadH - 6);
                    sb.Text = g.Sub;
                    sb.ForeColor = DescFg;
                    try { sb.Font = new Font("Microsoft YaHei UI", 7.5F); } catch { }
                    sb.TextAlign = ContentAlignment.MiddleRight;
                    body.Controls.Add(sb);
                }

                if (g.CanRename && onRename != null)
                {
                    LinkLabel rn = new LinkLabel();
                    rn.Text = "改名";
                    rn.AutoSize = true;
                    rn.Location = new Point(Width_ - PadX - 30, y + 3);
                    rn.LinkColor = LinkFg;
                    rn.LinkBehavior = LinkBehavior.HoverUnderline;
                    Group capturedG = g;
                    rn.Click += delegate { renameTarget = capturedG; try { pop.Close(); } catch { } };
                    body.Controls.Add(rn);
                }

                y += HeadH;

                // ---- 组内模型（单列）----
                for (int i = 0; i < g.Items.Count; i++)
                {
                    Item it = g.Items[i];
                    Label cell = new Label();
                    cell.AutoSize = false;
                    cell.Location = new Point(PadX, y);
                    cell.Size = new Size(Width_ - PadX * 2, ItemH - 4);
                    cell.Text = it.Text;
                    cell.TextAlign = ContentAlignment.MiddleLeft;
                    cell.Cursor = Cursors.Hand;
                    cell.BackColor = Bg;
                    cell.ForeColor = it.Checked ? ItemOn : ItemFg;
                    try { cell.Font = new Font("Microsoft YaHei UI", 9.5F, it.Checked ? FontStyle.Bold : FontStyle.Regular); } catch { }

                    if (it.Desc != null && it.Desc.Length > 0)
                    {
                        Label ds = new Label();
                        ds.AutoSize = false;
                        ds.Location = new Point(PadX + 120, y);
                        ds.Size = new Size(Width_ - PadX * 2 - 120, ItemH - 4);
                        ds.Text = it.Desc;
                        ds.ForeColor = DescFg;
                        try { ds.Font = new Font("Microsoft YaHei UI", 8F); } catch { }
                        ds.TextAlign = ContentAlignment.MiddleRight;
                        ds.Cursor = Cursors.Hand;
                        ds.BackColor = Bg;
                        Item cap2 = it;
                        Form of2 = pop;
                        ds.Click += delegate { picked = cap2; try { of2.Close(); } catch { } };
                        ds.MouseEnter += delegate { try { ds.BackColor = BgHover; } catch { } };
                        ds.MouseLeave += delegate { try { ds.BackColor = Bg; } catch { } };
                        body.Controls.Add(ds);
                    }

                    Item captured = it;
                    Form ownerForm = pop;
                    cell.Click += delegate
                    {
                        picked = captured;
                        try { ownerForm.Close(); } catch { }
                    };
                    cell.MouseEnter += delegate { try { cell.BackColor = BgHover; } catch { } };
                    cell.MouseLeave += delegate { try { cell.BackColor = Bg; } catch { } };
                    body.Controls.Add(cell);

                    y += ItemH;
                }

                // ---- 组分隔线（最后一组不画）----
                if (gi < groups.Count - 1)
                {
                    Panel line = new Panel();
                    line.Location = new Point(PadX, y + 6);
                    line.Size = new Size(Width_ - PadX * 2, 1);
                    line.BackColor = Line;
                    body.Controls.Add(line);
                    y += SepH;
                }
            }
            y += 4;

            // ---- 底部操作 ----
            Panel foot = new Panel();
            foot.Height = 32;
            foot.BackColor = FootBg;

            LinkLabel lkRefresh = new LinkLabel();
            lkRefresh.Text = "从服务刷新模型列表…";
            lkRefresh.AutoSize = true;
            lkRefresh.Location = new Point(PadX, 8);
            lkRefresh.LinkColor = LinkFg;
            lkRefresh.Click += delegate { act = 1; try { pop.Close(); } catch { } };
            foot.Controls.Add(lkRefresh);

            LinkLabel lkSet = new LinkLabel();
            lkSet.Text = "模型设置…";
            lkSet.AutoSize = true;
            lkSet.Location = new Point(Width_ - PadX - 72, 8);
            lkSet.LinkColor = LinkFg;
            lkSet.Click += delegate { act = 2; try { pop.Close(); } catch { } };
            foot.Controls.Add(lkSet);

            int bodyH = y;
            int maxH = Screen.FromControl(anchor).WorkingArea.Height - 90;
            if (bodyH > maxH) bodyH = maxH;
            body.Size = new Size(Width_ - 2, bodyH);
            body.Location = new Point(1, 1);

            pop.ClientSize = new Size(Width_, bodyH + foot.Height + 2);
            foot.Dock = DockStyle.Bottom;
            pop.Controls.Add(body);
            pop.Controls.Add(foot);

            // 定位：优先在锚点上方（输入框上沿），放不下转下方
            Point sp = anchor.PointToScreen(new Point(0, 0));
            Rectangle wa = Screen.FromControl(anchor).WorkingArea;
            int px = sp.X;
            int py = sp.Y - pop.Height - 2;
            if (py < wa.Top) py = sp.Y + anchor.Height + 2;
            if (px + pop.Width > wa.Right) px = wa.Right - pop.Width;
            if (px < wa.Left) px = wa.Left;
            if (py + pop.Height > wa.Bottom) py = Math.Max(wa.Top, wa.Bottom - pop.Height);
            pop.Location = new Point(px, py);

            pop.Deactivate += delegate { try { pop.Close(); } catch { } };
            pop.KeyDown += delegate(object s, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Escape) { try { pop.Close(); } catch { } }
            };

            // ★ 模态显示：ShowDialog 会阻塞到弹层关闭为止。
            //   调用方（ChatPanel.ShowModelMenu）在 Show 之后紧跟着 FocusInput()，
            //   非模态时那一步会立刻抢焦点把弹层关掉（0.9.2 回归根因）。
            //   owner 取**顶层窗体**（而非 UserControl），模态才会正确禁用主窗体。
            Form top = owner as Form;
            if (top == null && anchor != null) { try { top = anchor.FindForm(); } catch { } }
            if (top != null) pop.ShowDialog(top); else pop.ShowDialog();

            // 弹层已关闭，此时再执行后续动作（避免在模态循环里嵌套开新窗体）
            if (picked != null)
            {
                if (onPick != null) { try { onPick(picked); } catch { } }
            }
            else if (renameTarget != null)
            {
                if (onRename != null) { try { onRename(renameTarget); } catch { } }
            }
            else if (act == 1)
            {
                if (onRefresh != null) { try { onRefresh(); } catch { } }
            }
            else if (act == 2)
            {
                if (onSettings != null) { try { onSettings(); } catch { } }
            }
        }

        // 简易输入框（深色，与弹层一致）：返回新名称；取消返回 null
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
                f.ClientSize = new Size(400, 152);
                f.BackColor = Bg;
                try { f.Font = new Font("Microsoft YaHei UI", 9F); } catch { }

                Label lb = new Label();
                lb.Text = hint == null ? "给它起个名字（留空则恢复默认）：" : hint;
                lb.Location = new Point(18, 18);
                lb.AutoSize = true;
                lb.ForeColor = HeadFg;
                f.Controls.Add(lb);

                TextBox tb = new TextBox();
                tb.Location = new Point(20, 48);
                tb.Size = new Size(356, 26);
                tb.Text = initial == null ? "" : initial;
                tb.BorderStyle = BorderStyle.FixedSingle;
                tb.SelectAll();
                f.Controls.Add(tb);

                Button ok = new Button();
                ok.Text = "确定";
                ok.Location = new Point(196, 98);
                ok.Size = new Size(86, 30);
                ok.Cursor = Cursors.Hand;
                ok.FlatStyle = FlatStyle.Flat;
                ok.BackColor = Color.FromArgb(62, 99, 221);
                ok.ForeColor = Color.White;
                ok.FlatAppearance.BorderSize = 0;
                ok.Click += delegate { result = tb.Text == null ? "" : tb.Text.Trim(); f.Close(); };
                f.Controls.Add(ok);

                Button cancel = new Button();
                cancel.Text = "取消";
                cancel.Location = new Point(290, 98);
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
