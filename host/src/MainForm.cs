// 主窗体（M1.5）：WorkBuddy 风格布局 —— 左侧导航栏 + 主内容区（会话/任务台/预览/系统状态/审计）
// 首次启动弹出模型配置向导（SetupDialog）；导航自绘高亮；GDI+ 圆角与双缓冲。
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using OfficeAgent.Core;

namespace OfficeAgent.Host
{
    public partial class MainForm : Form
    {
        const string Version = "0.9.7";

        class NavEntry
        {
            public string Label; public int Page; public bool Header;
            public NavEntry(string label, int page, bool header) { Label = label; Page = page; Header = header; }
        }

        static readonly string PreviewDir = Path.Combine(Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeAgent"), "preview");

        string[] navItems = new string[] { "新建任务", "会话", "任务台", "表格核对", "汇总·发票", "预览", "系统状态", "审计", "内置插件" };
        NavEntry[] navEntries;      // 带分组头的导航条目（头不可点）
        int[] pageToNav;            // 页索引 → 导航列表索引
        bool navReentry = false;
        bool sessionReentry = false;
        ListBox sessionList;
        Button btnNewChat;
        int sessionHover = -1;
        bool showArchived = false;                  // 侧边栏是否显示已归档会话
        List<SessionInfo> sessionCache = new List<SessionInfo>();   // 列表数据快照（绘制/点击/右键共用一套索引）

        // 侧边栏「工作区（项目）」分组头：DSH 式，每个工作区挂自己的文件夹，会话按归属分组
        class WsHeader
        {
            public WorkspaceInfo Ws;
            public int Count;
            public bool Active;
        }
        string[] pageTitles = new string[] { "新任务", "会话", "任务台", "表格核对", "汇总·发票", "预览", "系统状态", "审计日志", "内置插件" };
        ListBox nav;
        Panel content; Label pageTitle;
        Panel[] pages;
        int currentPage = -1;
        int hoverIndex = -1;

        // 系统状态
        ListView lvEnv; Label lblEnvStatus; Label lblWorkspace;
        // 任务台
        ListView lvTasks; ComboBox cmbTarget; Label lblTaskStatus;
        // 预览
        Panel previewHost; Label previewInfo;
        DataGridView grid; PictureBox pdfBox; Label pdfPage; Button btnPrev; Button btnNext;
        // 引擎与配置
        ConvertEngine engine;
        AppConfig config;
        ChatPanel chat;
        List<string[]> previewRows = new List<string[]>();
        Pdfium.PdfDoc pdfDoc = null;
        int pdfPageIndex = 0;
        string tempPdf = null;
        string previewFolderPath = "";   // 当前预览文件的所在目录（「打开源文件所在地」按钮）
        delegate void VoidD();
        delegate void StrD(string s);
        delegate void FillEnvD(List<DetectItem> items);

        public MainForm()
        {
            Text = "OfficeAgent v" + Version;
            ClientSize = new Size(1120, 720);
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(940, 620);
            AllowDrop = true;
            DragEnter += new DragEventHandler(OnDragEnter);
            DragDrop += new DragEventHandler(OnDragDrop);
            BackColor = Color.White;
            Font = new Font("Microsoft YaHei UI", 9F);

            config = AppConfig.Load();
            WorkspaceStore.EnsureInit(config);   // 工作区（项目）注册表：首启自动建默认工作区
            engine = new ConvertEngine();
            engine.SofficePath = ConvertEngine.FindSoffice(EnvDetect.FindRoot());

            BuildSidebar();
            pageTitle = new Label();
            pageTitle.Font = new Font("Microsoft YaHei UI", 14F, FontStyle.Bold);
            pageTitle.Location = new Point(24, 14);
            pageTitle.AutoSize = true;
            Panel header = new Panel();
            header.Dock = DockStyle.Top;
            header.Height = 52;
            header.BackColor = Color.White;
            header.Controls.Add(pageTitle);

            content = new Panel();
            content.Dock = DockStyle.Fill;
            content.BackColor = Color.White;

            pages = new Panel[] { BuildNewTaskPage(), BuildChatPage(), BuildTaskPage(), BuildReconPage(), BuildExtractPage(), BuildPreviewPage(), BuildStatusPage(), BuildAuditPage(), BuildPluginsPage() };
            foreach (Panel p in pages) content.Controls.Add(p);

            Controls.Add(content);
            Controls.Add(header);
            Controls.Add(sidebar);   // 最后加入 = 停靠优先占左侧条，header 占剩余顶部，content 填充其余

            FormClosed += new FormClosedEventHandler(OnClosed);
            Shown += new EventHandler(OnShown);
            SelectPage(1); // 默认会话
        }

        // ================= 侧边栏 =================

        Panel sidebar;

        void BuildSidebar()
        {
            sidebar = new Panel();
            sidebar.Dock = DockStyle.Left;
            sidebar.Width = 228;
            sidebar.BackColor = Color.FromArgb(244, 244, 246);
            sidebar.Paint += Sidebar_Paint;

            Label brand = new Label();
            brand.Text = "OfficeAgent";
            brand.Font = new Font("Microsoft YaHei UI", 13F, FontStyle.Bold);
            brand.ForeColor = Color.FromArgb(24, 26, 32);
            brand.Location = new Point(18, 16);
            brand.AutoSize = true;
            sidebar.Controls.Add(brand);

            Label ver = new Label();
            ver.Text = Version;
            ver.Font = new Font("Microsoft YaHei UI", 8F);
            ver.ForeColor = Color.FromArgb(150, 153, 168);
            ver.Location = new Point(20, 44);
            ver.AutoSize = true;
            sidebar.Controls.Add(ver);

            // 侧边栏导航：任务页签（工作区 = 会话列表 + 新建对话按钮；系统组 = 状态/审计/插件）
            List<NavEntry> es = new List<NavEntry>();
            for (int i = 2; i <= 5; i++) es.Add(new NavEntry(navItems[i], i, false));
            es.Add(new NavEntry("系统", -1, true));
            for (int i = 6; i < navItems.Length; i++) es.Add(new NavEntry(navItems[i], i, false));
            navEntries = es.ToArray();
            pageToNav = new int[navItems.Length];
            for (int i = 0; i < pageToNav.Length; i++) pageToNav[i] = -1;
            for (int i = 0; i < navEntries.Length; i++) if (!navEntries[i].Header) pageToNav[navEntries[i].Page] = i;

            nav = new SideList();
            nav.BorderStyle = BorderStyle.None;
            nav.BackColor = sidebar.BackColor;
            nav.DrawMode = DrawMode.OwnerDrawFixed;
            nav.ItemHeight = 40;
            nav.Font = new Font("Microsoft YaHei UI", 10F);
            nav.Location = new Point(10, 296);
            nav.Size = new Size(208, navEntries.Length * 40 + 8);
            foreach (NavEntry en in navEntries) nav.Items.Add(en.Label);
            nav.DrawItem += Nav_DrawItem;
            nav.SelectedIndexChanged += delegate(object s, EventArgs e)
            {
                int i = nav.SelectedIndex;
                if (i < 0 || i >= navEntries.Length || navReentry) return;
                if (navEntries[i].Header)
                {   // 分组头不可选中：弹回当前页
                    navReentry = true;
                    nav.SelectedIndex = pageToNav[Math.Max(currentPage, 0)];
                    navReentry = false;
                    return;
                }
                SelectPage(navEntries[i].Page);
            };
            nav.MouseMove += new MouseEventHandler(delegate(object s, MouseEventArgs e) {
                int idx = nav.IndexFromPoint(e.Location);
                if (idx != hoverIndex) { hoverIndex = idx; nav.Invalidate(); }
            });
            nav.MouseLeave += delegate { hoverIndex = -1; nav.Invalidate(); };
            sidebar.Controls.Add(nav);

            // ===== 工作区：新建对话 + 会话列表（参照 DSH 式工作区分组） =====
            btnNewChat = new Button();
            btnNewChat.Text = "＋ 新建对话";
            btnNewChat.Font = new Font("Microsoft YaHei UI", 9.75F, FontStyle.Bold);
            btnNewChat.ForeColor = Color.White;
            btnNewChat.BackColor = Color.FromArgb(24, 26, 32);
            btnNewChat.FlatStyle = FlatStyle.Flat;
            btnNewChat.FlatAppearance.BorderSize = 0;
            btnNewChat.Location = new Point(18, 68);
            btnNewChat.Size = new Size(192, 34);
            btnNewChat.Cursor = Cursors.Hand;
            btnNewChat.Click += delegate { chat.NewConversation(); SelectPage(1); chat.FocusInput(); sessionList.Invalidate(); };
            sidebar.Controls.Add(btnNewChat);

            sessionList = new SideList();
            sessionList.BorderStyle = BorderStyle.None;
            sessionList.BackColor = sidebar.BackColor;
            sessionList.DrawMode = DrawMode.OwnerDrawFixed;
            sessionList.ItemHeight = 34;
            sessionList.Font = new Font("Microsoft YaHei UI", 9F);
            sessionList.Location = new Point(10, 112);
            sessionList.Size = new Size(208, 172);
            sessionList.DrawItem += SessionList_DrawItem;
            sessionList.SelectedIndexChanged += delegate(object s, EventArgs e)
            {
                int i = sessionList.SelectedIndex;
                sessionReentry = true;
                sessionList.SelectedIndex = -1;
                sessionReentry = false;
                if (i < 0 || sessionReentry) return;
                if (i == 0) { SelectPage(1); return; }   // 首行 = 会话段头
                object item = sessionList.Items[i];
                WsHeader wh = item as WsHeader;
                if (wh != null) { SetActiveWorkspace(wh.Ws); return; }   // 点工作区头 = 切换项目
                SessionInfo si = item as SessionInfo;
                if (si == null) return;                  // 分隔行等不可点
                chat.LoadSession(si.Id); SelectPage(1);
                sessionList.Invalidate();   // 重画：选中高亮外边框要跟着当前会话走
            };
            sessionList.MouseMove += new MouseEventHandler(delegate(object s, MouseEventArgs e) {
                int idx = sessionList.IndexFromPoint(e.Location);
                if (idx != sessionHover) { sessionHover = idx; sessionList.Invalidate(); }
            });
            sessionList.MouseLeave += delegate { sessionHover = -1; sessionList.Invalidate(); };
            // 右键菜单：会话 → 归档/删除；工作区头 → 新建/打开/改址/删除；任意位置 → 已归档开关
            sessionList.MouseUp += new MouseEventHandler(delegate(object s, MouseEventArgs e)
            {
                if (e.Button != MouseButtons.Right) return;
                int idx = sessionList.IndexFromPoint(e.Location);
                object item = (idx >= 0 && idx < sessionList.Items.Count) ? sessionList.Items[idx] : null;
                SessionInfo si = item as SessionInfo;
                WsHeader wh = item as WsHeader;
                ContextMenu m = new ContextMenu();
                if (si != null)
                {
                    string archTxt = si.Archived ? "恢复会话" : "归档会话";
                    m.MenuItems.Add(new MenuItem(archTxt, delegate
                    {
                        SessionStore.SetArchived(si.Id, !si.Archived);
                        RefreshSessions();
                    }));
                    m.MenuItems.Add(new MenuItem("删除会话", delegate
                    {
                        if (MessageBox.Show(this, "删除会话「" + (si.Title.Length == 0 ? "（无标题）" : si.Title) + "」？\n删除后无法恢复。",
                            "删除会话", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
                        SessionStore.Delete(si.Id);
                        if (chat.CurrentSessionId == si.Id) chat.NewConversation();   // 删的是当前会话 → 清空聊天区
                        RefreshSessions();
                    }));
                }
                if (wh != null)
                {
                    m.MenuItems.Add(new MenuItem("新建工作区…", delegate { NewWorkspace(); }));
                    m.MenuItems.Add(new MenuItem("打开工作区文件夹", delegate { OpenFolder.Open(wh.Ws.Dir); }));
                    m.MenuItems.Add(new MenuItem("更改此工作区的文件夹…", delegate { ChangeWorkspaceFolder(wh.Ws); }));
                    m.MenuItems.Add(new MenuItem("删除此工作区（会话移回默认）", delegate { DeleteWorkspace(wh.Ws); }));
                }
                else if (si == null)
                {
                    m.MenuItems.Add(new MenuItem("新建工作区…", delegate { NewWorkspace(); }));
                }
                MenuItem archToggle = new MenuItem(showArchived ? "隐藏已归档会话" : "显示已归档会话", delegate
                {
                    showArchived = !showArchived;
                    RefreshSessions();
                });
                m.MenuItems.Add(archToggle);
                m.Show(sessionList, new Point(e.X, e.Y));
            });
            sidebar.Controls.Add(sessionList);
            RefreshSessions();

            LinkLabel settings = new LinkLabel();
            settings.Text = "模型设置";
            settings.LinkBehavior = LinkBehavior.NeverUnderline;
            settings.LinkColor = Color.FromArgb(70, 73, 84);
            settings.Font = new Font("Microsoft YaHei UI", 9.5F);
            settings.AutoSize = true;
            settings.Location = new Point(18, sidebar.Height - 66);
            settings.Click += delegate { ShowSetup(false); };
            sidebar.Controls.Add(settings);

            Label account = new Label();
            account.Text = "本地模式 · 数据不出本机";
            account.Font = new Font("Microsoft YaHei UI", 8F);
            account.ForeColor = Color.FromArgb(150, 153, 168);
            account.AutoSize = true;
            account.Location = new Point(18, sidebar.Height - 40);
            sidebar.Controls.Add(account);
            sidebar.Resize += delegate
            {
                settings.Location = new Point(18, sidebar.Height - 66);
                account.Location = new Point(18, sidebar.Height - 40);
            };
        }

        void Sidebar_Paint(object sender, PaintEventArgs e)
        {
            using (Pen p = new Pen(Color.FromArgb(232, 233, 238)))
                e.Graphics.DrawLine(p, sidebar.Width - 1, 0, sidebar.Width - 1, sidebar.Height);
        }

        // ===== 会话列表（DSH 式工作区）=====

        void RefreshSessions()
        {
            if (sessionList == null) return;
            try
            {
                WorkspaceStore.EnsureInit(config);
                sessionCache = SessionStore.List();
                WorkspaceInfo act = WorkspaceStore.Active(config);
                sessionList.BeginUpdate();
                sessionList.Items.Clear();
                sessionList.Items.Add("会话");
                // DSH 式分组：工作区头 + 归属会话；全部展开，点击头切换激活工作区
                foreach (WorkspaceInfo w in WorkspaceStore.All())
                {
                    bool isActive = act != null && w.Id == act.Id;
                    int n = 0;
                    foreach (SessionInfo s in sessionCache)
                    {
                        if (!s.Archived && s.WorkspaceId == w.Id) n++;
                    }
                    sessionList.Items.Add(new WsHeader { Ws = w, Count = n, Active = isActive });
                    foreach (SessionInfo s in sessionCache)
                    {
                        if (s.Archived || s.WorkspaceId != w.Id) continue;
                        sessionList.Items.Add(s);
                    }
                    // 已归档仅在激活工作区下按开关展开，避免每组都堆一段归档
                    if (isActive && showArchived)
                    {
                        int na = 0;
                        foreach (SessionInfo s in sessionCache)
                        {
                            if (s.Archived && s.WorkspaceId == w.Id) na++;
                        }
                        if (na > 0) sessionList.Items.Add("─── 已归档（" + na + "）───");
                        foreach (SessionInfo s in sessionCache)
                        {
                            if (s.Archived && s.WorkspaceId == w.Id) sessionList.Items.Add(s);
                        }
                    }
                }
                sessionList.EndUpdate();
            }
            catch { }
        }

        // 点击工作区头 → 切换激活（agent 文件操作、新会话归属随之切换）
        void SetActiveWorkspace(WorkspaceInfo w)
        {
            if (w == null) return;
            WorkspaceInfo act = WorkspaceStore.Active(config);
            if (act != null && act.Id == w.Id) return;
            try
            {
                try { Directory.CreateDirectory(w.Dir); } catch { }
                WorkspaceStore.SetActive(config, w.Id);
                chat.ApplyConfig(config);   // 同步 AgentTools.WorkspaceRoot 与系统提示词
                RefreshSessions();
                AuditLog.Record("config_change", "切换工作区=" + w.Name);
            }
            catch { }
        }

        // 新建工作区：选一个文件夹（带磁盘维度），名字取文件夹名
        void NewWorkspace()
        {
            string picked = WorkspacePicker.Pick(this, WorkspaceStore.ActiveDir(config), "选择新工作区的文件夹");
            if (picked == null) return;
            string name = System.IO.Path.GetFileName(picked);
            if (name == null || name.Trim().Length == 0) name = "新建工作区";
            WorkspaceInfo w = WorkspaceStore.Add(name, picked);
            AuditLog.Record("file_write", "工作区创建 " + w.Name + " → " + w.Dir);
            SetActiveWorkspace(w);
        }

        void ChangeWorkspaceFolder(WorkspaceInfo w)
        {
            if (w == null) return;
            string picked = WorkspacePicker.Pick(this, w.Dir, "更改工作区「" + w.Name + "」的位置");
            if (picked == null) return;
            WorkspaceStore.UpdateDir(w.Id, picked);
            AuditLog.Record("file_write", "工作区改址 " + w.Name + " → " + picked);
            if (WorkspaceStore.Active(config).Id == w.Id) chat.ApplyConfig(config);
            RefreshSessions();
        }

        void DeleteWorkspace(WorkspaceInfo w)
        {
            if (w == null) return;
            List<WorkspaceInfo> all = WorkspaceStore.All();
            if (all.Count <= 1) { MessageBox.Show(this, "至少要保留一个工作区。", "无法删除", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            if (MessageBox.Show(this, "删除工作区「" + w.Name + "」？\n其中的会话会移到「默认工作区」，文件夹和文件不会被删。",
                "删除工作区", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            WorkspaceInfo act = WorkspaceStore.Active(config);
            WorkspaceStore.Delete(w.Id);
            SessionStore.MoveWorkspace(w.Id, "default");
            if (act != null && act.Id == w.Id)
            {
                WorkspaceInfo first = WorkspaceStore.All()[0];
                WorkspaceStore.SetActive(config, first.Id);
                chat.ApplyConfig(config);
            }
            RefreshSessions();
            AuditLog.Record("config_change", "删除工作区=" + w.Name);
        }

        void SessionList_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            e.DrawBackground();
            Graphics g = e.Graphics;
            bool hover = e.Index == sessionHover;
            Rectangle pill = new Rectangle(6, e.Bounds.Y + 2, e.Bounds.Width - 12, e.Bounds.Height - 4);
            if (hover)
            {
                using (GraphicsPath gp = RoundRect(pill, 10))
                using (SolidBrush b = new SolidBrush(Color.FromArgb(237, 238, 242)))
                    g.FillPath(b, gp);
            }
            if (e.Index == 0)
            {
                TextRenderer.DrawText(g, "会话", sessionList.Font, new Point(pill.X + 14, pill.Y + 7), Color.FromArgb(70, 73, 84));
                return;
            }
            WsHeader wh = sessionList.Items[e.Index] as WsHeader;
            if (wh != null)
            {
                // 工作区头：激活的给浅蓝底；名字 + 会话数
                if (wh.Active)
                {
                    using (GraphicsPath gp = RoundRect(pill, 10))
                    using (SolidBrush b = new SolidBrush(Color.FromArgb(213, 228, 255)))
                        g.FillPath(b, gp);
                }
                string mark = wh.Active ? "● " : "○ ";
                TextRenderer.DrawText(g, mark + wh.Ws.Name, sessionList.Font,
                    new Rectangle(pill.X + 10, pill.Y + 2, pill.Width - 46, pill.Height - 4),
                    wh.Active ? Color.FromArgb(24, 60, 160) : Color.FromArgb(70, 73, 84),
                    TextFormatFlags.EndEllipsis | TextFormatFlags.VerticalCenter);
                TextRenderer.DrawText(g, wh.Count.ToString() + " 项", new Font("Microsoft YaHei UI", 8F),
                    new Rectangle(pill.Right - 48, pill.Y + 7, 42, 18), Color.FromArgb(150, 153, 168), TextFormatFlags.Right);
                return;
            }
            string sep = sessionList.Items[e.Index] as string;
            if (sep != null)
            {
                TextRenderer.DrawText(g, sep, new Font("Microsoft YaHei UI", 8F),
                    pill, Color.FromArgb(150, 153, 168), TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                return;
            }
            SessionInfo s = sessionList.Items[e.Index] as SessionInfo;
            if (s == null) return;
            // ★ 当前选中会话用外边框高亮包围（用户要求：范围要明显）。
            //   注意**不能**用 sessionList.SelectedIndex 判断：该列表在点击后会立刻把
            //   SelectedIndex 重置为 -1（为避免系统蓝条），所以它恒为 -1，
            //   0.8.4 加的选中边框因此从来没画出来过（死代码）。改为按**当前会话 id** 比。
            string curSid = null;
            try { curSid = chat == null ? null : chat.CurrentSessionId; } catch { }
            bool sel = curSid != null && curSid.Length > 0 && s.Id == curSid;
            if (sel)
            {
                using (GraphicsPath gp = RoundRect(pill, 10))
                {
                    using (SolidBrush b = new SolidBrush(Color.FromArgb(226, 236, 254))) g.FillPath(b, gp);
                    // 2px 实线边框：比 1.6px 更醒目，选中范围一眼可辨
                    using (Pen p = new Pen(Color.FromArgb(62, 99, 221), 2.0F)) g.DrawPath(p, gp);
                }
            }
            TextRenderer.DrawText(g, s.Title.Length == 0 ? "（无标题会话）" : s.Title, sessionList.Font,
                new Rectangle(pill.X + 26, pill.Y + 2, pill.Width - 88, pill.Height - 4),
                sel ? Color.FromArgb(24, 60, 160) : (s.Archived ? Color.FromArgb(150, 153, 168) : Color.FromArgb(70, 73, 84)),
                TextFormatFlags.EndEllipsis | TextFormatFlags.VerticalCenter);
            TextRenderer.DrawText(g, s.Archived ? "已归档" : RelTime(s.Updated), new Font("Microsoft YaHei UI", 8F),
                new Rectangle(pill.Right - 58, pill.Y + 7, 50, 18), Color.FromArgb(150, 153, 168), TextFormatFlags.Right);
        }

        // "yyyy-MM-dd HH:mm" → 今天显示 HH:mm / 昨天 / MM-dd
        static string RelTime(string updated)
        {
            try
            {
                DateTime t;
                if (!DateTime.TryParse(updated, out t)) return "";
                DateTime now = DateTime.Now;
                if (t.Date == now.Date) return t.ToString("HH:mm");
                if (t.Date == now.Date.AddDays(-1)) return "昨天";
                return t.ToString("MM-dd");
            }
            catch { return ""; }
        }

        void Nav_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= navEntries.Length) return;
            NavEntry en = navEntries[e.Index];
            e.DrawBackground();
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            if (en.Header)
            {
                // 分组头：小号灰字 + 分隔线，无选中/悬停态
                TextRenderer.DrawText(g, en.Label, new Font("Microsoft YaHei UI", 8F, FontStyle.Bold),
                    new Point(20, e.Bounds.Y + 10), Color.FromArgb(150, 153, 168));
                using (Pen p = new Pen(Color.FromArgb(228, 229, 234)))
                    g.DrawLine(p, 16, e.Bounds.Bottom - 7, e.Bounds.Width - 16, e.Bounds.Bottom - 7);
                return;
            }
            bool selected = currentPage >= 0 && en.Page == currentPage;
            bool hover = e.Index == hoverIndex;
            Rectangle pill = new Rectangle(6, e.Bounds.Y + 3, e.Bounds.Width - 12, e.Bounds.Height - 6);
            if (selected || hover)
            {
                using (GraphicsPath gp = RoundRect(pill, 18))
                using (SolidBrush b = new SolidBrush(selected ? Color.FromArgb(228, 233, 255) : Color.FromArgb(237, 238, 242)))
                    g.FillPath(b, gp);
            }
            Color textColor = selected ? Color.FromArgb(46, 76, 196) : Color.FromArgb(70, 73, 84);
            TextRenderer.DrawText(g, en.Label, nav.Font, new Point(pill.X + 16, pill.Y + 8), textColor);
            if (en.Page == 0)
            {
                // 新建任务的 "+" 徽标
                using (SolidBrush b = new SolidBrush(Color.FromArgb(46, 76, 196)))
                    g.FillEllipse(b, pill.Right - 34, pill.Y + pill.Height / 2 - 7, 14, 14);
                TextRenderer.DrawText(g, "+", new Font("Microsoft YaHei UI", 9F, FontStyle.Bold),
                    new Rectangle(pill.Right - 35, pill.Y + pill.Height / 2 - 9, 16, 18), Color.White,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            }
        }

        static GraphicsPath RoundRect(Rectangle r, int rad)
        {
            GraphicsPath gp = new GraphicsPath();
            gp.AddArc(r.X, r.Y, rad * 2, rad * 2, 180, 90);
            gp.AddArc(r.Right - rad * 2, r.Y, rad * 2, rad * 2, 270, 90);
            gp.AddArc(r.Right - rad * 2, r.Bottom - rad * 2, rad * 2, rad * 2, 0, 90);
            gp.AddArc(r.X, r.Bottom - rad * 2, rad * 2, rad * 2, 90, 90);
            gp.CloseFigure();
            return gp;
        }

        // 侧边栏列表通用双缓冲（悬停逐帧 Invalidate，无缓冲会闪）
        class SideList : ListBox
        {
            public SideList()
            {
                SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);
            }
        }

        // ================= 页面切换 =================

        bool envAutoDetected = false;

        void SelectPage(int idx)
        {
            if (idx < 0 || idx >= pages.Length) idx = 1;
            currentPage = idx;
            foreach (Panel p in pages) p.Visible = false;
            pages[idx].Visible = true;
            pageTitle.Text = pageTitles[idx];
            // 会话/新建任务不在页签导航里（由会话列表承担）→ 清除导航选中
            nav.SelectedIndex = pageToNav[idx] >= 0 ? pageToNav[idx] : -1;
            nav.Invalidate();
            // 切到会话页把焦点还给输入框（否则第一次打字全丢——焦点停在导航/列表上）
            if (idx <= 1 && chat != null) chat.FocusInput();
            if (idx == 6 && !envAutoDetected) { envAutoDetected = true; RunDetect(); }   // 低配机：自动检测每次会话只跑一次，可手动强制
        }

        void OnShown(object sender, EventArgs e)
        {
            AuditLog.Record("app_start", Version);
            RefreshAudit();
            RefreshSessions();
            chat.ApplyConfig(config);
            chat.RestoreLastSession();
            bool skipWizard = false;
            try { skipWizard = Environment.GetEnvironmentVariable("OFFICEAGENT_NO_WIZARD") == "1"; } catch { }
            if (!config.WizardDone && !skipWizard) ShowSetup(true);
        }

        void ShowSetup(bool firstRun)
        {
            using (SetupDialog dlg = new SetupDialog(config, firstRun))
            {
                dlg.ShowDialog(this);
                if (dlg.Saved)
                {
                    config = AppConfig.Load();
                    chat.ApplyConfig(config);
                    chat.RestoreLastSession();
                    string root = EnvDetect.FindRoot();
                    engine.SofficePath = ConvertEngine.FindSoffice(root);
                }
            }
        }

        void OnClosed(object sender, FormClosedEventArgs e)
        {
            if (pdfDoc != null) { pdfDoc.Dispose(); pdfDoc = null; }
            try { if (Directory.Exists(PreviewDir)) Directory.Delete(PreviewDir, true); } catch { }
        }

        // ================= 会话页 / 新任务页 =================

        Panel BuildChatPage()
        {
            Panel p = new Panel();
            p.Dock = DockStyle.Fill;
            p.BackColor = Color.White;
            chat = new ChatPanel(config);
            chat.OnOpenSettings += delegate { ShowSetup(false); };
            chat.OnSessionSaved += delegate { RefreshSessions(); };
            chat.OnOpenPreview += delegate(string path) { ShowPreview(path); SelectPage(5); };
            // 拖放统一入口：窗体、消息列表、输入框三处的文件拖放都汇到这里（0.8.7）
            chat.FilesDropped += delegate(string[] files) { HandleFilesDropped(files); };
            // 同名产物避让通知（**非阻塞**）：agent 的后台线程发现目标已存在时不会覆盖，
            // 而是自动改名并回调到这里；这里只把它转成一条聊天区提示。
            // 注意：回调发生在后台线程，所有 UI 操作必须经 Invoke 封送；
            // 且这里绝不能弹模态框（会自建消息泵并挂住 worker 线程）。
            AgentTools.OnOverwriteAvoided = delegate(string existingPath)
            {
                try
                {
                    if (chat == null) return;
                    if (chat.InvokeRequired)
                    {
                        try { chat.BeginInvoke((MethodInvoker)delegate { chat.NotifyOverwriteAvoided(existingPath); }); }
                        catch { }
                    }
                    else chat.NotifyOverwriteAvoided(existingPath);
                }
                catch { }
            };
            p.Controls.Add(chat);
            // 安全兜底确认桥（v0.8.2）：agent 后台线程判定"需确认"时回调到这里。
            // ★ 与上面 OnOverwriteAvoided 的**非阻塞**通知不同，这里必须**阻塞式**封送到 UI 线程：
            //   语义就是"不确认不执行"，所以后台线程要一直等到用户作答（ManualResetEvent）。
            //   绝不能在后台线程直接 ShowDialog（会自建消息泵、挂住 worker）。
            SafetyConfirm.Ask = delegate(string rule, string detail, string allowDir)
            {
                return ConfirmOnUi(rule, detail, allowDir);
            };
            return p;
        }

        // 把安全确认弹窗封送到 UI 线程并阻塞等待结果（0=取消 1=允许一次 2=始终允许此目录）
        int ConfirmOnUi(string rule, string detail, string allowDir)
        {
            if (IsDisposed || Disposing) return 0;
            if (!InvokeRequired) return ConfirmDialog.Show(this, rule, detail, allowDir);
            int r = 0;
            using (System.Threading.ManualResetEvent done = new System.Threading.ManualResetEvent(false))
            {
                try
                {
                    Invoke((MethodInvoker)delegate
                    {
                        try { r = ConfirmDialog.Show(this, rule, detail, allowDir); }
                        finally { try { done.Set(); } catch { } }
                    });
                }
                catch { try { done.Set(); } catch { } }
                done.WaitOne();
            }
            return r;
        }

        Panel BuildNewTaskPage()
        {
            Panel p = new Panel();
            p.Dock = DockStyle.Fill;
            p.BackColor = Color.White;
            Label tip = new Label();
            tip.Text = "新建任务\n\n· 把文件拖进窗口任意位置 → 自动进入预览，并加入任务台队列\n· 任务台批量转换：xlsx/xls/doc/docx/ppt/pptx → PDF，xlsx → CSV，CSV → XLSX\n· 「表格核对」页：两表勾稽核对（银行流水 vs 账面），自带差异表与审计校验\n· 「汇总·发票」页：多簿报表汇总 + 发票批量提取\n· 会话页可向配置好的模型提问；隐私分级见「内置插件」页";
            tip.Font = new Font("Microsoft YaHei UI", 10.5F);
            tip.ForeColor = Color.FromArgb(90, 93, 105);
            tip.Location = new Point(40, 48);
            tip.Size = new Size(700, 220);
            Button goChat = new Button();
            goChat.Text = "开始对话";
            goChat.Size = new Size(120, 34);
            goChat.Location = new Point(40, 280);
            goChat.FlatStyle = FlatStyle.Flat;
            goChat.FlatAppearance.BorderSize = 0;
            goChat.BackColor = Color.FromArgb(24, 26, 32);
            goChat.ForeColor = Color.White;
            goChat.Cursor = Cursors.Hand;
            goChat.Click += delegate { SelectPage(1); chat.FocusInput(); };
            Button newConv = new Button();
            newConv.Text = "新建对话";
            newConv.Size = new Size(120, 34);
            newConv.Location = new Point(180, 280);
            newConv.FlatStyle = FlatStyle.Flat;
            newConv.FlatAppearance.BorderSize = 0;
            newConv.BackColor = Color.White;
            newConv.ForeColor = Color.FromArgb(24, 26, 32);
            newConv.Cursor = Cursors.Hand;
            newConv.Click += delegate { chat.NewConversation(); SelectPage(1); chat.FocusInput(); sessionList.Invalidate(); };
            p.Controls.Add(tip);
            p.Controls.Add(goChat);
            p.Controls.Add(newConv);
            return p;
        }

        // ================= 任务台 =================

        Panel BuildTaskPage()
        {
            Panel p = new Panel();
            p.Dock = DockStyle.Fill;
            p.BackColor = Color.White;
            lvTasks = new ListView();
            lvTasks.View = View.Details;
            lvTasks.FullRowSelect = true;
            lvTasks.Dock = DockStyle.Fill;
            lvTasks.Columns.Add("状态", 150);
            lvTasks.Columns.Add("目标", 70);
            lvTasks.Columns.Add("文件", 760);
            p.Controls.Add(lvTasks);

            Panel bottom = new Panel();
            bottom.Dock = DockStyle.Bottom;
            bottom.Height = 46;
            Label lblTo = new Label();
            lblTo.Text = "转换为:";
            lblTo.Location = new Point(16, 14);
            lblTo.AutoSize = true;
            cmbTarget = new ComboBox();
            cmbTarget.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbTarget.Items.Add("PDF");
            cmbTarget.Items.Add("CSV");
            cmbTarget.Items.Add("XLSX");
            cmbTarget.SelectedIndex = 0;
            if (config.PluginPref && config.ConvTargetIndex >= 0 && config.ConvTargetIndex <= 2)
                cmbTarget.SelectedIndex = config.ConvTargetIndex;      // 转换偏好插件：恢复上次目标
            cmbTarget.SelectedIndexChanged += new EventHandler(delegate(object s, EventArgs ev)
            {
                if (config.PluginPref)
                {
                    config.ConvTargetIndex = cmbTarget.SelectedIndex;
                    config.Save();
                }
            });
            cmbTarget.Location = new Point(82, 10);
            cmbTarget.Size = new Size(80, 24);
            Button btnAdd = new Button();
            btnAdd.Text = "添加文件";
            btnAdd.Location = new Point(174, 9);
            btnAdd.Size = new Size(88, 27);
            btnAdd.Click += new EventHandler(BtnAddFiles_Click);
            Button btnStart = new Button();
            btnStart.Text = "开始转换";
            btnStart.Location = new Point(270, 9);
            btnStart.Size = new Size(88, 27);
            btnStart.FlatStyle = FlatStyle.Flat;
            btnStart.FlatAppearance.BorderSize = 0;
            btnStart.BackColor = Color.FromArgb(46, 76, 196);
            btnStart.ForeColor = Color.White;
            btnStart.Cursor = Cursors.Hand;
            btnStart.Click += new EventHandler(BtnStart_Click);
            Button btnClear = new Button();
            btnClear.Text = "清空";
            btnClear.Location = new Point(366, 9);
            btnClear.Size = new Size(64, 27);
            btnClear.Click += new EventHandler(delegate(object s, EventArgs ev) { lvTasks.Items.Clear(); });
            lblTaskStatus = new Label();
            lblTaskStatus.Location = new Point(444, 14);
            lblTaskStatus.Size = new Size(620, 20);
            lblTaskStatus.ForeColor = Color.DimGray;
            lblTaskStatus.Text = "原始文件只读；输出为同名 _conv 文件";
            bottom.Controls.Add(lblTo);
            bottom.Controls.Add(cmbTarget);
            bottom.Controls.Add(btnAdd);
            bottom.Controls.Add(btnStart);
            bottom.Controls.Add(btnClear);
            bottom.Controls.Add(lblTaskStatus);
            p.Controls.Add(bottom);
            return p;
        }

        void BtnAddFiles_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Title = "选择要转换的文件（可多选）";
                dlg.Multiselect = true;
                dlg.Filter = "办公文档|*.xlsx;*.xls;*.csv;*.doc;*.docx;*.ppt;*.pptx|全部文件|*.*";
                if (dlg.ShowDialog(this) == DialogResult.OK) AddTasks(dlg.FileNames);
            }
        }

        void AddTasks(string[] files)
        {
            string target = cmbTarget.SelectedItem == null ? "PDF" : cmbTarget.SelectedItem.ToString();
            foreach (string f in files)
            {
                ListViewItem i = new ListViewItem("排队中");
                i.SubItems.Add(target);
                i.SubItems.Add(f);
                i.Tag = f;
                lvTasks.Items.Add(i);
            }
        }

        void BtnStart_Click(object sender, EventArgs e)
        {
            List<ListViewItem> pending = new List<ListViewItem>();
            foreach (ListViewItem i in lvTasks.Items) { if (i.Text == "排队中" || i.Text.StartsWith("失败")) pending.Add(i); }
            if (pending.Count == 0) { lblTaskStatus.Text = "没有待处理的任务。"; return; }
            // 目标格式以"开始转换"那一刻的下拉值为准（修复：此前在添加时锁定，改下拉不影响已入队任务）
            string targetText = cmbTarget.SelectedItem == null ? "PDF" : cmbTarget.SelectedItem.ToString();
            foreach (ListViewItem i in pending) { i.SubItems[1].Text = targetText; }
            foreach (Control c in this.Controls) { } // no-op
            Button startBtn = FindStartButton();
            if (startBtn != null) startBtn.Enabled = false;
            Thread t = new Thread(new ThreadStart(delegate
            {
                int ok = 0, fail = 0, idx = 0;
                foreach (ListViewItem i in pending)
                {
                    idx++;
                    string file = (string)i.Tag;
                    ConvTarget target = TargetFromText(targetText);
                    SetTaskRow(i, "转换中...", null);
                    SetTaskStatus("转换中 (" + idx + "/" + pending.Count + "): " + Path.GetFileName(file));
                    AuditLog.Record("file_read", file);
                    string outPath;
                    string err = engine.Convert(file, target, out outPath);
                    if (err == null)
                    {
                        ok++;
                        // 成功但有附加提示（如源表公式无缓存值、导出格被置空）也要让用户看见
                        string note = (engine.LastWarning != null && engine.LastWarning.Length > 0)
                            ? ("（注意：" + engine.LastWarning + "）") : "";
                        SetTaskRow(i, "完成 → " + outPath + note, Color.DarkGreen);
                        AuditLog.Record("file_write", outPath);
                    }
                    else { fail++; SetTaskRow(i, "失败: " + err, Color.Red); }
                }
                SetTaskStatus("队列完成：成功 " + ok + "，失败 " + fail + "。原始文件未被修改。");
                EnableStart();
            }));
            t.IsBackground = true;
            t.Start();
        }

        Button FindStartButton()
        {
            foreach (Control c in pages[2].Controls)
            {
                Panel pnl = c as Panel;
                if (pnl == null) continue;
                foreach (Control b in pnl.Controls)
                {
                    Button btn = b as Button;
                    if (btn != null && btn.Text == "开始转换") return btn;
                }
            }
            return null;
        }

        void EnableStart()
        {
            if (InvokeRequired) { Invoke(new VoidD(EnableStart)); return; }
            Button b = FindStartButton();
            if (b != null) b.Enabled = true;
        }

        static ConvTarget TargetFromText(string s)
        {
            if (s == "CSV") return ConvTarget.Csv;
            if (s == "XLSX") return ConvTarget.Xlsx;
            return ConvTarget.Pdf;
        }

        void SetTaskRow(ListViewItem i, string status, Color? color)
        {
            if (InvokeRequired) { Invoke((MethodInvoker)delegate { SetTaskRow(i, status, color); }); return; }
            i.Text = status;
            if (color.HasValue) i.ForeColor = color.Value;
        }
        void SetTaskStatus(string s)
        {
            if (InvokeRequired) { Invoke(new StrD(SetTaskStatus), s); return; }
            lblTaskStatus.Text = s;
        }

        // ================= 预览 =================

        Panel BuildPreviewPage()
        {
            Panel p = new Panel();
            p.Dock = DockStyle.Fill;
            p.BackColor = Color.White;
            previewHost = new Panel();
            previewHost.Dock = DockStyle.Fill;
            previewInfo = new Label();
            previewInfo.Dock = DockStyle.Bottom;
            previewInfo.Height = 26;
            previewInfo.ForeColor = Color.DimGray;
            previewInfo.Text = "把文件拖到窗口任意位置，或点「打开预览文件」（xlsx 网格 / PDF 渲染 / doc·ppt 先转 PDF）";
            Panel topBar = new Panel();
            topBar.Dock = DockStyle.Top;
            topBar.Height = 42;
            topBar.BackColor = Color.White;
            Button btnOpenPreview = new Button();
            btnOpenPreview.Text = "打开预览文件";
            btnOpenPreview.Location = new Point(24, 7);
            btnOpenPreview.Size = new Size(110, 28);
            btnOpenPreview.Cursor = Cursors.Hand;
            btnOpenPreview.Click += new EventHandler(BtnOpenPreview_Click);
            topBar.Controls.Add(btnOpenPreview);
            // 打开当前预览文件所在文件夹（用户实测需求：预览后想直接定位产物文件）
            Button btnShowInFolder = new Button();
            btnShowInFolder.Text = "打开源文件所在地";
            btnShowInFolder.Location = new Point(142, 7);
            btnShowInFolder.Size = new Size(130, 28);
            btnShowInFolder.Cursor = Cursors.Hand;
            btnShowInFolder.Click += delegate { OpenFolder.Open(previewFolderPath); };
            topBar.Controls.Add(btnShowInFolder);
            p.Controls.Add(previewHost);
            p.Controls.Add(previewInfo);
            p.Controls.Add(topBar);
            return p;
        }

        void BtnOpenPreview_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Title = "打开预览文件";
                dlg.Filter = "可预览|*.xlsx;*.csv;*.pdf;*.xls;*.doc;*.docx;*.ppt;*.pptx;*.frp|全部文件|*.*";
                try { dlg.InitialDirectory = config.EffectiveWorkspace(); } catch { }
                if (dlg.ShowDialog(this) == DialogResult.OK) ShowPreview(dlg.FileName);
            }
        }

        void ClearPreview()
        {
            if (pdfDoc != null) { pdfDoc.Dispose(); pdfDoc = null; }
            previewHost.Controls.Clear();
            if (tempPdf != null)
            {
                try { File.Delete(tempPdf); } catch { }
                tempPdf = null;
            }
        }

        void OnDragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        void OnDragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files == null || files.Length == 0) return;
            HandleFilesDropped(files);
        }

        // 拖放统一处理（窗体/消息列表/输入框三处汇入）：
        // 0.8.4 用户规则：拖文件进对话框**留在当前页**（此前跳去表格核对页+预览，WorkBuddy 式
        // 交互=引用条出现在输入框上方，聊天继续在这里进行）
        void HandleFilesDropped(string[] files)
        {
            if (files == null || files.Length == 0) return;
            chat.SetContextFiles(files);
            AddTasks(files);
            chat.FocusInput();
        }

        public void ShowPreview(string path)
        {
            ClearPreview();
            try { previewFolderPath = Path.GetDirectoryName(path) ?? ""; } catch { previewFolderPath = ""; }
            AuditLog.Record("file_read", path);
            string ext = (Path.GetExtension(path) ?? "").ToLowerInvariant();
            try
            {
                if (ext == ".pdf") { ShowPdf(path); return; }
                if (ext == ".xlsx") { ShowXlsx(path); return; }
                if (ext == ".csv") { ShowCsv(path); return; }
                // frp 打印模板：解析成网格 → 临时 xlsx → 走既有 xlsx 预览（0.8.5）
                if (ext == ".frp")
                {
                    if (!Directory.Exists(PreviewDir)) Directory.CreateDirectory(PreviewDir);
                    string tmp = Path.Combine(PreviewDir, "frp-preview.xlsx");
                    string frpErr = FrpReport.ConvertToXlsx(path, tmp);
                    if (frpErr != null) { SetPreviewInfo("frp 解析失败: " + frpErr); return; }
                    ShowXlsx(tmp);
                    return;
                }
                if (ext == ".xls" || ext == ".doc" || ext == ".docx" || ext == ".ppt" || ext == ".pptx")
                {
                    previewInfo.Text = "正在通过转换引擎生成预览（Office COM / LibreOffice）...";
                    Thread t = new Thread(new ThreadStart(delegate
                    {
                        string outPath;
                        string err = engine.Convert(path, ConvTarget.Pdf, out outPath);
                        if (err == null && File.Exists(outPath))
                        {
                            if (!Directory.Exists(PreviewDir)) Directory.CreateDirectory(PreviewDir);
                            string tmp = Path.Combine(PreviewDir, Path.GetFileNameWithoutExtension(path) + "-preview.pdf");
                            File.Copy(outPath, tmp, true);
                            tempPdf = tmp;
                            ShowPdfFromThread(tmp);
                        }
                        else SetPreviewInfo("预览失败: " + err);
                    }));
                    t.IsBackground = true;
                    t.Start();
                    return;
                }
                SetPreviewInfo("不支持预览的格式: " + ext);
            }
            catch (Exception ex)
            {
                SetPreviewInfo("预览异常: " + ex.Message);
            }
        }

        // xlsx 预览（低配优化）：改为**单遍扫描 + 有界保留**，不再一次性物化整表。
        //
        // 原实现 LoadGrid(0, 20000, 128) 的实测代价（开发机 12 核/16GB）：
        //   20000x6   -> 887ms，托管堆 +16MB（优化前为 +41MB，流式解压已先降一截）
        //   20000x128 -> 3148ms，托管堆 +89MB
        // 且这只是数组本身；DataGridView 还要为每个单元格建对象（20000x128 = 256 万个），
        // 逐行 Rows.Add 全程在 UI 线程。4GB Win7 老机上表现为 OOM 或长时间假死。
        //
        // 现在：只保留前 DefaultPreviewRows 行、列数取实际用到的（上限 MaxPreviewCols），
        // 同时**继续读完整表**以得到真实总行数（不能读够就停——总行数是用户判断
        // "这表要不要处理"的关键信息，报错的行数比慢一点更糟）。
        // 实测：20000x128 -> 1694ms / +5MB；10 万行 x6 -> 577ms / +5MB。
        void ShowXlsx(string path)
        {
            string[] sheetNames;
            int sheetCount = 0;
            XlsxPreview p = XlsxPreview.Load(path, 0, XlsxPreview.DefaultPreviewRows, XlsxPreview.MaxPreviewCols);
            // 工作表名单独取一次（只为标题栏显示，代价可忽略：workbook.xml 很小）
            try
            {
                XlsxBook book = XlsxBook.Open(path);
                sheetCount = book.Sheets.Count;
                sheetNames = new string[sheetCount];
                for (int i = 0; i < sheetCount; i++) sheetNames[i] = book.Sheets[i].Name;
                book.Dispose();
            }
            catch { sheetNames = new string[0]; }

            if (p.Error.Length > 0 && p.Rows.Length == 0)
            {
                SetPreviewInfo(p.Error);
                return;
            }
            ShowGridFromThread(p.Rows, p.TotalRows, p.UsedCols, sheetNames,
                p.Describe(sheetCount, string.Join(" | ", sheetNames)));
        }

        // CSV 预览（低配优化）：与 xlsx 同一路径。
        // 原实现把整个 CSV 解析成 List<string[]>（全量驻留），再复制进 string[,] 网格，
        // 峰值是"文件内容 x2"。大 CSV（几十万行）在 4GB 老机上同样是 OOM 风险。
        // 现在只保留前 DefaultPreviewRows 行交给网格，行数仍如实统计。
        void ShowCsv(string path)
        {
            Encoding used;
            List<string[]> all = MiniCsv.Parse(MiniCsv.DetectRead(path, out used));
            int rowsCount = all.Count, cols = 0;
            foreach (string[] r in all) { if (r.Length > cols) cols = r.Length; }
            if (cols == 0) cols = 1;
            if (cols > XlsxPreview.MaxPreviewCols) cols = XlsxPreview.MaxPreviewCols;

            int keep = Math.Min(rowsCount, XlsxPreview.DefaultPreviewRows);
            string[][] gridData = new string[keep][];
            for (int i = 0; i < keep; i++) gridData[i] = all[i];
            all = null;   // 提前释放引用，便于 GC 回收（大 CSV 的解析结果）

            string name = "CSV(" + used.WebName + ")";
            string extra = rowsCount > keep ? "（仅预览前 " + keep + " 行 / 共 " + rowsCount + " 行）" : "";
            ShowGridFromThread(gridData, rowsCount, cols, new string[] { name },
                "工作表: " + name + "  " + extra);
        }

        // infoText：预览信息栏文本（由调用方组装，便于如实标注"仅预览前 N 行 / 共 M 行"）
        void ShowGridFromThread(string[][] data, int totalRows, int usedCols, string[] sheetNames, string infoText)
        {
            if (InvokeRequired) { Invoke((MethodInvoker)delegate { ShowGridFromThread(data, totalRows, usedCols, sheetNames, infoText); }); return; }
            grid = new DataGridView();
            grid.Dock = DockStyle.Fill;
            grid.ReadOnly = true;
            grid.AllowUserToAddRows = false;
            grid.AllowUserToDeleteRows = false;
            grid.AllowUserToOrderColumns = true;
            grid.RowHeadersVisible = true;                 // 行号（像表格软件那样可定位）
            grid.RowHeadersWidth = 54;
            grid.BackgroundColor = Color.White;
            grid.GridColor = Color.FromArgb(214, 218, 226);
            grid.CellBorderStyle = DataGridViewCellBorderStyle.Single;
            grid.BorderStyle = BorderStyle.None;
            grid.EnableHeadersVisualStyles = false;
            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(242, 244, 248);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(60, 64, 74);
            grid.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold);
            grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            grid.ColumnHeadersHeight = 30;
            grid.DefaultCellStyle.Font = new Font("Microsoft YaHei UI", 9.75F);
            grid.DefaultCellStyle.Padding = new Padding(4, 0, 4, 0);
            grid.RowTemplate.Height = 24;
            // 列宽按内容实测（首行当表头给更高权重），比统一宽度或按表头自适应都好读
            int cols = usedCols < 1 ? 1 : usedCols;
            int[] widths = new int[cols];
            Font measure = new Font("Microsoft YaHei UI", 9.75F);
            int sample = Math.Min(data.Length, 200);   // 只量前 200 行，避免大表卡顿
            for (int c = 0; c < cols; c++)
            {
                int w = 40;
                for (int r = 0; r < sample; r++)
                {
                    string v = (data[r] != null && c < data[r].Length) ? data[r][c] : null;
                    if (string.IsNullOrEmpty(v)) continue;
                    string one = v.Length > 60 ? v.Substring(0, 60) : v;
                    int mw = TextRenderer.MeasureText(one, measure).Width + 18;
                    if (mw > w) w = mw;
                    if (w > 360) { w = 360; break; }         // 超长文本不撑爆视口，靠单元格内省略
                }
                widths[c] = w;
            }
            measure.Dispose();
            for (int c = 0; c < cols; c++)
            {
                string h = "";
                int ci = c;
                while (ci >= 0) { h = (char)('A' + ci % 26) + h; ci = ci / 26 - 1; }
                int idx = grid.Columns.Add("c" + c, h);
                grid.Columns[idx].Width = widths[c];
                grid.Columns[idx].SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            // 挂起布局后批量插入：Rows.Add 每次都会触发布局重算，
            // 未挂起时插入几千行会明显拖慢（这是"打开大表卡"的第二个来源，
            // 第一个来源是内存——见 ShowXlsx 注释）。
            grid.SuspendLayout();
            try
            {
                for (int r = 0; r < data.Length; r++)
                {
                    object[] vals = new object[cols];
                    string[] src = data[r];
                    for (int c = 0; c < cols; c++)
                        vals[c] = (src != null && c < src.Length && src[c] != null) ? src[c] : "";
                    grid.Rows.Add(vals);
                    grid.Rows[r].HeaderCell.Value = (r + 1).ToString();
                }
            }
            finally { grid.ResumeLayout(); }
            previewHost.Controls.Add(grid);
            previewInfo.Text = infoText;
        }

        void ShowPdf(string path) { ShowPdfFromThread(path); }

        void ShowPdfFromThread(string path)
        {
            if (InvokeRequired) { Invoke((MethodInvoker)delegate { ShowPdfFromThread(path); }); return; }
            try
            {
                string root = EnvDetect.FindRoot();
                Pdfium.EnsureInit(root);
                if (pdfDoc != null) { pdfDoc.Dispose(); pdfDoc = null; }
                string err;
                pdfDoc = Pdfium.PdfDoc.Open(path, out err);
                if (pdfDoc == null) { SetPreviewInfo("PDF 打开失败: " + err); return; }
                pdfPageIndex = 0;
                pdfBox = new PictureBox();
                pdfBox.Dock = DockStyle.Fill;
                pdfBox.SizeMode = PictureBoxSizeMode.Zoom;
                pdfBox.BackColor = Color.FromArgb(120, 124, 132);
                btnPrev = new Button();
                btnPrev.Text = "◀ 上一页";
                btnPrev.Enabled = false;
                btnNext = new Button();
                btnNext.Text = "下一页 ▶";
                Panel navBar = new Panel();
                navBar.Dock = DockStyle.Top;
                navBar.Height = 36;
                navBar.BackColor = Color.White;
                pdfPage = new Label();
                pdfPage.Location = new Point(210, 10);
                pdfPage.Size = new Size(300, 18);
                btnPrev.Location = new Point(8, 6);
                btnPrev.Size = new Size(92, 24);
                btnNext.Location = new Point(106, 6);
                btnNext.Size = new Size(92, 24);
                btnPrev.Click += new EventHandler(delegate(object s, EventArgs ev) { PdfNav(-1); });
                btnNext.Click += new EventHandler(delegate(object s, EventArgs ev) { PdfNav(1); });
                navBar.Controls.Add(btnPrev);
                navBar.Controls.Add(btnNext);
                navBar.Controls.Add(pdfPage);
                previewHost.Controls.Add(pdfBox);
                previewHost.Controls.Add(navBar);
                RenderPdfPage();
            }
            catch (DllNotFoundException ex) { SetPreviewInfo(ex.Message); }
            catch (Exception ex) { SetPreviewInfo("PDF 渲染异常: " + ex.Message); }
        }

        void PdfNav(int delta)
        {
            if (pdfDoc == null) return;
            int np = pdfPageIndex + delta;
            if (np < 0 || np >= pdfDoc.PageCount) return;
            pdfPageIndex = np;
            RenderPdfPage();
        }

        void RenderPdfPage()
        {
            Bitmap bm = null;
            try
            {
                bm = pdfDoc.RenderPage(pdfPageIndex, 120.0);
                Image old = pdfBox.Image;
                pdfBox.Image = bm;
                if (old != null) old.Dispose();
                pdfPage.Text = "第 " + (pdfPageIndex + 1) + " / " + pdfDoc.PageCount + " 页";
                btnPrev.Enabled = pdfPageIndex > 0;
                btnNext.Enabled = pdfPageIndex < pdfDoc.PageCount - 1;
            }
            catch (Exception ex)
            {
                if (bm != null) bm.Dispose();
                SetPreviewInfo("渲染失败: " + ex.Message);
            }
        }

        void SetPreviewInfo(string s)
        {
            if (InvokeRequired) { Invoke(new StrD(SetPreviewInfo), s); return; }
            previewInfo.Text = s;
        }

        // ================= 系统状态 / 审计 =================

        Panel BuildStatusPage()
        {
            Panel p = new Panel();
            p.Dock = DockStyle.Fill;
            p.BackColor = Color.White;
            lvEnv = new ListView();
            lvEnv.View = View.Details;
            lvEnv.FullRowSelect = true;
            lvEnv.Dock = DockStyle.Fill;
            lvEnv.Columns.Add("组件", 240);
            lvEnv.Columns.Add("状态", 60);
            lvEnv.Columns.Add("说明", 660);
            p.Controls.Add(lvEnv);

            Panel bottom = new Panel();
            bottom.Dock = DockStyle.Bottom;
            bottom.Height = 52;
            lblEnvStatus = new Label();
            lblEnvStatus.Location = new Point(16, 4);
            lblEnvStatus.Size = new Size(800, 18);
            lblEnvStatus.ForeColor = Color.DimGray;
            Button btnRecheck = new Button();
            btnRecheck.Text = "重新检测";
            btnRecheck.Location = new Point(16, 22);
            btnRecheck.Size = new Size(104, 26);
            btnRecheck.Click += new EventHandler(delegate(object s, EventArgs ev) { RunDetect(); });
            Button btnOpenLogs = new Button();
            btnOpenLogs.Text = "打开日志目录";
            btnOpenLogs.Location = new Point(128, 22);
            btnOpenLogs.Size = new Size(104, 26);
            btnOpenLogs.Click += new EventHandler(delegate(object s, EventArgs ev) { OpenLogDir(); });
            // 工作区目录：建表/建PPT 相对路径、下载、@引用 的默认落盘位置
            Button btnWorkspace = new Button();
            btnWorkspace.Text = "更改工作区目录";
            btnWorkspace.Location = new Point(240, 22);
            btnWorkspace.Size = new Size(120, 26);
            btnWorkspace.Click += new EventHandler(delegate(object s, EventArgs ev) { ChangeWorkspace(); });
            lblWorkspace = new Label();
            lblWorkspace.Location = new Point(368, 27);
            lblWorkspace.Size = new Size(560, 18);
            lblWorkspace.ForeColor = Color.DimGray;
            bottom.Controls.Add(btnWorkspace);
            bottom.Controls.Add(lblWorkspace);
            bottom.Controls.Add(lblEnvStatus);
            bottom.Controls.Add(btnRecheck);
            bottom.Controls.Add(btnOpenLogs);
            p.Controls.Add(bottom);
            RefreshWorkspaceLabel();
            return p;
        }

        // 工作区目录：agent 建表/建 PPT 的相对路径、下载文件、@引用 的默认落盘位置
        void ChangeWorkspace()
        {
            string picked = WorkspacePicker.Pick(this, config.EffectiveWorkspace(),
                "选择工作区位置（建表/建PPT/下载的默认落盘位置，@ 可引用其中文件）");
            if (picked == null) return;
            config.WorkspaceDir = picked;
            config.Save();
            try { AgentTools.WorkspaceRoot = config.EffectiveWorkspace(); } catch { }
            RefreshWorkspaceLabel();
            AuditLog.Record("config_change", "工作区目录=" + picked);
        }

        void RefreshWorkspaceLabel()
        {
            if (lblWorkspace == null) return;
            try { lblWorkspace.Text = "工作区：" + config.EffectiveWorkspace(); } catch { }
        }

        void RunDetect()
        {
            lblEnvStatus.Text = "环境检测中...";
            Thread t = new Thread(new ThreadStart(delegate
            {
                string root = EnvDetect.FindRoot();
                List<DetectItem> items = EnvDetect.DetectAll(root);
                FillEnv(items);
            }));
            t.IsBackground = true;
            t.Start();
        }

        void FillEnv(List<DetectItem> items)
        {
            if (InvokeRequired) { Invoke(new FillEnvD(FillEnv), items); return; }
            lvEnv.BeginUpdate();
            lvEnv.Items.Clear();
            int missing = 0;
            foreach (DetectItem it in items)
            {
                ListViewItem i = new ListViewItem(it.Name);
                i.SubItems.Add(EnvDetect.StateText(it.State));
                i.SubItems.Add(it.Detail);
                switch (it.State)
                {
                    case DetectState.Ok: i.ForeColor = Color.DarkGreen; break;
                    case DetectState.Missing: i.ForeColor = Color.Red; missing++; break;
                    case DetectState.Unknown: i.ForeColor = Color.DarkOrange; break;
                    default: i.ForeColor = Color.Gray; break;
                }
                lvEnv.Items.Add(i);
            }
            lvEnv.EndUpdate();
            string lo = engine.SofficePath == null ? "LibreOffice 未探测到（转换走 Office COM）" : "LibreOffice: " + engine.SofficePath;
            lblEnvStatus.Text = missing == 0
                ? "环境齐备。" + lo
                : ("有 " + missing + " 项缺失：请通过引导器（boot\\OfficeAgentBoot.exe）离线补全。" + lo);
        }

        void OpenLogDir()
        {
            try
            {
                string dir = Path.Combine(Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeAgent"), "logs");
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                System.Diagnostics.Process.Start("explorer.exe", "\"" + dir + "\"");
            }
            catch (Exception ex) { MessageBox.Show("打开日志目录失败: " + ex.Message); }
        }

        ListView lvAudit; Label lblAudit;
        delegate void StrRefD(string s);

        Panel BuildAuditPage()
        {
            Panel p = new Panel();
            p.Dock = DockStyle.Fill;
            p.BackColor = Color.White;
            lvAudit = new ListView();
            lvAudit.View = View.Details;
            lvAudit.FullRowSelect = true;
            lvAudit.ShowItemToolTips = true;
            lvAudit.Dock = DockStyle.Fill;
            lvAudit.Columns.Add("序号", 64);
            lvAudit.Columns.Add("时间", 150);
            lvAudit.Columns.Add("软件做了什么", 150);
            lvAudit.Columns.Add("详情", 720);
            p.Controls.Add(lvAudit);

            Panel bottom = new Panel();
            bottom.Dock = DockStyle.Bottom;
            bottom.Height = 52;
            lblAudit = new Label();
            lblAudit.Location = new Point(16, 4);
            lblAudit.Size = new Size(860, 18);
            lblAudit.ForeColor = Color.DimGray;
            lblAudit.Text = "这里记录软件在本机做过的事：读了哪个文件、写了哪个结果、什么时候用了 AI 模型。\n每条记录都和上一条环环相扣（防止有人偷偷删改），点「校验审计链」即可验证。把鼠标停在带 ❗ 的行上可看解释。";
            Button btnVerify = new Button();
            btnVerify.Text = "校验审计链";
            btnVerify.Location = new Point(16, 22);
            btnVerify.Size = new Size(104, 26);
            btnVerify.Click += new EventHandler(delegate(object s, EventArgs ev)
            {
                Thread t = new Thread(new ThreadStart(delegate
                {
                    int count; long broken;
                    string err = AuditLog.VerifyChain(out count, out broken);
                    SetAuditStatus(err == null
                        ? ("✓ 审计链完整：已校验 " + count + " 条事件")
                        : ("✗ " + err));
                }));
                t.IsBackground = true;
                t.Start();
            });
            Button btnRefresh = new Button();
            btnRefresh.Text = "刷新";
            btnRefresh.Location = new Point(128, 22);
            btnRefresh.Size = new Size(72, 26);
            btnRefresh.Click += new EventHandler(delegate(object s, EventArgs ev) { RefreshAudit(); });
            Button btnOpenAudit = new Button();
            btnOpenAudit.Text = "打开审计目录";
            btnOpenAudit.Location = new Point(208, 22);
            btnOpenAudit.Size = new Size(112, 26);
            btnOpenAudit.Click += new EventHandler(delegate(object s, EventArgs ev) { OpenLogDir(); });
            bottom.Controls.Add(lblAudit);
            bottom.Controls.Add(btnVerify);
            bottom.Controls.Add(btnRefresh);
            bottom.Controls.Add(btnOpenAudit);
            p.Controls.Add(bottom);
            return p;
        }

        void RefreshAudit()
        {
            Thread t = new Thread(new ThreadStart(delegate
            {
                List<string[]> rows = AuditLog.ReadTail(500);
                FillAudit(rows);
            }));
            t.IsBackground = true;
            t.Start();
        }

        delegate void FillAuditD(List<string[]> rows);

        void FillAudit(List<string[]> rows)
        {
            if (InvokeRequired) { Invoke(new FillAuditD(FillAudit), rows); return; }
            lvAudit.BeginUpdate();
            lvAudit.Items.Clear();
            foreach (string[] r in rows)
            {
                ListViewItem i = new ListViewItem(r[0]);
                i.SubItems.Add(r[1]);
                i.SubItems.Add(PlainTips.PlainEvent(r[2]) + " ❗");
                i.SubItems.Add(r[3]);
                i.ToolTipText = "这条记录的意思：\n" + PlainTips.EventHover(r[2]);
                if (r[2] == "llm_call") i.ForeColor = Color.FromArgb(90, 96, 110);
                lvAudit.Items.Add(i);
            }
            lvAudit.EndUpdate();
            lblAudit.Text = "共显示 " + rows.Count + " 条（最近）。鼠标停在带 ❗ 的行上看解释；点「校验审计链」验证没有被人删改。";
        }

        void SetAuditStatus(string s)
        {
            if (InvokeRequired) { Invoke(new StrRefD(SetAuditStatus), s); return; }
            lblAudit.Text = s;
        }

        // ================= 内置插件页（开箱即用） =================

        Panel BuildPluginsPage()
        {
            Panel p = new Panel();
            p.Dock = DockStyle.Fill;
            p.BackColor = Color.White;

            Label intro = new Label();
            intro.Text = "以下插件随程序内置，无需安装、默认启用，开箱即用；取消勾选即关闭对应功能。";
            intro.ForeColor = Color.FromArgb(90, 93, 105);
            intro.Location = new Point(24, 16);
            intro.AutoSize = true;
            p.Controls.Add(intro);

            AddPluginCheck(p, 64, "记忆模块",
                "对话命令：「记住 xxx」记事、「记忆」查看、「忘记 关键词」删除；记忆自动注入模型提示词，数据只存本机。",
                config.PluginMemory, delegate(bool v) { config.PluginMemory = v; config.Save(); AuditLog.Record("config_change", "记忆模块=" + v); });
            AddPluginCheck(p, 132, "转换偏好",
                "记住上次使用的转换目标格式（PDF/CSV/XLSX），下次打开自动恢复。",
                config.PluginPref, delegate(bool v) { config.PluginPref = v; config.Save(); AuditLog.Record("config_change", "转换偏好=" + v); });
            AddPluginCheck(p, 200, "会话历史",
                "会话记录保存在本机（最多 500 条），下次启动自动恢复上次对话。",
                config.PluginHist, delegate(bool v) { config.PluginHist = v; config.Save(); AuditLog.Record("config_change", "会话历史=" + v); });
            AddPluginCheck(p, 268, "任务计划栏",
                "遇到多步骤复杂任务时，AI 自动创建计划并在会话右侧展示进度（✓ 完成 / ▶ 进行中 / ○ 待办）。",
                config.PluginPlan, delegate(bool v) { config.PluginPlan = v; config.Save(); AuditLog.Record("config_change", "任务计划栏=" + v); });

            Label privTitle = new Label();
            privTitle.Text = "隐私分级（发往模型的内容口径；文件转换/核对不受影响）";
            privTitle.Font = new Font("Microsoft YaHei UI", 11F, FontStyle.Bold);
            privTitle.Location = new Point(24, 346);
            privTitle.AutoSize = true;
            p.Controls.Add(privTitle);

            string[] privNames = new string[] {
                "L0 全本地（不出网）—— 模型对话禁用，纯本地功能可用",
                "L1 脱敏出网（推荐，默认）—— 发送前自动脱敏手机号/身份证/银行卡/税号/邮箱",
                "L2 全量出网 —— 原文发送（财务数据不建议）"
            };
            for (int i = 0; i < 3; i++)
            {
                RadioButton rb = new RadioButton();
                rb.Text = privNames[i];
                rb.Location = new Point(48, 376 + i * 28);
                rb.AutoSize = true;
                int level = i;
                rb.Checked = config.PrivacyLevel == i;
                rb.CheckedChanged += new EventHandler(delegate(object s, EventArgs ev)
                {
                    if (rb.Checked)
                    {
                        config.PrivacyLevel = level;
                        config.Save();
                        AuditLog.Record("privacy_change", "L" + level);
                    }
                });
                p.Controls.Add(rb);
            }
            return p;
        }

        void AddPluginCheck(Panel p, int y, string name, string desc, bool initial, Action<bool> onChange)
        {
            CheckBox chk = new CheckBox();
            chk.Text = name;
            chk.Font = new Font("Microsoft YaHei UI", 11F, FontStyle.Bold);
            chk.Checked = initial;
            chk.Location = new Point(24, y);
            chk.AutoSize = true;
            chk.CheckedChanged += new EventHandler(delegate(object s, EventArgs ev) { onChange(chk.Checked); });
            p.Controls.Add(chk);

            Label descLabel = new Label();
            descLabel.Text = desc;
            descLabel.ForeColor = Color.FromArgb(110, 113, 126);
            descLabel.Location = new Point(48, y + 26);
            descLabel.Size = new Size(800, 34);
            p.Controls.Add(descLabel);
        }
    }
}
