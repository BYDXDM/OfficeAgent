// ChatPanel —— WorkBuddy 风格对话页：气泡消息列表 + 圆角输入区 + 圆形发送按钮 + 模型入口
// Win7 WinForms：GDI+ 自绘圆角（Region/Path），双缓冲防闪烁。
// Agent 能力（0.4.1）：多轮历史上下文 + 消息内文件读取 + function calling 工具循环
//   （读文件/列目录/下载/转换/环境修复）+ 长回复分块显示（修复 ListBox 255px 单项高度裁剪）。
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

using OfficeAgent.Core;

namespace OfficeAgent.Host
{
    public class ChatMsg
    {
        public bool User;
        public string Text = "";
        public string AttachPath = "";   // 附件卡片（工具产物，点击进预览）
        public string Time = "";
        public bool Pending;
        public bool Error;
        public bool First = true;      // 分块组第一条（带头部、上圆角）
        public bool Last = true;       // 分块组最后一条（下圆角、底距）
        public int CachedTextH = -1;   // 气泡文本高度缓存（低配机：避免每次重绘全文测量）
        public int CachedHeight = -1;  // 整项高度缓存
        public int CachedTextW = -1;   // 气泡宽度缓存
        public int CachedHeaderW = -1; // 名字+时间宽度缓存
        public int CachedAvail = -1;   // 测量时的可用宽度（变宽后需重测）
    }

    public partial class ChatPanel : Panel
    {
        public event Action OnOpenSettings;
        public event Action OnSessionSaved;   // 会话落盘后通知（侧边栏刷新列表）
        public event Action<string> OnOpenPreview;   // 点击附件卡片 → 主窗体打开预览

        BufferedMsgList list;
        ForwardWheelBox input;
        Label placeholder;
        Button send;
        LinkLabel modelLink;
        Label ctxLabel;
        Panel inputBorder;
        Panel planPanel;                 // 任务计划栏（右侧小窗，DSH/ZCode 式）
        Label planTitle;
        ListBox planStepsList;
        List<string[]> planSteps = new List<string[]>();   // [文本, 是否完成("1"/"0")]
        List<ChatMsg> msgs = new List<ChatMsg>();
        List<LlmTurn> llmHistory = new List<LlmTurn>();   // 发给模型的多轮历史（role/content）
        List<string[]> sessionRows = new List<string[]>(); // 会话存档（显示文本；与 llmHistory 的问答一一对应）
        string lastUserDisplay = "";                       // 当前这条用户消息的原文（会话存档用，不含文件上下文/脱敏）
        string sessionId = null;                          // 当前会话（null = 新会话未落盘，首次发言时创建）
        AppConfig config;
        LlmClient client;
        bool busy = false;
        List<string> contextFiles = new List<string>();
        Label fileChip;                 // 输入框左侧的「引用文件」提示条
        Form atPopup;                   // @ 引用文件弹层（无边框小窗，不抢焦点）
        ListBox atList;
        bool atActive = false;
        bool atSuspend = false;         // 程序改写输入文本时挂起触发
        int atTokenStart = -1;          // '@' 在输入框中的起始位置
        int atCaret = -1;               // 弹出时输入框光标位置
        ActionPlan pendingPlan = null;
        System.Windows.Forms.Timer flushTimer;     // 流式刷新节流：脏标记 + 100ms 定时（上限 10fps，替代每增量整表重绘）
        System.Windows.Forms.Timer resizeDebounce; // 拖动窗口防抖：停止拖动 200ms 后重排一次
        bool dirty = false;
        int pendingIdx = -1;            // 等待中的气泡（显示已等待秒数）
        DateTime pendingStart;
        int pendingShownSecs = -1;

        const int MaxTextWidth = 560;
        const int MaxChunkTextH = 440;   // 单条气泡文本区最大高度（ListBox 单项高度被系统钳制在 255px，长文必须分块）
        const int MaxHistoryTurns = 24;  // 发给模型的最大历史消息条数（user+assistant 合计）
        const int MaxHistoryChars = 24000; // 历史字符预算：文件上下文可能很大，超预算时丢最旧的

        [DllImport("user32.dll")]
        static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wp, IntPtr lp);

        const int WM_MOUSEWHEEL = 0x020A;

        List<string> fetchedModels = null;   // 从服务拉到的模型列表（内联下拉用）

        int LeftMargin() { return (list.Width > 900) ? 170 : 20; }
        int AvailWidth() { return Math.Min(MaxTextWidth, Math.Max(200, list.Width - LeftMargin() - 60)); }

        static Font NameFont = new Font("Microsoft YaHei UI", 8.25F);
        static Font TextFont = new Font("Microsoft YaHei UI", 9.75F);

        public ChatPanel(AppConfig cfg)
        {
            config = cfg;
            BackColor = Color.White;
            Dock = DockStyle.Fill;
            DoubleBuffered = true;

            list = new BufferedMsgList();
            list.Dock = DockStyle.Fill;
            list.BorderStyle = BorderStyle.None;
            list.BackColor = Color.White;
            list.DrawMode = DrawMode.OwnerDrawVariable;
            list.Font = new Font("Microsoft YaHei UI", 9F);
            list.DrawItem += List_DrawItem;
            list.MeasureItem += List_MeasureItem;
            list.MouseClick += delegate(object s, MouseEventArgs e)
            {
                int idx = list.IndexFromPoint(e.Location);
                if (idx < 0 || idx >= msgs.Count) return;
                if (msgs[idx].AttachPath.Length > 0 && OnOpenPreview != null) OnOpenPreview(msgs[idx].AttachPath);
            };
            Controls.Add(list);

            // 任务计划栏（右侧，任务计划插件）：模型经 task_plan 工具驱动，✓=已完成 ▶=进行中 ○=待办
            planPanel = new Panel();
            planPanel.Dock = DockStyle.Right;
            planPanel.Width = 0;               // 无计划时零宽隐藏
            planPanel.BackColor = Color.FromArgb(248, 249, 252);
            planTitle = new Label();
            planTitle.Text = "任务计划";
            planTitle.Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold);
            planTitle.ForeColor = Color.FromArgb(46, 76, 196);
            planTitle.Dock = DockStyle.Top;
            planTitle.Height = 34;
            planTitle.Padding = new Padding(12, 8, 0, 0);
            planStepsList = new ListBox();
            planStepsList.Dock = DockStyle.Fill;
            planStepsList.BorderStyle = BorderStyle.None;
            planStepsList.BackColor = planPanel.BackColor;
            planStepsList.DrawMode = DrawMode.OwnerDrawFixed;
            planStepsList.ItemHeight = 34;
            planStepsList.Font = new Font("Microsoft YaHei UI", 9F);
            planStepsList.DrawItem += PlanStep_DrawItem;
            planPanel.Controls.Add(planStepsList);
            planPanel.Controls.Add(planTitle);
            planPanel.Visible = false;
            Controls.Add(planPanel);
            list.BringToFront();   // Fill 最后停靠，吃掉剩余宽度

            Panel bottom2 = new Panel();
            bottom = bottom2;
            bottom2.Dock = DockStyle.Bottom;
            bottom2.Height = 118;
            bottom2.BackColor = Color.White;
            bottom2.Paint += Bottom_Paint;

            inputBorder = new Panel();
            inputBorder.Location = new Point(150, 12);
            inputBorder.Size = new Size(700, 66);
            inputBorder.Paint += InputBorder_Paint;

            input = new ForwardWheelBox();
            input.Target = list;      // 焦点在输入框时滚轮转发给消息列表（Win7 滚轮只作用于焦点控件）
            input.BorderStyle = BorderStyle.None;
            input.Multiline = true;
            input.Font = new Font("Microsoft YaHei UI", 9.75F);
            input.Location = new Point(12, 8);
            input.Size = new Size(676, 50);
            input.ScrollBars = ScrollBars.Vertical;
            input.KeyDown += Input_KeyDown;
            input.TextChanged += delegate { UpdatePlaceholder(); HandleAtCaret(); };
            // Win7 默认 1px 光标几乎不可见 → 获得焦点时重建为 3px、比行高略高。
            // 时序关键：WinForms 先触发 GotFocus、控件默认 WM_SETFOCUS 处理在后——
            // 直接重建会被控件的默认细光标覆盖（用户实测"光标有点小"），
            // BeginInvoke 把重建排到默认处理之后；位置拍自控件已放好的系统光标，
            // 打字/点击/IME 一律交给控件原生维护（见 CaretHelper.Build 注释）。
            input.HandleCreated += delegate { BoldCaret(); };
            input.GotFocus += delegate { BeginInvoke((MethodInvoker)delegate { BoldCaret(); }); };

            placeholder = new Label();
            placeholder.Text = "今天帮你做些什么？可直接发文件路径让我读取，或拖入文件预览/转换";
            placeholder.ForeColor = Color.FromArgb(160, 163, 178);
            placeholder.Font = new Font("Microsoft YaHei UI", 9.75F);
            placeholder.Location = new Point(14, 10);
            placeholder.Size = new Size(650, 20);
            placeholder.BackColor = Color.Transparent;
            placeholder.Enabled = false;

            inputBorder.Controls.Add(input);
            inputBorder.Controls.Add(placeholder);
            placeholder.BringToFront();   // 否则会被输入框盖住

            send = new Button();
            send.Text = "发送";
            send.ForeColor = Color.White;
            send.BackColor = Color.FromArgb(24, 26, 32);
            send.FlatStyle = FlatStyle.Flat;
            send.FlatAppearance.BorderSize = 0;
            send.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold);
            send.Size = new Size(64, 30);
            send.Location = new Point(786, 24);
            send.Cursor = Cursors.Hand;
            send.Click += delegate { DoSend(); };
            SizeRound(send, 15);

            modelLink = new LinkLabel();
            modelLink.Text = "模型：未配置 ▾";
            modelLink.LinkColor = Color.FromArgb(90, 96, 110);
            modelLink.ActiveLinkColor = Color.FromArgb(62, 99, 221);
            modelLink.LinkBehavior = LinkBehavior.HoverUnderline;
            modelLink.AutoSize = true;
            modelLink.Location = new Point(560, 84);
            modelLink.Font = new Font("Microsoft YaHei UI", 9F);
            modelLink.Click += delegate { ShowModelMenu(); };

            ctxLabel = new Label();
            ctxLabel.Text = "上下文：空（新对话）";
            ctxLabel.ForeColor = Color.FromArgb(150, 153, 168);
            ctxLabel.Font = new Font("Microsoft YaHei UI", 8.25F);
            ctxLabel.AutoSize = true;
            ctxLabel.Location = new Point(24, 84);

            // 引用文件提示条（输入框左侧空白区）：@ 引用或拖入文件后可见
            fileChip = new Label();
            fileChip.ForeColor = Color.FromArgb(70, 73, 84);
            fileChip.Font = new Font("Microsoft YaHei UI", 8F);
            fileChip.Location = new Point(12, 10);
            fileChip.Size = new Size(132, 66);
            fileChip.Visible = false;

            bottom.Controls.Add(inputBorder);
            bottom.Controls.Add(send);
            bottom.Controls.Add(modelLink);
            bottom.Controls.Add(ctxLabel);
            bottom.Controls.Add(fileChip);
            Controls.Add(bottom);
            bottom.SendToBack();

            // @ 弹层在输入框失焦时收起（点击弹层内的列表时靠坐标判定不收）
            input.LostFocus += delegate { TryCloseAtPopupOnBlur(); };
            Disposed += delegate { if (atPopup != null) { try { atPopup.Dispose(); } catch { } atPopup = null; } };

            // 低配机优化：拖动窗口防抖（停止 200ms 后重排一次）；流式刷新走 100ms 节流定时器
            Resize += delegate
            {
                LayoutWidth();
                resizeDebounce.Stop();
                resizeDebounce.Start();
            };
            resizeDebounce = new System.Windows.Forms.Timer();
            resizeDebounce.Interval = 200;
            resizeDebounce.Tick += delegate
            {
                resizeDebounce.Stop();
                ReMeasure();
            };

            flushTimer = new System.Windows.Forms.Timer();
            flushTimer.Interval = 100;
            flushTimer.Tick += delegate
            {
                // 流式最终答复：把后台线程累积的文本刷进气泡（与下面的整表重绘节流同一节拍）
                try { FlushStream(); } catch { }
                // 等待中气泡显示已等待秒数（非流式请求体感优化，1s 粒度）
                if (busy && pendingIdx >= 0 && pendingIdx < msgs.Count && msgs[pendingIdx].Pending
                    && msgs[pendingIdx].Text.StartsWith("思考中"))
                {
                    int secs = (int)((DateTime.Now - pendingStart).TotalSeconds);
                    if (secs != pendingShownSecs && secs > 0)
                    {
                        pendingShownSecs = secs;
                        msgs[pendingIdx].Text = "思考中… " + secs + "s";
                        msgs[pendingIdx].CachedTextH = -1;
                        msgs[pendingIdx].CachedHeight = -1;
                        dirty = true;
                    }
                }
                if (!dirty) return;
                dirty = false;
                // 始终跟随到最新（用户要求：AI 回复后窗口自动下拉）
                try
                {
                    int last = list.Items.Count - 1;
                    if (last >= 0)
                    {
                        list.TopIndex = last;
                        Rectangle r = list.GetItemRectangle(last);
                        Rectangle tail = Rectangle.FromLTRB(r.Left, Math.Max(0, r.Top - 40),
                            list.ClientRectangle.Right, list.ClientRectangle.Bottom);
                        list.Invalidate(tail);
                    }
                    else list.Invalidate();
                }
                catch { list.Invalidate(); }
            };
            flushTimer.Start();
        }

        // 焦点在输入框/按钮上时，ChatPanel 也能收到滚轮并转发给消息列表
        protected override void OnMouseWheel(MouseEventArgs e)
        {
            base.OnMouseWheel(e);
            try
            {
                if (list != null && !list.IsDisposed && list.IsHandleCreated)
                {
                    // 重建 wParam：高 16 位为滚轮增量（WHEEL_DELTA=120）
                    IntPtr wp = (IntPtr)(((e.Delta & 0xFFFF) << 16));
                    SendMessage(list.Handle, WM_MOUSEWHEEL, wp, (IntPtr)0);
                }
            }
            catch { }
        }

        // 标记脏：不立即重绘，由 flushTimer 节流（最多 10 次/秒）
        void MarkDirty()
        {
            dirty = true;
        }

        // 尺寸变化后重算高度缓存（OwnerDrawVariable 的高度查询依赖此缓存，本身代价低）
        void ReMeasure()
        {
            if (list == null || msgs.Count == 0) return;
            list.BeginUpdate();
            list.Items.Clear();
            foreach (ChatMsg m in msgs) list.Items.Add(m);
            list.EndUpdate();
            list.Invalidate();
        }

        static void SizeRound(Control c, int radius)
        {
            GraphicsPath gp = new GraphicsPath();
            Rectangle r = new Rectangle(0, 0, c.Width, c.Height);
            gp.AddArc(r.X, r.Y, radius * 2, radius * 2, 180, 90);
            gp.AddArc(r.Right - radius * 2, r.Y, radius * 2, radius * 2, 270, 90);
            gp.AddArc(r.Right - radius * 2, r.Bottom - radius * 2, radius * 2, radius * 2, 0, 90);
            gp.AddArc(r.X, r.Bottom - radius * 2, radius * 2, radius * 2, 90, 90);
            gp.CloseFigure();
            c.Region = new Region(gp);
        }

        void LayoutWidth()
        {
            int w = Width;
            if (w < 600) return;
            int inputW = Math.Max(300, w - 150 - 90 - 60);
            inputBorder.Width = inputW;
            input.Width = inputW - 24;
            send.Location = new Point(150 + inputW + 12, 24);
            modelLink.Location = new Point(150 + inputW - modelLink.Width - 6, 84);
        }

        public void ApplyConfig(AppConfig cfg)
        {
            config = cfg;
            client = IsConfigured() ? new LlmClient(config) : null;
            modelLink.Text = IsConfigured() ? ("模型：" + config.Model + (config.WebSearch ? "（联网） ▾" : " ▾")) : "模型：未配置，点击设置 ▾";
            // 工具层根目录 = 当前激活工作区的文件夹（建表/建PPT 相对路径、下载、@引用 的默认落盘位置）
            try { AgentTools.WorkspaceRoot = WorkspaceStore.ActiveDir(config); } catch { }
            LayoutWidth();
        }

        // 当前会话 id（侧边栏删除/归档会话时判断是否正在使用该会话）
        public string CurrentSessionId { get { return sessionId; } }

        // 新会话归属的当前工作区 id（侧边栏切换工作区后，新会话落到对应项目下）
        string CurrentWorkspaceId()
        {
            try
            {
                WorkspaceInfo w = WorkspaceStore.Active(config);
                if (w != null) return w.Id;
            }
            catch { }
            return "default";
        }

        public bool IsConfigured()
        {
            return config != null && config.BaseUrl != null && config.BaseUrl.Length > 0
                && config.Model != null && config.Model.Length > 0 && config.GetKey() != null;
        }

        public void NewConversation()
        {
            // 先作废流式状态：msgs/list.Items 即将被清空，若不复位 streamingBubble，
            // 残留的旧索引会在新会话里"看起来合法"，把流式文本写进新会话的无关消息。
            CancelStreamBubble();
            sessionId = null;            // 新会话：首次发言时才创建 id
            msgs.Clear();
            list.Items.Clear();
            llmHistory.Clear();          // 新会话 = 模型上下文一并清空
            sessionRows.Clear();
            planSteps.Clear();
            planPanel.Visible = false;   // 计划栏随新会话收起
            contextFiles.Clear();        // 新会话不带旧文件上下文
            RefreshFileChip();
            if (IsConfigured())
                Append("assistant", "就绪。可以：直接提问；发我文件路径（如 D:\\报告.txt）我会读取内容；让我上网查资料、下载文件、转换文档，或启动环境修复。");
            else
                Append("assistant", "当前未配置模型：填好 API 地址与密钥后即可对话。文件预览与批量转换无需模型，现在就能用。");
            UpdateCtxLabel();
        }

        // 恢复最近一次会话（无历史时开新会话）
        public void RestoreLastSession()
        {
            List<SessionInfo> all = SessionStore.List();
            if (all.Count == 0) { NewConversation(); return; }
            LoadSession(all[0].Id);
        }

        // 加载历史会话（侧边栏点击）：重建气泡 + 模型上下文
        public void LoadSession(string id)
        {
            List<string[]> rows = SessionStore.Load(id);
            if (rows.Count == 0) { NewConversation(); return; }
            CancelStreamBubble();   // 同上：清列表前必须作废流式索引
            sessionId = id;
            msgs.Clear();
            list.Items.Clear();
            lock (llmHistory) { llmHistory.Clear(); }
            sessionRows = new List<string[]>(rows);
            foreach (string[] r in rows)
            {
                if (r[0] == "file")
                {
                    // 附件卡片行：不进模型上下文，只还原卡片
                    ChatMsg fm = new ChatMsg();
                    fm.AttachPath = r[2];
                    fm.Time = r[1];
                    msgs.Add(fm);
                    list.Items.Add(fm);
                    continue;
                }
                ChatMsg m = new ChatMsg();
                m.User = r[0] == "user";
                m.Time = r[1];
                m.Text = r[2];
                msgs.Add(m);
                list.Items.Add(m);
                lock (llmHistory) { llmHistory.Add(new LlmTurn(m.User ? "user" : "assistant", r[2])); }
            }
            // 长消息重建分块（倒序插入不影响前面的索引）
            for (int i = msgs.Count - 1; i >= 0; i--)
            {
                MeasureMsg(msgs[i], AvailWidth());
                if (msgs[i].CachedTextH > MaxChunkTextH) SplitLongFinal(i);
            }
            UpdateCtxLabel();
            MarkDirty();
        }

        // 会话落盘：存「显示文本」而非模型载荷——重载后气泡所见即所存，
        // 且不会把文件上下文/脱敏产物写进磁盘；模型上下文重载时以显示文本重建（旧文件可让模型用工具重读）。
        void SaveSession()
        {
            try
            {
                if (sessionId == null || sessionRows.Count == 0) return;
                string title = "";
                foreach (string[] r in sessionRows)
                {
                    if (r[0] == "user" && r[2].Length > 0) { title = r[2].Length > 24 ? r[2].Substring(0, 24) : r[2]; break; }
                }
                SessionStore.Save(sessionId, title, new List<string[]>(sessionRows), CurrentWorkspaceId());
                if (OnSessionSaved != null) { try { OnSessionSaved(); } catch { } }
            }
            catch { }
        }

        public void FocusInput() { input.Focus(); }

        public void SetContextFiles(string[] files)
        {
            contextFiles.Clear();
            if (files != null)
            {
                foreach (string f in files)
                {
                    if (f != null && f.Length > 0 && File.Exists(f) && !contextFiles.Contains(f)) contextFiles.Add(f);
                }
            }
            RefreshFileChip();
        }

        // 追加一个引用文件（@ 引用弹层 / 拖拽路径共用），去重
        public void AddContextFile(string path)
        {
            if (path == null || path.Length == 0 || !File.Exists(path)) return;
            if (!contextFiles.Contains(path)) contextFiles.Add(path);
            RefreshFileChip();
        }

        void RefreshFileChip()
        {
            if (fileChip == null) return;
            if (contextFiles.Count == 0) { fileChip.Visible = false; return; }
            StringBuilder sb = new StringBuilder();
            sb.Append("引用 ").Append(contextFiles.Count).Append(" 个文件\n");
            int shown = 0;
            foreach (string f in contextFiles)
            {
                if (shown >= 3) { sb.Append("…"); break; }
                sb.Append("· ").Append(Path.GetFileName(f)).Append("\n");
                shown++;
            }
            fileChip.Text = sb.ToString().TrimEnd();
            fileChip.Visible = true;
        }

        void Input_KeyDown(object sender, KeyEventArgs e)
        {
            // @ 引用弹层可见时：方向键选词、回车插入、Esc 收起（不发送消息）
            if (atActive)
            {
                if (e.KeyCode == Keys.Down) { MoveAtSelection(1); e.SuppressKeyPress = true; e.Handled = true; return; }
                if (e.KeyCode == Keys.Up) { MoveAtSelection(-1); e.SuppressKeyPress = true; e.Handled = true; return; }
                if (e.KeyCode == Keys.Enter) { InsertAtSelection(); e.SuppressKeyPress = true; e.Handled = true; return; }
                if (e.KeyCode == Keys.Escape) { CloseAtPopup(); e.SuppressKeyPress = true; e.Handled = true; return; }
            }
            if (e.KeyCode == Keys.Enter && !e.Shift)
            {
                e.SuppressKeyPress = true;
                DoSend();
            }
        }

        void UpdatePlaceholder()
        {
            placeholder.Visible = input.Text.Length == 0;
        }

        void DoSend()
        {
            string text = input.Text.Trim();
            if (text.Length == 0 || busy) return;
            if (sessionId == null) sessionId = SessionStore.NewId();   // 会话首次发言时创建

            // ---- 已确认动作计划 ----
            if (pendingPlan != null && IsConfirmText(text))
            {
                input.Text = "";
                ActionPlan confirmed = pendingPlan;
                pendingPlan = null;
                Append("user", text);
                ExecutePlan(confirmed);
                return;
            }
            if (pendingPlan != null && IsCancelText(text))
            {
                pendingPlan.State = ActionState.Cancelled;
                pendingPlan = null;
                input.Text = "";
                Append("user", text);
                Append("assistant", "已取消，不会创建或修改任何输出文件。");
                AuditLog.Record("confirmation", "cancelled");
                return;
            }

            // ---- 规则优先的自然语言动作路由 ----
            // 可执行 → 走确认流程；输入齐了只缺列映射 → 建议通道自动补参；
            // 其余缺参情况不再死挡（此前任何提到「汇总/核对」的消息都被模板接住回
            // 「请补齐参数」，模型根本收不到——用户要的是 agent 不是表单），转给 agent 处理，
            // 路由诊断作为提示随请求发给模型（只进载荷，不显示、不进存档）。
            ActionPlan routed;
            string agentHint = "";
            if (IntentRouter.TryCreate(text, contextFiles, out routed))
            {
                ActionPlan plan = EnrichPlan(routed);
                bool mappingOnly = !plan.IsExecutable && plan.Kind == ActionKind.Recon && plan.Inputs.Count >= 2;
                if (plan.IsExecutable || mappingOnly)
                {
                    pendingPlan = plan;
                    input.Text = "";
                    Append("user", text);
                    string confirm1 = "\u201C确认\u201D";
                    string cancel1 = "\u201C取消\u201D";
                    string prompt = plan.IsExecutable
                        ? "我准备执行以下任务：\n" + plan.DisplayText() + "\n\n回复" + confirm1 + "执行，回复" + cancel1 + "。"
                        : "我识别到一个任务，但还不能执行：\n" + plan.DisplayText() + "\n\n请先补齐参数，或回复" + cancel1 + "。";
                    Append("assistant", prompt);
                    AuditLog.Record("plan_created", plan.Kind.ToString() + "; executable=" + plan.IsExecutable);
                    if (!plan.IsExecutable && IsConfigured() && config.PrivacyLevel > 0)
                    {
                        StartColumnSuggest(plan);
                    }
                    return;
                }
                // 规划层第三期：把规则命中作为**提示**（而非指令）喂给模型。
                // 第二期这里是"命中模板但缺参数"的诊断文本，只有缺参时才产生；
                // 第三期改为**无论是否可执行**都产生提示——因为规则的价值在于
                // "用户换了个说法时给模型一个先验"，而这恰恰发生在规则**不完全命中**的时候。
                // 提示的措辞与免责声明见 RouteHint 类注释（核心：模型有权推翻它）。
                agentHint = RouteHint.Build(plan, text);
            }

            // ---- 内置记忆模块命令（离线可用，开箱即用）----
            if (config.PluginMemory)
            {
                if (text == "记忆" || text == "记忆列表")
                {
                    List<MemoryEntry> list = MemoryStore.LoadMemories();
                    if (list.Count == 0) Append("assistant", "长期记忆为空。用「记住 xxx」可以让我记住事项，例如：记住 我是一名会计。");
                    else
                    {
                        StringBuilder sb = new StringBuilder();
                        sb.Append("当前长期记忆 ").Append(list.Count).Append(" 条：\n");
                        foreach (MemoryEntry m in list) sb.Append("· ").Append(m.Text).Append("（").Append(m.Time).Append("）\n");
                        sb.Append("用「忘记 关键词」可删除。");
                        Append("assistant", sb.ToString());
                    }
                    input.Text = "";
                    return;
                }
                if (text.StartsWith("记住 "))
                {
                    string what = text.Substring(3).Trim();
                    if (what.Length > 0)
                    {
                        MemoryStore.AddMemory(what);
                        Append("assistant", "已记住：" + what + "\n之后对话我会自动带上这条记忆（可用「记忆」查看，离线也可用）。");
                    }
                    input.Text = "";
                    return;
                }
                if (text.StartsWith("忘记 "))
                {
                    int n = MemoryStore.RemoveMemoryLike(text.Substring(3).Trim());
                    Append("assistant", n > 0 ? ("已删除 " + n + " 条相关记忆。") : "没有找到包含该关键词的记忆。");
                    input.Text = "";
                    return;
                }
            }

            if (!IsConfigured())
            {
                Append("assistant", "尚未配置模型。请点击右下角「模型：未配置 ▾」填写 API 地址与密钥。", false, true);
                if (OnOpenSettings != null) OnOpenSettings();
                return;
            }
            input.Text = "";
            Append("user", text);
            if (config.PluginHist) MemoryStore.AppendHistory("user", text);
            int idx = Append("assistant", "思考中…", true);
            pendingIdx = idx;
            pendingStart = DateTime.Now;
            pendingShownSecs = -1;
            busy = true;
            send.Enabled = false;
            ResetStreamState();   // 新一轮：允许创建流式预览气泡
            string question = text;

            // 隐私分级（设计方案 §7.2）：L0 禁止出网；L1 出网前规则脱敏（命中规则交审计）
            List<string> maskHits = new List<string>();
            if (config.PrivacyLevel == 0)
            {
                FinishReply(idx, "当前隐私等级为「全本地（不出网）」，模型对话不可用。\n" +
                    "如需对话，请到「内置插件」页把隐私等级调整为「脱敏出网」或「全量出网」。" +
                    "文件预览、转换与表格核对等本地功能不受影响。", false, 0);
                return;
            }

            // 消息里/拖拽上下文中的文本文件 → 直接读出内容附进提问（修：AI 看不到 txt 等基本文件）
            string fileCtx = BuildFileContext(text, maskHits);
            if (config.PrivacyLevel == 1) question = MaskEngine.Mask(question, maskHits);
            AuditLog.Record("llm_call",
                (maskHits.Count > 0 ? "masked[" + string.Join(",", maskHits.ToArray()) + "] " : "plain ") +
                "len=" + (question.Length + fileCtx.Length));

            string sysPrompt = BuildSystemPrompt();
            string userPayload = question + fileCtx + agentHint;
            userPayloadForHistory = userPayload;
            lastUserDisplay = text;   // 会话存档用原文（不含脱敏与文件上下文）
            Thread t = new Thread(new ThreadStart(delegate
            {
                System.Diagnostics.Stopwatch sw = System.Diagnostics.Stopwatch.StartNew();
                RunAgentLoop(idx, sysPrompt, userPayload, sw);
            }));
            t.IsBackground = true;
            t.Start();
        }

        // 抓取消息文本与拖拽上下文里的文本文件内容（最多 3 个，每个截 6000 字符；L1 时随问题一起脱敏）
        string BuildFileContext(string text, List<string> maskHits)
        {
            List<string> paths = new List<string>();
            try
            {
                foreach (string p in IntentRouter.ExtractPaths(text)) { if (!paths.Contains(p)) paths.Add(p); }
                foreach (string f in contextFiles)
                {
                    if (f != null && File.Exists(f) && !paths.Contains(f)) paths.Add(f);
                }
            }
            catch { }
            if (paths.Count == 0) return "";
            StringBuilder sb = new StringBuilder();
            int used = 0;
            foreach (string p in paths)
            {
                if (used >= 3) break;
                string ext = (Path.GetExtension(p) ?? "").ToLowerInvariant();
                if (ext == ".xlsx" || ext == ".xls" || ext == ".pdf" || ext == ".doc" || ext == ".docx" ||
                    ext == ".ppt" || ext == ".pptx" || ext == ".png" || ext == ".jpg")
                {
                    sb.Append("\n\n【附带文件 ").Append(p).Append("】（二进制文件，内容未附带；可让我用 convert_document 转换后阅读）");
                    used++;
                    continue;
                }
                bool ok;
                string content = AgentTools.ReadTextFilePublic(p, out ok);
                if (!ok || content.Length == 0) continue;
                if (content.Length > 6000) content = content.Substring(0, 6000) + "\n…[截断，可用工具 read_text_file 读更多]";
                if (config.PrivacyLevel == 1) content = MaskEngine.Mask(content, maskHits);
                sb.Append("\n\n【附带文件 ").Append(p).Append("】\n").Append(content);
                used++;
            }
            return sb.ToString();
        }

        // Agent 主循环（共享实现见 AgentLoop.Run）：模型 ↔ 白名单工具，取最终答复后回填气泡
        void RunAgentLoop(int idx, string sysPrompt, string userPayload, System.Diagnostics.Stopwatch sw)
        {
            // 历史快照（后台线程读，UI 线程写，加锁拷贝）
            List<LlmTurn> hist;
            lock (llmHistory) { hist = new List<LlmTurn>(llmHistory); }
            // 流式接收槽：最终答复逐段到达时，先落到这里再由 UI 定时刷新，
            // 避免每个 token 都 Invoke 一次把 UI 线程压满。
            streamText = "";
            streamDirty = false;
            AgentLoop.Result r = AgentLoop.Run(client, config, sysPrompt, userPayload,
                delegate(string log) { SetBubbleText(idx, TailOfLog(log) + "\n…"); },
                HandlePlanTool, hist,
                delegate(string acc) { streamText = acc; streamDirty = true; });
            sw.Stop();
            if (InvokeRequired) { try { Invoke((MethodInvoker)delegate { AgentDone(idx, r, sw.ElapsedMilliseconds); }); } catch { } }
            else AgentDone(idx, r, sw.ElapsedMilliseconds);
        }

        // 流式预览：后台线程写 streamText，UI 定时器读并刷新气泡（见 flushTimer）
        string streamText = "";
        volatile bool streamDirty;
        int streamingBubble = -1;
        // 收尾后置 true：阻止定时器在 AgentDone 之后又重建一条预览气泡（会变成重复答复）
        volatile bool streamClosed;

        // 开一条"正在输入"的气泡，后续流式文本回填到它（没有流式时保持隐藏）
        void BeginStreamBubble()
        {
            if (streamingBubble >= 0) return;
            streamingBubble = Append("assistant", "", true, false);
        }

        // UI 定时器：把最新流式文本刷进气泡（增量更新，只在有变化时重绘）
        void FlushStream()
        {
            if (streamClosed) { streamDirty = false; return; }
            if (!streamDirty) return;
            // 再次复查关闸：本方法在 UI 线程由定时器驱动，DiscardStreamBubble/CancelStreamBubble
            // 也在 UI 线程，二者理论上不会交错；但"先检查、后使用"之间隔着一次 BeginStreamBubble
            // 调用，未来若有人在其中插入消息泵就会变成真实竞态。这里廉价地再挡一次，
            // 保证关闸后绝不会重建预览气泡（重建会导致正文重复显示）。
            if (streamClosed) { streamDirty = false; return; }
            streamDirty = false;
            if (streamingBubble < 0) BeginStreamBubble();
            if (streamingBubble >= 0 && streamingBubble < msgs.Count)
            {
                msgs[streamingBubble].Text = streamText;
                msgs[streamingBubble].Pending = false;
                msgs[streamingBubble].CachedTextH = -1; msgs[streamingBubble].CachedHeight = -1;
                dirty = true;   // 交给同一定时器里的滚动/重绘逻辑
            }
        }

        // 新一轮对话开始前复位（DoSend 里调用），允许再次创建预览气泡
        void ResetStreamState()
        {
            streamClosed = false;
            streamingBubble = -1;
            streamText = "";
            streamDirty = false;
        }

        void EndStreamBubble()
        {
            streamingBubble = -1;
            streamText = "";
            streamDirty = false;
        }

        // 移除流式预览气泡（收尾时调用；正式答复会重新创建一条，避免重复显示）
        void DiscardStreamBubble()
        {
            streamClosed = true;   // 先关闸，防定时器在删除后又重建
            if (streamingBubble >= 0 && streamingBubble < msgs.Count && streamingBubble < list.Items.Count)
            {
                msgs.RemoveAt(streamingBubble);
                list.Items.RemoveAt(streamingBubble);
            }
            EndStreamBubble();
        }

        // 放弃流式状态但不动 msgs/list（调用方马上就要整体 Clear 的场景：新建/切换会话）。
        // 必须先关闸再清列表：否则定时器可能在列表清空后按残留索引写入无关消息。
        void CancelStreamBubble()
        {
            streamClosed = true;
            EndStreamBubble();
        }

        // 流式期间的日志视图：ListBox 单条目高度超过视口后底部不可见（TopIndex 只能钉住条目顶部），
        // 表现为"AI 回复时不自动滚动"。进行中只显示最近几步保持气泡矮小，完整日志结束后再分块。
        static string TailOfLog(string log)
        {
            if (log == null || log.Length == 0) return log;
            string[] lines = log.Split('\n');
            const int Keep = 6;
            if (lines.Length <= Keep) return log;
            StringBuilder sb = new StringBuilder();
            sb.Append("…前 ").Append(lines.Length - Keep).Append(" 步");
            for (int i = lines.Length - Keep; i < lines.Length; i++) sb.Append('\n').Append(lines[i]);
            return sb.ToString();
        }

        // UI 线程收尾：写最终答复 + 历史落账
        void AgentDone(int idx, AgentLoop.Result r, long elapsedMs)
        {
            busy = false;
            send.Enabled = true;
            RememberTurn(r);   // 目标继承：记下本轮真实用过的技能与产物（供下一轮指代使用）
            // 流式预览气泡只用于"边写边看"，收尾时移除，由下面的正式答复路径重建，
            // 否则会出现同一条答复显示两次。
            DiscardStreamBubble();
            if (r.Error.Length > 0)
            {
                FinishReply(idx, "✗ " + r.Error, true, 0);
                return;
            }
            if (r.UsedTools && r.ToolLog.Length > 0)
            {
                // 工具调用记录留在原气泡；最终答复另起一条（顺序：日志在上、答复在下）
                if (idx >= 0 && idx < msgs.Count)
                {
                    msgs[idx].Text = r.ToolLog;
                    msgs[idx].Pending = false;
                    msgs[idx].CachedTextH = -1; msgs[idx].CachedHeight = -1;
                    SplitLongFinal(idx);   // 完整日志按高度分块，避免单条超出视口又看不到底
                }
                int fidx = Append("assistant", r.FinalText, false, false);
                if (config.PluginHist) MemoryStore.AppendHistory("assistant", r.FinalText);
                if (fidx >= 0 && fidx < msgs.Count)
                {
                    msgs[fidx].Time = msgs[fidx].Time + " · " + (elapsedMs / 1000.0).ToString("0.0") + "s";
                    SplitLongFinal(fidx);
                }
            }
            else
            {
                FinishReply(idx, r.FinalText, false, elapsedMs);
            }
            lock (llmHistory)
            {
                llmHistory.Add(new LlmTurn("user", userPayloadForHistory));
                llmHistory.Add(new LlmTurn("assistant", r.FinalText));
                TrimHistory();
            }
            // 会话存档记「显示文本」：用户存原文（非脱敏载荷），助手存最终答复
            string now = DateTime.Now.ToString("HH:mm");
            sessionRows.Add(new string[] { "user", now, lastUserDisplay.Length > 0 ? lastUserDisplay : userPayloadForHistory });
            sessionRows.Add(new string[] { "assistant", now, r.FinalText });
            UpdateCtxLabel();
            SaveSession();
            // 工具产物 → 附件卡片（跟在回复后面，点击进预览；去重避免重复转换出重复卡）
            if (r.Products != null)
            {
                List<string> seen = new List<string>();
                foreach (string fp in r.Products)
                {
                    if (fp == null || fp.Length == 0 || seen.Contains(fp)) continue;
                    seen.Add(fp);
                    if (File.Exists(fp)) AddAttachment(fp);
                }
            }
            MarkDirty();
        }

        // 同名产物避让提示（UI 线程调用；由 MainForm 从后台线程 BeginInvoke 过来）。
        // 只提示不阻塞：旧文件已保留，新产物落在不重名的路径上，用户无需做任何决定。
        public void NotifyOverwriteAvoided(string existingPath)
        {
            try
            {
                string nm = existingPath == null ? "" : Path.GetFileName(existingPath);
                Append("assistant", "注意：工作区里已经有「" + nm + "」，我没有覆盖它（旧文件保持不动），" +
                    "新产物会自动存成不重名的文件名。如果你确实想覆盖，先把旧文件删掉再让我重做。", false, false);
                MarkDirty();
            }
            catch { }
        }

        // 追加附件卡片气泡（并持久化到会话）
        void AddAttachment(string path)
        {
            ChatMsg m = new ChatMsg();
            m.AttachPath = path;
            m.Time = DateTime.Now.ToString("HH:mm");
            msgs.Add(m);
            list.Items.Add(m);
            sessionRows.Add(new string[] { "file", m.Time, path });
            SaveSession();
            MarkDirty();
        }

        // 上下文占用指示：条数 + token 粗估 + 模型上下文窗口占比（≥80% 变红警告）
        void UpdateCtxLabel()
        {
            if (ctxLabel == null) return;
            int turns = 0; long chars = 0, cjk = 0;
            lock (llmHistory)
            {
                turns = llmHistory.Count;
                foreach (LlmTurn t in llmHistory)
                {
                    string c = t.Content == null ? "" : t.Content;
                    chars += c.Length;
                    foreach (char ch in c)
                    {
                        if (ch >= 0x2E80) cjk++;   // CJK/全角区按 1 token 计
                    }
                }
            }
            long tokens = cjk + (chars - cjk) / 3;   // 粗估：中文约 1 token/字，其余约 1/3
            int win = ModelWindow(config == null ? "" : config.Model);
            if (win > 0)
            {
                long pct = tokens * 100 / win;
                ctxLabel.Text = "上下文 " + turns + " 条 · 约 " + tokens + " / " + (win / 1000) + "K tokens（" + pct + "%）";
                ctxLabel.ForeColor = pct >= 80 ? Color.FromArgb(200, 60, 40) : Color.FromArgb(150, 153, 168);
            }
            else
            {
                string size = chars >= 10000 ? (chars / 1000.0).ToString("0.#") + "k" : chars.ToString();
                ctxLabel.Text = turns == 0 ? "上下文：空（新对话）" : "上下文 " + turns + " 条 · 约 " + size + " 字";
            }
        }

        // 常见模型的上下文窗口（tokens）；未知模型返回 0 = 不显示占比
        static int ModelWindow(string model)
        {
            string m = (model ?? "").ToLowerInvariant();
            if (m.Length == 0) return 0;
            if (m.Contains("glm-4-long")) return 1000000;
            if (m.Contains("glm")) return 128000;
            if (m.Contains("deepseek")) return 64000;
            if (m.Contains("qwen")) return 128000;
            if (m.Contains("gpt")) return 128000;
            return 0;
        }

        // 当前历史 token 粗估（须在 lock(llmHistory) 内调用）
        long EstTokensLocked()
        {
            long chars = 0, cjk = 0;
            foreach (LlmTurn t in llmHistory)
            {
                string c = t.Content == null ? "" : t.Content;
                chars += c.Length;
                foreach (char ch in c) { if (ch >= 0x2E80) cjk++; }
            }
            return cjk + (chars - cjk) / 3;
        }

        // RunAgentLoop 与 AgentDone 之间的用户消息载荷（含文件上下文）
        string userPayloadForHistory;

        void SetBubbleText(int idx, string text)
        {
            if (InvokeRequired) { try { Invoke((MethodInvoker)delegate { SetBubbleText(idx, text); }); } catch { } return; }
            if (idx >= 0 && idx < msgs.Count)
            {
                msgs[idx].Text = text;
                msgs[idx].CachedTextH = -1; msgs[idx].CachedHeight = -1;
                MarkDirty();
            }
        }

        void TrimHistory()
        {
            while (llmHistory.Count > MaxHistoryTurns) llmHistory.RemoveAt(0);
            // 字符预算：附了文件上下文的历史可能极大，超预算从最旧开始丢（至少保留最近一问一答）
            while (llmHistory.Count > 2)
            {
                long total = 0;
                foreach (LlmTurn t in llmHistory) total += (t.Content == null ? 0 : t.Content.Length);
                if (total <= MaxHistoryChars) break;
                llmHistory.RemoveAt(0);
            }
            // 窗口预算：token 估算达到模型窗口 85% 时自动裁剪到 60% 以下，防止上下文爆窗
            int win = ModelWindow(config == null ? "" : config.Model);
            if (win > 0)
            {
                lock (llmHistory)
                {
                    while (llmHistory.Count > 2 && EstTokensLocked() * 100 >= (long)win * 85)
                    {
                        llmHistory.RemoveAt(0);
                    }
                }
            }
        }

        ActionPlan EnrichPlan(ActionPlan plan)
        {
            if (plan.Kind == ActionKind.Recon)
            {
                ReconTemplate t = ReconTemplateStore.Get(ReconTemplateStore.LastName);
                if (t != null && t.KeyColsA.Length > 0 && t.KeyColsB.Length > 0)
                {
                    plan.Recon = t.ToMapping();
                    plan.TemplateName = t.Name;
                    plan.Missing = plan.Inputs.Count >= 2 ? new string[0] : new string[] { "两个输入文件" };
                }
                else if (plan.Inputs.Count >= 2) plan.Missing = new string[] { "核对列映射" };
            }
            else if (plan.Kind == ActionKind.Merge)
            {
                List<MergeTemplate> all = MergeTemplateStore.LoadAll();
                if (all.Count > 0)
                {
                    plan.Merge = all[0];
                    plan.TemplateName = all[0].Name;
                    plan.Missing = plan.Inputs.Count > 0 ? new string[0] : new string[] { "至少一个输入文件" };
                }
            }
            return plan;
        }
        bool IsConfirmText(string s) { return s == "确认" || s == "执行" || s == "好" || s == "开始" || s == "确认执行"; }
        bool IsCancelText(string s) { return s == "取消" || s == "不要执行" || s == "算了" || s == "放弃"; }

        // 列映射建议：只发表头（L1 先脱敏），建议列名超白名单即整体拒绝；回填后仍需用户"确认"
        void StartColumnSuggest(ActionPlan plan)
        {
            string fileA = plan.Inputs[0], fileB = plan.Inputs[1];
            Thread t = new Thread(new ThreadStart(delegate
            {
                string[] ha = ColumnSuggest.ReadHeaders(fileA, "", 1);
                string[] hb = ColumnSuggest.ReadHeaders(fileB, "", 1);
                SuggestResult sug;
                if (ha.Length == 0 || hb.Length == 0)
                {
                    sug = new SuggestResult();
                    sug.Error = "读取表头失败";
                }
                else
                {
                    sug = ColumnSuggest.ForRecon(client, ColumnSuggest.HeadersText(ha),
                        ColumnSuggest.HeadersText(hb), config.PrivacyLevel);
                }
                if (InvokeRequired) { try { Invoke((MethodInvoker)delegate { ApplySuggestion(plan, sug); }); } catch { } }
                else ApplySuggestion(plan, sug);
            }));
            t.IsBackground = true;
            t.Start();
        }

        void ApplySuggestion(ActionPlan plan, SuggestResult sug)
        {
            if (pendingPlan != plan) return;   // 用户已开新计划，过期建议作废
            if (sug == null || !sug.Ok)
            {
                string reason = sug == null ? "未知原因" : sug.Error;
                Append("assistant", "模型无法给出可靠的列映射建议（" + reason + "）。\n请打开「表格核对」页手工配置。");
                return;
            }
            plan.Recon = sug.Mapping;
            plan.Missing = new string[0];
            pendingPlan = plan;
            StringBuilder sb = new StringBuilder();
            sb.Append("模型建议的列映射（请人工核对列含义后回复「确认」执行）：\n");
            sb.Append("键列 A：").Append(string.Join(",", sug.Mapping.KeyColsA))
              .Append("  B：").Append(string.Join(",", sug.Mapping.KeyColsB)).Append("\n");
            sb.Append("金额 A：").Append(sug.Mapping.DebitA).Append(" / ").Append(sug.Mapping.CreditA)
              .Append("  B：").Append(sug.Mapping.DebitB).Append(" / ").Append(sug.Mapping.CreditB).Append("\n");
            sb.Append("容差：±").Append(sug.Mapping.Tolerance.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture));
            if (sug.Explain != null && sug.Explain.Length > 0) sb.Append("\n理由：").Append(sug.Explain);
            Append("assistant", sb.ToString());
        }
        void ExecutePlan(ActionPlan plan)
        {
            if (!plan.IsExecutable)
            {
                Append("assistant", "当前计划参数还不完整，不能执行。请先补齐参数或打开对应任务页。", false, true);
                return;
            }
            int idx = Append("assistant", "执行中…", true);
            busy = true;
            send.Enabled = false;
            Thread t = new Thread(new ThreadStart(delegate
            {
                ConvertEngine conv = new ConvertEngine();
                conv.SofficePath = ConvertEngine.FindSoffice(EnvDetect.FindRoot());
                ActionExecutor executor = new ActionExecutor(conv);
                ActionExecutionResult result = executor.Run(plan);
                StringBuilder sb = new StringBuilder();
                if (result.State == ActionState.Succeeded || result.State == ActionState.Partial)
                    sb.Append("已完成：").Append(result.Message);
                else
                    sb.Append("执行失败：").Append(result.Message);
                if (result.OutputPath != null && result.OutputPath.Length > 0)
                    sb.Append("\n输出：").Append(result.OutputPath);
                FinishReply(idx, sb.ToString(), result.State == ActionState.Failed, result.ElapsedMs);
            }));
            t.IsBackground = true;
            t.Start();
        }

        // 内联模型下拉（用户要求：点右下角直接选模型，不必进设置页）
        void ShowModelMenu()
        {
            ContextMenu m = new ContextMenu();
            string cur = config == null ? "" : (config.Model ?? "");
            List<string> items = new List<string>();
            if (fetchedModels != null) { foreach (string f in fetchedModels) if (!items.Contains(f)) items.Add(f); }
            foreach (string p in LlmClient.ModelPresets) if (!items.Contains(p)) items.Add(p);
            if (cur.Length > 0 && !items.Contains(cur)) items.Insert(0, cur);
            foreach (string s in items)
            {
                MenuItem mi = new MenuItem(s);
                mi.Checked = s == cur;
                string val = s;
                mi.Click += delegate { SwitchModel(val); };
                m.MenuItems.Add(mi);
            }
            m.MenuItems.Add("-");
            MenuItem refresh = new MenuItem("从服务刷新模型列表…");
            refresh.Click += delegate { FetchModelsThenReopen(); };
            MenuItem more = new MenuItem("打开模型设置…");
            more.Click += delegate { if (OnOpenSettings != null) OnOpenSettings(); };
            m.MenuItems.Add(refresh);
            m.MenuItems.Add(more);
            m.Show(modelLink, new Point(0, modelLink.Height));
        }

        void SwitchModel(string model)
        {
            if (config == null || model == null || model.Length == 0 || model == config.Model) return;
            config.Model = model;
            config.Save();
            client = IsConfigured() ? new LlmClient(config) : null;
            ApplyConfig(config);
            Append("assistant", "已切换模型：" + model + "。");
        }

        void FetchModelsThenReopen()
        {
            if (client == null) { Append("assistant", "尚未配置服务地址与密钥，无法获取模型列表。", false, true); return; }
            Thread t = new Thread(new ThreadStart(delegate
            {
                string err;
                List<string> models = client.ListModels(out err);
                if (InvokeRequired)
                {
                    Invoke((MethodInvoker)delegate
                    {
                        if (models != null) { fetchedModels = models; ShowModelMenu(); }
                        else Append("assistant", "获取模型列表失败：" + err + "\n可直接从预置列表选择，或打开模型设置手动填写。", false, true);
                    });
                }
                else if (models != null) { fetchedModels = models; ShowModelMenu(); }
            }));
            t.IsBackground = true;
            t.Start();
        }

        // 光标管理：仅在获得焦点/句柄创建时重建一次（2px 宽、与字号等高），见 CaretHelper.Build。
        // 打字/点击/输入法过程的光标位置由编辑控件原生维护，此处不再插手。
        void BoldCaret() { CaretHelper.Build(input, TextFont); }

        // ===== @ 引用文件：输入 @ 弹出工作区文件列表（上下键选、回车插入）=====

        // 弹层条目
        class AtItem
        {
            public string Path = "";
            public string Label = "";
            public bool IsDir = false;
            public bool IsBrowse = false;
            public override string ToString() { return Label; }
        }

        string atRelDir = "";   // 弹层当前浏览的子目录（相对工作区）

        // 输入框文本/光标变化时检测 @ 触发：'@' 位于行首或空白之后，到光标间无空白
        void HandleAtCaret()
        {
            try
            {
                if (atSuspend) return;
                string t = input.Text ?? "";
                int caret = input.SelectionStart;
                int at = -1;
                for (int i = caret - 1; i >= 0; i--)
                {
                    char c = t[i];
                    if (c == '@') { at = i; break; }
                    if (char.IsWhiteSpace(c)) break;
                }
                if (at < 0 || (at > 0 && !char.IsWhiteSpace(t[at - 1]))) { CloseAtPopup(); return; }
                atTokenStart = at;
                atCaret = caret;
                ShowAtPopup(t.Substring(at + 1, caret - at - 1));
            }
            catch { CloseAtPopup(); }
        }

        // '@' 后支持「子目录\文件名」渐进浏览；目录限定在工作区内（拒绝 ..）
        void ShowAtPopup(string filter)
        {
            try
            {
                string root = AgentTools.WorkspaceRoot;
                if (root == null || root.Length == 0) root = config.EffectiveWorkspace();
                if (root == null || root.Length == 0) { CloseAtPopup(); return; }

                string relDir = "";
                string namePart = filter == null ? "" : filter;
                int slash = namePart.LastIndexOf('\\');
                if (slash >= 0) { relDir = namePart.Substring(0, slash); namePart = namePart.Substring(slash + 1); }
                string baseDir = root;
                if (relDir.Length > 0)
                {
                    foreach (string seg in relDir.Split('\\', '/'))
                    {
                        if (seg.Length == 0 || seg == "..") { CloseAtPopup(); return; }
                        baseDir = Path.Combine(baseDir, seg);
                    }
                }
                if (!Directory.Exists(baseDir)) { CloseAtPopup(); return; }
                atRelDir = relDir;

                List<AtItem> items = new List<AtItem>();
                if (relDir.Length == 0)
                {
                    AtItem br = new AtItem();
                    br.Label = "浏览…（打开文件选择器）";
                    br.IsBrowse = true;
                    items.Add(br);
                }
                string lower = namePart.ToLowerInvariant();
                List<string> dirs = new List<string>();
                List<string> files = new List<string>();
                foreach (string d in Directory.GetDirectories(baseDir))
                {
                    if (lower.Length > 0 && !Path.GetFileName(d).ToLowerInvariant().Contains(lower)) continue;
                    dirs.Add(d);
                }
                foreach (string f in Directory.GetFiles(baseDir))
                {
                    if (lower.Length > 0 && !Path.GetFileName(f).ToLowerInvariant().Contains(lower)) continue;
                    files.Add(f);
                }
                dirs.Sort(delegate(string a, string b) { return string.Compare(a, b, true); });
                files.Sort(delegate(string a, string b) { return string.Compare(a, b, true); });
                foreach (string d in dirs)
                {
                    AtItem it = new AtItem(); it.Path = d; it.Label = Path.GetFileName(d) + "\\"; it.IsDir = true; items.Add(it);
                }
                foreach (string f in files)
                {
                    AtItem it = new AtItem(); it.Path = f; it.Label = Path.GetFileName(f); items.Add(it);
                }
                if (items.Count == 0) { CloseAtPopup(); return; }

                EnsureAtPopup();
                atList.BeginUpdate();
                atList.Items.Clear();
                foreach (AtItem it in items) atList.Items.Add(it);
                atList.EndUpdate();
                atList.SelectedIndex = 0;

                Point scr = inputBorder.PointToScreen(new Point(0, inputBorder.Height + 2));
                atPopup.Location = scr;
                atPopup.Width = Math.Max(320, inputBorder.Width);
                atPopup.Height = Math.Min(244, atList.Items.Count * 30 + 10);
                if (!atPopup.Visible) atPopup.Show(this); else atPopup.BringToFront();
                atActive = true;
                input.Focus();   // 保险：焦点保持在输入框
            }
            catch { CloseAtPopup(); }
        }

        // 不抢焦点的无边框弹层窗体（WinForms 无 ShowActivated，以 ShowWithoutActivation 实现）
        class NoActivateForm : Form
        {
            protected override bool ShowWithoutActivation { get { return true; } }
        }

        void EnsureAtPopup()
        {
            if (atPopup != null && !atPopup.IsDisposed) return;
            atPopup = new NoActivateForm();
            atPopup.FormBorderStyle = FormBorderStyle.None;
            atPopup.ShowInTaskbar = false;
            atPopup.StartPosition = FormStartPosition.Manual;
            atPopup.TopMost = true;
            atPopup.BackColor = Color.White;
            atList = new ListBox();
            atList.BorderStyle = BorderStyle.None;
            atList.Dock = DockStyle.Fill;
            atList.DrawMode = DrawMode.OwnerDrawFixed;
            atList.ItemHeight = 30;
            atList.Font = new Font("Microsoft YaHei UI", 9.5F);
            atList.DrawItem += AtList_DrawItem;
            atList.MouseClick += delegate(object s, MouseEventArgs e)
            {
                int idx = atList.IndexFromPoint(e.Location);
                if (idx >= 0) { atList.SelectedIndex = idx; InsertAtSelection(); }
            };
            atPopup.Controls.Add(atList);
        }

        void AtList_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || atList == null) return;
            AtItem it = atList.Items[e.Index] as AtItem;
            if (it == null) return;
            bool sel = (e.State & DrawItemState.Selected) != 0;
            using (SolidBrush b = new SolidBrush(sel ? Color.FromArgb(62, 99, 221) : Color.White))
                e.Graphics.FillRectangle(b, e.Bounds);
            Color fc = it.IsBrowse ? Color.FromArgb(120, 124, 136)
                : (it.IsDir ? Color.FromArgb(46, 76, 196) : Color.FromArgb(50, 54, 64));
            TextRenderer.DrawText(e.Graphics, it.Label, atList.Font,
                new Rectangle(e.Bounds.X + 10, e.Bounds.Y, e.Bounds.Width - 16, e.Bounds.Height),
                fc, TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        }

        void MoveAtSelection(int delta)
        {
            try
            {
                if (atList == null || atList.Items.Count == 0) return;
                int n = atList.SelectedIndex + delta;
                if (n < 0) n = atList.Items.Count - 1;
                if (n >= atList.Items.Count) n = 0;
                atList.SelectedIndex = n;
            }
            catch { }
        }

        // 选中条目：目录 → 进入子目录继续列；文件 → 插入路径并加入引用；浏览 → 文件选择器
        void InsertAtSelection()
        {
            try
            {
                AtItem it = (atList != null && atList.SelectedItem != null) ? atList.SelectedItem as AtItem : null;
                if (it == null) { CloseAtPopup(); return; }
                if (it.IsBrowse)
                {
                    CloseAtPopup();
                    using (OpenFileDialog dlg = new OpenFileDialog())
                    {
                        dlg.Title = "选择要引用的文件";
                        try { dlg.InitialDirectory = AgentTools.WorkspaceRoot; } catch { }
                        if (dlg.ShowDialog(FindForm()) == DialogResult.OK)
                        {
                            AddContextFile(dlg.FileName);
                            InsertAtPath(dlg.FileName);
                        }
                    }
                    return;
                }
                if (it.IsDir)
                {
                    atSuspend = true;
                    string rel = (atRelDir.Length > 0 ? atRelDir + "\\" : "") + it.Label.TrimEnd('\\');
                    string t = input.Text ?? "";
                    int end = Math.Min(atCaret, t.Length);
                    string newText = t.Substring(0, atTokenStart) + "@" + rel + "\\" + t.Substring(end);
                    int newCaret = atTokenStart + rel.Length + 2;
                    input.Text = newText;
                    input.SelectionStart = newCaret;
                    input.Focus();
                    atSuspend = false;
                    HandleAtCaret();   // 立即列出子目录内容
                    return;
                }
                CloseAtPopup();
                AddContextFile(it.Path);
                InsertAtPath(it.Path);
            }
            catch { CloseAtPopup(); }
        }

        // 把 token「@筛选词」替换为实际路径（含空格加引号），光标停在路径后
        void InsertAtPath(string path)
        {
            try
            {
                atSuspend = true;
                string quoted = path.Contains(" ") ? "\"" + path + "\"" : path;
                string t = input.Text ?? "";
                int end = Math.Min(atCaret, t.Length);
                if (atTokenStart < 0) atTokenStart = end;
                string newText = t.Substring(0, atTokenStart) + quoted + " " + t.Substring(end);
                input.Text = newText;
                input.SelectionStart = atTokenStart + quoted.Length + 1;
                input.Focus();
            }
            finally { atSuspend = false; }
        }

        void CloseAtPopup()
        {
            atActive = false;
            try { if (atPopup != null && atPopup.Visible) atPopup.Hide(); } catch { }
        }

        // 输入框失焦时收起；焦点若正落向弹层列表（用户要点击）则保留
        void TryCloseAtPopupOnBlur()
        {
            try
            {
                if (atPopup == null || !atPopup.Visible) { atActive = false; return; }
                if (atPopup.Bounds.Contains(Cursor.Position)) return;
                CloseAtPopup();
            }
            catch { CloseAtPopup(); }
        }

        // ===== 任务计划栏插件：模型经 task_plan 工具驱动 =====

        // 返回给模型的应答文本；同时刷新右侧计划栏
        string HandlePlanTool(string argsJson)
        {
            if (InvokeRequired)
            {
                string r = null;
                try { Invoke((MethodInvoker)delegate { r = HandlePlanTool(argsJson); }); } catch { }
                return r == null ? "计划更新失败" : r;
            }
            string action = "", data = "";
            try
            {
                List<Dictionary<string, string>> objs = MiniJson.ParseObjects(argsJson == null ? "{}" : argsJson);
                if (objs.Count > 0)
                {
                    action = MiniJson.Get(objs[0], "action");
                    data = MiniJson.Get(objs[0], "data");
                }
            }
            catch { }
            if (action == "start")
            {
                planSteps.Clear();
                foreach (string raw in (data ?? "").Split('\n', ';', '；'))
                {
                    string t = raw.Trim();
                    if (t.Length == 0) continue;
                    t = System.Text.RegularExpressions.Regex.Replace(t, "^[0-9]+[.、)．]\\s*", "");  // 去掉手写序号
                    if (t.Length > 0) planSteps.Add(new string[] { t, "0" });
                }
                if (planSteps.Count == 0) { planPanel.Visible = false; return "计划为空，请提供步骤列表"; }
                planPanel.Visible = true;
                planPanel.Width = 264;
            }
            else if (action == "done")
            {
                string key = (data ?? "").Trim();
                int hit = -1;
                for (int i = 0; i < planSteps.Count; i++)
                {
                    if (planSteps[i][1] == "1") continue;
                    if ((key.Length > 0 && planSteps[i][0].IndexOf(key, StringComparison.OrdinalIgnoreCase) >= 0) ||
                        (key == (i + 1).ToString())) { hit = i; break; }
                }
                if (hit >= 0) planSteps[hit][1] = "1";
                else return "未找到匹配的步骤: " + key;
            }
            else if (action == "finish")
            {
                for (int i = 0; i < planSteps.Count; i++) planSteps[i][1] = "1";
            }
            else return "未知 action: " + action;
            RefreshPlanPanel();
            int done = 0;
            foreach (string[] st in planSteps) { if (st[1] == "1") done++; }
            if (action == "start") return "计划已创建，共 " + planSteps.Count + " 步，已展示给用户。每完成一步记得用 done 勾选。";
            if (action == "finish") return "计划全部完成（" + done + "/" + planSteps.Count + "）。";
            return "已勾选（" + done + "/" + planSteps.Count + " 完成）";
        }

        void RefreshPlanPanel()
        {
            int done = 0;
            foreach (string[] st in planSteps) { if (st[1] == "1") done++; }
            planTitle.Text = "任务计划（" + done + "/" + planSteps.Count + "）";
            planStepsList.BeginUpdate();
            planStepsList.Items.Clear();
            foreach (string[] st in planSteps) planStepsList.Items.Add(st);
            planStepsList.EndUpdate();
        }

        void PlanStep_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            string[] st = planSteps[e.Index];
            using (SolidBrush bg = new SolidBrush(planPanel.BackColor))
                e.Graphics.FillRectangle(bg, e.Bounds);
            bool done = st[1] == "1";
            bool current = !done;
            for (int i = 0; i < e.Index; i++) { if (planSteps[i][1] != "1") { current = false; break; } }
            string mark = done ? "✓" : (current ? "▶" : "○");
            Color mc = done ? Color.FromArgb(34, 139, 84) : (current ? Color.FromArgb(62, 99, 221) : Color.FromArgb(170, 173, 184));
            Color tc = done ? Color.FromArgb(150, 153, 168) : Color.FromArgb(28, 30, 38);
            TextRenderer.DrawText(e.Graphics, mark, new Font("Microsoft YaHei UI", 10F, FontStyle.Bold),
                new Rectangle(e.Bounds.X + 10, e.Bounds.Y + 4, 24, 24), mc,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            TextRenderer.DrawText(e.Graphics, st[0], new Font("Microsoft YaHei UI", 9F),
                new Rectangle(e.Bounds.X + 38, e.Bounds.Y + 4, e.Bounds.Width - 48, e.Bounds.Height - 8),
                tc, TextFormatFlags.WordBreak | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        }

        // 技能提示段（规划层第一期/第二期）：把当前注册的技能概要写进系统提示。
        // 为什么不在提示里逐条列出全部 19 个动作：那会让每次请求都背上冗长的固定文本，
        // 而工具 schema 本身已经带了每个动作的完整描述（模型本来就看得到）。
        // 这里只需要让模型知道「有技能这一类工具、什么时候该用」，具体选哪个由 schema 决定。
        // 动态生成而非硬编码：用户装了新技能后提示自动跟上。
        string SkillHint()
        {
            try
            {
                string root = OfficeAgent.Core.EnvDetect.FindRoot();
                List<SkillActionSpec> all = SkillToolBridge.Collect(root);
                if (all == null || all.Count == 0) return "";
                // 按技能聚合，保持注册顺序
                List<string> ids = new List<string>();
                Dictionary<string, int> counts = new Dictionary<string, int>();
                foreach (SkillActionSpec sp in all)
                {
                    if (!counts.ContainsKey(sp.SkillId)) { counts[sp.SkillId] = 0; ids.Add(sp.SkillId); }
                    counts[sp.SkillId] = counts[sp.SkillId] + 1;
                }
                StringBuilder sb = new StringBuilder();
                sb.Append("另外你还有一组专门的办公/会计技能工具（工具名以 skill_ 开头），共 ")
                  .Append(ids.Count).Append(" 项技能、").Append(all.Count).Append(" 个动作：");
                for (int i = 0; i < ids.Count; i++)
                {
                    if (i > 0) sb.Append("、");
                    sb.Append(ids[i]).Append("(").Append(counts[ids[i]]).Append("个动作)");
                }
                sb.Append("。涉及会计计算（试算平衡、账龄、折旧、个税、增值税、报表、合并、凭证）、")
                  .Append("Excel 高级操作（看结构、加表、写单元格、写公式、冻结、调列宽）、")
                  .Append("生成 Word 报告时，**优先用这些技能工具**——它们比你自己拼 CSV 文本更准确、更专业。")
                  .Append("技能名称与参数说明见工具定义，按需直接调用即可。");
                sb.Append(BadActionHint(ids));
                return sb.ToString();
            }
            catch { return ""; }
        }

        // 历史表现差的动作提示（第三期 · 成功率台账）。
        // 为什么值得写进提示：实测里最贵的不是"算得慢"，而是"反复失败重试"——
        // 一次失败的技能调用要起一次 Python sidecar（~650ms），模型再想一轮、再调一次，
        // 等效跳预算就这么被烧掉。把"这个动作历史上总失败"直接告诉模型，
        // 比让它自己撞一次墙再换路便宜得多。
        // 只列 BadActionCount 个，且**措辞留余地**（环境问题修好后是会转好的）。
        const int BadActionCount = 3;
        string BadActionHint(List<string> skillIds)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                int shown = 0;
                foreach (string sid in skillIds)
                {
                    if (shown >= BadActionCount) break;
                    foreach (string[] bad in SkillStats.BadActions(sid, BadActionCount - shown))
                    {
                        if (shown == 0)
                            sb.Append("提醒：以下动作在你之前的使用中经常失败（可能是当前环境缺依赖，也可能参数不合适），")
                              .Append("调用它们前先确认参数与文件；若再次失败就换别的做法，不要重复重试：");
                        sb.Append(" ").Append(sid).Append("/").Append(bad[0])
                          .Append("（成功 ").Append(bad[1]).Append(" 次、失败 ").Append(bad[2]).Append(" 次）");
                        if (shown < BadActionCount - 1) sb.Append("；");
                        shown++;
                    }
                }
                if (shown > 0) sb.Append("。");
                return sb.ToString();
            }
            catch { return ""; }
        }

        // 目标继承（规划层第三期）：把**上一轮的意图**压缩成一句话带进本轮。
        //
        // 问题：多轮会话里用户会说"再来一次，换成上个月的"、"第二个文件也要"、
        //      "把它转成 PDF"。这些话本身**不含任何可路由的线索**——IntentRouter 抓不到
        //      场景词、路径也没有，于是模型只能看到一句孤零零的指代，
        //      经常答"请问您指的是哪个文件"。
        //      而 llmHistory 里虽然有上一轮，但历史只存**最终答复文本**，
        //      任务的结构化信息（用了哪些技能、产出了什么文件）已经丢了。
        //
        // 方案：记住上一轮真实用过的技能与产物，在本轮拼成简短上下文。
        //      ★ 只带"上一轮"、只在有实质内容时带，且**明确标注这是上一轮的事**——
        //        否则模型会把历史目标当成当前指令，在用户开了个全新话题时答非所问。
        //
        // 不做的事：不做"指代消解"（不把"它"替换成具体路径）。
        //      那属于"解释用户指令"，与 ArtifactRegistry 不做别名还原是同一条红线：
        //      我们只提供事实（上一轮做了什么、产出了什么），判断与替换交给模型。
        string GoalInheritance()
        {
            try
            {
                // 文本拼装放在 TurnMemory（可直测）；这里只负责取当前会话的上一轮快照
                return TurnMemory.Build(lastTurnSkills, lastTurnProducts, 4);
            }
            catch { return ""; }
        }

        // 上一轮真实用过的技能（skill_tool 审计的轻量内存副本）与产物
        readonly List<string> lastTurnSkills = new List<string>();
        readonly List<string> lastTurnProducts = new List<string>();

        // 回合结束后记录"这轮到底干了什么"，供下一轮的目标继承使用。
        // 数据源是 AgentLoop.Result（技能从 ToolLog 解析、产物是 r.Products），
        // 而不是让模型自我汇报——实测模型会在总结里给出与产物不符的读数，
        // 拿它的叙述当"上一轮事实"会把错误继承下去。
        void RememberTurn(AgentLoop.Result r)
        {
            try
            {
                lastTurnSkills.Clear();
                lastTurnProducts.Clear();
                if (r == null) return;
                // 解析逻辑放在 TurnMemory（可被 /plantest 直测，不必启动 WinForms 消息循环）
                lastTurnSkills.AddRange(TurnMemory.SkillsFromLog(r.ToolLog));
                lastTurnProducts.AddRange(TurnMemory.Products(r.Products));
            }
            catch { }
        }

        string BuildSystemPrompt()
        {
            // 记忆段放在最前（优先进入模型上下文；同时便于 E2E 在请求体前段断言）
            string mem = MemoryStore.ForPrompt(config.PluginMemory);
            string ws = "";
            try
            {
                string w = WorkspaceStore.ActiveDir(config);
                if (w != null && w.Length > 0) ws = "用户当前的工作区目录是 " + w +
                    "（用户说的文件默认先在这里找；用 list_directory 列出后确认，不要凭空猜文件名）。";
            }
            catch { }
            string basePrompt = "你是 OfficeAgent，一款运行在 Windows 7 上的办公 AI 助手（agent），面向会计与办公人员。" +
                "你具备一组可以直接调用执行的工具：读取本地文本文件和 PDF、列出目录、联网下载文件、文档格式转换（pdf/csv/xlsx）、" +
                "创建 Excel 表格、创建 PPT 演示文稿、弹出环境修复器安装系统组件、维护面向用户的任务计划清单。" +
                SkillHint() +
                GoalInheritance() +
                "规则：凡是需要上述能力的请求，一律直接发起工具调用去完成；绝不输出代码示例、调用语法或操作步骤说明，" +
                "也不让用户自己去处理。缺信息时先自己用工具查证（列目录/读文件），查不到再向用户追问，一次只问最关键的一项。" +
                ws +
                // 效率策略：明确"先定位→再一次做对"，抑制重复列目录/反复试探
                "执行效率要求（重要）：" +
                "① 先用 list_directory 定位目标一次，记住结果，不要为了确认同一件事反复列同一个目录；" +
                "② 目标文件确认后，直接一次调用就把动作做完（读就读、转就转、建就建），不要分成多次试探性调用；" +
                "③ 同一个工具用完全相同的参数不要调用第二次——那不会得到新信息，只会浪费时间；" +
                "④ 参数不确定时用一次工具查证后再动手，不要靠连续试错碰运气；" +
                "⑤ 信息足够时立即给出最终答复，不要为了「再确认一下」多跑工具。" +
                "三步以上的多步骤任务，先创建任务计划，随着执行逐项更新，全部完成后再标记完成。" +
                // 多步串联（规划层第二期）：上一版只说了"要建计划"，没说"怎么把多步串起来"。
                // 实测模型会把每步当成独立任务，做完一步就停下来问用户，而不是自己接着做下一步。
                "多步骤任务要**连续执行到底**：不要每完成一步就停下来问我，应该根据上一步的结果直接进行下一步，" +
                "直到整个任务完成或确实缺关键信息时才回复。" +
                "上一步产出的文件路径会随工具结果返回（并标注为「产物N」），后续步骤直接使用该完整路径，" +
                "不要凭猜测拼路径，也不要重新读取同一份文件来确认它是否生成。" +
                "如果某个工具调用失败，先看错误信息里的提示：若有推荐的替代工具就换一个重试，" +
                "若没有就直接说明失败原因，不要用完全相同的参数反复重试。" +
                "系统还内置：文件预览（xlsx/csv/pdf）、批量转换队列、两表核对、报表汇总、发票提取、长期记忆（记住/记忆/忘记）。" +
                "用简体中文回答；回答使用纯文本，不要用 markdown 语法（不要 ** 星号加粗、# 标题、表格线）；" +
                "回答时明确说出产物保存在哪个路径，方便用户直接打开；" +
                "不要编造用户未提供的文件内容；涉及金额计算时提醒用户以软件的核对功能结果为准。";
            return mem.Length > 0 ? mem + "\n" + basePrompt : basePrompt;
        }

        void FinishReply(int idx, string text, bool isError, long elapsedMs)
        {
            if (InvokeRequired) { Invoke((MethodInvoker)delegate { FinishReply(idx, text, isError, elapsedMs); }); return; }
            busy = false;
            send.Enabled = true;
            if (idx >= 0 && idx < msgs.Count)
            {
                msgs[idx].Text = text;
                msgs[idx].Pending = false;
                msgs[idx].Error = isError;
                msgs[idx].CachedTextH = -1;
                msgs[idx].CachedHeight = -1;
                if (!isError && elapsedMs > 0)
                    msgs[idx].Time = msgs[idx].Time + " · " + (elapsedMs / 1000.0).ToString("0.0") + "s";
                if (config.PluginHist && !isError && text != null && !text.StartsWith("✗"))
                    MemoryStore.AppendHistory("assistant", text);
                if (!isError) { SplitLongFinal(idx); SaveSession(); }
            }
            MarkDirty();
        }

        // 长回复分块：ListBox 单项高度被系统钳制在 255px，超高文本会被裁掉 → 拆成连续多条气泡
        void SplitLongFinal(int idx)
        {
            if (idx < 0 || idx >= msgs.Count) return;
            ChatMsg head = msgs[idx];
            List<string> parts = SplitText(head.Text, AvailWidth());
            if (parts.Count <= 1)
            {
                head.First = true; head.Last = true;
                return;
            }
            head.Text = parts[0];
            head.First = true; head.Last = false;
            head.CachedTextH = -1; head.CachedHeight = -1;
            int insertAt = idx + 1;
            for (int i = 1; i < parts.Count; i++)
            {
                ChatMsg c = new ChatMsg();
                c.User = head.User;
                c.Error = head.Error;
                c.Text = parts[i];
                c.Time = "";
                c.First = false;
                c.Last = i == parts.Count - 1;
                msgs.Insert(insertAt, c);
                insertAt++;
            }
            ReMeasure();
        }

        // 把超长文本按测量高度切成 ≤ MaxChunkTextH 的段（优先在换行/标点处断开）
        static List<string> SplitText(string text, int avail)
        {
            List<string> parts = new List<string>();
            if (text == null || text.Length == 0) { parts.Add(""); return parts; }
            Size ts = TextRenderer.MeasureText(text, TextFont, new Size(avail, 100000), TextFormatFlags.WordBreak);
            if (ts.Height <= MaxChunkTextH) { parts.Add(text); return parts; }
            int lineH = TextRenderer.MeasureText("测Mg", TextFont).Height;
            if (lineH < 8) lineH = 17;
            int linesTotal = Math.Max(1, ts.Height / lineH);
            int charsPerLine = Math.Max(10, text.Length / linesTotal);
            int chunkLines = Math.Max(4, MaxChunkTextH / lineH - 1);
            int target = chunkLines * charsPerLine;
            string rest = text;
            int guard = 0;
            while (guard++ < 200)
            {
                int rh = TextRenderer.MeasureText(rest, TextFont, new Size(avail, 100000), TextFormatFlags.WordBreak).Height;
                if (rh <= MaxChunkTextH || rest.Length < 40) { parts.Add(rest); break; }
                int cut = Math.Min(target, rest.Length - 1);
                int brk = -1;
                int floor = cut / 2;
                for (int i = cut; i > floor; i--)
                {
                    char c = rest[i];
                    if (c == '\n' || c == ' ' || c == '。' || c == '，' || c == '；' || c == '、' || c == '.' || c == ',')
                    {
                        brk = i + 1;
                        break;
                    }
                }
                if (brk <= 0) brk = cut;
                string part = rest.Substring(0, brk);
                int ph = TextRenderer.MeasureText(part, TextFont, new Size(avail, 100000), TextFormatFlags.WordBreak).Height;
                while (ph > MaxChunkTextH && brk > 40)
                {
                    brk = brk * 3 / 4;
                    part = rest.Substring(0, brk);
                    ph = TextRenderer.MeasureText(part, TextFont, new Size(avail, 100000), TextFormatFlags.WordBreak).Height;
                }
                parts.Add(part.TrimEnd('\r', '\n'));
                rest = rest.Substring(brk).TrimStart('\n');
            }
            return parts;
        }

        int Append(string role, string text)
        {
            return Append(role, text, false, false);
        }

        int Append(string role, string text, bool pending)
        {
            return Append(role, text, pending, false);
        }

        int Append(string role, string text, bool pending, bool error)
        {
            ChatMsg m = new ChatMsg();
            m.User = role == "user";
            m.Text = text;
            m.Pending = pending;
            m.Error = error;
            m.Time = DateTime.Now.ToString("HH:mm");
            msgs.Add(m);
            list.Items.Add(m);
            // 超长消息（如用户粘贴大段文本）同样分块，避免 255px 裁剪
            if (!pending && text != null && text.Length > 0)
            {
                MeasureMsg(m, AvailWidth());
                if (m.CachedTextH > MaxChunkTextH) SplitLongFinal(msgs.Count - 1);
            }
            MarkDirty();
            return msgs.Count - 1;
        }

        // 测量单条消息（带缓存）；宽高变更后由 ReMeasure 清缓存
        static void MeasureMsg(ChatMsg m, int avail)
        {
            if (m.CachedHeight >= 0 && m.CachedAvail == avail) return;
            if (m.AttachPath.Length > 0)
            {
                string nm = Path.GetFileName(m.AttachPath);
                Size ns = TextRenderer.MeasureText(nm, TextFont);
                m.CachedTextH = 46;
                m.CachedTextW = Math.Max(220, Math.Min(avail, ns.Width + 130));
                string attachHeader = (m.User ? "你" : "助手") + "  " + m.Time;
                m.CachedHeaderW = TextRenderer.MeasureText(attachHeader, NameFont).Width;
                m.CachedHeight = 34 + m.CachedTextH + 14;
                m.CachedAvail = avail;
                return;
            }
            Size ts = TextRenderer.MeasureText(m.Text == null ? "" : m.Text, TextFont,
                new Size(avail, 100000), TextFormatFlags.WordBreak);
            m.CachedTextH = ts.Height;
            // +36：左右内边距与边框余量（修复长英文单词/标点贴边破出气泡）
            m.CachedTextW = Math.Max(60, Math.Min(avail, ts.Width + 36));
            string header = (m.User ? "你" : "助手") + "  " + m.Time;
            m.CachedHeaderW = TextRenderer.MeasureText(header, NameFont).Width;
            if (m.First && m.Last) m.CachedHeight = 34 + m.CachedTextH + 16;
            else if (m.First) m.CachedHeight = 34 + m.CachedTextH + 4;
            else if (m.Last) m.CachedHeight = m.CachedTextH + 14;
            else m.CachedHeight = m.CachedTextH + 4;
            m.CachedAvail = avail;
        }

        void List_MeasureItem(object sender, MeasureItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= msgs.Count) { e.ItemHeight = 30; return; }
            ChatMsg m = msgs[e.Index];
            MeasureMsg(m, AvailWidth());
            e.ItemHeight = m.CachedHeight;
        }

        void List_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= msgs.Count) return;
            ChatMsg m = msgs[e.Index];
            // 穿模修复：不用 e.DrawBackground()（选中态会把系统蓝画满整行），
            // 一律白底铺满行区域 + 关闭列表选中模式
            Graphics g = e.Graphics;
            using (SolidBrush bg = new SolidBrush(BackColor))
                g.FillRectangle(bg, e.Bounds);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            MeasureMsg(m, AvailWidth());
            int left = LeftMargin();
            int textH = m.CachedTextH;
            int bubbleW = m.CachedTextW;
            int top = e.Bounds.Y + (m.First ? 22 : 2);
            Rectangle bubble;

            if (m.AttachPath.Length > 0)
            {
                // 附件卡片：类型徽标 + 文件名 + 大小；点击进预览
                string attachHeader = "助手  " + m.Time;
                TextRenderer.DrawText(g, attachHeader, NameFont, new Point(left + 2, e.Bounds.Y + 4), Color.FromArgb(150, 153, 168));
                Rectangle card = new Rectangle(left, top, bubbleW, 46);
                using (GraphicsPath gp = RoundRect(card, 8))
                {
                    using (SolidBrush b = new SolidBrush(Color.White)) g.FillPath(b, gp);
                    using (Pen p2 = new Pen(Color.FromArgb(222, 225, 233)))
                        g.DrawPath(p2, gp);
                }
                string ext = (Path.GetExtension(m.AttachPath) ?? "").Trim('.').ToUpperInvariant();
                Color ic = ext.StartsWith("XLS") ? Color.FromArgb(34, 139, 84)
                    : (ext.StartsWith("PPT") ? Color.FromArgb(234, 121, 12)
                    : (ext == "PDF" ? Color.FromArgb(220, 60, 50) : Color.FromArgb(59, 110, 220)));
                Rectangle icon = new Rectangle(left + 10, top + 5, 36, 36);
                using (GraphicsPath gp = RoundRect(icon, 6))
                using (SolidBrush b = new SolidBrush(ic)) g.FillPath(b, gp);
                string tag = ext.Length == 0 ? "FILE" : (ext.Length > 3 ? ext.Substring(0, 3) : ext);
                TextRenderer.DrawText(g, tag, new Font("Microsoft YaHei UI", 8F, FontStyle.Bold), icon, Color.White,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                string nm = Path.GetFileName(m.AttachPath);
                string sz = "";
                try { long b2 = new FileInfo(m.AttachPath).Length; sz = b2 >= 1024 ? (b2 / 1024.0).ToString("0.#") + " KB" : b2 + " B"; } catch { }
                TextRenderer.DrawText(g, nm, new Font("Microsoft YaHei UI", 9F, FontStyle.Bold),
                    new Rectangle(left + 56, top + 4, bubbleW - 68, 20), Color.FromArgb(28, 30, 38),
                    TextFormatFlags.EndEllipsis);
                TextRenderer.DrawText(g, sz + "  ·  点击预览", new Font("Microsoft YaHei UI", 8F),
                    new Rectangle(left + 56, top + 24, bubbleW - 68, 18), Color.FromArgb(150, 153, 168),
                    TextFormatFlags.EndEllipsis);
                return;
            }
            if (m.User)
            {
                if (m.First)
                {
                    string header = "你  " + m.Time;
                    TextRenderer.DrawText(g, header, NameFont, new Point(e.Bounds.Right - 20 - m.CachedHeaderW, e.Bounds.Y + 4), Color.FromArgb(150, 153, 168));
                }
                bubble = new Rectangle(e.Bounds.Right - 20 - bubbleW, top, bubbleW, textH + 20);
                using (GraphicsPath gp = RoundRectEx(bubble, CornerRad(m, true)))
                {
                    using (SolidBrush b = new SolidBrush(Color.FromArgb(213, 228, 255)))     // 用户：实心浅蓝（与助手白底拉开）
                        g.FillPath(b, gp);
                    using (Pen p = new Pen(Color.FromArgb(150, 175, 225)))
                        g.DrawPath(p, gp);
                }
                TextRenderer.DrawText(g, m.Text, TextFont, bubble, Color.FromArgb(20, 40, 90),
                    TextFormatFlags.WordBreak | TextFormatFlags.VerticalCenter | TextFormatFlags.LeftAndRightPadding);
            }
            else
            {
                if (m.First)
                {
                    string header = "助手  " + m.Time;
                    TextRenderer.DrawText(g, header, NameFont, new Point(left + 2, e.Bounds.Y + 4), Color.FromArgb(150, 153, 168));
                }
                bubble = new Rectangle(left, top, bubbleW, textH + 20);
                using (GraphicsPath gp = RoundRectEx(bubble, CornerRad(m, false)))
                {
                    using (SolidBrush b = new SolidBrush(m.Error ? Color.FromArgb(253, 232, 232) : Color.White))
                        g.FillPath(b, gp);
                    using (Pen p = new Pen(m.Error ? Color.FromArgb(230, 160, 160) : Color.FromArgb(222, 225, 233)))
                        g.DrawPath(p, gp);
                }
                TextRenderer.DrawText(g, m.Text, TextFont, bubble,
                    m.Error ? Color.FromArgb(178, 58, 58) : Color.FromArgb(28, 30, 38),
                    TextFormatFlags.WordBreak | TextFormatFlags.VerticalCenter | TextFormatFlags.LeftAndRightPadding);
            }
        }

        // 分块圆角：首条上圆下小、尾条上小下圆、中间全小
        static int[] CornerRad(ChatMsg m, bool user)
        {
            int big = 10, small = 3;
            if (m.First && m.Last) return new int[] { big, big, big, big };
            if (m.First) return new int[] { big, big, small, small };
            if (m.Last) return new int[] { small, small, big, big };
            return new int[] { small, small, small, small };
        }

        // radii = {左上, 右上, 右下, 左下}
        static GraphicsPath RoundRectEx(Rectangle r, int[] rad)
        {
            GraphicsPath gp = new GraphicsPath();
            int tl = rad[0], tr = rad[1], br = rad[2], bl = rad[3];
            gp.AddArc(r.X, r.Y, tl * 2, tl * 2, 180, 90);
            gp.AddArc(r.Right - tr * 2, r.Y, tr * 2, tr * 2, 270, 90);
            gp.AddArc(r.Right - br * 2, r.Bottom - br * 2, br * 2, br * 2, 0, 90);
            gp.AddArc(r.X, r.Bottom - bl * 2, bl * 2, bl * 2, 90, 90);
            gp.CloseFigure();
            return gp;
        }

        static GraphicsPath RoundRect(Rectangle r, int rad)
        {
            return RoundRectEx(r, new int[] { rad, rad, rad, rad });
        }

        Panel bottom;

        void InputBorder_Paint(object sender, PaintEventArgs e)
        {
            using (GraphicsPath gp = RoundRect(new Rectangle(0, 0, inputBorder.Width - 1, inputBorder.Height - 1), 12))
            using (Pen p = new Pen(Color.FromArgb(222, 224, 230)))
                e.Graphics.DrawPath(p, gp);
        }

        void Bottom_Paint(object sender, PaintEventArgs e)
        {
            using (Pen p = new Pen(Color.FromArgb(238, 239, 243)))
                e.Graphics.DrawLine(p, 0, 0, bottom.Width, 0);
        }

        class BufferedMsgList : ListBox
        {
            public BufferedMsgList()
            {
                SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);
                ItemHeight = 60;
                IntegralHeight = false;
                SelectionMode = SelectionMode.None;   // 自绘气泡无选中概念；禁用后系统蓝条不再穿模
            }
        }

        // 输入框持有焦点时把滚轮转发给消息列表（Win7 滚轮只作用于焦点控件，不转发就永远滚不动历史）
        class ForwardWheelBox : TextBox
        {
            public Control Target;
            protected override void WndProc(ref Message m)
            {
                if (m.Msg == WM_MOUSEWHEEL && Target != null && !Target.IsDisposed && Target.IsHandleCreated)
                {
                    try { SendMessage(Target.Handle, WM_MOUSEWHEEL, m.WParam, m.LParam); } catch { }
                    return;
                }
                base.WndProc(ref m);
            }
        }
    }
}
