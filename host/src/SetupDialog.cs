// SetupDialog —— 首次启动/设置向导：填 API 地址与密钥 → 自动获取模型列表 → 测试连通
// 密钥经 DPAPI 加密保存（AppConfig）；出网经 HostGuard 守卫。
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Threading;
using System.Windows.Forms;

namespace OfficeAgent.Host
{
    public class SetupDialog : Form
    {
        TextBox txtUrl, txtKey, txtWorkspace;
        ComboBox cmbModel;
        // 密钥框占位符：已存密钥时显示，表示"不修改"。用户一旦输入真实内容即视为要替换。
        const string KeyPlaceholder = "●●●●●●●●（已保存，留空则沿用；要更换请直接输入新密钥）";
        bool hasStoredKey = false;
        Button btnFetch, btnTest;
        Label lblResult;
        CheckBox chkLan;
        CheckBox chkWeb;
        Button btnSave, btnSkip;
        AppConfig config;
        bool saved = false;
        bool busy = false;

        public bool Saved { get { return saved; } }

        public SetupDialog(AppConfig cfg, bool firstRun)
        {
            config = cfg;
            Text = firstRun ? "欢迎使用 OfficeAgent" : "模型设置";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(560, 496);
            Font = new Font("Microsoft YaHei UI", 9F);
            BackColor = Color.White;

            Label title = new Label();
            title.Text = firstRun ? "欢迎使用 OfficeAgent" : "模型设置";
            title.Font = new Font("Microsoft YaHei UI", 15F, FontStyle.Bold);
            title.Location = new Point(28, 22);
            title.AutoSize = true;
            Controls.Add(title);

            Label sub = new Label();
            sub.Text = firstRun
                ? "三步开始：① 填 OpenAI 兼容服务地址 ② 填 API 密钥 ③ 获取模型并测试。\n密钥只保存在本机（DPAPI 加密），不会随文件外发。也可以先跳过，离线使用预览与转换。"
                : "修改模型服务地址、密钥或选择的模型。";
            sub.ForeColor = Color.FromArgb(110, 113, 126);
            sub.Location = new Point(28, 62);
            sub.Size = new Size(508, 44);
            Controls.Add(sub);

            AddLabel("服务地址（填到 /v1 或 /v4 即可；粘成完整 /chat/completions 也会自动识别）", 118);
            txtUrl = new TextBox();
            txtUrl.Location = new Point(28, 140);
            txtUrl.Size = new Size(504, 26);
            txtUrl.Text = cfg.BaseUrl == null ? "" : cfg.BaseUrl;
            Controls.Add(txtUrl);

            AddLabel("API 密钥（每家服务分开保存，换服务商不用重填；只存本机，DPAPI 加密）", 176);
            txtKey = new TextBox();
            txtKey.Location = new Point(28, 198);
            txtKey.Size = new Size(504, 26);
            txtKey.PasswordChar = '●';
            // 已存密钥时预填占位符：让用户看到"已经存过"，不填即为不修改。
            // 此前该框恒为空 → 用户每次都要重填，且以为是没保存（旧行为还会让空值覆盖已存密钥）。
            // 密钥按服务地址（host）归属：地址栏指向哪家就显示哪家的存储状态（见 RefreshKeyState）。
            hasStoredKey = cfg.HasKeyForHost(AppConfig.HostOf(cfg.BaseUrl));
            // 先挂 TextChanged 会被下面的占位赋值误判成"用户输入过"，故占位赋值放在事件挂接之前
            if (hasStoredKey) { txtKey.Text = KeyPlaceholder; txtKey.ForeColor = Color.FromArgb(120, 120, 130); }
            txtKey.GotFocus += delegate
            {
                if (hasStoredKey && txtKey.Text == KeyPlaceholder)
                {
                    txtKey.Text = "";                    // 首次进入即清空占位，直接输入新密钥
                    txtKey.ForeColor = Color.Black;
                }
            };
            txtKey.TextChanged += delegate
            {
                if (txtKey.Text == KeyPlaceholder) return;   // 占位本身不算用户输入
                if (txtKey.ForeColor != Color.Black) txtKey.ForeColor = Color.Black;
            };
            Controls.Add(txtKey);
            // 服务地址改动 → 密钥框切换到对应服务的存储状态（修"一家 key 到处用"）
            txtUrl.TextChanged += delegate { RefreshKeyState(); };

            btnFetch = AddBtn("获取模型列表", 28, 236, 120, delegate { FetchModels(); });
            AddLabel("模型", 246 - 2, 168);
            cmbModel = new ComboBox();
            cmbModel.Location = new Point(168, 236);
            cmbModel.Size = new Size(364, 26);
            cmbModel.Text = cfg.Model == null ? "" : cfg.Model;
            // 预置常用模型（/models 拉取失败时直接选；拉取成功会被覆盖）
            foreach (string pm in LlmClient.ModelPresets) if (!cmbModel.Items.Contains(pm)) cmbModel.Items.Add(pm);
            Controls.Add(cmbModel);

            btnTest = AddBtn("测试连接", 28, 274, 120, delegate { TestConn(); });
            lblResult = new Label();
            lblResult.Location = new Point(168, 280);
            lblResult.Size = new Size(364, 40);
            lblResult.ForeColor = Color.FromArgb(110, 113, 126);
            lblResult.Text = "";
            Controls.Add(lblResult);

            // 布局约束：chkLan 必须保持在 y=326（E2E 按坐标勾选它），chkWeb 在其后
            chkLan = new CheckBox();
            chkLan.Text = "允许内网/本机端点（仅企业自建网关时勾选；默认拒绝私网地址）";
            chkLan.Location = new Point(28, 326);
            chkLan.AutoSize = true;
            chkLan.Checked = cfg.AllowLan;
            Controls.Add(chkLan);

            chkWeb = new CheckBox();
            chkWeb.Text = "联网搜索（模型联网查资料/下载；智谱系端点生效）";
            chkWeb.Location = new Point(28, 352);
            chkWeb.AutoSize = true;
            chkWeb.Checked = cfg.WebSearch;
            Controls.Add(chkWeb);

            // 工作区目录：全局默认值；侧边栏可为不同项目分别建工作区（各挂一个文件夹）
            AddLabel("默认工作区目录（新建表格/PPT/下载默认保存在这里；不同项目可在侧边栏右键分别建工作区）", 382);
            txtWorkspace = new TextBox();
            txtWorkspace.Location = new Point(28, 402);
            txtWorkspace.Size = new Size(444, 26);
            try { txtWorkspace.Text = cfg.EffectiveWorkspace(); } catch { txtWorkspace.Text = ""; }
            Controls.Add(txtWorkspace);
            Button btnBrowseWs = AddBtn("浏览…", 478, 401, 54, delegate
            {
                using (FolderBrowserDialog fb = new FolderBrowserDialog())
                {
                    fb.Description = "选择工作区目录";
                    try { fb.SelectedPath = txtWorkspace.Text; } catch { }
                    if (fb.ShowDialog(this) == DialogResult.OK) txtWorkspace.Text = fb.SelectedPath;
                }
            });

            btnSkip = AddBtn(firstRun ? "暂不配置，离线使用" : "取消", 28, 440, 170, delegate { Close(); });
            btnSave = AddBtn("保存并开始使用", 362, 440, 170, delegate { Save(); });
            if (btnSave != null) { btnSave.BackColor = Color.FromArgb(24, 26, 32); btnSave.ForeColor = Color.White; btnSave.FlatStyle = FlatStyle.Flat; btnSave.FlatAppearance.BorderSize = 0; }
        }

        void AddLabel(string text, int y)
        {
            Label l = new Label();
            l.Text = text;
            l.ForeColor = Color.FromArgb(90, 93, 105);
            l.Location = new Point(28, y);
            l.AutoSize = true;
            Controls.Add(l);
        }

        void AddLabel(string text, int y, int x)
        {
            Label l = new Label();
            l.Text = text;
            l.ForeColor = Color.FromArgb(90, 93, 105);
            l.Location = new Point(x, y);
            l.AutoSize = true;
            Controls.Add(l);
        }

        Button AddBtn(string text, int x, int y, int w, EventHandler onClick)
        {
            Button b = new Button();
            b.Text = text;
            b.Location = new Point(x, y);
            b.Size = new Size(w, 32);
            b.Cursor = Cursors.Hand;
            b.Click += onClick;
            Controls.Add(b);
            return b;
        }

        // 取用户真正输入的密钥；占位符/留空都表示"沿用已存密钥"（返回 null）
        string EnteredKey()
        {
            string k = txtKey.Text == null ? "" : txtKey.Text.Trim();
            if (k.Length == 0 || k == KeyPlaceholder) return null;
            return k;
        }

        // 服务地址变化 → 密钥框跟随该地址（host）的存储状态：
        // 已存 → 占位符「留空沿用」；未存 → 空框等待输入。只动占位，不碰用户已输入的真实内容。
        void RefreshKeyState()
        {
            bool stored = config != null && config.HasKeyForHost(AppConfig.HostOf(txtUrl.Text));
            hasStoredKey = stored;
            if (stored)
            {
                if (txtKey.Text.Length == 0)
                {
                    txtKey.Text = KeyPlaceholder;
                    txtKey.ForeColor = Color.FromArgb(120, 120, 130);
                }
            }
            else if (txtKey.Text == KeyPlaceholder)
            {
                txtKey.Text = "";
                txtKey.ForeColor = Color.Black;
            }
        }

        LlmClient MakeClient()
        {
            AppConfig tmp = new AppConfig();
            tmp.BaseUrl = txtUrl.Text.Trim();
            tmp.Model = cmbModel.Text.Trim();
            tmp.AllowLan = chkLan.Checked;
            tmp.WebSearch = chkWeb.Checked;
            string key = EnteredKey();
            if (key == null) key = config.KeyForHost(AppConfig.HostOf(txtUrl.Text.Trim()));   // 未重填则沿用该服务已存密钥
            if (key != null) tmp.SetKey(key);
            return new LlmClient(tmp);
        }

        void Guarded(Action work)
        {
            if (busy) return;
            string guard = HostGuard.Check(txtUrl.Text.Trim(), chkLan.Checked);
            if (guard != null) { lblResult.ForeColor = Color.FromArgb(178, 58, 58); lblResult.Text = guard; return; }
            busy = true;
            btnFetch.Enabled = false; btnTest.Enabled = false; btnSave.Enabled = false;
            lblResult.ForeColor = Color.FromArgb(110, 113, 126);
            lblResult.Text = "请求中…";
            Thread t = new Thread(new ThreadStart(delegate
            {
                work();
                if (InvokeRequired) Invoke((MethodInvoker)delegate { busy = false; btnFetch.Enabled = true; btnTest.Enabled = true; btnSave.Enabled = true; });
                else { busy = false; btnFetch.Enabled = true; btnTest.Enabled = true; btnSave.Enabled = true; }
            }));
            t.IsBackground = true;
            t.Start();
        }

        void SetResult(string text, bool ok)
        {
            if (InvokeRequired) { Invoke((MethodInvoker)delegate { SetResult(text, ok); }); return; }
            lblResult.ForeColor = ok ? Color.FromArgb(22, 130, 93) : Color.FromArgb(178, 58, 58);
            lblResult.Text = text;
        }

        void FetchModels()
        {
            Guarded(delegate
            {
                LlmClient c = MakeClient();
                string err;
                System.Collections.Generic.List<string> models = c.ListModels(out err);
                if (models == null) { SetResult("✗ " + err, false); return; }
                if (InvokeRequired)
                {
                    Invoke((MethodInvoker)delegate
                    {
                        cmbModel.Items.Clear();
                        foreach (string m in models) cmbModel.Items.Add(m);
                        if (cmbModel.Items.Contains(config.Model)) cmbModel.SelectedItem = config.Model;
                        else if (cmbModel.Items.Count > 0) cmbModel.SelectedIndex = 0;
                    });
                }
                SetResult("✓ 获取到 " + models.Count + " 个模型，请选择后点「测试连接」", true);
            });
        }

        void TestConn()
        {
            Guarded(delegate
            {
                LlmClient c = MakeClient();
                string err;
                System.Diagnostics.Stopwatch sw = System.Diagnostics.Stopwatch.StartNew();
                string reply = c.Chat("ping", "You are a connectivity test. Reply with exactly: pong", out err);
                sw.Stop();
                if (reply == null) { SetResult("✗ " + err, false); return; }
                SetResult("✓ 模型可用：" + c.Model + "，回复 " + reply.Trim().Length + " 字符，耗时 " + (sw.ElapsedMilliseconds / 1000.0).ToString("0.0") + "s", true);
            });
        }

        void Save()
        {
            string url = txtUrl.Text.Trim();
            string key = EnteredKey();
            // 地址与密钥都空 = 离线模式。只清当前指向，不删任何一家已存密钥
            //（下次把地址填回来，该服务的密钥自动生效——此前空值覆盖导致 config 被清空的教训）。
            // 各家已存密钥原样保留。
            if (url.Length == 0 && key == null)
            {
                config.BaseUrl = ""; config.Model = ""; config.WizardDone = true;
                config.Save();
                saved = true;
                Close();
                return;
            }
            if (url.Length == 0) { lblResult.ForeColor = Color.FromArgb(178, 58, 58); lblResult.Text = "服务地址为空；若要离线使用请清空密钥后保存"; return; }
            string guard = HostGuard.Check(url, chkLan.Checked);
            if (guard != null) { lblResult.ForeColor = Color.FromArgb(178, 58, 58); lblResult.Text = guard; return; }
            if (cmbModel.Text.Trim().Length == 0) { lblResult.ForeColor = Color.FromArgb(178, 58, 58); lblResult.Text = "请先获取并选择模型"; return; }
            config.BaseUrl = url;
            config.Model = cmbModel.Text.Trim();
            config.AllowLan = chkLan.Checked;
            config.WebSearch = chkWeb.Checked;
            config.WorkspaceDir = txtWorkspace.Text.Trim();   // 工作区目录随设置一起保存
            config.WizardDone = true;
            if (key != null) config.SetKey(key);   // 只有真正输入了新密钥才覆盖
            config.Save();
            saved = true;
            Close();
        }
    }
}
