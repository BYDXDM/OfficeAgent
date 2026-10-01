// 引导器主界面：检测清单 + 一键补全 + 日志
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using OfficeAgent.Core;

namespace OfficeAgent.Boot
{
    public class BootForm : Form
    {
        List<DetectItem> items;
        ListView lv;
        TextBox txtLog;
        Button btnDetect, btnFix, btnReport, btnLaunch;
        Label lblStatus;
        delegate void FillListD(List<DetectItem> its);
        delegate void VoidD();

        public BootForm(List<DetectItem> initialItems)
        {
            items = initialItems;
            Text = "OfficeAgent 环境引导 v0.1.0 (M0)";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(780, 606);

            Label header = new Label();
            header.Text = "首次运行自检：缺失组件将按依赖顺序从本地 payload 离线补全（无需联网）";
            header.Location = new Point(12, 12);
            header.Size = new Size(756, 18);
            Controls.Add(header);

            lv = new ListView();
            lv.View = View.Details;
            lv.FullRowSelect = true;
            lv.Location = new Point(12, 36);
            lv.Size = new Size(756, 262);
            lv.Columns.Add("组件", 200);
            lv.Columns.Add("状态", 55);
            lv.Columns.Add("处理", 105);
            lv.Columns.Add("说明", 380);
            Controls.Add(lv);

            btnDetect = MkBtn("重新检测", 12, 306, BtnDetect_Click);
            btnFix = MkBtn("安装缺失组件", 120, 306, BtnFix_Click);
            btnReport = MkBtn("生成报告", 240, 306, BtnReport_Click);
            btnLaunch = MkBtn("启动主程序", 350, 306, BtnLaunch_Click);

            lblStatus = new Label();
            lblStatus.Location = new Point(12, 342);
            lblStatus.Size = new Size(756, 18);
            lblStatus.ForeColor = Color.DimGray;
            Controls.Add(lblStatus);

            txtLog = new TextBox();
            txtLog.Multiline = true;
            txtLog.ReadOnly = true;
            txtLog.ScrollBars = ScrollBars.Vertical;
            txtLog.Font = new Font("Consolas", 9F);
            txtLog.Location = new Point(12, 366);
            txtLog.Size = new Size(756, 228);
            Controls.Add(txtLog);

            FillList(items);
            Log.Emitted += new Action<string>(OnLog);
            foreach (DetectItem it in items) { if (it.State == DetectState.Missing) { Log.Line("检测完成，存在待补全项，点击“安装缺失组件”。"); break; } }
            if (Program.AdminFix) BeginFix(true);
        }

        Button MkBtn(string text, int x, int y, EventHandler h)
        {
            Button b = new Button();
            b.Text = text;
            b.Location = new Point(x, y);
            b.Size = new Size(104, 28);
            b.Click += h;
            Controls.Add(b);
            return b;
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            Log.Emitted -= new Action<string>(OnLog);
            base.OnFormClosed(e);
        }

        // ---------- 列表 ----------

        void FillList(List<DetectItem> its)
        {
            if (InvokeRequired) { Invoke(new FillListD(FillList), its); return; }
            items = its;
            lv.BeginUpdate();
            lv.Items.Clear();
            foreach (DetectItem it in its)
            {
                ListViewItem i = new ListViewItem(it.Name);
                i.SubItems.Add(EnvDetect.StateText(it.State));
                i.SubItems.Add(it.FixKey == null ? "-" : (it.NeedAdmin ? "提权安装" : "自动补全"));
                i.SubItems.Add(it.Detail);
                i.ForeColor = ColorOf(it.State);
                lv.Items.Add(i);
            }
            lv.EndUpdate();
            UpdateStatus();
        }

        void UpdateStatus()
        {
            if (InvokeRequired) { Invoke(new VoidD(UpdateStatus)); return; }
            int missing = 0, ok = 0;
            foreach (DetectItem it in items)
            {
                if (it.State == DetectState.Missing) missing++;
                else if (it.State == DetectState.Ok) ok++;
            }
            lblStatus.Text = missing == 0
                ? "结论: P0 组件齐备，可以启动主程序。"
                : ("结论: " + missing + " 项缺失，点击“安装缺失组件”离线补全。");
        }

        static Color ColorOf(DetectState s)
        {
            switch (s)
            {
                case DetectState.Ok: return Color.DarkGreen;
                case DetectState.Missing: return Color.Red;
                case DetectState.Unknown: return Color.DarkOrange;
                default: return Color.Gray;
            }
        }

        // ---------- 按钮 ----------

        void BtnDetect_Click(object sender, EventArgs e)
        {
            RunDetect();
        }

        void BtnFix_Click(object sender, EventArgs e)
        {
            bool needAdmin = false;
            foreach (DetectItem it in items)
            {
                if (it.State == DetectState.Missing && it.NeedAdmin) { needAdmin = true; break; }
            }
            if (needAdmin && !EnvDetect.IsAdmin())
            {
                // 提权：启动 requireAdministrator 变体（Windows 自行弹 UAC）；拒绝后继续无管理员路径
                string err = BootLaunch.RestartElevated(Program.Root);
                if (err == null) { Close(); return; }
                Log.Line(err);
            }
            BeginFix(needAdmin && EnvDetect.IsAdmin());
        }

        void BtnReport_Click(object sender, EventArgs e)
        {
            string report = Report.Build(items, Program.Root);
            string path = Report.Save(report);
            Log.Line(report);
            Log.Line("报告: " + path);
        }

        void BtnLaunch_Click(object sender, EventArgs e)
        {
            // 按 .NET 环境选主程序：有 4.8 启 OfficeAgent.exe，否则启 3.5 降级壳 OfficeAgent35.exe
            bool usedLite;
            string err = BootLaunch.LaunchHost(Program.Root, out usedLite);
            if (err != null) { Log.Line(err); return; }
            Log.Line(usedLite ? "已启动主程序（.NET 3.5 降级壳）。" : "已启动主程序。");
        }

        // ---------- 后台动作 ----------

        void RunDetect()
        {
            SetStatus("检测中...");
            Thread t = new Thread(new ThreadStart(delegate
            {
                List<DetectItem> its = EnvDetect.DetectAll(Program.Root);
                FillList(its);
                SetStatusDone();
            }));
            t.IsBackground = true;
            t.Start();
        }

        void BeginFix(bool includeAdmin)
        {
            SetStatus("补全中（离线载荷 → 固定暂存 → 白名单执行）...");
            btnFix.Enabled = false;
            Thread t = new Thread(new ThreadStart(delegate
            {
                foreach (DetectItem it in items)
                {
                    if (it.FixKey == null || it.State == DetectState.Ok) continue;
                    if (it.NeedAdmin && !includeAdmin) { Log.Line("跳过(需管理员): " + it.Name); continue; }
                    string err = Installer.DoFix(it, Program.Root);
                    Log.Line(err == null ? "完成: " + it.Name : "失败: " + it.Name + " — " + err);
                }
                List<DetectItem> its = EnvDetect.DetectAll(Program.Root);
                FillList(its);
                SetStatusDone();
                if (InvokeRequired) { Invoke(new VoidD(delegate { btnFix.Enabled = true; })); }
                else btnFix.Enabled = true;
            }));
            t.IsBackground = true;
            t.Start();
        }

        void SetStatus(string s)
        {
            if (InvokeRequired) { Invoke(new SetStatusD(SetStatus), s); return; }
            lblStatus.Text = s;
        }
        delegate void SetStatusD(string s);

        void SetStatusDone()
        {
            if (InvokeRequired) { Invoke(new VoidD(SetStatusDone)); return; }
            UpdateStatus();
        }

        // ---------- 日志 ----------

        void OnLog(string s)
        {
            if (InvokeRequired) { Invoke(new AppendLogD(AppendLog), s); return; }
            AppendLog(s);
        }
        delegate void AppendLogD(string s);

        void AppendLog(string s)
        {
            try
            {
                txtLog.AppendText(s + Environment.NewLine);
            }
            catch { }
        }
    }
}
