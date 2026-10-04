// ConfirmDialog —— 安全兜底确认弹窗（自建轻量窗体）
//
// 为什么不用 MessageBox：需要三按钮（允许一次 / 始终允许此目录 / 取消）与自定义文案，
// MessageBox 做不到；且我们要把"触发规则 + 完整目标路径"清楚地摆给用户。
//
// 交互约定（安全默认）：
//   * 默认焦点在「取消」，Esc = 取消；Enter 不触发"允许"（AcceptButton 不设为允许按钮）
//   * 返回：0=取消  1=允许一次  2=始终允许此目录（allowDir 为空时该按钮不显示，恒为 0/1）
//
// 红线：C# 3.0 语法；本窗体只在 UI 线程创建（由 SafetyConfirm 负责封送）。
using System;
using System.Drawing;
using System.Windows.Forms;

namespace OfficeAgent.Host
{
    static class ConfirmDialog
    {
        public static int Show(IWin32Window owner, string rule, string detail, string allowDir)
        {
            int result = 0;
            bool canAlways = allowDir != null && allowDir.Trim().Length > 0;

            using (Form f = new Form())
            {
                f.Text = "安全确认";
                f.FormBorderStyle = FormBorderStyle.FixedDialog;
                f.MinimizeBox = false;
                f.MaximizeBox = false;
                f.ShowInTaskbar = false;
                f.StartPosition = FormStartPosition.CenterParent;
                f.ClientSize = new Size(520, 306);
                try { f.Font = new Font("Microsoft YaHei UI", 9F); } catch { }
                try { f.Icon = owner is Form ? ((Form)owner).Icon : null; } catch { }

                Label lblRule = new Label();
                lblRule.Text = "⚠ " + (rule == null || rule.Length == 0 ? "高风险操作" : rule);
                try { lblRule.Font = new Font("Microsoft YaHei UI", 11F, FontStyle.Bold); } catch { }
                lblRule.ForeColor = Color.FromArgb(178, 34, 34);
                lblRule.Location = new Point(20, 16);
                lblRule.AutoSize = true;
                f.Controls.Add(lblRule);

                Label lblHint = new Label();
                lblHint.Text = "该操作可能影响系统盘文件或删除表格，请确认后继续：";
                lblHint.ForeColor = Color.FromArgb(90, 93, 105);
                lblHint.Location = new Point(22, 48);
                lblHint.AutoSize = true;
                f.Controls.Add(lblHint);

                TextBox tb = new TextBox();
                tb.Multiline = true;
                tb.ReadOnly = true;
                tb.ScrollBars = ScrollBars.Vertical;
                tb.BackColor = Color.FromArgb(250, 250, 252);
                tb.BorderStyle = BorderStyle.FixedSingle;
                tb.Location = new Point(22, 74);
                tb.Size = new Size(476, 158);
                tb.Text = detail == null ? "" : detail;
                f.Controls.Add(tb);

                Button bCancel = new Button();
                bCancel.Text = "取消";
                bCancel.Location = new Point(22, 250);
                bCancel.Size = new Size(110, 32);
                bCancel.Cursor = Cursors.Hand;
                bCancel.Click += delegate { result = 0; f.Close(); };
                f.Controls.Add(bCancel);

                Button bOnce = new Button();
                bOnce.Text = "允许一次";
                bOnce.Location = new Point(250, 250);
                bOnce.Size = new Size(110, 32);
                bOnce.Cursor = Cursors.Hand;
                bOnce.Click += delegate { result = 1; f.Close(); };
                f.Controls.Add(bOnce);

                if (canAlways)
                {
                    Button bAlways = new Button();
                    bAlways.Text = "始终允许此目录";
                    bAlways.Location = new Point(368, 250);
                    bAlways.Size = new Size(130, 32);
                    bAlways.Cursor = Cursors.Hand;
                    bAlways.Click += delegate { result = 2; f.Close(); };
                    f.Controls.Add(bAlways);
                }

                f.CancelButton = bCancel;   // Esc → 取消
                f.AcceptButton = null;      // Enter 不触发"允许"（防误触放行）
                f.ActiveControl = bCancel;  // 默认焦点在取消

                if (owner != null) f.ShowDialog(owner);
                else f.ShowDialog();
            }
            return result;
        }
    }
}
