// SafetyConfirm —— 安全确认桥：把 SafetyGuard 的判定接到"人"身上。
//
// 分工：
//   SafetyGuard   = 纯判定（Allow / Confirm / Deny），无 UI
//   SafetyConfirm = 交互（本文件）：Confirm 时弹窗问人；无界面则按"拒绝"处理
//   UI 实现       = MainForm 注入 Ask（负责把弹窗封送到 UI 线程并阻塞等待）
//
// ★ 线程模型（关键）：工具在 agent 的后台线程执行
//   （ChatPanel.DoSend → new Thread → AgentLoop.Run → AgentTools.Dispatch → 这里）。
//   因此 Ask 的实现必须用 Control.Invoke 封送到 UI 线程，并用 ManualResetEvent 阻塞
//   后台线程直到用户作答——这正是"不确认不执行"的语义。绝不能在后台线程直接
//   ShowDialog（会自建消息泵，行为不可预期）。
//
// 安全默认：Ask == null（CLI/无界面）→ Confirm 一律按拒绝处理。
// 唯一例外是自测入口显式设置 AllowAllForTest（仅 Program.cs 自测使用）。
//
// 红线：C# 3.0 语法。
using System;
using System.IO;
using System.Text;
using OfficeAgent.Core;

namespace OfficeAgent.Host
{
    public static class SafetyConfirm
    {
        // UI 注入。参数：(rule, detail, allowDir)；返回 0=取消 1=允许一次 2=始终允许此目录。
        // null = 无界面。
        public static Func<string, string, string, int> Ask = null;

        // 仅自测/CLI 自检入口置 true，跳过一切判定（不弹窗、不拒绝）。
        // 生产 CLI 不设置它——那正是"CLI 自动拒绝"的体现。
        public static bool AllowAllForTest = false;

        // 测试钩子：置 false 时"始终允许"只改内存不落盘（/safetytest 用，避免污染真实 config.json）。
        public static bool PersistAllowList = true;

        // 工具实现调用：true=放行；false=拒绝（err 为用户/模型可读的原因）
        public static bool Ensure(SafetyOp op, string path, AppConfig cfg, out string err)
        {
            err = null;
            if (AllowAllForTest) return true;

            SafetyCheck c = SafetyGuard.Classify(op, path, cfg);
            if (c.Verdict == SafetyVerdict.Allow) return true;

            if (c.Verdict == SafetyVerdict.Deny)
            {
                Record("deny", op, c, "");
                err = "安全兜底拒绝：" + c.Reason + "。请改选一个非系统关键目录的目标。";
                return false;
            }

            // ---- Confirm ----
            string allowDir = AllowDirOf(c.NormalizedPath);
            if (Ask == null)
            {
                Record("deny-noui", op, c, "");
                err = "该操作需要用户确认（" + c.Reason + "），但当前是非交互环境，已拒绝执行。" +
                      "请在有界面的客户端里操作，或由用户明确指示后改用非系统盘目标。";
                return false;
            }

            int r;
            try { r = Ask(c.Rule, BuildDetail(op, c), allowDir); }
            catch { r = 0; }   // 弹窗自身异常 → 保守按取消

            if (r == 2 && allowDir.Length > 0)
            {
                AddAllowPath(cfg, allowDir);
                Record("always", op, c, allowDir);
                return true;
            }
            if (r == 1)
            {
                Record("once", op, c, "");
                return true;
            }
            Record("cancelled", op, c, "");
            err = "用户取消了该操作（" + c.Reason + "），未执行。请不要擅自重试，先向用户确认是否继续或换目标。";
            return false;
        }

        // "始终允许"的粒度 = 目标文件所在目录；但**盘根目录不提供该选项**
        //（允许整盘过于宽泛，容易把兜底彻底架空）。
        static string AllowDirOf(string normalizedPath)
        {
            if (normalizedPath == null || normalizedPath.Length == 0) return "";
            string dir = "";
            try { dir = Path.GetDirectoryName(normalizedPath); } catch { return ""; }
            if (dir == null || dir.Length == 0) return "";
            dir = SafetyGuard.Norm(dir);
            if (dir.Length == 0) return "";
            // 盘根（如 "C:" 或 "C:\"）→ 不提供"始终允许"
            string root = "";
            try { root = Path.GetPathRoot(dir); } catch { }
            if (root != null && root.Length > 0 &&
                string.Equals(dir.TrimEnd('\\'), root.TrimEnd('\\'), StringComparison.OrdinalIgnoreCase))
                return "";
            return dir;
        }

        static void AddAllowPath(AppConfig cfg, string dir)
        {
            if (cfg == null) return;
            string n = SafetyGuard.Norm(dir);
            if (n.Length == 0) return;
            string cur = cfg.SafetyAllowPaths == null ? "" : cfg.SafetyAllowPaths;
            string[] parts = cur.Split(';');
            for (int i = 0; i < parts.Length; i++)
            {
                string r = SafetyGuard.Norm(parts[i]);
                if (r.Length > 0 && string.Equals(r, n, StringComparison.OrdinalIgnoreCase)) return;   // 已存在
            }
            cfg.SafetyAllowPaths = cur.Length == 0 ? n : (cur + ";" + n);
            if (PersistAllowList) { try { cfg.Save(); } catch { } }
        }

        static string BuildDetail(SafetyOp op, SafetyCheck c)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("操作：").Append(OpName(op)).Append("\n");
            sb.Append("目标：").Append(c.NormalizedPath).Append("\n");
            sb.Append("原因：").Append(c.Reason).Append("\n");
            if (op == SafetyOp.Delete)
                sb.Append("\n⚠ 删除后本程序无法恢复该文件，请确认不是误删。\n");
            sb.Append("\n「始终允许此目录」= 该目录下的同类操作今后不再询问（可在设置页清除）。");
            return sb.ToString();
        }

        static string OpName(SafetyOp op)
        {
            if (op == SafetyOp.Read) return "读取";
            if (op == SafetyOp.Delete) return "删除";
            return "写入";
        }

        static void Record(string verdict, SafetyOp op, SafetyCheck c, string extra)
        {
            try
            {
                AuditLog.Record("safety_" + verdict,
                    OpName(op) + "; rule=" + (c.Rule == null ? "" : c.Rule) +
                    "; path=" + c.NormalizedPath + (extra.Length > 0 ? "; dir=" + extra : ""));
            }
            catch { }
        }
    }
}
