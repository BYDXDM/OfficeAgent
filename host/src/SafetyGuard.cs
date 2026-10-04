// SafetyGuard —— 文件操作安全兜底判定层（纯函数，无 UI 依赖，可单测）
//
// 职责：给定「操作类型 + 目标路径」，返回三态判定 Allow / Confirm / Deny。
// 本类**不做任何交互**——弹窗由 SafetyConfirm（UI 桥）负责，便于：
//   * /safetytest 无头单测（判定分支全覆盖）
//   * CLI 场景复用同一判定，交互退化为"拒绝"
//
// 判定优先级（先严后宽，见 Classify）：
//   0) 总开关关闭 → Allow
//   1) 系统关键目录（Windows / Program Files / ProgramData）→ Deny（不可放行）
//   2) 读操作 → Allow（读无破坏性，不打断）
//   3) 排除目录（系统临时目录 / OfficeAgent 自管目录）→ Allow
//   4) "始终允许"白名单 → Allow
//   5) 删除 + 表格扩展名 → Confirm（"删除表格"）
//   6) 写/删 + 系统盘 → Confirm（"涉及C盘文件"）
//   7) 其余 → Allow
//
// 红线：C# 3.0 语法（双目标含 .NET 3.5）；本文件不参与任何命令行/shell 构造。
using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeAgent.Host
{
    public enum SafetyOp { Read, Write, Delete }
    public enum SafetyVerdict { Allow, Confirm, Deny }

    public class SafetyCheck
    {
        public SafetyVerdict Verdict = SafetyVerdict.Allow;
        public string Reason = "";          // 触发理由（多条以 "；" 连接），供弹窗与审计展示
        public string NormalizedPath = "";  // 归一化后的目标路径（可能为空：无法解析时）
        public string Rule = "";            // 命中的规则名（"删除表格" / "涉及C盘文件" / "系统关键目录"）
    }

    public static class SafetyGuard
    {
        // "任意表格"：删除这些扩展名的文件必须确认
        static readonly string[] TableExts =
            { ".xlsx", ".xls", ".xlsm", ".xlsb", ".csv", ".tsv", ".ods" };

        // ---------- 路径工具 ----------

        // 归一化：GetFullPath + 去尾部分隔符；失败返回 ""
        public static string Norm(string path)
        {
            if (path == null) return "";
            string p = path.Trim();
            if (p.Length == 0) return "";
            try { p = Path.GetFullPath(p); } catch { return ""; }
            if (p.Length > 3 && (p[p.Length - 1] == '\\' || p[p.Length - 1] == '/'))
                p = p.Substring(0, p.Length - 1);
            return p;
        }

        // target 是否位于 root 之下（含 root 本身）。按路径分隔符边界比较，避免
        // C:\Windows 误匹配 C:\WindowsApps。
        static bool UnderRoot(string target, string root)
        {
            if (target.Length == 0 || root.Length == 0) return false;
            string r = root;
            if (r.Length > 0 && r[r.Length - 1] != '\\') r += "\\";
            if (target.Length == r.Length - 1)
                return string.Equals(target, r.Substring(0, r.Length - 1), StringComparison.OrdinalIgnoreCase);
            return target.StartsWith(r, StringComparison.OrdinalIgnoreCase);
        }

        // ---------- 单点判定 ----------

        public static bool IsTableFile(string path)
        {
            string p = path == null ? "" : path;
            string ext = "";
            try { ext = Path.GetExtension(p); } catch { return false; }
            if (ext == null) return false;
            ext = ext.ToLowerInvariant();
            for (int i = 0; i < TableExts.Length; i++)
                if (ext == TableExts[i]) return true;
            return false;
        }

        // 目标是否在系统盘（通常是 C:）
        public static bool IsSystemDrive(string path)
        {
            string p = Norm(path);
            if (p.Length == 0) return false;
            string root = "";
            try { root = Path.GetPathRoot(p); } catch { return false; }
            if (root == null || root.Length == 0) return false;
            string sysRoot = "";
            try { sysRoot = Path.GetPathRoot(Environment.SystemDirectory); } catch { }
            if (sysRoot == null || sysRoot.Length == 0) return false;
            return string.Equals(root.TrimEnd('\\'), sysRoot.TrimEnd('\\'), StringComparison.OrdinalIgnoreCase);
        }

        // 系统关键目录（写/删一律拒绝，不可放行）
        static List<string> ForbiddenRoots()
        {
            List<string> list = new List<string>();
            // 注意：SpecialFolder.Windows / ProgramFilesX86 均为 .NET 4.0+，本项目双目标含 3.5，
            // 故统一改用环境变量（Win7 上 SystemRoot / ProgramFiles(x86) 均存在）。
            AddIf(list, delegate { return Environment.GetEnvironmentVariable("SystemRoot"); });
            AddIf(list, delegate { return Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles); });
            AddIf(list, delegate { return Environment.GetEnvironmentVariable("ProgramFiles(x86)"); });
            AddIf(list, delegate { return Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData); });
            return list;
        }

        // 排除目录：系统临时目录 + OfficeAgent 自管目录（写/删无需确认）
        // 依据：这些是应用自产的中间产物/预览缓存，弹窗只会造成打扰且无安全价值。
        static List<string> ExcludedRoots()
        {
            List<string> list = new List<string>();
            AddIf(list, delegate { return Path.GetTempPath(); });
            AddIf(list, delegate
            {
                return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeAgent");
            });
            return list;
        }

        static void AddIf(List<string> list, Func<string> get)
        {
            try
            {
                string s = get();
                string n = Norm(s);
                if (n.Length > 0) list.Add(n);
            }
            catch { }
        }

        public static bool IsForbidden(string path)
        {
            string p = Norm(path);
            if (p.Length == 0) return false;
            List<string> roots = ForbiddenRoots();
            for (int i = 0; i < roots.Count; i++)
                if (UnderRoot(p, roots[i])) return true;
            return false;
        }

        public static bool IsExcluded(string path)
        {
            string p = Norm(path);
            if (p.Length == 0) return false;
            List<string> roots = ExcludedRoots();
            for (int i = 0; i < roots.Count; i++)
                if (UnderRoot(p, roots[i])) return true;
            return false;
        }

        // "始终允许"白名单：cfg.SafetyAllowPaths 为分号分隔的目录前缀
        static bool InAllowList(string path, AppConfig cfg)
        {
            if (cfg == null) return false;
            string s = cfg.SafetyAllowPaths;
            if (s == null || s.Trim().Length == 0) return false;
            string p = Norm(path);
            if (p.Length == 0) return false;
            string[] parts = s.Split(';');
            for (int i = 0; i < parts.Length; i++)
            {
                string r = Norm(parts[i]);
                if (r.Length == 0) continue;
                if (UnderRoot(p, r)) return true;
            }
            return false;
        }

        // ---------- 主判定 ----------

        public static SafetyCheck Classify(SafetyOp op, string path, AppConfig cfg)
        {
            SafetyCheck c = new SafetyCheck();

            // 0) 总开关（默认开；用户显式关闭则整层放行）
            if (cfg != null && !cfg.SafetyGuard) { c.Verdict = SafetyVerdict.Allow; return c; }

            c.NormalizedPath = Norm(path);
            if (c.NormalizedPath.Length == 0)
            {
                // 无法解析路径（或未给路径）→ 不在本层拦截，交由调用方
                c.Verdict = SafetyVerdict.Allow;
                return c;
            }

            // 1) 系统关键目录 → Deny（不可放行）
            if (IsForbidden(c.NormalizedPath))
            {
                c.Verdict = SafetyVerdict.Deny;
                c.Rule = "系统关键目录";
                c.Reason = "目标是系统关键目录（Windows / Program Files / ProgramData），禁止写入或删除";
                return c;
            }

            // 2) 读 → Allow
            if (op == SafetyOp.Read) { c.Verdict = SafetyVerdict.Allow; return c; }

            // 3) 排除目录（临时/自管）→ Allow
            if (IsExcluded(c.NormalizedPath)) { c.Verdict = SafetyVerdict.Allow; return c; }

            // 4) 白名单 → Allow
            if (InAllowList(c.NormalizedPath, cfg)) { c.Verdict = SafetyVerdict.Allow; return c; }

            // 5) 删表格 / 6) 写删系统盘 → Confirm
            bool del = op == SafetyOp.Delete;
            bool isTable = IsTableFile(c.NormalizedPath);
            bool sysDrive = IsSystemDrive(c.NormalizedPath);

            string rule = "";
            List<string> reasons = new List<string>();
            if (del && isTable)
            {
                rule = "删除表格";
                reasons.Add("将删除表格文件：" + Path.GetFileName(c.NormalizedPath));
            }
            if (sysDrive)
            {
                if (rule.Length == 0) rule = "涉及C盘文件";
                reasons.Add("目标位于系统盘（" + (Path.GetPathRoot(c.NormalizedPath) ?? "") + "）");
            }

            if (reasons.Count > 0)
            {
                c.Verdict = SafetyVerdict.Confirm;
                c.Rule = rule;
                c.Reason = string.Join("；", reasons.ToArray());
            }
            return c;
        }
    }
}
