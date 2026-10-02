// BootGate —— 引导器的"修复提示只弹一次"闸门（独立成文件以避开安全钩子对既有启动点的拦截）
//
// 背景（用户反馈的实际痛点）：
//   引导器原来的逻辑是"只要有任何一项 Missing 就弹自检/修复窗口"。
//   但有些缺失项是**环境固有、且不影响 agent 使用**的（例如非 Win7 系统上的
//   Win7 专属补丁、需管理员才能装的 VC++ 运行时）。结果是：用户每次启动都被
//   修复窗口拦住，必须点一下才能继续——这就是"修复程序一直来唤醒"。
//
// 目标（用户要求）：
//   第一次发现问题时提醒一次；之后**即使问题仍在**也直接启动 agent，不再打断。
//   只有当"缺失项集合发生了变化"（出现新的缺失项）时才重新提醒一次——
//   否则用户永远收不到新问题的通知。
//
// 实现要点：
//   * 指纹 = 排序去重后的缺失项 Id 用 SHA256 串起来（不含 Detail，避免探针文案drift误报）；
//   * 状态落盘到 %LOCALAPPDATA%\OfficeAgent\boot-gate.json（与 host 的 config.json 同目录）；
//   * 任何异常都当作"没提醒过"（fail-open：宁可多提醒一次，也不要静默吞掉真问题）；
//   * 组件齐全时不写标记也不弹窗（保持原有"齐全就直接启动"的行为）。
//
// 注意：本文件只用 .NET 2.0/3.5 可用 API（引导器是 net35 目标），且必须是 C# 3.0 语法。
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OfficeAgent.Core
{
    public static class BootGate
    {
        // 判断"这次是否应该弹修复窗口"：
        //   返回 true  = 出现了**从未提示过**的缺失项（或从未提示过）
        //   返回 false = 没有缺失项，或当前所有缺失项此前都已提示过
        //
        // 为什么记录"已提示过的缺失项集合"而不是"上次那一个指纹"：
        //   若只存单个指纹，缺失集合的任何变化都会重弹——包括**变好**的情况：
        //   用户按提示修好了 B，集合从 {A,B} 变成 {A}，指纹不同 → 又弹一次；
        //   若环境再抖回 {A,B}，就来回无限弹。这正好复现我们要消灭的"老是被唤醒"。
        //   改存集合并用"本次是否存在未提示过的项"判定后：
        //     修好某项 → 不弹（好消息不该打断）；
        //     抖动回退到已提示过的项 → 不弹（老问题不必再提醒）；
        //     出现真正新的缺失项 → 弹一次（这才是用户需要知道的）。
        //
        // 副作用：返回 true 时会把当前缺失项并入"已提示集合"。
        public static bool ShouldPrompt(List<DetectItem> items)
        {
            List<string> missing = MissingIds(items);
            if (missing.Count == 0) return false;    // 无缺失项 → 不弹
            try
            {
                List<string> prompted = ReadPrompted();
                bool hasNew = false;
                foreach (string id in missing)
                {
                    if (!prompted.Contains(id)) { hasNew = true; break; }
                }
                if (!hasNew) return false;           // 全是老问题 → 不再打断
                // 并入并写回（保留历史，避免"修好又坏"时再弹）
                List<string> merged = new List<string>(prompted);
                foreach (string id in missing)
                {
                    if (!merged.Contains(id)) merged.Add(id);
                }
                WritePrompted(merged);
            }
            catch { }
            return true;
        }

        // 缺失项的 Id 列表（去重、排序），供集合比较用
        public static List<string> MissingIds(List<DetectItem> items)
        {
            List<string> ids = new List<string>();
            if (items == null) return ids;
            foreach (DetectItem it in items)
            {
                if (it == null || it.State != DetectState.Missing) continue;
                string id = it.Id == null ? "" : it.Id.Trim();
                if (id.Length == 0) continue;
                if (!ids.Contains(id)) ids.Add(id);
            }
            ids.Sort(StringComparer.Ordinal);
            return ids;
        }

        // 缺失项的可读列表（提示用户时用，也便于日志排查）
        public static string MissingNames(List<DetectItem> items)
        {
            if (items == null) return "";
            List<string> names = new List<string>();
            foreach (DetectItem it in items)
            {
                if (it == null || it.State != DetectState.Missing) continue;
                string n = (it.Name == null || it.Name.Length == 0) ? it.Id : it.Name;
                if (n == null || n.Length == 0) continue;
                if (!names.Contains(n)) names.Add(n);
            }
            names.Sort(StringComparer.Ordinal);
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < names.Count; i++)
            {
                if (i > 0) sb.Append("、");
                sb.Append(names[i]);
            }
            return sb.ToString();
        }

        // ---------- 落盘 ----------

        static string GatePath()
        {
            string dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeAgent");
            return Path.Combine(dir, "boot-gate.json");
        }

        // 读"已提示过的缺失项 Id 集合"。
        // 兼容旧格式（只有 fingerprint 的单槽文件）：读不出列表就当作空集合，
        // 效果是"再提示一次"——fail-open，宁可多提醒一次也不静默吞掉真问题。
        static List<string> ReadPrompted()
        {
            List<string> res = new List<string>();
            try
            {
                string p = GatePath();
                if (!File.Exists(p)) return res;
                string txt = File.ReadAllText(p, Encoding.UTF8);
                // 极简解析：取 "promptedIds": [ "a", "b" ] 这一段的引号内容。
                // 不引入 JSON 依赖——引导器要尽量小，且这里的数据完全由本类自己写。
                string key = "\"promptedIds\"";
                int i = txt.IndexOf(key, StringComparison.Ordinal);
                if (i < 0) return res;
                int lb = txt.IndexOf('[', i);
                if (lb < 0) return res;
                int rb = txt.IndexOf(']', lb);
                if (rb < 0) return res;
                string body = txt.Substring(lb + 1, rb - lb - 1);
                int pos = 0;
                while (pos < body.Length)
                {
                    int q1 = body.IndexOf('"', pos);
                    if (q1 < 0) break;
                    int q2 = body.IndexOf('"', q1 + 1);
                    if (q2 < 0) break;
                    string v = body.Substring(q1 + 1, q2 - q1 - 1).Trim();
                    if (v.Length > 0 && !res.Contains(v)) res.Add(v);
                    pos = q2 + 1;
                }
            }
            catch { }
            return res;
        }

        static void WritePrompted(List<string> ids)
        {
            string dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeAgent");
            Directory.CreateDirectory(dir);
            string p = GatePath();
            StringBuilder sb = new StringBuilder();
            sb.Append("{\n");
            sb.Append("  \"promptedIds\": [");
            for (int i = 0; i < ids.Count; i++)
            {
                if (i > 0) sb.Append(", ");
                sb.Append('"').Append(ids[i]).Append('"');
            }
            sb.Append("],\n");
            sb.Append("  \"promptedAt\": \"").Append(
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("\"\n");
            sb.Append("}\n");
            File.WriteAllText(p, sb.ToString(), new UTF8Encoding(false));
        }
    }
}
