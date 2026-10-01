// MemoryStore —— 内置"记忆模块"存储（开箱即用，零依赖：JSON 落盘于 %LOCALAPPDATA%\OfficeAgent）
//   memory.json  长期记忆条目（用户"记住 xxx"命令或插件写入），注入对话 system prompt
//   history.jsonl 会话历史（每行一条 JSON，跨会话恢复）
// 凭据与隐私：只存用户显式让记的内容与会话文本，不出网（HostGuard 管出网，本模块不触网）。
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OfficeAgent.Host
{
    public class MemoryEntry
    {
        public string Time = "";
        public string Text = "";
    }

    public static class MemoryStore
    {
        const int MaxMemories = 200;
        const int MaxHistoryLines = 500;

        static string Dir()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeAgent");
        }
        static string MemoryFile() { return Path.Combine(Dir(), "memory.json"); }
        static string HistoryFile() { return Path.Combine(Dir(), "history.jsonl"); }

        public static string Esc(string s)
        {
            if (s == null) return "";
            StringBuilder sb = new StringBuilder();
            foreach (char ch in s)
            {
                switch (ch)
                {
                    case '\\': sb.Append("\\\\"); break;
                    case '"': sb.Append("\\\""); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default: sb.Append(ch); break;
                }
            }
            return sb.ToString();
        }

        // ---------- 长期记忆 ----------

        public static List<MemoryEntry> LoadMemories()
        {
            List<MemoryEntry> list = new List<MemoryEntry>();
            try
            {
                if (!File.Exists(MemoryFile())) return list;
                string json = File.ReadAllText(MemoryFile(), Encoding.UTF8);
                // 极简解析：顺序扫描 "t":" 与 "x":" 对
                int i = 0;
                while (i < json.Length)
                {
                    int t = json.IndexOf("\"t\":\"", i, StringComparison.Ordinal);
                    if (t < 0) break;
                    int te = json.IndexOf('"', t + 5);
                    if (te < 0) break;
                    string time = json.Substring(t + 5, te - t - 5);
                    int x = json.IndexOf("\"x\":\"", te, StringComparison.Ordinal);
                    if (x < 0) break;
                    MemoryEntry m = new MemoryEntry();
                    m.Time = time;
                    m.Text = ReadRaw(json, x + 5, out i);
                    list.Add(m);
                }
            }
            catch { }
            return list;
        }

        static string ReadRaw(string json, int start, out int next)
        {
            StringBuilder sb = new StringBuilder();
            int i = start;
            while (i < json.Length && json[i] != '"')
            {
                if (json[i] == '\\' && i + 1 < json.Length)
                {
                    i++;
                    char c = json[i];
                    if (c == 'n') sb.Append('\n');
                    else if (c == 'r') sb.Append('\r');
                    else if (c == 't') sb.Append('\t');
                    else sb.Append(c);
                }
                else sb.Append(json[i]);
                i++;
            }
            next = i + 1;
            return sb.ToString();
        }

        public static void SaveMemories(List<MemoryEntry> list)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("[\n");
            for (int i = 0; i < list.Count; i++)
            {
                sb.Append("  {\"t\":\"").Append(Esc(list[i].Time)).Append("\",\"x\":\"").Append(Esc(list[i].Text)).Append("\"}");
                if (i < list.Count - 1) sb.Append(",");
                sb.Append("\n");
            }
            sb.Append("]\n");
            Directory.CreateDirectory(Dir());
            File.WriteAllText(MemoryFile(), sb.ToString(), new UTF8Encoding(false));
        }

        public static void AddMemory(string text)
        {
            text = (text == null ? "" : text.Trim());
            if (text.Length == 0) return;
            List<MemoryEntry> list = LoadMemories();
            foreach (MemoryEntry m in list) { if (m.Text == text) return; }  // 去重
            MemoryEntry e = new MemoryEntry();
            e.Time = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
            e.Text = text;
            list.Add(e);
            while (list.Count > MaxMemories) list.RemoveAt(0);
            SaveMemories(list);
        }

        // 删除包含关键词的条目，返回删除数
        public static int RemoveMemoryLike(string keyword)
        {
            keyword = (keyword == null ? "" : keyword.Trim());
            if (keyword.Length == 0) return 0;
            List<MemoryEntry> list = LoadMemories();
            List<MemoryEntry> kept = new List<MemoryEntry>();
            int removed = 0;
            foreach (MemoryEntry m in list)
            {
                if (m.Text.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0) removed++;
                else kept.Add(m);
            }
            if (removed > 0) SaveMemories(kept);
            return removed;
        }

        // 注入 system prompt 的记忆段（无记忆返回空串；放在提示词最前，小上下文模型也能看到）
        public static string ForPrompt(bool enabled)
        {
            if (!enabled) return "";
            List<MemoryEntry> list = LoadMemories();
            if (list.Count == 0) return "";
            StringBuilder sb = new StringBuilder();
            sb.Append("用户长期记忆（用户明确要求记住的事实）：\n");
            for (int i = 0; i < list.Count; i++)
            {
                sb.Append("· ").Append(list[i].Text).Append("\n");
            }
            return sb.ToString();
        }

        // ---------- 会话历史（jsonl 追加） ----------

        public static void AppendHistory(string role, string text)
        {
            try
            {
                Directory.CreateDirectory(Dir());
                string line = "{\"role\":\"" + Esc(role) + "\",\"t\":\"" + Esc(DateTime.Now.ToString("HH:mm")) +
                              "\",\"x\":\"" + Esc(text) + "\"}\n";
                File.AppendAllText(HistoryFile(), line, Encoding.UTF8);
                TrimHistory();
            }
            catch { }
        }

        static void TrimHistory()
        {
            try
            {
                string[] lines = File.ReadAllLines(HistoryFile(), Encoding.UTF8);
                if (lines.Length <= MaxHistoryLines) return;
                StringBuilder sb = new StringBuilder();
                for (int i = lines.Length - MaxHistoryLines; i < lines.Length; i++) sb.Append(lines[i]).Append("\n");
                File.WriteAllText(HistoryFile(), sb.ToString(), Encoding.UTF8);
            }
            catch { }
        }

        public static bool HasHistory()
        {
            try { return File.Exists(HistoryFile()) && new FileInfo(HistoryFile()).Length > 0; }
            catch { return false; }
        }

        // 读最近 max 条历史：string[3] = {role, time, text}
        public static List<string[]> ReadHistoryTail(int max)
        {
            List<string[]> result = new List<string[]>();
            try
            {
                if (!File.Exists(HistoryFile())) return result;
                string[] lines = File.ReadAllLines(HistoryFile(), Encoding.UTF8);
                for (int i = lines.Length - 1; i >= 0 && result.Count < max; i--)
                {
                    string line = lines[i];
                    if (line.Trim().Length == 0) continue;
                    string role = RawField(line, "role");
                    string time = RawField(line, "t");
                    string text = RawField(line, "x");
                    if (text == null) continue;
                    result.Add(new string[] { role == null ? "assistant" : role, time == null ? "" : time, text });
                }
                result.Reverse();
            }
            catch { }
            return result;
        }

        static string RawField(string line, string name)
        {
            string pat = "\"" + name + "\":\"";
            int i = line.IndexOf(pat, StringComparison.Ordinal);
            if (i < 0) return null;
            string v = "";
            int pos = i + pat.Length;
            while (pos < line.Length && line[pos] != '"')
            {
                if (line[pos] == '\\' && pos + 1 < line.Length) { pos++; v += line[pos]; }
                else v += line[pos];
                pos++;
            }
            return v;
        }
    }
}
