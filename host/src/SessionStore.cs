// SessionStore —— 会话持久化（侧边栏工作区/会话列表的数据源）
// 目录：%LOCALAPPDATA%\OfficeAgent\sessions\
//   sessions-index.json  → [{"id":"...","title":"...","updated":"yyyy-MM-dd HH:mm"}]
//   <id>.json            → [{"role":"user|assistant","time":"HH:mm","text":"..."}]
// 每次助手答复后整体重写该会话文件（条目上限 200，防无限膨胀）。
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OfficeAgent.Host
{
    public class SessionInfo
    {
        public string Id = "";
        public string Title = "";
        public string Updated = "";
        public bool Archived = false;     // 归档标记：true=已归档（侧边栏默认隐藏，可恢复/删除）
        public string WorkspaceId = "default";   // 归属工作区（项目），空值按默认工作区处理
    }

    public static class SessionStore
    {
        static string Dir()
        {
            return Path.Combine(Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeAgent"), "sessions");
        }

        static string IndexPath()
        {
            return Path.Combine(Dir(), "sessions-index.json");
        }

        static string SessionPath(string id)
        {
            return Path.Combine(Dir(), SafeId(id) + ".json");
        }

        static string SafeId(string id)
        {
            if (id == null) return "";
            StringBuilder sb = new StringBuilder();
            foreach (char c in id)
            {
                bool ok = (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9') || c == '-' || c == '_';
                if (ok) sb.Append(c);
            }
            return sb.ToString();
        }

        public static string NewId()
        {
            return "s" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" +
                DateTime.Now.Millisecond.ToString("x8");
        }

        // 全部会话（按 updated 倒序）
        public static List<SessionInfo> List()
        {
            List<SessionInfo> list = new List<SessionInfo>();
            try
            {
                if (!File.Exists(IndexPath())) return list;
                foreach (Dictionary<string, string> o in MiniJson.ParseObjects(File.ReadAllText(IndexPath(), Encoding.UTF8)))
                {
                    SessionInfo s = new SessionInfo();
                    s.Id = MiniJson.Get(o, "id");
                    s.Title = MiniJson.Get(o, "title");
                    s.Updated = MiniJson.Get(o, "updated");
                    s.Archived = MiniJson.Get(o, "archived") == "true";
                    s.WorkspaceId = MiniJson.Get(o, "ws");
                    if (s.WorkspaceId == null || s.WorkspaceId.Length == 0) s.WorkspaceId = "default";
                    if (s.Id.Length > 0) list.Add(s);
                }
                list.Sort(delegate(SessionInfo a, SessionInfo b) { return string.Compare(b.Updated, a.Updated, StringComparison.Ordinal); });
            }
            catch { }
            return list;
        }

        // 读取会话消息：每行 [role, time, text]
        public static List<string[]> Load(string id)
        {
            List<string[]> rows = new List<string[]>();
            try
            {
                string f = SessionPath(id);
                if (!File.Exists(f)) return rows;
                foreach (Dictionary<string, string> o in MiniJson.ParseObjects(File.ReadAllText(f, Encoding.UTF8)))
                {
                    string role = MiniJson.Get(o, "role");
                    string time = MiniJson.Get(o, "time");
                    string text = MiniJson.Get(o, "text");
                    if (role == "user" || role == "assistant" || role == "file") rows.Add(new string[] { role, time == null ? "" : time, text == null ? "" : text });
                }
            }
            catch { }
            return rows;
        }

        // 保存会话 + 刷新索引。workspaceId：新会话首次落盘时归属的工作区（已有会话保持原归属）
        public static void Save(string id, string title, List<string[]> rows, string workspaceId)
        {
            try
            {
                if (rows == null || rows.Count == 0) return;
                Directory.CreateDirectory(Dir());
                string updated = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
                StringBuilder sb = new StringBuilder();
                sb.Append("[\n");
                for (int i = 0; i < rows.Count; i++)
                {
                    sb.Append("  {\"role\":\"").Append(MiniJson.Esc(rows[i][0]))
                      .Append("\",\"time\":\"").Append(MiniJson.Esc(rows[i][1]))
                      .Append("\",\"text\":\"").Append(MiniJson.Esc(rows[i][2])).Append("\"}");
                    if (i < rows.Count - 1) sb.Append(",");
                    sb.Append("\n");
                }
                sb.Append("]\n");
                File.WriteAllText(SessionPath(id), sb.ToString(), new UTF8Encoding(false));
                TouchIndex(id, title, updated, workspaceId);
            }
            catch { }
        }

        public static void Save(string id, string title, List<string[]> rows)
        {
            Save(id, title, rows, null);
        }

        // 把某工作区的全部会话移到另一个工作区（删除工作区时用）
        public static void MoveWorkspace(string fromId, string toId)
        {
            try
            {
                List<SessionInfo> all = List();
                bool changed = false;
                foreach (SessionInfo s in all)
                {
                    if (s.WorkspaceId == fromId) { s.WorkspaceId = toId; changed = true; }
                }
                if (changed) WriteIndex(all);
            }
            catch { }
        }

        public static void Delete(string id)
        {
            try { File.Delete(SessionPath(id)); } catch { }
            try
            {
                List<SessionInfo> all = List();
                List<SessionInfo> keep = new List<SessionInfo>();
                foreach (SessionInfo s in all) { if (s.Id != id) keep.Add(s); }
                WriteIndex(keep);
            }
            catch { }
        }

        // 归档/恢复：只改索引里的 archived 标记，会话文件不动
        public static void SetArchived(string id, bool archived)
        {
            try
            {
                List<SessionInfo> all = List();
                foreach (SessionInfo s in all) { if (s.Id == id) { s.Archived = archived; break; } }
                WriteIndex(all);
            }
            catch { }
        }

        // 统一的索引写盘（List/Save/Delete/SetArchived 都走这里，保证 archived/ws 字段不丢）
        static void WriteIndex(List<SessionInfo> all)
        {
            Directory.CreateDirectory(Dir());
            all.Sort(delegate(SessionInfo a, SessionInfo b) { return string.Compare(b.Updated, a.Updated, StringComparison.Ordinal); });
            StringBuilder sb = new StringBuilder();
            sb.Append("[\n");
            for (int i = 0; i < all.Count; i++)
            {
                sb.Append("  {\"id\":\"").Append(MiniJson.Esc(all[i].Id)).Append("\",\"title\":\"").Append(MiniJson.Esc(all[i].Title))
                  .Append("\",\"updated\":\"").Append(MiniJson.Esc(all[i].Updated)).Append("\"");
                if (all[i].Archived) sb.Append(",\"archived\":\"true\"");
                if (all[i].WorkspaceId != null && all[i].WorkspaceId.Length > 0)
                    sb.Append(",\"ws\":\"").Append(MiniJson.Esc(all[i].WorkspaceId)).Append("\"");
                sb.Append("}");
                if (i < all.Count - 1) sb.Append(",");
                sb.Append("\n");
            }
            sb.Append("]\n");
            File.WriteAllText(IndexPath(), sb.ToString(), new UTF8Encoding(false));
        }

        static void TouchIndex(string id, string title, string updated, string workspaceId)
        {
            List<SessionInfo> all = List();
            SessionInfo mine = null;
            foreach (SessionInfo s in all) { if (s.Id == id) { mine = s; break; } }
            if (mine == null)
            {
                mine = new SessionInfo();
                mine.Id = id;
                mine.WorkspaceId = workspaceId == null || workspaceId.Length == 0 ? "default" : workspaceId;
                all.Add(mine);
            }
            if (title != null && title.Length > 0) mine.Title = title;
            mine.Updated = updated;
            WriteIndex(all);
        }
    }
}
