// WorkspaceStore —— 工作区（项目）注册表：每个工作区挂一个磁盘文件夹，
// 会话归属工作区，agent 的建表/建PPT/下载/@引用 默认落到当前工作区的文件夹。
// 目录：%LOCALAPPDATA%\OfficeAgent\workspaces.json
//   [{"id":"...","name":"...","dir":"D:\\..."}]
// 当前激活工作区 id 存 config.activeWorkspaceId（随 config.json 持久化）。
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OfficeAgent.Host
{
    public class WorkspaceInfo
    {
        public string Id = "";
        public string Name = "";
        public string Dir = "";
    }

    public static class WorkspaceStore
    {
        static List<WorkspaceInfo> list = null;   // 进程内缓存（UI 线程访问）

        static string FilePath()
        {
            return Path.Combine(AppConfig.Dir(), "workspaces.json");
        }

        // 首次调用时加载；没有任何工作区时创建默认工作区（目录沿用全局配置的默认值）
        public static void EnsureInit(AppConfig cfg)
        {
            if (list != null) return;
            list = new List<WorkspaceInfo>();
            try
            {
                if (File.Exists(FilePath()))
                {
                    foreach (Dictionary<string, string> o in MiniJson.ParseObjects(File.ReadAllText(FilePath(), Encoding.UTF8)))
                    {
                        WorkspaceInfo w = new WorkspaceInfo();
                        w.Id = MiniJson.Get(o, "id");
                        w.Name = MiniJson.Get(o, "name");
                        w.Dir = MiniJson.Get(o, "dir");
                        if (w.Id.Length > 0) list.Add(w);
                    }
                }
            }
            catch { }
            if (list.Count == 0)
            {
                WorkspaceInfo w = new WorkspaceInfo();
                w.Id = "default";
                w.Name = "默认工作区";
                try { w.Dir = cfg.EffectiveWorkspace(); } catch { }
                list.Add(w);
                SaveToDisk();
            }
        }

        public static List<WorkspaceInfo> All()
        {
            return list == null ? new List<WorkspaceInfo>() : list;
        }

        public static WorkspaceInfo Find(string id)
        {
            if (id == null || id.Length == 0) return null;
            foreach (WorkspaceInfo w in All()) { if (w.Id == id) return w; }
            return null;
        }

        // 当前激活的工作区（config.activeWorkspaceId；缺失/未知时回退第一个）
        public static WorkspaceInfo Active(AppConfig cfg)
        {
            WorkspaceInfo w = Find(cfg.ActiveWorkspaceId);
            if (w != null) return w;
            List<WorkspaceInfo> all = All();
            return all.Count > 0 ? all[0] : null;
        }

        // 当前工作区实际生效的目录（空目录回退全局默认）
        public static string ActiveDir(AppConfig cfg)
        {
            WorkspaceInfo w = Active(cfg);
            if (w != null && w.Dir != null && w.Dir.Trim().Length > 0) return w.Dir;
            try { return cfg.EffectiveWorkspace(); } catch { return ""; }
        }

        // 激活工作区（写回 config 持久化）
        public static void SetActive(AppConfig cfg, string id)
        {
            cfg.ActiveWorkspaceId = id;
            cfg.Save();
        }

        public static WorkspaceInfo Add(string name, string dir)
        {
            EnsureInit(null);
            WorkspaceInfo w = new WorkspaceInfo();
            w.Id = "w" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + DateTime.Now.Millisecond.ToString("x4");
            w.Name = name == null || name.Trim().Length == 0 ? "新建工作区" : name.Trim();
            w.Dir = dir == null ? "" : dir.Trim();
            list.Add(w);
            SaveToDisk();
            return w;
        }

        public static void UpdateDir(string id, string dir)
        {
            WorkspaceInfo w = Find(id);
            if (w == null) return;
            w.Dir = dir == null ? "" : dir.Trim();
            SaveToDisk();
        }

        public static void Rename(string id, string name)
        {
            WorkspaceInfo w = Find(id);
            if (w == null) return;
            w.Name = name == null || name.Trim().Length == 0 ? w.Name : name.Trim();
            SaveToDisk();
        }

        public static void Delete(string id)
        {
            if (list == null) return;
            List<WorkspaceInfo> keep = new List<WorkspaceInfo>();
            foreach (WorkspaceInfo w in list) { if (w.Id != id) keep.Add(w); }
            if (keep.Count == 0) return;   // 至少保留一个工作区
            list = keep;
            SaveToDisk();
        }

        static void SaveToDisk()
        {
            try
            {
                Directory.CreateDirectory(AppConfig.Dir());
                StringBuilder sb = new StringBuilder();
                sb.Append("[\n");
                for (int i = 0; i < list.Count; i++)
                {
                    sb.Append("  {\"id\":\"").Append(MiniJson.Esc(list[i].Id))
                      .Append("\",\"name\":\"").Append(MiniJson.Esc(list[i].Name))
                      .Append("\",\"dir\":\"").Append(MiniJson.Esc(list[i].Dir)).Append("\"}");
                    if (i < list.Count - 1) sb.Append(",");
                    sb.Append("\n");
                }
                sb.Append("]\n");
                File.WriteAllText(FilePath(), sb.ToString(), new UTF8Encoding(false));
            }
            catch { }
        }
    }
}
