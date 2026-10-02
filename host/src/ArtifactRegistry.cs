// ArtifactRegistry —— 回合级产物登记表（规划层第二期）
//
// 问题：多步任务里，第 N 步要引用第 N-1 步产出的文件。产物路径目前只在工具返回文本里
// 出现一次（"已保存到 C:\...\工资表.xlsx"），模型得从上下文里把路径抄出来。
// 实测里简单场景模型抄得对，但产物一多（本会话实测一次回合产了 3 个文件）就开始错位：
// 把"账龄分析.xlsx"和"账龄分析_conv.pdf"混为一谈，或引用到中间产物。
//
// 方案：把本回合产出的文件登记成编号别名（产物1、产物2…），并在每次工具结果后
// 附上"当前可用产物"清单。模型可以直接说"把产物1转成 PDF"，而不必抄绝对路径。
//
// 设计取舍：
//   * **不解析模型输出里的别名**。别名只是给模型的便利说法，真正传给工具的还是绝对路径。
//     若要支持"产物1"直接当参数，就得在 AgentTools 里做文本替换——那会让"模型给的
//     参数只作数据用"这条红线变模糊（参数从数据变成了要解释的指令）。
//     故别名只作为**提示信息**回传，路径仍由模型显式给出，零解释、零注入面。
//   * 只登记确实存在的文件，避免给模型一个不存在的别名。
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OfficeAgent.Host
{
    public static class ArtifactRegistry
    {
        // 有序产物表：绝对路径（去重，保留首现顺序）
        public static List<string> Items = new List<string>();

        public static void Reset() { Items.Clear(); }

        // 登记一个产物；返回它的别名（"产物N"）。已登记过的返回原别名。
        // 不存在的文件不登记（返回空串）。
        public static string Add(string path)
        {
            if (path == null) return "";
            string p = path.Trim();
            if (p.Length == 0) return "";
            try { if (!File.Exists(p)) return ""; }
            catch { return ""; }
            for (int i = 0; i < Items.Count; i++)
            {
                if (string.Equals(Items[i], p, StringComparison.OrdinalIgnoreCase))
                    return "产物" + (i + 1);
            }
            if (Items.Count >= 20) return "";   // 上限：避免清单长到吃掉上下文
            Items.Add(p);
            return "产物" + Items.Count;
        }

        // 供回灌的清单文本；无产物时返回空串（调用方据此不追加任何内容）
        public static string Describe()
        {
            if (Items.Count == 0) return "";
            StringBuilder sb = new StringBuilder();
            sb.Append("【本回合已产出的文件】");
            for (int i = 0; i < Items.Count; i++)
            {
                sb.Append("\n  产物").Append(i + 1).Append(": ").Append(Items[i]);
            }
            sb.Append("\n（后续步骤要引用这些文件时，直接使用上面的完整路径。）");
            return sb.ToString();
        }

        // 产物文件是否都属于本回合（自检用）
        public static bool AllExist()
        {
            foreach (string p in Items)
            {
                try { if (!File.Exists(p)) return false; }
                catch { return false; }
            }
            return true;
        }
    }
}
