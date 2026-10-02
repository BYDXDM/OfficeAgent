// TurnMemory —— 上一轮任务的目标继承（规划层第三期）
//
// 问题：多轮会话里用户会说"再来一次，换成上个月的"、"第二个文件也要"、"把它转成 PDF"。
//   这些话本身**不含任何可路由线索**：IntentRouter 抓不到场景词，也没有路径 token，
//   于是模型只看到一句孤零零的指代，经常反问"请问您指的是哪个文件"。
//   而 llmHistory 里虽然存着上一轮，但历史只保存**最终答复文本**，
//   任务的结构化事实（用了哪些技能、产出了哪些文件）早就丢了。
//
// 方案：从上一轮的工具执行记录里提取（技能、产物），压成一小段上下文喂给下一轮。
//
// ★ 数据源是**工具执行记录**，不是模型的自我汇报。
//   实测模型会在总结里写出与产物不符的读数（设计文档 §7 已把这条列为风险），
//   拿它的叙述当"上一轮事实"会把错误继承下去。工具日志是引擎自己写的，可信。
//
// ★ 只带上一轮、且明确标注"仅在指代时才参考"。
//   否则模型会把历史目标当成当前指令——用户开了个全新话题时就会答非所问。
//
// ★ 不做指代消解（不把"它"替换成具体路径）。
//   那属于"解释用户指令"，与 ArtifactRegistry 不还原别名是同一条红线：
//   我们只提供**事实**，判断与替换交给模型。
//
// 之所以从 ChatPanel 里抽出来：它是纯字符串/列表处理，抽成独立类才能在 /plantest 里直测，
// 而不必启动 WinForms 消息循环（否则这段解析逻辑只能靠手点 GUI 验证，回归时形同虚设）。
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeAgent.Host
{
    public static class TurnMemory
    {
        // 从工具日志里解析出**技能工具名**列表（去重，保持出现顺序）。
        //
        // 日志行形如 "🔧 skill_acct_tools_vat ✓" 或 "🔧 skill_a_b ✗" / "🔧 skill_a_b ↷跳过"。
        // 注意：只认 skill_ 前缀的 token；普通工具（read_text_file 等）不入此表——
        // 技能才是"上一轮在做什么"的语义信号，读写文件只是手段。
        public static List<string> SkillsFromLog(string toolLog)
        {
            List<string> skills = new List<string>();
            if (toolLog == null || toolLog.Length == 0) return skills;
            foreach (string line in toolLog.Split('\n'))
            {
                if (line == null) continue;
                // 状态符号（✓/✗/↷跳过）可能紧贴工具名，故按空白切分即可
                foreach (string tok in line.Trim().Split(' '))
                {
                    if (tok.Length == 0) continue;
                    if (!tok.StartsWith(SkillToolBridge.Prefix, StringComparison.Ordinal)) continue;
                    if (!skills.Contains(tok)) skills.Add(tok);
                }
            }
            return skills;
        }

        // 产物路径去重（保持顺序，跳过空值）
        public static List<string> Products(IEnumerable<string> products)
        {
            List<string> list = new List<string>();
            if (products == null) return list;
            foreach (string p in products)
            {
                if (p == null || p.Length == 0) continue;
                if (!list.Contains(p)) list.Add(p);
            }
            return list;
        }

        // 生成目标继承段。无内容时返回空串（调用方据此不追加任何文本）。
        // maxProducts：清单里最多列几个文件（超出只报数量，避免提示膨胀）。
        public static string Build(List<string> skills, List<string> products, int maxProducts)
        {
            if ((skills == null || skills.Count == 0) && (products == null || products.Count == 0))
                return "";
            if (maxProducts <= 0) maxProducts = 4;

            StringBuilder sb = new StringBuilder();
            sb.Append("\n\n[上一轮任务·仅在用户用「再来一次/换成…/它」等指代时才参考] ");
            if (skills != null && skills.Count > 0)
            {
                sb.Append("上一轮你调用了这些技能：");
                for (int i = 0; i < skills.Count; i++)
                {
                    if (i > 0) sb.Append("、");
                    sb.Append(skills[i]);
                }
                sb.Append("。");
            }
            if (products != null && products.Count > 0)
            {
                sb.Append("上一轮产出的文件：");
                int n = products.Count > maxProducts ? maxProducts : products.Count;
                for (int i = 0; i < n; i++)
                {
                    if (i > 0) sb.Append("；");
                    sb.Append(products[i]);
                }
                if (products.Count > n) sb.Append("（还有 ").Append(products.Count - n).Append(" 个）");
                sb.Append("。");
            }
            sb.Append("若用户这句话是在指代上一轮（例如「换成上个月的」「第二个也要」），")
              .Append("就沿用上面的技能与文件继续做；若用户开启了新话题，直接忽略本段。")
              .Append("不要向用户复述本段。");
            return sb.ToString();
        }
    }
}
