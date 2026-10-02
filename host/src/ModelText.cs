// ModelText —— 模型回复的文本后处理（纯函数，无副作用）
//
// 背景：部分兼容端点（实测 zhipu glm-4.5-air）会把思考过程直接写进 message.content，
// 以 </think>（偶见 <｜end▁of▁thinking｜>）收尾，然后紧接着重写一遍正式答复。若原样展示，
// 用户会先看到一段自我矛盾、错别字横生的草稿，再看到真正的答案——实测里甚至出现
// 数字前后不一致（同一个人"应发 8,800"又在下一段写成"8,500"）和姓名错乱（"名五"/"王五"）。
//
// 处理策略：**只剥思考段，不动正式答复**。
//   * 有闭合标签（</think> 等）→ 丢弃标签及之前的全部内容，保留其后文本。
//   * 只有开标签、无闭合 → 说明思考尚未结束（多半是被截断），整段丢弃并给占位提示，
//     这比把半截思考当答案更诚实。
//   * 无任何标签 → 原样返回（绝大多数模型走这条，零影响）。
//
// 之所以放在渲染层而不是请求层：模型返回什么我们无法控制，但展示什么可以。
// 且此函数是纯函数（可单测），不依赖网络/配置，便于纳入自检。
using System;
using System.Text;

namespace OfficeAgent.Host
{
    public static class ModelText
    {
        // 闭合标记：不同模型/版本用词不一，全部覆盖（比较时统一小写）
        static readonly string[] CloseTags = new string[] {
            "</think>", "</thinking>", "</reasoning>", "<｜end▁of▁thinking｜>",
            "</analysis>", "</scratchpad>"
        };

        // 起始标记（仅用于判断"有开标签但没闭合"的情形）
        static readonly string[] OpenTags = new string[] {
            "<think>", "<thinking>", "<reasoning>", " thinking",
            "<analysis>", "<scratchpad>"
        };

        // 清理后的文本；changed 为 true 表示确实剥掉了内容（供调用方决定是否记审计）
        public static string Clean(string text, out bool changed)
        {
            changed = false;
            if (text == null || text.Length == 0) return text == null ? "" : text;

            // 找**最后一个**闭合标签：模型可能分段思考多次（草稿1</think>改稿2</think>答案3），
            // 取最后一个才能保证只留最终答复。
            int cut = -1; int cutLen = 0;
            string low = text.ToLowerInvariant();
            foreach (string tag in CloseTags)
            {
                int at = low.LastIndexOf(tag, StringComparison.Ordinal);
                if (at >= 0 && at + tag.Length > cut + cutLen - 1)
                {
                    if (at > cut) { cut = at; cutLen = tag.Length; }
                }
            }

            if (cut >= 0)
            {
                string rest = text.Substring(cut + cutLen).Trim();
                changed = true;
                // 闭合标签之后必须还有内容才有意义；否则视为"整段都是思考"
                if (rest.Length == 0)
                    return "（本次模型只返回了思考过程，没有给出正式答复。请重试，或把问题说得更具体一些。）";
                return rest;
            }

            // 无闭合标签：检查是否只有开标签（被截断的思考）
            foreach (string tag in OpenTags)
            {
                int at = low.IndexOf(tag, StringComparison.Ordinal);
                if (at >= 0)
                {
                    changed = true;
                    return "（本次模型返回的内容似乎被截断在思考阶段，没有给出正式答复。请重试，或把问题说得更具体一些。）";
                }
            }

            return text;
        }

        public static string Clean(string text)
        {
            bool ignored;
            return Clean(text, out ignored);
        }

        // 流式回调用的增量过滤器：流式场景下标签可能被拆到两次回调里
        // （第一次 "...</thi"，第二次 "nk>答案"），逐块过滤会漏判。
        // 故流式路径只做"整块里含完整标签"的粗过滤，最终仍以 Clean 处理后的文本为准。
        public static string CleanChunk(string chunk)
        {
            if (chunk == null || chunk.Length == 0) return chunk == null ? "" : chunk;
            string low = chunk.ToLowerInvariant();
            foreach (string tag in OpenTags)
            {
                if (low.IndexOf(tag, StringComparison.Ordinal) >= 0)
                {
                    // 开标签出现 → 从标签处截断（后续 delta 由调用方累积后统一 Clean）
                    int at = low.IndexOf(tag, StringComparison.Ordinal);
                    return chunk.Substring(0, at);
                }
            }
            foreach (string tag in CloseTags)
            {
                if (low.IndexOf(tag, StringComparison.Ordinal) >= 0)
                {
                    int at = low.IndexOf(tag, StringComparison.Ordinal);
                    return chunk.Substring(at + tag.Length);
                }
            }
            return chunk;
        }
    }
}
