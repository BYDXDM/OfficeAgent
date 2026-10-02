// HopBudget —— Agent 循环的加权步数预算（规划层第二期）
//
// 问题：此前循环用固定计数 for (hop = 0; hop < MaxHops+1; hop++)，每一跳不管做什么都算 1。
// 但多步办公任务的成本分布极不均匀：
//   * 一次 list_directory / read_text_file 很快，也几乎没有副作用；
//   * 一次技能调用（acct-tools/payroll 实测 ~650ms，内部要起 Python sidecar）或
//     一次 convert_document（LibreOffice 转换，数秒）重得多，且真的产出文件。
// 等价计费的后果是：模型"列几次目录、读几个文件"就能把预算耗光，
// 真正干活的那一跳反而没机会执行——实测里长任务常在第 10~11 跳才走到关键步骤。
//
// 方案：按"这一跳做了什么"加权计费。
//   读型工具（list_directory / read_text_file）      0.5
//   写型工具（convert / download / create_* / repair） 1.0
//   技能调用（skill_ 前缀）                            1.0   ← 与写型同价：都要起 sidecar 且产文件
//   task_plan                                          0.0   ← 纯记账，不该占预算
//   纯文本回复（模型不再调工具）                        0.0   ← 这一跳就是收尾
//
// 预算上限从"12 跳"变为"20 等效跳"，但**读型仍然受限**：
// 见 ReadCap——单回合读型调用超过该次数后，即便预算没耗尽也不再计 0.5 而计 1.0，
// 防止模型退化成"疯狂列目录"来规避预算设计。
//
// 本类刻意做成不依赖 AgentLoop 的独立单元（纯计数 + 分类），便于自检直测。
using System;

namespace OfficeAgent.Host
{
    public static class HopBudget
    {
        // 等效跳上限：12 → 20。多步任务（列目录→建表→算税→转 PDF→写报告）
        // 在不加权时轻易超过 12，加权后 20 等效跳对应的真实跳数通常更多。
        public const double MaxEquivalent = 20.0;

        // 读型豁免额度：前 N 次读型调用按 0.5 计。超过后按 1.0 计，
        // 避免"读型便宜"被滥用成无限探测。
        public const int ReadCap = 6;

        public class State
        {
            public double Spent;        // 已消耗等效跳
            public int ReadCalls;       // 读型调用次数（用于 ReadCap）

            public double Remaining { get { double r = MaxEquivalent - Spent; return r < 0 ? 0 : r; } }
        }

        // 工具名 → 等效跳成本。budget 会被就地更新（读型计数）。
        public static double Cost(string toolName, State budget)
        {
            if (budget == null) return 1.0;
            if (toolName == null) return 1.0;

            // 计划工具：纯记账，不占预算（否则"先列计划"反而惩罚了规范用法）
            if (toolName == "task_plan") return 0.0;

            // 技能调用：与写型同价
            if (toolName.StartsWith(SkillToolBridge.Prefix, StringComparison.Ordinal)) return 1.0;

            // 读型：前 ReadCap 次半价
            if (IsReadTool(toolName))
            {
                budget.ReadCalls++;
                return budget.ReadCalls <= ReadCap ? 0.5 : 1.0;
            }

            return 1.0;
        }

        // 读型工具表：**与 AgentTools 白名单一一对应**，新增读型工具时须同步这里，
        // 否则默认按写型计 1.0（保守：宁可多算，不可少算）。
        public static bool IsReadTool(string toolName)
        {
            return toolName == "read_text_file" || toolName == "list_directory";
        }

        // 是否应当在执行完当前这一跳后强制收尾。
        // 收尾跳本身不计成本（模型不再调工具）。
        public static bool ShouldWrapUp(State budget)
        {
            if (budget == null) return false;
            return budget.Spent >= MaxEquivalent;
        }

        public static string Describe(State budget)
        {
            if (budget == null) return "";
            string s = "预算 " + budget.Spent.ToString("0.#") + "/" + MaxEquivalent.ToString("0.#") + " 等效跳";
            if (budget.ReadCalls > ReadCap)
                s += "（读型 " + budget.ReadCalls + " 次，已超出半价额度 " + ReadCap + "，后续读型按全价计）";
            return s;
        }
    }
}
