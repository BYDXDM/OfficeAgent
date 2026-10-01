// PlainTips —— 界面通俗化映射：把会计/安全术语翻译成大白话，配合"❗"悬停提示
// 只影响界面显示；Excel 报告与审计日志内容保持专业口径不变。
using System;
using System.Collections.Generic;

namespace OfficeAgent.Host
{
    public static class PlainTips
    {
        class Tip
        {
            public string Plain = "";
            public string Hover = "";
            public Tip(string plain, string hover) { Plain = plain; Hover = hover; }
        }

        static readonly Dictionary<string, Tip> Checks = BuildChecks();
        static readonly Dictionary<string, Tip> Events = BuildEvents();

        static Dictionary<string, Tip> BuildChecks()
        {
            Dictionary<string, Tip> d = new Dictionary<string, Tip>();
            d["合计勾稽"] = new Tip("总数对得上",
                "两侧的金额总数相减，应该正好等于「差异明细里各条差额加起来」再加「容差内的微小误差」。对不上说明统计有漏洞，请不要直接使用本次结果。");
            d["行数守恒"] = new Tip("一行都没漏",
                "配对成功的行 + 只在一侧出现的行 = 各自的总行数。这保证每一行都被核对过，没有多算也没有漏算。");
            d["抽样重算"] = new Tip("抽查复核",
                "程序从差异行里抽出最多 20 行，用原始单元格里的数字重新算一遍差额，防止程序自己算错。");
            d["重复键提示"] = new Tip("有重复的账号/单号",
                "同一个键值出现多次时，程序按先后顺序一一配对。如果业务里不该有重复，请人工再核对一遍。");
            d["金额解析错误"] = new Tip("读不懂的金额",
                "这些单元格不是标准数字（可能是文字、批注或乱码），已被跳过并单独列成清单，请人工检查后再确认结果。");
            d["文件装载"] = new Tip("文件都能打开",
                "汇总时如果有文件读取失败，会被跳过并注明原因，其余文件正常参与汇总。");
            d["单票勾稽"] = new Tip("每张票金额自洽",
                "每张发票应满足：金额 + 税额 = 价税合计。不等的票请人工核实，清单里已标出。");
            d["提取成功率"] = new Tip("识别成功了多少",
                "成功提取的发票数与失败数对比。失败的票列在「未识别清单」，常见原因是扫描件没有文字层。");
            return d;
        }

        static Dictionary<string, Tip> BuildEvents()
        {
            Dictionary<string, Tip> d = new Dictionary<string, Tip>();
            d["app_start"] = new Tip("启动程序", "软件启动的时间记录。");
            d["llm_call"] = new Tip("使用 AI 模型",
                "向配置的 AI 模型发送了问题。若开了「脱敏出网」，会标注命中了哪些脱敏规则；问题原文不写入审计。");
            d["file_read"] = new Tip("读取文件", "程序读取了这个原始文件（只读，绝不修改原文件）。");
            d["file_write"] = new Tip("写出结果", "程序生成了这个结果文件（核对差异表 / 汇总底稿 / 发票清单等）。");
            d["plan_created"] = new Tip("生成任务计划",
                "根据你说的话生成了任务计划，正在等你确认。这一步还没有动任何文件。");
            d["confirmation"] = new Tip("你确认/取消了任务", "记录你确认或取消了待执行任务。");
            d["action_started"] = new Tip("开始执行任务", "你确认之后，任务真正开始执行。");
            d["action_finished"] = new Tip("任务执行完成", "任务结束，这里记录成功或失败以及结果摘要。");
            d["recon_run"] = new Tip("表格核对", "执行了一次两表核对。");
            d["merge_run"] = new Tip("报表汇总", "执行了一次多文件报表汇总。");
            d["invoice_run"] = new Tip("发票提取", "执行了一次发票批量提取。");
            d["config_change"] = new Tip("修改设置", "某个内置插件的开关被打开或关闭。");
            d["privacy_change"] = new Tip("调整隐私等级",
                "隐私分级（L0 全本地 / L1 脱敏出网 / L2 全量出网）被修改，决定发往 AI 模型的内容口径。");
            return d;
        }

        // 检查项：前缀匹配（如「合计勾稽：金额」带列名后缀也能命中）
        public static string PlainCheck(string name)
        {
            Tip t = FindByPrefix(Checks, name);
            return t == null ? name : t.Plain;
        }

        public static string CheckHover(string name)
        {
            Tip t = FindByPrefix(Checks, name);
            return t == null ? "程序自动做的检查。" : t.Hover;
        }

        // 审计事件名 → 通俗名
        public static string PlainEvent(string type)
        {
            if (type == null) return "";
            Tip t;
            if (Events.TryGetValue(type, out t)) return t.Plain;
            return type;
        }

        public static string EventHover(string type)
        {
            if (type == null) return "程序事件记录。";
            Tip t;
            if (Events.TryGetValue(type, out t)) return t.Hover;
            return "程序事件记录。";
        }

        static Tip FindByPrefix(Dictionary<string, Tip> d, string name)
        {
            if (name == null) return null;
            foreach (KeyValuePair<string, Tip> kv in d)
            {
                if (name.StartsWith(kv.Key, StringComparison.Ordinal)) return kv.Value;
            }
            return null;
        }
    }
}
