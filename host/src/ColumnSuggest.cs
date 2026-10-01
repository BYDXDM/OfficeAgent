// ColumnSuggest —— M3 二期：LLM 列映射建议通道（设计方案 §6.6 / §7.1 上下文节流）
// 安全边界：
//   * 只把【表头名】发给模型（绝无数据行、绝无路径、绝无金额）；
//   * 模型只能从名单里挑列 —— 每个建议列名必须在真实表头里存在（归一化比对），否则整体拒绝；
//   * 建议仅作为"待确认"填充进 ActionPlan，用户仍须显式"确认"才会执行；
//   * L0 全本地 / 模型未配置 / 网络失败 / JSON 解析失败 → 一律回退手工配置提示。
// 响应格式（固定扁平 schema，逗号分隔而非数组，兼容 MiniJson）：
//   {"keyA":"账号","keyB":"对方账号","debitA":"收入","creditA":"支出",
//    "debitB":"借方","creditB":"贷方","tolerance":0.01,"explain":"一句话"}
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace OfficeAgent.Host
{
    public class SuggestResult
    {
        public bool Ok;
        public string Error = "";
        public string Explain = "";
        public ReconMapping Mapping = null;
    }

    public static class ColumnSuggest
    {
        const int MaxHeaderCols = 64;

        // 读取表头（xlsx 第一个/指定 sheet 或 csv），失败返回空数组
        public static string[] ReadHeaders(string path, string sheetSpec, int headerRow)
        {
            try
            {
                if (path == null || !File.Exists(path)) return new string[0];
                if (headerRow < 1) headerRow = 1;
                string ext = (Path.GetExtension(path) ?? "").ToLowerInvariant();
                if (ext == ".xlsx")
                {
                    using (XlsxBook book = XlsxBook.Open(path))
                    {
                        int idx = 0;
                        string s = (sheetSpec ?? "").Trim();
                        int n;
                        if (s.Length > 0 && int.TryParse(s, out n) && n >= 1 && n <= book.Sheets.Count) idx = n - 1;
                        if (idx >= book.Sheets.Count) return new string[0];
                        string[] header = null;
                        book.StreamRows(idx, MaxHeaderCols, delegate(string[] cells, int excelRow)
                        {
                            if (header == null && excelRow == headerRow) header = (string[])cells.Clone();
                        });
                        return header ?? new string[0];
                    }
                }
                if (ext == ".csv")
                {
                    Encoding used;
                    List<string[]> rows = MiniCsv.Parse(MiniCsv.DetectRead(path, out used));
                    if (rows.Count >= headerRow) return rows[headerRow - 1];
                }
                return new string[0];
            }
            catch { return new string[] { }; }
        }

        public static string HeadersText(string[] headers)
        {
            StringBuilder sb = new StringBuilder();
            int shown = 0;
            for (int i = 0; i < headers.Length && shown < MaxHeaderCols; i++)
            {
                string h = headers[i] == null ? "" : headers[i].Trim();
                if (h.Length == 0) continue;
                if (sb.Length > 0) sb.Append("|");
                sb.Append(h);
                shown++;
            }
            return sb.Length == 0 ? "(空)" : sb.ToString();
        }

        // 拉取建议。client 为已配置的 LlmClient；headersA/B 为两侧真实表头。
        // 出网口径：只发表头（无数据行/路径/金额）。表头是列名文字，不做值级脱敏——
        // 对列名套 MaskEngine 会把它改写成 [银行卡] 占位符，既无保护意义（列名非敏感值），
        // 又会破坏下面的白名单比对（脱敏后的名字不在真实表头里，建议会被整体拒绝）。
        // L0（privacyLevel==0）在调用方已拦下，不到这里。
        public static SuggestResult ForRecon(LlmClient client, string headersA, string headersB, int privacyLevel)
        {
            SuggestResult r = new SuggestResult();
            if (client == null) { r.Error = "模型未配置"; return r; }

            AuditLog.Record("llm_call", "column_suggest len=" + (headersA.Length + headersB.Length) +
                " headers_only");   // 只发表头；如需值级脱敏应在这之前对数据行做，不在此通道

            StringBuilder prompt = new StringBuilder();
            prompt.Append("你是表格核对助手的列映射建议器。只输出一个 JSON 对象，不要输出其他任何文字。\n");
            prompt.Append("任务：根据两侧表头，建议两表核对的键列与借/贷金额列。\n");
            prompt.Append("A 侧表头（用|分隔）：").Append(headersA).Append("\n");
            prompt.Append("B 侧表头（用|分隔）：").Append(headersB).Append("\n");
            prompt.Append("输出格式：{\"keyA\":\"列名,可多个用逗号分隔\",\"keyB\":\"同上\",\"debitA\":\"列名或空\",\"creditA\":\"列名或空\",\"debitB\":\"列名或空\",\"creditB\":\"列名或空\",\"tolerance\":0.01,\"explain\":\"不超过30字的理由\"}\n");
            prompt.Append("规则：列名必须逐字来自上面的表头；键列通常是账号/单号/凭证号；借-贷口径金额差为零视为匹配。");

            string sys = "你是 OfficeAgent 的列映射建议器。只输出 JSON，字段名与要求完全一致，列名逐字来自用户提供的表头。";
            string err;
            string answer = client.Chat(prompt.ToString(), sys, out err);
            if (answer == null) { r.Error = "模型请求失败: " + err; return r; }

            return ParseSuggestion(answer, headersA.Split('|'), headersB.Split('|'));
        }

        // 解析与白名单校验（模型输出的每个列名都必须存在于真实表头）
        public static SuggestResult ParseSuggestion(string answer, string[] headersA, string[] headersB)
        {
            SuggestResult r = new SuggestResult();
            int i = answer.IndexOf('{');
            int j = answer.LastIndexOf('}');
            if (i < 0 || j <= i) { r.Error = "模型未返回 JSON"; return r; }
            List<Dictionary<string, string>> objs = MiniJson.ParseObjects(answer.Substring(i, j - i + 1));
            if (objs.Count == 0) { r.Error = "JSON 解析失败"; return r; }
            Dictionary<string, string> o = objs[0];

            ReconMapping m = new ReconMapping();
            m.KeyColsA = SplitValidated(MiniJson.Get(o, "keyA"), headersA, "A 键列", r);
            if (!r.Ok && r.Error.Length > 0) return r;
            m.KeyColsB = SplitValidated(MiniJson.Get(o, "keyB"), headersB, "B 键列", r);
            if (!r.Ok && r.Error.Length > 0) return r;
            m.DebitA = ValidatedSingle(MiniJson.Get(o, "debitA"), headersA, "A 借方列", r);
            m.CreditA = ValidatedSingle(MiniJson.Get(o, "creditA"), headersA, "A 贷方列", r);
            m.DebitB = ValidatedSingle(MiniJson.Get(o, "debitB"), headersB, "B 借方列", r);
            m.CreditB = ValidatedSingle(MiniJson.Get(o, "creditB"), headersB, "B 贷方列", r);
            if (m.KeyColsA.Length == 0 || m.KeyColsB.Length == 0) { r.Error = "建议缺少键列"; return r; }
            if (m.DebitA.Length == 0 && m.CreditA.Length == 0) { r.Error = "建议缺少 A 侧金额列"; return r; }
            if (m.DebitB.Length == 0 && m.CreditB.Length == 0) { r.Error = "建议缺少 B 侧金额列"; return r; }
            double tol;
            string tolText = MiniJson.Get(o, "tolerance");
            if (double.TryParse(tolText, NumberStyles.Float, CultureInfo.InvariantCulture, out tol) && tol >= 0 && tol <= 1000)
                m.Tolerance = tol;
            r.Explain = MiniJson.Get(o, "explain");
            r.Mapping = m;
            r.Ok = true;
            return r;
        }

        // 逗号分隔的多个列名，逐个校验必须在表头中存在（归一化比对）
        static string[] SplitValidated(string value, string[] headers, string what, SuggestResult r)
        {
            List<string> parts = new List<string>();
            if (value != null)
            {
                foreach (string p in value.Split(','))
                {
                    string t = p.Trim();
                    if (t.Length == 0) continue;
                    if (!HeaderExists(t, headers))
                    {
                        r.Error = what + "「" + t + "」不在表头名单中（模型越权，整体拒绝）";
                        return parts.ToArray();
                    }
                    parts.Add(t);
                }
            }
            return parts.ToArray();
        }

        // 单列名校验；空串合法（表示"无此列"）
        static string ValidatedSingle(string value, string[] headers, string what, SuggestResult r)
        {
            string t = (value ?? "").Trim();
            if (t.Length == 0) return "";
            if (!HeaderExists(t, headers))
            {
                r.Error = what + "「" + t + "」不在表头名单中（模型越权，整体拒绝）";
                return "";
            }
            return t;
        }

        static bool HeaderExists(string name, string[] headers)
        {
            string want = ReconEngine.NormHeader(name);
            for (int i = 0; i < headers.Length; i++)
            {
                if (ReconEngine.NormHeader(headers[i]) == want) return true;
            }
            return false;
        }
    }
}
