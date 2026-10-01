// AuditLog —— 防篡改事件审计（设计方案 §7.3 哈希链）
// 形态说明：M2 以零依赖 JSONL 落地（SHA256 链式签名），SQLite 迁移随团队版（M4）评估；
// 语义与设计一致：每事件含前一条哈希，任何删改/插入都会被 Verify 发现。
//   行格式：{"seq":N,"t":"yyyy-MM-dd HH:mm:ss","type":"…","detail":"…","prev":"…","hash":"…"}
//   hash = SHA256(prev|seq|t|type|detail)；首行 prev = GENESIS。
// 文件：%LOCALAPPDATA%\OfficeAgent\audit.jsonl（>5MB 轮转为 audit-<时间戳>.jsonl）
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace OfficeAgent.Host
{
    public static class AuditLog
    {
        const long MaxBytes = 5L * 1024 * 1024;
        static object gate = new object();
        static bool loaded = false;
        static long lastSeq = 0;
        static string lastHash = "GENESIS";

        public static string Dir()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeAgent");
        }

        public static string FilePath()
        {
            return Path.Combine(Dir(), "audit.jsonl");
        }

        static string NowText()
        {
            return DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        }

        static string ComputeHash(string prev, long seq, string t, string type, string detail)
        {
            string payload = prev + "|" + seq.ToString(CultureInfo.InvariantCulture) + "|" + t + "|" + type + "|" + detail;
            using (SHA256 sha = SHA256.Create())
            {
                byte[] b = sha.ComputeHash(Encoding.UTF8.GetBytes(payload));
                StringBuilder sb = new StringBuilder(b.Length * 2);
                foreach (byte x in b) sb.Append(x.ToString("x2", CultureInfo.InvariantCulture));
                return sb.ToString();
            }
        }

        static void LoadTail()
        {
            lastSeq = 0;
            lastHash = "GENESIS";
            try
            {
                if (!File.Exists(FilePath())) return;
                using (FileStream fs = new FileStream(FilePath(), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (StreamReader sr = new StreamReader(fs, Encoding.UTF8))
                {
                    string line, lastLine = null;
                    while ((line = sr.ReadLine()) != null) { if (line.Trim().Length > 0) lastLine = line; }
                    if (lastLine != null)
                    {
                        lastSeq = GetLong(lastLine, "seq");
                        lastHash = GetStr(lastLine, "hash");
                        if (lastHash == null || lastHash.Length != 64) lastHash = "GENESIS";
                    }
                }
            }
            catch { }
            loaded = true;
        }

        // 记录一条事件（任何线程可调；失败静默——审计不阻断业务）
        public static void Record(string type, string detail)
        {
            try
            {
                lock (gate)
                {
                    if (!loaded) LoadTail();
                    string dir = Dir();
                    if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                    string path = FilePath();
                    try
                    {
                        FileInfo fi = new FileInfo(path);
                        if (fi.Exists && fi.Length > MaxBytes)
                        {
                            string archive = Path.Combine(Dir(), "audit-" +
                                DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture) + ".jsonl");
                            File.Move(path, archive);
                            LoadTail();
                        }
                    }
                    catch { }
                    long seq = lastSeq + 1;
                    string t = NowText();
                    string h = ComputeHash(lastHash, seq, t, type, detail);
                    string line = "{\"seq\":" + seq.ToString(CultureInfo.InvariantCulture) +
                                  ",\"t\":\"" + t +
                                  "\",\"type\":\"" + Esc(type) +
                                  "\",\"detail\":\"" + Esc(detail) +
                                  "\",\"prev\":\"" + lastHash +
                                  "\",\"hash\":\"" + h + "\"}";
                    File.AppendAllText(path, line + "\n", new UTF8Encoding(false));
                    lastSeq = seq;
                    lastHash = h;
                }
            }
            catch { }
        }

        // 校验全链（含轮转前的当前文件）：返回 null=通过，否则错误描述；brokenSeq=断点
        public static string VerifyChain(out int checkedCount, out long brokenSeq)
        {
            checkedCount = 0;
            brokenSeq = -1;
            try
            {
                if (!File.Exists(FilePath())) return null;
                string prev = "GENESIS";
                using (StreamReader sr = new StreamReader(FilePath(), Encoding.UTF8))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        if (line.Trim().Length == 0) continue;
                        long seq = GetLong(line, "seq");
                        string t = GetStr(line, "t");
                        string type = GetStr(line, "type");
                        string detail = GetStr(line, "detail");
                        string p = GetStr(line, "prev");
                        string h = GetStr(line, "hash");
                        checkedCount++;
                        if (p != prev || h != ComputeHash(prev, seq, t, type, detail))
                        {
                            brokenSeq = seq;
                            return "第 " + seq + " 条链校验失败（审计日志可能被篡改）";
                        }
                        prev = h;
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                return "校验异常: " + ex.Message;
            }
        }

        // 读最近 max 条：string[4] = {seq, time, type, detail}
        public static List<string[]> ReadTail(int max)
        {
            List<string[]> result = new List<string[]>();
            try
            {
                if (!File.Exists(FilePath())) return result;
                string[] lines = File.ReadAllLines(FilePath(), Encoding.UTF8);
                for (int i = lines.Length - 1; i >= 0 && result.Count < max; i--)
                {
                    string line = lines[i];
                    if (line.Trim().Length == 0) continue;
                    result.Add(new string[] {
                        GetLong(line, "seq").ToString(CultureInfo.InvariantCulture),
                        GetStr(line, "t"),
                        GetStr(line, "type"),
                        GetStr(line, "detail") });
                }
                result.Reverse();
            }
            catch { }
            return result;
        }

        static string Esc(string s)
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

        static string GetStr(string line, string key)
        {
            string pat = "\"" + key + "\":\"";
            int i = line.IndexOf(pat, StringComparison.Ordinal);
            if (i < 0) return null;
            int pos = i + pat.Length;
            StringBuilder sb = new StringBuilder();
            while (pos < line.Length && line[pos] != '"')
            {
                if (line[pos] == '\\' && pos + 1 < line.Length)
                {
                    pos++;
                    char c = line[pos];
                    if (c == 'n') sb.Append('\n');
                    else if (c == 'r') sb.Append('\r');
                    else if (c == 't') sb.Append('\t');
                    else sb.Append(c);
                }
                else sb.Append(line[pos]);
                pos++;
            }
            return sb.ToString();
        }

        static long GetLong(string line, string key)
        {
            string pat = "\"" + key + "\":";
            int i = line.IndexOf(pat, StringComparison.Ordinal);
            if (i < 0) return 0;
            i += pat.Length;
            int end = i;
            while (end < line.Length && line[end] >= '0' && line[end] <= '9') end++;
            long v;
            return long.TryParse(line.Substring(i, end - i), NumberStyles.Integer,
                CultureInfo.InvariantCulture, out v) ? v : 0;
        }
    }
}
