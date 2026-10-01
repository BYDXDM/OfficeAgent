// SkillRunner —— python38 技能执行桥（设计方案 §4.5 执行桥接）
// 红线 #1 合规形态：
//   * 进程启动复用 JobProbeRunner.RunInJob（ProcRunner.RunVersionProbe 同形态 + Job 隔离）：
//     exe 来自安装目录 runtime\py38，参数是编译期字面量 "-m oa_skill_main"——技能清单内容不参与命令行构造；
//   * 宿主把技能的 main.py 复制为暂存目录里的固定文件名 _skill_main.py，并写入 request.json
//     （任务输入/输出路径全部走文件，不过命令行）；
//   * 暂存目录固定为 %LOCALAPPDATA%\OfficeAgent\stage-skill\，技能脚本按同一约定读它；
//   * 协议：stdout 最后一行 JSON = 执行结果（{"ok":true/false,"message":"…","data":{…}}）；
//   * 进程入 Job（1.5GB 上限 + 超时整树击杀）；技能脚本必须是纯 stdlib（依赖经 store\wheels 安装，见 README）。
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OfficeAgent.Host
{
    public class SkillRunResult
    {
        public bool Ok;
        public string Message = "";
        public string DataJson = "";     // 结果 data 段原始 JSON（宿主按需解析）
        public long ElapsedMs;
        public int ExitCode;
    }

    public static class SkillRunner
    {
        public static string StageDir()
        {
            return Path.Combine(Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeAgent"),
                "stage-skill");
        }

        static string PythonExe(string root)
        {
            return Path.Combine(Path.Combine(Path.Combine(root, "runtime"), "py38"), "python.exe");
        }

        // ._pth 追加暂存目录（幂等）；embedded 隔离模式下这是唯一生效的搜索路径注入方式
        static void EnsureStageOnPath(string pthFile, string stage)
        {
            try
            {
                if (!File.Exists(pthFile)) return;
                string[] lines = File.ReadAllLines(pthFile);
                bool have = false;
                foreach (string l in lines)
                {
                    if (string.Equals(l.Trim(), stage, StringComparison.OrdinalIgnoreCase)) { have = true; break; }
                }
                if (have) return;
                StringBuilder sb = new StringBuilder();
                foreach (string l in lines) sb.Append(l).Append("\r\n");
                sb.Append(stage).Append("\r\n");
                File.WriteAllText(pthFile, sb.ToString(), new UTF8Encoding(false));
            }
            catch { }
        }

        // 执行一个 python38 技能：skillDir 内必须有 main.py；requestJson 为任务请求（含输入路径与参数）
        public static SkillRunResult Run(string root, string skillDir, string requestJson, int timeoutSeconds)
        {
            SkillRunResult r = new SkillRunResult();
            System.Diagnostics.Stopwatch sw = System.Diagnostics.Stopwatch.StartNew();
            try
            {
                string mainSrc = Path.Combine(skillDir, "main.py");
                if (!File.Exists(mainSrc)) { r.Message = "技能缺少 main.py"; return r; }

                // §4.5 权限红线：声明需要网络/Office COM 的技能，执行桥当前无沙箱代理可放行 → 拒绝执行。
                // 宁可显式失败，也不让"声明了权限"的技能在无约束下跑起来（此前两字段解析后无人读）。
                SkillManifest mf = SkillSystem.ParseManifest(skillDir);
                if (mf.Network || mf.OfficeCom)
                {
                    r.Message = "技能声明了未支持的能力（" +
                        (mf.Network ? "network " : "") + (mf.OfficeCom ? "officeCom" : "") +
                        "），执行桥无法在沙箱内放行，已拒绝执行";
                    return r;
                }

                string stage = StageDir();
                if (!Directory.Exists(stage)) Directory.CreateDirectory(stage);

                // 固定文件名暂存（清单内容只进文件，不进命令行）；
                // 以固定模块名导入：embeddable python 的 ._pth 隔离模式忽略 PYTHONPATH，
                // 故把暂存目录幂等追加进 python38._pth（boot PatchPth 同族手法），命令行保持字面量。
                string pythonExe = PythonExe(root);
                File.Copy(mainSrc, Path.Combine(stage, "oa_skill_main.py"), true);
                File.WriteAllText(Path.Combine(stage, "request.json"), requestJson, new UTF8Encoding(false));
                EnsureStageOnPath(Path.Combine(Path.GetDirectoryName(pythonExe), "python38._pth"), stage);

                string output;
                int code = OfficeAgent.Core.JobProbeRunner.RunInJob(pythonExe, "-m oa_skill_main", out output, timeoutSeconds);
                r.ExitCode = code;
                if (code == OfficeAgent.Core.ProcRunner.ExitTimeout)
                {
                    r.Message = "技能执行超时（已终止）";
                    return r;
                }
                if (code == OfficeAgent.Core.ProcRunner.ExitRejected)
                {
                    r.Message = "技能执行启动失败";
                    return r;
                }

                // stdout 最后一行以 { 开头的作为结果 JSON；stderr 末尾 300 字符用于失败诊断
                string dataJson = null;
                if (output != null)
                {
                    string[] lines = output.Split('\n');
                    for (int i = lines.Length - 1; i >= 0; i--)
                    {
                        string t = lines[i].Trim();
                        if (t.StartsWith("{"))
                        {
                            dataJson = t;
                            break;
                        }
                    }
                }
                if (dataJson == null)
                {
                    r.Message = "技能未输出结果 JSON（exit=" + code + "）";
                    r.DataJson = Tail(output, 300);
                    return r;
                }
                List<Dictionary<string, string>> objs = MiniJson.ParseObjects(dataJson);
                if (objs.Count == 0)
                {
                    r.Message = "结果 JSON 解析失败";
                    r.DataJson = Tail(dataJson, 300);
                    return r;
                }
                Dictionary<string, string> o = objs[0];
                r.Ok = MiniJson.Get(o, "ok") == "true";
                r.Message = MiniJson.Get(o, "message");
                r.DataJson = dataJson;
                if (r.Message.Length == 0) r.Message = r.Ok ? "完成" : "技能报告失败";
                return r;
            }
            catch (Exception ex)
            {
                r.Message = "技能执行异常: " + ex.Message;
                return r;
            }
            finally { sw.Stop(); r.ElapsedMs = sw.ElapsedMilliseconds; }
        }

        static string Tail(string s, int max)
        {
            if (s == null) return "";
            string t = s.Trim();
            return t.Length <= max ? t : t.Substring(t.Length - max);
        }
    }
}
