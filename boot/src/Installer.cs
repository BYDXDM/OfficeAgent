// 组件补全：离线 payload 校验（SHA256）→ 复制到固定暂存路径 → ProcRunner 白名单执行
using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using OfficeAgent.Core;

namespace OfficeAgent.Boot
{
    public static class Installer
    {
        // 返回 null = 成功；否则错误信息
        public static string DoFix(DetectItem item, string root)
        {
            LogProxyWrite("补全: " + item.Name + " ...");
            try
            {
                switch (item.FixKey)
                {
                    case "install_kb4490628":
                        // 按系统位宽选变体（payload 双架构齐备；x86 变体来自 Microsoft Update Catalog）
                        return EnvDetect.Is64OS()
                            ? InstallMsu(root, "payload/kb/kb4490628-x64.msu",
                                "8075f6d889bcb27be6f52ed47081675e5bb8a5390f2f5bfe4ec27a2bb70cbf5e")
                            : InstallMsu(root, "payload/kb/windows6.1-kb4490628-x86_3cdb3df55b9cd7ef7fcb24fc4e237ea287ad0992.msu",
                                "b86849fda570f906012992a033f6342373c00a3b3ec0089780eabb084a9a0182");
                    case "install_kb4474419":
                        return EnvDetect.Is64OS()
                            ? InstallMsu(root, "payload/kb/kb4474419-v3-x64.msu",
                                "99312df792b376f02e25607d2eb3355725c47d124d8da253193195515fe90213")
                            : InstallMsu(root, "payload/kb/windows6.1-kb4474419-v3-x86_0f687d50402790f340087c576886501b3223bec6.msu",
                                "8cf49fc7ac61e0b217859313a96337b149ab41b3307eb0d9529615142ea34c6c");
                    case "install_kb3020369":
                        // 2015-04 服务堆栈更新；仅 Win7 SP1 x64（本项目唯一支持目标）
                        if (!EnvDetect.Is64OS()) return "KB3020369 需要 64 位系统（本项目仅支持 Win7 SP1 x64）";
                        return InstallMsu(root, "payload/kb/windows6.1-kb3020369-x64.msu",
                            "e9fbb6b43c6fb9c1dcc8864bb7a35f29e4002830f3acbb93e0e717ae653fe76a");
                    case "install_kb3140245":
                        // KB3140245 的硬前置是 KB3020369（缺它 wusa 报 0x80242017）——此处自愈兜底。
                        // 只在**确认已装**时跳过；Absent（确实没装）与 Unavailable（WMI 查不出来）
                        // 都进入安装分支：重装已存在的 KB 是幂等安全的（wusa 返回 2359302 视为成功，
                        // 见 InterpretInstallExit），而漏装前置会真失败。所以这里保守装一下是对的。
                        // 注意不要写成 !HasHotfix(...)：新语义下 false 同时代表"没装"和"查不出来"，
                        // 虽然在本处结论相同，但显式三态更不容易被后人改错。
                        if (EnvDetect.QueryHotfix("KB3020369") != EnvDetect.HotfixState.Present && EnvDetect.Is64OS())
                        {
                            string ssuErr = InstallMsu(root, "payload/kb/windows6.1-kb3020369-x64.msu",
                                "e9fbb6b43c6fb9c1dcc8864bb7a35f29e4002830f3acbb93e0e717ae653fe76a");
                            if (ssuErr != null) return "前置 KB3020369 安装失败：" + ssuErr;
                        }
                        return EnvDetect.Is64OS()
                            ? InstallMsu(root, "payload/kb/kb3140245-x64.msu",
                                "bc35e63a125eb2dce45971bedc7ef8a78706b7e26286aed0cf664999ad833eda")
                            : InstallMsu(root, "payload/kb/windows6.1-kb3140245-x86_cdafb409afbe28db07e2254f40047774a0654f18.msu",
                                "0b08390e2f2b12e4e89af45829c0645081920fb93abe604d6d850a1f0a92be43");
                    case "install_kb2670838":
                        return EnvDetect.Is64OS()
                            ? InstallMsu(root, "payload/kb/windows6.1-kb2670838-x64.msu",
                                "9fe71e7dcd2280ce323880b075ade6e56c49b68fc702a9b4c0a635f0f1fb9db8")
                            : InstallMsu(root, "payload/kb/windows6.1-kb2670838-x86_984b8d122a688d917f81c04155225b3ef31f012e.msu",
                                "a43037dd15993273e6dd7398345e2bd0424225be81eb8acfaa1361442ef56fce");
                    case "install_netfx48":
                        return InstallExe(root, "payload/ndp48/ndp48-x86x64allos-enu.exe", "netfx48",
                            "0a3a390c47e639d0f7fc65b21195fee6b7f65b066f80f70c60fab191d14b7e40");
                    case "install_vcredist_x86":
                        return InstallExe(root, "payload/vcrt/VC_redist.x86.exe", "vcredist-x86",
                            "0c09f2611660441084ce0df425c51c11e147e6447963c3690f97e0b25c55ed64");
                    case "install_vcredist_x64":
                        return InstallExe(root, "payload/vcrt/VC_redist.x64.exe", "vcredist-x64",
                            "cc0ff0eb1dc3f5188ae6300faef32bf5beeba4bdd6e8e445a9184072096b713b");
                    case "unpack_py38":
                        return UnpackPy38(root);
                    case "unpack_lo76":
                        return UnpackLo76(root);
                    default:
                        return "未知修复动作 " + item.FixKey;
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        static void LogProxyWrite(string s)
        {
            Log.Line(s);
        }

        // ---------- msu / exe 安装器 ----------

        static string InstallMsu(string root, string rel, string sha256)
        {
            string src = PayloadPath(root, rel);
            string miss = CheckPayload(src, sha256);
            if (miss != null) return miss;
            Directory.CreateDirectory(ProcRunner.StageDir);
            File.Copy(src, ProcRunner.StageMsu, true);
            int code = ProcRunner.RunWusaMsuStage();
            return InterpretInstallExit(code);
        }

        static string InstallExe(string root, string rel, string whitelistId, string sha256)
        {
            string src = PayloadPath(root, rel);
            string miss = CheckPayload(src, sha256);
            if (miss != null) return miss;
            Directory.CreateDirectory(ProcRunner.StageDir);
            File.Copy(src, ProcRunner.StageExe, true);
            int code = ProcRunner.RunStageExe(whitelistId);
            return InterpretInstallExit(code);
        }

        static string PayloadPath(string root, string rel)
        {
            return Path.Combine(root, rel.Replace('/', Path.DirectorySeparatorChar));
        }

        static string CheckPayload(string src, string sha256)
        {
            if (!File.Exists(src))
            {
                return "离线载荷缺失: " + Path.GetFileName(src) + "（见 payload\\README.md；离线安装必须提供）";
            }
            if (sha256 != null && sha256.Length > 0 && !VerifySha256(src, sha256))
            {
                return "SHA256 校验失败: " + Path.GetFileName(src) + "（载荷被篡改或不完整，拒绝执行）";
            }
            return null;
        }

        static string InterpretInstallExit(int code)
        {
            if (code == 0) return null;
            if (code == 1641 || code == 3010) { Log.Line("安装完成 exit=" + code + "（需要重启后生效）"); return null; }
            if (code == 2359302) { Log.Line("已安装（重复安装被系统跳过）"); return null; }
            if (code == ProcRunner.ExitRejected) return "安装器被白名单拒绝或启动失败";
            if (code == ProcRunner.ExitTimeout) return "安装超时被终止";
            if (code == -2145124329) return "安装失败 0x80242017：需先装服务堆栈更新 KB3020369（正常会自动先装，请重试）";
            return "安装器退出码 " + code + "（0x" + (code < 0 ? (code + 4294967296L) : code).ToString("x8") + "）";
        }

        static bool VerifySha256(string file, string expected)
        {
            try
            {
                using (SHA256Managed sha = new SHA256Managed())
                using (FileStream fs = File.OpenRead(file))
                {
                    byte[] hash = sha.ComputeHash(fs);
                    StringBuilder sb = new StringBuilder(hash.Length * 2);
                    foreach (byte b in hash) sb.Append(b.ToString("x2"));
                    return sb.ToString().Equals(expected.ToLowerInvariant(), StringComparison.OrdinalIgnoreCase);
                }
            }
            catch { return false; }
        }

        // ---------- 免安装解压组件 ----------

        static string UnpackPy38(string root)
        {
            string arch = EnvDetect.Is64OS() ? "python-3.8.10-embed-amd64.zip" : "python-3.8.10-embed-win32.zip";
            string zip = PayloadPath(root, "payload/py38/" + arch);
            if (!File.Exists(zip)) return "离线载荷缺失: payload\\py38\\" + arch;
            string target = Path.Combine(Path.Combine(root, "runtime"), "py38");
            string err = Unzip.Extract(zip, target, 180000);
            if (err != null) return err;
            Flatten(target);   // embeddable zip 内含一层同名目录，需平铺
            PatchPth(target);  // 启用 import site（后续 site-packages/wheel 安装依赖）
            return null;
        }

        static string UnpackLo76(string root)
        {
            string dir = Path.Combine(Path.Combine(root, "payload"), "lo76");
            if (!Directory.Exists(dir)) return "离线载荷缺失: payload\\lo76（M1 组件，暂可跳过）";
            string[] zips = Directory.GetFiles(dir, "*.zip");
            if (zips.Length == 0) return "payload\\lo76 下没有 zip（M1 组件，暂可跳过）";
            string target = Path.Combine(root, "lo76");
            return Unzip.Extract(zips[0], target, 600000);
        }

        // zip 顶层若是单一子目录，则把其内容上移一层
        static void Flatten(string dir)
        {
            string py = FindFile(dir, "python.exe");
            if (py == null) return;
            string sub = Path.GetDirectoryName(py);
            if (string.Equals(sub, dir, StringComparison.OrdinalIgnoreCase)) return;
            if (sub == null || !sub.StartsWith(dir, StringComparison.OrdinalIgnoreCase)) return;
            MoveContents(sub, dir);
            try { Directory.Delete(sub, true); } catch { }
        }

        static string FindFile(string dir, string name)
        {
            try
            {
                string direct = Path.Combine(dir, name);
                if (File.Exists(direct)) return direct;
                foreach (string d in Directory.GetDirectories(dir))
                {
                    string r = FindFile(d, name);
                    if (r != null) return r;
                }
            }
            catch { }
            return null;
        }

        static void MoveContents(string fromDir, string toDir)
        {
            foreach (string f in Directory.GetFiles(fromDir))
            {
                string dst = Path.Combine(toDir, Path.GetFileName(f));
                if (File.Exists(dst)) File.Delete(dst);
                File.Move(f, dst);
            }
            foreach (string d in Directory.GetDirectories(fromDir))
            {
                string dst = Path.Combine(toDir, Path.GetFileName(d));
                if (Directory.Exists(dst))
                {
                    MoveContents(d, dst);
                }
                else
                {
                    Directory.Move(d, dst);
                }
            }
        }

        // embeddable python 的 ._pth 默认禁用 site；启用以便后续 pip/site-packages 生效
        static void PatchPth(string target)
        {
            try
            {
                foreach (string f in Directory.GetFiles(target, "*._pth"))
                {
                    string[] lines = File.ReadAllLines(f);
                    bool changed = false;
                    for (int i = 0; i < lines.Length; i++)
                    {
                        string t = lines[i].TrimStart();
                        if (t.StartsWith("#import site"))
                        {
                            lines[i] = t.Remove(0, 1);
                            changed = true;
                        }
                    }
                    if (changed) File.WriteAllLines(f, lines);
                }
            }
            catch { }
        }
    }
}
