# run-capture.ps1 -- 启动子进程并把 stdout/stderr 抓到文件，规避 PowerShell 的
# Start-Process -RedirectStandardOutput 在"大小写重复环境变量"机器上的崩溃。
#
# 背景（实测确认的 Windows 环境缺陷）：
#   Windows 允许同名不同大小写的环境变量共存，代理软件会同时写入
#   HTTP_PROXY/http_proxy、HTTPS_PROXY/https_proxy、NO_PROXY/no_proxy。
#   PowerShell 的 Start-Process 一旦带 -RedirectStandardOutput / -RedirectStandardError，
#   内部就会触碰 ProcessStartInfo.EnvironmentVariables（大小写不敏感字典），
#   在建环境块时抛：
#       ArgumentException: Item has already been added.
#       Key in dictionary: 'NO_PROXY'  Key being added: 'no_proxy'
#   于是整个 Start-Process 调用失败（**不带重定向时反而正常**，因为那条路径不建环境块）。
#
# 本脚本改用 .NET Process 直接启动，并在启动前用反射把那个会抛的字典换成一个
# 干净的（先灌入当前进程环境、同名仅大小写不同只留一个），因此不受该缺陷影响。
#
# 用法（结果写入 -OutFile，退出码随进程返回）：
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\e2e\run-capture.ps1 `
#       -Exe <可执行文件> -ArgLine "<参数1>|<参数2>|..." -OutFile <输出文件>
#   注意：参数用单个字符串 + "|" 分隔，**不要**用 PowerShell 数组传参——
#   从外层 PS 用 -File 传数组时会被编组成一个字符串，导致参数绑定失败（实测踩到）。
# 需要中文时本文件必须带 UTF-8 BOM（PS 5.1 否则按 GBK 解析）。

param(
  [Parameter(Mandatory=$true)][string]$Exe,
  [string]$ArgLine = "",
  [Parameter(Mandatory=$true)][string]$OutFile,
  [int]$TimeoutSec = 300
)

$ErrorActionPreference = "Continue"

$ArgList = @()
if ($ArgLine -ne "") { $ArgList = $ArgLine.Split('|') }

# 把参数安全地拼成命令行：逐项加引号并转义内部引号（不做 shell 解析，仅给 CreateProcess）
function Quote-Arg([string]$a) {
  if ($null -eq $a) { return '""' }
  if ($a -match '[\s"]') {
    return '"' + ($a -replace '(\\*)"', '$1$1\"' -replace '(\\+)$', '$1$1') + '"'
  }
  return $a
}

Add-Type -TypeDefinition @"
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Text;

public class SafeRun {
    // 判断当前进程环境是否存在"仅大小写不同"的重复键
    static bool HasDup() {
        try {
            IDictionary raw = Environment.GetEnvironmentVariables();
            var seen = new List<string>();
            foreach (object k in raw.Keys) {
                string key = k as string;
                if (key == null) continue;
                foreach (string s in seen) {
                    if (string.Equals(s, key, StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(s, key, StringComparison.Ordinal)) return true;
                }
                seen.Add(key);
            }
        } catch { }
        return false;
    }

    // 用反射把会抛异常的 environmentVariables 字段换成一个干净字典
    static void PatchEnv(ProcessStartInfo si) {
        try {
            Type t = si.GetType();
            FieldInfo f = t.GetField("environmentVariables", BindingFlags.Instance | BindingFlags.NonPublic);
            if (f == null) return;
            object sd = null;
            try { sd = f.GetValue(si); } catch { sd = null; }
            if (sd == null) {
                try { sd = Activator.CreateInstance(f.FieldType); } catch { return; }
                PropertyInfo idx = sd.GetType().GetProperty("Item", new Type[] { typeof(string) });
                try {
                    IDictionary src = Environment.GetEnvironmentVariables();
                    var added = new List<string>();
                    foreach (object k in src.Keys) {
                        string key = k as string;
                        if (key == null) continue;
                        bool dup = false;
                        foreach (string s in added) {
                            if (string.Equals(s, key, StringComparison.OrdinalIgnoreCase)) { dup = true; break; }
                        }
                        if (dup) continue;
                        if (idx != null && idx.CanWrite) idx.SetValue(sd, src[k], new object[] { key });
                        added.Add(key);
                    }
                } catch { }
                try { f.SetValue(si, sd); } catch { }
            }
        } catch { }
    }

    // 返回 exit code；输出（stdout+stderr）写入 outFile（UTF-8 无 BOM）
    public static int Run(string exe, string args, string outFile, int timeoutSec) {
        try {
            var si = new ProcessStartInfo(exe, args);
            si.UseShellExecute = false;
            si.CreateNoWindow = true;
            si.RedirectStandardOutput = true;
            si.RedirectStandardError = true;
            si.StandardOutputEncoding = Encoding.UTF8;
            si.StandardErrorEncoding = Encoding.UTF8;
            if (HasDup()) PatchEnv(si);
            using (Process p = Process.Start(si)) {
                StringBuilder so = new StringBuilder(), se = new StringBuilder();
                p.OutputDataReceived += delegate(object s, DataReceivedEventArgs e) { if (e.Data != null) so.AppendLine(e.Data); };
                p.ErrorDataReceived += delegate(object s, DataReceivedEventArgs e) { if (e.Data != null) se.AppendLine(e.Data); };
                p.BeginOutputReadLine();
                p.BeginErrorReadLine();
                if (!p.WaitForExit(timeoutSec * 1000)) {
                    try { p.Kill(); } catch { }
                    p.WaitForExit();
                    System.IO.File.WriteAllText(outFile, so.ToString() + se.ToString(), new UTF8Encoding(false));
                    return -999;
                }
                p.WaitForExit();
                System.IO.File.WriteAllText(outFile, so.ToString() + se.ToString(), new UTF8Encoding(false));
                return p.ExitCode;
            }
        } catch (Exception ex) {
            try { System.IO.File.WriteAllText(outFile, "run-capture 启动失败: " + ex.Message, new UTF8Encoding(false)); } catch { }
            return -998;
        }
    }
}
"@

$argsStr = ($ArgList | ForEach-Object { Quote-Arg $_ }) -join ' '
$rc = [SafeRun]::Run($Exe, $argsStr, $OutFile, $TimeoutSec)
exit $rc
