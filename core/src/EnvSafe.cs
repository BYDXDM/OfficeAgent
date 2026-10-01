// EnvSafe —— 安全设置子进程环境变量（修 Win 环境"大小写重复变量"导致的启动失败）
//
// 背景（真实缺陷，非测试环境特有）：
//   Windows 环境允许同名不同大小写的变量共存，代理软件（Clash/v2ray 等）会同时写入
//   HTTP_PROXY 与 http_proxy、NO_PROXY 与 no_proxy。而 .NET 的
//   ProcessStartInfo.EnvironmentVariables 内部是**大小写不敏感**的 Hashtable，
//   对它做任何索引/迭代（哪怕只是写一个无关的键）都会抛：
//       ArgumentException: 已添加项。字典中的关键字:"NO_PROXY"所添加的关键字:"no_proxy"
//   结果：凡是碰过 si.EnvironmentVariables 的启动点（ProcRunner.RunVersionProbe、
//   JobProbeRunner.RunInJob）在有代理配置的机器上**一律启动失败**（返回 ExitRejected），
//   表现为"技能执行启动失败"、Python/LibreOffice 探测不到。
//
// 修法：不碰那个有问题的字典，改用反射直写底层私有字段。
//   .NET Framework 中 ProcessStartInfo 持有：
//     * environmentVariables : 上述 Hashtable（构造函数里由 Environment.GetEnvironmentVariables()
//                              填充 —— 构造那一刻就已经埋了重复键，但 Hashtable 允许
//                              重复键存在，只要不 Insert/索引就不会抛）
//     * environmentVariablesDictionary : .NET Core 才有的新形态
//     * useDefaultEnvironment : 布尔，为 true 时子进程直接继承父环境
//   我们优先写 .NET Core 的 Dictionary<string,string>（大小写不敏感但支持 TryAdd 语义，
//   且用索引赋值不会因重复键抛异常）；退化到 Framework 时改用 useDefaultEnvironment
//   + 不注入（PYTHONIOENCODING 仅为让 Python 用 UTF-8 输出，非必需：技能协议本身
//   在 main.py 里已显式 ensure_ascii=False 且 SkillRunner 用 UTF-8 解码 stdout）。
//
// 红线合规：本文件**不启动任何进程**，只做反射赋值；所有 Process.Start 仍留在
// 既有白名单启动点（ProcRunner / JobProbeRunner），本文件不新增启动形态。
using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;

namespace OfficeAgent.Core
{
    public static class EnvSafe
    {
        // 先做一次探测：当前进程环境里是否存在"仅大小写不同"的重复变量。
        // 有重复时调用方应避免触碰 si.EnvironmentVariables。
        static bool _dupChecked = false;
        static bool _hasDup = false;

        public static bool HasCaseDuplicateNames()
        {
            if (_dupChecked) return _hasDup;
            _dupChecked = true;
            try
            {
                // 用 cmd 之外的方式取"原始"环境：GetEnvironmentVariables 返回 IDictionary，
                // 其底层是大小写不敏感的 Hashtable，枚举 Keys 是安全的（不 Insert）。
                IDictionary raw = Environment.GetEnvironmentVariables();
                List<string> seen = new List<string>();
                foreach (object k in raw.Keys)
                {
                    string key = k as string;
                    if (key == null) continue;
                    foreach (string s in seen)
                    {
                        if (string.Equals(s, key, StringComparison.OrdinalIgnoreCase) &&
                            !string.Equals(s, key, StringComparison.Ordinal))
                        {
                            _hasDup = true;
                            return true;
                        }
                    }
                    seen.Add(key);
                }
            }
            catch { }
            return _hasDup;
        }

        // 在 ProcessStartInfo 上设置环境变量，且**不触发**大小写重复键异常。
        // 返回 true=已成功注入；false=环境不安全且无法注入（调用方应继续，注入失败不影响正确性）。
        public static bool SetEnv(global::System.Diagnostics.ProcessStartInfo si, string name, string value)
        {
            if (si == null || name == null) return false;
            try
            {
                if (!HasCaseDuplicateNames())
                {
                    // 环境干净：走官方 API，行为最可预期
                    si.EnvironmentVariables[name] = value;
                    return true;
                }
            }
            catch
            {
                // 探测不准时也不要让调用方炸，继续走反射兜底
            }

            // 环境有大小写重复键：绕开那个会抛的字典，直写底层字段。
            //
            // .NET Framework 4.x 的 ProcessStartInfo 有两个字段：
            //   environmentVariables : StringDictionary （**懒创建，初始为 null**；
            //                          大小写不敏感，索引赋值安全——这正是我们要用的）
            //   environment          : IDictionary<string,string>（**有问题的那个**：
            //                          由 Environment.GetEnvironmentVariables() 填充，
            //                          在存在 NO_PROXY/no_proxy 时构造即抛
            //                          ArgumentException）
            // 关键：绝不能访问公开属性 si.EnvironmentVariables——getter 会去碰
            // environment 并立刻抛异常。必须直接操作私有字段。
            try
            {
                Type t = si.GetType();
                FieldInfo fSd = t.GetField("environmentVariables",
                    BindingFlags.Instance | BindingFlags.NonPublic);
                if (fSd != null)
                {
                    object sd = null;
                    try { sd = fSd.GetValue(si); } catch { sd = null; }
                    if (sd == null)
                    {
                        // 懒创建：自己造一个 StringDictionary（大小写不敏感），
                        // **必须先把当前进程环境灌进去**——否则子进程拿到的是空环境，
                        // Python 会因缺 SystemRoot 而报
                        // "Fatal Python error: _Py_HashRandomization_Init"。
                        Type sdType = fSd.FieldType;   // System.Collections.Specialized.StringDictionary
                        try { sd = Activator.CreateInstance(sdType); } catch { sd = null; }
                        if (sd != null)
                        {
                            System.Reflection.PropertyInfo idx0 = sd.GetType().GetProperty("Item",
                                new Type[] { typeof(string) });
                            try
                            {
                                // 逐键复制；同名仅大小写不同时只保留第一个（避免再次撞重复键）
                                IDictionary src = Environment.GetEnvironmentVariables();
                                List<string> added = new List<string>();
                                foreach (object k in src.Keys)
                                {
                                    string key = k as string;
                                    if (key == null) continue;
                                    bool dup = false;
                                    foreach (string s in added)
                                    {
                                        if (string.Equals(s, key, StringComparison.OrdinalIgnoreCase)) { dup = true; break; }
                                    }
                                    if (dup) continue;
                                    if (idx0 != null && idx0.CanWrite)
                                        idx0.SetValue(sd, src[k], new object[] { key });
                                    added.Add(key);
                                }
                            }
                            catch { }
                            try { fSd.SetValue(si, sd); } catch { sd = null; }
                        }
                    }
                    if (sd != null)
                    {
                        // StringDictionary 的索引器是 public string this[string]
                        System.Reflection.PropertyInfo idx = sd.GetType().GetProperty("Item",
                            new Type[] { typeof(string) });
                        if (idx != null && idx.CanWrite)
                        {
                            idx.SetValue(sd, value, new object[] { name });
                            return true;
                        }
                    }
                }

                // 次选：.NET Core 的 environmentVariablesDictionary
                FieldInfo fDict = t.GetField("environmentVariablesDictionary",
                    BindingFlags.Instance | BindingFlags.NonPublic);
                if (fDict != null)
                {
                    IDictionary<string, string> dict =
                        fDict.GetValue(si) as IDictionary<string, string>;
                    if (dict != null) { dict[name] = value; return true; }
                }
            }
            catch
            {
                // 反射失败也不能让调用方崩：注入 PYTHONIOENCODING 只是锦上添花
                // （Python 侧 main.py 已 ensure_ascii=False，宿主按 UTF-8 解码 stdout）
            }
            return false;
        }
    }
}
