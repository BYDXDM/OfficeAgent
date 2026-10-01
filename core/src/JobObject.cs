// JobObject —— Win32 Job Object 封装（设计方案 §4.5 运行隔离）
// 工具/技能进程一律入 Job：内存上限（默认 1.5GB/进程）+ KILL_ON_JOB_CLOSE + 超时整树击杀。
// 仅用于不可信输入处理器（soffice、后续 python sidecar）；系统安装器（wusa/离线包）不进 Job。
// 平台 x86：结构体按 32 位对齐（Sequential 默认即可）。
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace OfficeAgent.Core
{
    public static class JobObject
    {
        const int JobObjectExtendedLimitInformation = 9;
        const uint JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE = 0x2000;
        const uint JOB_OBJECT_LIMIT_PROCESS_MEMORY = 0x100;

        [StructLayout(LayoutKind.Sequential)]
        struct IO_COUNTERS
        {
            public ulong ReadOperationCount, WriteOperationCount, OtherOperationCount;
            public ulong ReadTransferCount, WriteTransferCount, OtherTransferCount;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct JOBOBJECT_BASIC_LIMIT_INFORMATION
        {
            public long PerProcessUserTimeLimit;
            public long PerJobUserTimeLimit;
            public uint LimitFlags;
            public UIntPtr MinimumWorkingSetSize;
            public UIntPtr MaximumWorkingSetSize;
            public uint ActiveProcessLimit;
            public UIntPtr Affinity;
            public uint PriorityClass;
            public uint SchedulingClass;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct JOBOBJECT_EXTENDED_LIMIT_INFORMATION
        {
            public JOBOBJECT_BASIC_LIMIT_INFORMATION BasicLimitInformation;
            public IO_COUNTERS IoInfo;
            public UIntPtr ProcessMemoryLimit;
            public UIntPtr JobMemoryLimit;
            public UIntPtr PeakProcessMemoryUsed;
            public UIntPtr PeakJobMemoryUsed;
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern IntPtr CreateJobObject(IntPtr attrs, string name);
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool SetInformationJobObject(IntPtr job, int infoClass,
            ref JOBOBJECT_EXTENDED_LIMIT_INFORMATION info, int len);
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);
        [DllImport("kernel32.dll")]
        static extern bool TerminateJobObject(IntPtr job, uint exitCode);
        [DllImport("kernel32.dll")]
        static extern bool CloseHandle(IntPtr h);

        public const ulong DefaultMemoryLimit = 1536UL * 1024 * 1024;   // 1.5GB（设计方案 §4.5）

        // 创建带内存上限 + 句柄关闭即杀的 Job；失败返回 IntPtr.Zero（调用方降级为无 Job 运行）
        public static IntPtr Create(ulong memoryLimitBytes)
        {
            try
            {
                IntPtr job = CreateJobObject(IntPtr.Zero, null);
                if (job == IntPtr.Zero) return IntPtr.Zero;
                JOBOBJECT_EXTENDED_LIMIT_INFORMATION info = new JOBOBJECT_EXTENDED_LIMIT_INFORMATION();
                info.BasicLimitInformation.LimitFlags =
                    JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE | JOB_OBJECT_LIMIT_PROCESS_MEMORY;
                info.ProcessMemoryLimit = new UIntPtr(memoryLimitBytes);
                if (!SetInformationJobObject(job, JobObjectExtendedLimitInformation, ref info,
                    Marshal.SizeOf(typeof(JOBOBJECT_EXTENDED_LIMIT_INFORMATION))))
                {
                    CloseHandle(job);
                    return IntPtr.Zero;
                }
                return job;
            }
            catch { return IntPtr.Zero; }
        }

        public static bool Assign(IntPtr job, Process p)
        {
            if (job == IntPtr.Zero || p == null) return false;
            try { return AssignProcessToJobObject(job, p.Handle); }
            catch { return false; }
        }

        // 整树击杀（Job 内所有进程，含 soffice.bin 等子进程）
        public static void KillTree(IntPtr job)
        {
            if (job == IntPtr.Zero) return;
            try { TerminateJobObject(job, 1); } catch { }
        }

        public static void Close(IntPtr job)
        {
            if (job == IntPtr.Zero) return;
            try { CloseHandle(job); } catch { }
        }
    }
}
