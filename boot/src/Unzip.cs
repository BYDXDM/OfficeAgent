// ZIP 解压器（.NET 3.5 原生实现：中央目录解析 + DeflateStream + CRC32 校验）
// 不依赖 Shell COM（其 CopyHere 在非交互会话下会静默失败）、不依赖外部 exe（tar/unzip）。
// 支持方法 0(store)/8(deflate)、UTF-8 文件名标志位；含 zip-slip 防护与 CRC 完整性校验。
// 局限：不支持 Zip64（>4GB 单包/65535 条目；python/LO 便携包远小于此，M1 如需再扩展）。
using System;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace OfficeAgent.Boot
{
    public static class Unzip
    {
        const uint EOCD_SIGNATURE = 0x06054b50;
        const uint CDH_SIGNATURE = 0x02014b50;
        const uint LFH_SIGNATURE = 0x04034b50;
        const long MAX_UNCOMPRESSED = 1024L * 1024 * 1024; // 单条目上限 1GB，防炸弹

        // 返回 null=成功，否则错误信息
        public static string Extract(string zipPath, string destDir, int timeoutMs)
        {
            try
            {
                if (!File.Exists(zipPath)) return "压缩包不存在: " + zipPath;
                if (!Directory.Exists(destDir)) Directory.CreateDirectory(destDir);

                using (FileStream fs = File.OpenRead(zipPath))
                {
                    byte[] eocd = FindEocd(fs);
                    if (eocd == null) return "无效的 ZIP（未找到中央目录结尾记录）";
                    int entryCount = ReadU16(eocd, 10);
                    uint cdOffset = ReadU32(eocd, 16);
                    int extracted = 0;

                    fs.Seek(cdOffset, SeekOrigin.Begin);
                    for (int i = 0; i < entryCount; i++)
                    {
                        long entryStart = fs.Position;
                        byte[] cdh = ReadExact(fs, 46);
                        if (ReadU32(cdh, 0) != CDH_SIGNATURE) return "中央目录损坏（条目 " + i + "）";
                        ushort flags = ReadU16(cdh, 8);
                        ushort method = ReadU16(cdh, 10);
                        uint crcExpected = ReadU32(cdh, 16);
                        uint compSize = ReadU32(cdh, 20);
                        uint uncompSize = ReadU32(cdh, 24);
                        ushort nameLen = ReadU16(cdh, 28);
                        ushort extraLen = ReadU16(cdh, 30);
                        ushort commentLen = ReadU16(cdh, 32);
                        uint localOffset = ReadU32(cdh, 42);
                        long nextEntry = entryStart + 46 + nameLen + extraLen + commentLen;

                        byte[] nameBytes = ReadExact(fs, nameLen);
                        string name = DecodeName(nameBytes, flags);
                        ReadExact(fs, extraLen);
                        ReadExact(fs, commentLen);

                        if (name.Length > 0 && !name.EndsWith("/"))
                        {
                            string target = SafeJoin(destDir, name);
                            if (target == null) return "拒绝不安全路径: " + name;
                            if (uncompSize > MAX_UNCOMPRESSED) return "条目过大，疑似 ZIP 炸弹: " + name;

                            byte[] data = ReadEntryData(fs, localOffset, method, compSize, uncompSize);
                            if (data == null) return "条目解压失败: " + name;
                            uint crcActual = Crc32(data, 0, data.Length);
                            if (crcActual != crcExpected) return "CRC 校验失败: " + name;

                            string parent = Path.GetDirectoryName(target);
                            if (parent != null && parent.Length > destDir.Length && !Directory.Exists(parent)) Directory.CreateDirectory(parent);
                            File.WriteAllBytes(target, data);
                            extracted++;
                        }
                        else if (name.Length > 0)
                        {
                            string dir = SafeJoin(destDir, name);
                            if (dir == null) return "拒绝不安全路径: " + name;
                            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                        }
                        // 条目数据读取会移动流位置，必须回位到中央目录的下一记录
                        fs.Seek(nextEntry, SeekOrigin.Begin);
                    }
                    if (extracted <= 0) return "未提取到任何文件";
                    return null;
                }
            }
            catch (Exception ex) { return ex.Message; }
        }

        // ---------- 解析辅助 ----------

        static byte[] FindEocd(FileStream fs)
        {
            long len = fs.Length;
            int window = (int)Math.Min(len, 66000);
            fs.Seek(len - window, SeekOrigin.Begin);
            byte[] buf = ReadExact(fs, window);
            for (int i = buf.Length - 22; i >= 0; i--)
            {
                if (ReadU32(buf, i) == EOCD_SIGNATURE)
                {
                    // 只返回 22 字节 EOCD 记录本身；后续字段按记录内偏移解析
                    byte[] rec = new byte[22];
                    Array.Copy(buf, i, rec, 0, 22);
                    return rec;
                }
            }
            return null;
        }

        static byte[] ReadEntryData(FileStream fs, uint localOffset, ushort method, uint compSize, uint uncompSize)
        {
            fs.Seek(localOffset, SeekOrigin.Begin);
            byte[] lfh = ReadExact(fs, 30);
            if (ReadU32(lfh, 0) != LFH_SIGNATURE) return null;
            ushort nameLen = ReadU16(lfh, 26);
            ushort extraLen = ReadU16(lfh, 28);
            fs.Seek(nameLen + extraLen, SeekOrigin.Current);
            byte[] comp = ReadExact(fs, (int)compSize);

            if (method == 0)
            {
                if (comp.Length != (int)uncompSize) return null;
                return comp;
            }
            if (method != 8) return null; // 仅支持 store/deflate

            using (MemoryStream src = new MemoryStream(comp))
            using (DeflateStream ds = new DeflateStream(src, CompressionMode.Decompress))
            {
                byte[] outBuf = new byte[uncompSize];
                int total = 0;
                while (total < (int)uncompSize)
                {
                    int n = ds.Read(outBuf, total, (int)uncompSize - total);
                    if (n <= 0) break;
                    total += n;
                }
                if (total != (int)uncompSize) return null;
                return outBuf;
            }
        }

        static string DecodeName(byte[] bytes, ushort flags)
        {
            if ((flags & 0x800) != 0) return Encoding.UTF8.GetString(bytes);
            // 非标志位时按系统 ANSI 代码页（Win7 中文=GBK）；名字为 ASCII 时与 UTF-8 等价
            return Encoding.Default.GetString(bytes);
        }

        // zip-slip 防护：拒绝绝对路径、盘符、.. 上跳；返回限定在 destDir 内的目标路径
        static string SafeJoin(string destDir, string entryName)
        {
            string norm = entryName.Replace('/', Path.DirectorySeparatorChar);
            if (norm.IndexOf(":") >= 0) return null;
            if (norm.StartsWith("\\")) return null;
            string[] parts = norm.Split(Path.DirectorySeparatorChar);
            StringBuilder sb = new StringBuilder(destDir);
            foreach (string p in parts)
            {
                if (p.Length == 0 || p == ".") continue;
                if (p == "..") return null;
                sb.Append(Path.DirectorySeparatorChar);
                sb.Append(p);
            }
            string result = sb.ToString();
            if (!result.StartsWith(destDir, StringComparison.OrdinalIgnoreCase)) return null;
            return result;
        }

        static byte[] ReadExact(FileStream fs, int count)
        {
            byte[] buf = new byte[count];
            int off = 0;
            while (off < count)
            {
                int n = fs.Read(buf, off, count - off);
                if (n <= 0) throw new IOException("文件意外结束");
                off += n;
            }
            return buf;
        }

        static ushort ReadU16(byte[] b, int off) { return (ushort)(b[off] | (b[off + 1] << 8)); }
        static uint ReadU32(byte[] b, int off) { return (uint)(b[off] | (b[off + 1] << 8) | (b[off + 2] << 16) | (b[off + 3] << 24)); }

        // ---------- CRC32 ----------

        static uint[] crcTable;

        static uint[] CrcTable()
        {
            if (crcTable == null)
            {
                uint[] t = new uint[256];
                for (uint i = 0; i < 256; i++)
                {
                    uint c = i;
                    for (int k = 0; k < 8; k++)
                    {
                        if ((c & 1) != 0) c = 0xEDB88320u ^ (c >> 1);
                        else c = c >> 1;
                    }
                    t[i] = c;
                }
                crcTable = t;
            }
            return crcTable;
        }

        static uint Crc32(byte[] data, int offset, int count)
        {
            uint[] t = CrcTable();
            uint crc = 0xFFFFFFFFu;
            for (int i = offset; i < offset + count; i++)
            {
                crc = t[(crc ^ data[i]) & 0xFF] ^ (crc >> 8);
            }
            return crc ^ 0xFFFFFFFFu;
        }
    }
}
