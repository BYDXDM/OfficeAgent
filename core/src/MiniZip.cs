// MiniZip —— 纯 .NET 3.5 的 ZIP 读/写实现（替代 System.IO.Compression.ZipArchive）
// 目的：让 host 全源码可在 csc 3.5 下编译出降级壳（设计方案 §8.4：.NET 4.8 装不上时
// 转 3.5 降级壳，核对/转换核心可用）。读侧逻辑与 boot\src\Unzip.cs 同源；
// 写侧生成 deflate 条目（xlsx OOXML 包，Excel/LO 均接受）。
// 红线：C# 3.0 语法；不支持 Zip64（xlsx 条目远小于上限）；名字匹配大小写不敏感（对齐 ZipArchive）。
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace OfficeAgent.Core
{
    // 只暴露底层流的一段（[start, start+length)）的只读包装。
    // 用于把 DeflateStream 限制在**单个条目**的压缩数据范围内：
    // 若不限制，DeflateStream 会一直读到文件尾，把压缩包里后续条目也当数据解压，结果错乱。
    internal class MiniZipBoundedStream : Stream
    {
        readonly Stream inner;
        readonly long length;
        long pos;

        public MiniZipBoundedStream(Stream inner, long length, long ignored)
        {
            this.inner = inner;
            this.length = length;
        }

        public override bool CanRead { get { return true; } }
        public override bool CanSeek { get { return false; } }
        public override bool CanWrite { get { return false; } }
        public override long Length { get { return length; } }
        public override long Position
        {
            get { return pos; }
            set { throw new NotSupportedException(); }
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            if (pos >= length) return 0;
            long remain = length - pos;
            if (count > remain) count = (int)remain;
            int n = inner.Read(buffer, offset, count);
            if (n > 0) pos += n;
            return n;
        }

        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) { throw new NotSupportedException(); }
        public override void SetLength(long value) { throw new NotSupportedException(); }
        public override void Write(byte[] buffer, int offset, int count) { throw new NotSupportedException(); }

        protected override void Dispose(bool disposing)
        {
            // 关闭有界流 = 关闭底层文件句柄（这个流是句柄的唯一持有者）
            if (disposing) { try { inner.Close(); } catch { } }
            base.Dispose(disposing);
        }
    }

    // 校验"解压出来的字节数"与中央目录声明的 UncompressedSize 一致。
    // 为什么值得单独做一层：流式读取放弃了 CRC32 校验（要校验就得缓存全部数据，
    // 那正是要避免的）。而 ZIP 截断/损坏最典型的症状就是**解压字节数不足**，
    // 这一层能把它抓成明确异常，而不是让上层 XmlReader 报一句看不懂的 XML 错误。
    internal class MiniZipVerifyingStream : Stream
    {
        readonly Stream inner;
        readonly long expected;
        long seen;

        public MiniZipVerifyingStream(Stream inner, long expected)
        {
            this.inner = inner;
            this.expected = expected;
        }

        public override bool CanRead { get { return true; } }
        public override bool CanSeek { get { return false; } }
        public override bool CanWrite { get { return false; } }
        public override long Length { get { return expected; } }
        public override long Position
        {
            get { return seen; }
            set { throw new NotSupportedException(); }
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            int remain = expected - seen > int.MaxValue ? int.MaxValue : (int)(expected - seen);
            if (remain <= 0) return 0;
            if (count > remain) count = remain;
            int n = inner.Read(buffer, offset, count);
            if (n > 0) seen += n;
            return n;
        }

        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) { throw new NotSupportedException(); }
        public override void SetLength(long value) { throw new NotSupportedException(); }
        public override void Write(byte[] buffer, int offset, int count) { throw new NotSupportedException(); }

        protected override void Dispose(bool disposing)
        {
            bool truncated = false;
            try
            {
                if (disposing && seen < expected) truncated = true;
            }
            catch { }
            if (disposing)
            {
                // 先排空剩余字节：DeflateStream 只有在读到末尾时才会发现数据损坏，
                // 提前关闭会把"损坏"静默吞掉（表现为表格少了最后几行而不是报错）。
                try
                {
                    byte[] sink = new byte[8192];
                    int guard = 0;
                    while (seen < expected && guard++ < 1000000)
                    {
                        int n = Read(sink, 0, sink.Length);
                        if (n <= 0) break;
                    }
                }
                catch { }
                if (seen < expected) truncated = true;
                try { inner.Close(); } catch { }
            }
            base.Dispose(disposing);
            if (disposing && truncated)
                throw new IOException("条目解压长度不足（数据可能损坏或被截断）");
        }
    }

    public class MiniZipEntryInfo
    {
        public string FullName = "";
        public long UncompressedSize;
        internal uint LocalOffset;
        internal ushort Method;
        internal uint CompressedSize;
        internal uint Crc;
        internal ushort Flags;
    }

    public class MiniZipFile : IDisposable
    {
        string path;
        List<MiniZipEntryInfo> entries = new List<MiniZipEntryInfo>();
        Dictionary<string, MiniZipEntryInfo> byName = new Dictionary<string, MiniZipEntryInfo>(StringComparer.OrdinalIgnoreCase);

        MiniZipFile(string zipPath) { path = zipPath; }

        public static MiniZipFile OpenRead(string zipPath)
        {
            MiniZipFile z = new MiniZipFile(zipPath);
            using (FileStream fs = File.OpenRead(zipPath))
            {
                byte[] eocd = MiniZip.FindEocd(fs);
                if (eocd == null) throw new IOException("无效的 ZIP（未找到中央目录结尾记录）: " + zipPath);
                int entryCount = MiniZip.ReadU16(eocd, 10);
                uint cdOffset = MiniZip.ReadU32(eocd, 16);
                fs.Seek(cdOffset, SeekOrigin.Begin);
                for (int i = 0; i < entryCount; i++)
                {
                    long entryStart = fs.Position;
                    byte[] cdh = MiniZip.ReadExact(fs, 46);
                    if (MiniZip.ReadU32(cdh, 0) != MiniZip.CDH_SIGNATURE) throw new IOException("中央目录损坏（条目 " + i + "）");
                    MiniZipEntryInfo e = new MiniZipEntryInfo();
                    e.Flags = MiniZip.ReadU16(cdh, 8);
                    e.Method = MiniZip.ReadU16(cdh, 10);
                    e.Crc = MiniZip.ReadU32(cdh, 16);
                    e.CompressedSize = MiniZip.ReadU32(cdh, 20);
                    e.UncompressedSize = MiniZip.ReadU32(cdh, 24);
                    ushort nameLen = MiniZip.ReadU16(cdh, 28);
                    ushort extraLen = MiniZip.ReadU16(cdh, 30);
                    ushort commentLen = MiniZip.ReadU16(cdh, 32);
                    e.LocalOffset = MiniZip.ReadU32(cdh, 42);
                    byte[] nameBytes = MiniZip.ReadExact(fs, nameLen);
                    e.FullName = MiniZip.DecodeName(nameBytes, e.Flags);
                    MiniZip.ReadExact(fs, extraLen);
                    MiniZip.ReadExact(fs, commentLen);
                    long nextEntry = entryStart + 46 + nameLen + extraLen + commentLen;
                    if (e.FullName.Length > 0 && !e.FullName.EndsWith("/"))
                    {
                        z.entries.Add(e);
                        z.byName[e.FullName] = e;
                    }
                    fs.Seek(nextEntry, SeekOrigin.Begin);
                }
            }
            return z;
        }

        // 大小写不敏感查找；不存在返回 null（与原 ZipArchive 遍历语义一致）
        public MiniZipEntryInfo FindEntry(string name)
        {
            MiniZipEntryInfo e;
            return byName.TryGetValue(name, out e) ? e : null;
        }

        // 全部条目（模板替换式生成需要枚举包内所有部件）
        public List<MiniZipEntryInfo> Entries { get { return entries; } }

        // 解压条目内容（CRC 校验后以内存流返回；调用方负责 Dispose）
        //
        // ⚠️ 本方法会把**整个条目解压进内存**（一次 `new byte[uncompSize]`）。
        //    对小条目（sharedStrings/styles/workbook.xml）没问题，
        //    但对工作表 XML 是致命的：实测一张 20000x128 的表，
        //    `xl/worksheets/sheet1.xml` 压缩后 9.7MB、**解压后 78MB**，
        //    这一处分配就占了"打开大表"内存峰值的绝大部分。
        //    需要按行流式读取工作表时请改用 OpenEntryStream。
        public Stream OpenEntry(string name)
        {
            MiniZipEntryInfo e = FindEntry(name);
            if (e == null) throw new IOException("xlsx 包内缺少条目: " + name);
            using (FileStream fs = File.OpenRead(path))
            {
                byte[] data = MiniZip.ReadEntryData(fs, e.LocalOffset, e.Method, e.CompressedSize, e.UncompressedSize, e.FullName);
                if (data == null) throw new IOException("条目解压失败: " + e.FullName);
                if (MiniZip.Crc32(data, 0, data.Length) != e.Crc) throw new IOException("CRC 校验失败: " + e.FullName);
                return new MemoryStream(data);
            }
        }

        // 流式读取条目：不全量解压进内存，边解压边交给调用方消费。
        //
        // 用途：工作表 XML 只需要**顺序**读一遍（XmlReader 本身就是前向只读），
        // 把 78MB 的中间 byte[] 省掉，内存占用与条目大小解耦。
        //
        // ★ 关于 CRC：本方法**不**校验 CRC。CRC32 在 ZIP 尾部，要校验就得先缓存全部数据，
        //   那正是这里要避免的。完整性由三层保障兜底：
        //     ① DeflateStream 解压时自带结构校验，坏数据会抛异常；
        //     ② 解压长度必须等于中央目录声明的 uncompSize，少一个字节即报错；
        //     ③ 上层用 XmlReader 解析，截断/损坏的 XML 会解析失败而不是静默给出错数据。
        //   对"读一个表格给自己看"的用途，这三层足够；需要强校验的场景仍走 OpenEntry。
        //
        // 返回值：调用方 Dispose 时一并关闭文件句柄与解压流。
        public Stream OpenEntryStream(string name)
        {
            MiniZipEntryInfo e = FindEntry(name);
            if (e == null) throw new IOException("xlsx 包内缺少条目: " + name);
            if (e.UncompressedSize > MiniZip.MAX_UNCOMPRESSED)
                throw new IOException("条目过大，疑似 ZIP 炸弹: " + e.FullName);
            FileStream fs = File.OpenRead(path);
            try
            {
                fs.Seek(e.LocalOffset, SeekOrigin.Begin);
                byte[] lfh = MiniZip.ReadExact(fs, 30);
                if (MiniZip.ReadU32(lfh, 0) != MiniZip.LFH_SIGNATURE)
                    throw new IOException("本地文件头损坏: " + e.FullName);
                ushort nameLen = MiniZip.ReadU16(lfh, 26);
                ushort extraLen = MiniZip.ReadU16(lfh, 28);
                fs.Seek(nameLen + extraLen, SeekOrigin.Current);

                // 存储方式（method=0）：数据本身就是明文，直接交给调用方，零拷贝
                if (e.Method == 0)
                    return new MiniZipBoundedStream(fs, e.CompressedSize, e.CompressedSize);

                // deflate：在**有界子流**上解压。用有界流而不是裸 FileStream，
                // 是因为 DeflateStream 会一直读到文件尾——那样会把压缩包后面的
                // 其他条目也当数据读进来，解压结果就是错的。
                Stream bounded = new MiniZipBoundedStream(fs, e.CompressedSize, e.CompressedSize);
                return new MiniZipVerifyingStream(
                    new DeflateStream(bounded, CompressionMode.Decompress), e.UncompressedSize);
            }
            catch
            {
                try { fs.Close(); } catch { }
                throw;
            }
        }

        public void Dispose() { entries = null; byName = null; }
    }

    // 写侧：entries 一次性写出（局部文件头 + 数据 + 中央目录 + EOCD）
    public class MiniZipEntryOut
    {
        public string Name;
        public byte[] Data;
        public MiniZipEntryOut(string name, byte[] data) { Name = name; Data = data; }
    }

    public static class MiniZipWriter
    {
        public static void Write(Stream output, List<MiniZipEntryOut> items)
        {
            List<uint> crcs = new List<uint>();
            List<uint> compSizes = new List<uint>();
            List<uint> uncompSizes = new List<uint>();
            List<uint> offsets = new List<uint>();
            List<ushort> methods = new List<ushort>();

            for (int i = 0; i < items.Count; i++)
            {
                byte[] raw = items[i].Data == null ? new byte[0] : items[i].Data;
                uint crc = MiniZip.Crc32(raw, 0, raw.Length);
                byte[] comp = Deflate(raw);
                ushort method = comp.Length < raw.Length ? (ushort)8 : (ushort)0;
                if (method == 0) comp = raw;
                offsets.Add((uint)output.Position);
                crcs.Add(crc);
                compSizes.Add((uint)comp.Length);
                uncompSizes.Add((uint)raw.Length);
                methods.Add(method);

                byte[] name = Encoding.UTF8.GetBytes(items[i].Name);
                ushort dosTime, dosDate;
                MiniZip.DosDateTime(out dosTime, out dosDate);
                MemoryStream h = new MemoryStream(30 + name.Length);
                MiniZip.WriteU32(h, 0x04034b50);
                MiniZip.WriteU16(h, 20);            // version needed
                MiniZip.WriteU16(h, 0x0800);        // flags: UTF-8 名
                MiniZip.WriteU16(h, method);
                MiniZip.WriteU16(h, dosTime);
                MiniZip.WriteU16(h, dosDate);
                MiniZip.WriteU32(h, crc);
                MiniZip.WriteU32(h, (uint)comp.Length);
                MiniZip.WriteU32(h, (uint)raw.Length);
                MiniZip.WriteU16(h, (ushort)name.Length);
                MiniZip.WriteU16(h, 0);
                h.Write(name, 0, name.Length);
                byte[] hb = h.ToArray();
                output.Write(hb, 0, hb.Length);
                output.Write(comp, 0, comp.Length);
            }

            uint cdStart = (uint)output.Position;
            for (int i = 0; i < items.Count; i++)
            {
                byte[] name = Encoding.UTF8.GetBytes(items[i].Name);
                MemoryStream c = new MemoryStream(46 + name.Length);
                MiniZip.WriteU32(c, 0x02014b50);
                MiniZip.WriteU16(c, 20);            // version made by
                MiniZip.WriteU16(c, 20);            // version needed
                MiniZip.WriteU16(c, 0x0800);
                MiniZip.WriteU16(c, methods[i]);
                MiniZip.WriteU16(c, 0);             // time（与 LFH 一致性非必需，读侧不校验）
                MiniZip.WriteU16(c, 0);
                MiniZip.WriteU32(c, crcs[i]);
                MiniZip.WriteU32(c, compSizes[i]);
                MiniZip.WriteU32(c, uncompSizes[i]);
                MiniZip.WriteU16(c, (ushort)name.Length);
                MiniZip.WriteU16(c, 0);
                MiniZip.WriteU16(c, 0);
                MiniZip.WriteU16(c, 0);
                MiniZip.WriteU16(c, 0);
                MiniZip.WriteU32(c, 0);
                MiniZip.WriteU32(c, offsets[i]);
                c.Write(name, 0, name.Length);
                byte[] cb = c.ToArray();
                output.Write(cb, 0, cb.Length);
            }
            uint cdSize = (uint)output.Position - cdStart;

            MemoryStream e = new MemoryStream(22);
            MiniZip.WriteU32(e, 0x06054b50);
            MiniZip.WriteU16(e, 0);
            MiniZip.WriteU16(e, 0);
            MiniZip.WriteU16(e, (ushort)items.Count);
            MiniZip.WriteU16(e, (ushort)items.Count);
            MiniZip.WriteU32(e, cdSize);
            MiniZip.WriteU32(e, cdStart);
            MiniZip.WriteU16(e, 0);
            byte[] eb = e.ToArray();
            output.Write(eb, 0, eb.Length);
            output.Flush();
        }

        static byte[] Deflate(byte[] raw)
        {
            using (MemoryStream src = new MemoryStream(raw))
            using (MemoryStream dst = new MemoryStream())
            {
                using (DeflateStream ds = new DeflateStream(dst, CompressionMode.Compress))
                {
                    byte[] buf = new byte[8192];
                    while (true)
                    {
                        int n = src.Read(buf, 0, buf.Length);
                        if (n <= 0) break;
                        ds.Write(buf, 0, n);
                    }
                }
                return dst.ToArray();
            }
        }
    }

    // 共享底层解析/校验（boot\Unzip.cs 与 MiniZipFile、写侧共用）
    public static class MiniZip
    {
        public const uint EOCD_SIGNATURE = 0x06054b50;
        public const uint CDH_SIGNATURE = 0x02014b50;
        public const uint LFH_SIGNATURE = 0x04034b50;
        public const long MAX_UNCOMPRESSED = 1024L * 1024 * 1024;

        public static byte[] FindEocd(FileStream fs)
        {
            long len = fs.Length;
            int window = (int)Math.Min(len, 66000);
            fs.Seek(len - window, SeekOrigin.Begin);
            byte[] buf = ReadExact(fs, window);
            for (int i = buf.Length - 22; i >= 0; i--)
            {
                if (ReadU32(buf, i) == EOCD_SIGNATURE)
                {
                    byte[] rec = new byte[22];
                    Array.Copy(buf, i, rec, 0, 22);
                    return rec;
                }
            }
            return null;
        }

        public static byte[] ReadEntryData(FileStream fs, uint localOffset, ushort method, uint compSize, long uncompSize, string name)
        {
            if (uncompSize > MAX_UNCOMPRESSED) throw new IOException("条目过大，疑似 ZIP 炸弹: " + name);
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
            if (method != 8) return null;
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

        public static string DecodeName(byte[] bytes, ushort flags)
        {
            if ((flags & 0x800) != 0) return Encoding.UTF8.GetString(bytes);
            return Encoding.Default.GetString(bytes);
        }

        public static void DosDateTime(out ushort time, out ushort date)
        {
            DateTime now = DateTime.Now;
            time = (ushort)((now.Hour << 11) | (now.Minute << 5) | (now.Second / 2));
            date = (ushort)(((now.Year - 1980) << 9) | (now.Month << 5) | now.Day);
        }

        public static byte[] ReadExact(FileStream fs, int count)
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

        public static ushort ReadU16(byte[] b, int off) { return (ushort)(b[off] | (b[off + 1] << 8)); }
        public static uint ReadU32(byte[] b, int off) { return (uint)(b[off] | (b[off + 1] << 8) | (b[off + 2] << 16) | (b[off + 3] << 24)); }

        public static void WriteU16(MemoryStream s, ushort v)
        {
            s.WriteByte((byte)(v & 0xFF));
            s.WriteByte((byte)((v >> 8) & 0xFF));
        }
        public static void WriteU32(MemoryStream s, uint v)
        {
            s.WriteByte((byte)(v & 0xFF));
            s.WriteByte((byte)((v >> 8) & 0xFF));
            s.WriteByte((byte)((v >> 16) & 0xFF));
            s.WriteByte((byte)((v >> 24) & 0xFF));
        }

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

        public static uint Crc32(byte[] data, int offset, int count)
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

    // XlsxSkeleton —— 最小 xlsx 包骨架的共享构造（消除 ConvertEngine 与 MiniXlsxWrite 的逐字重复）
    // 两处此前各写一遍 [Content_Types].xml / _rels/.rels / workbook.xml / workbook.xml.rels 样板，
    // 且已开始漂移（单 sheet 硬编码 vs 多 sheet 循环）。统一到这里，各调用方只保留 sheet 内容差异。
    public static class XlsxSkeleton
    {
        public const string XmlHead = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";

        public static void Add(List<MiniZipEntryOut> ents, string name, string content)
        {
            ents.Add(new MiniZipEntryOut(name, new UTF8Encoding(false).GetBytes(content)));
        }

        // 固定的包级关系（与 sheet 数无关）
        public static string RootRels()
        {
            return XmlHead +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                "</Relationships>";
        }

        // workbook.xml：sheetsXml 为各 <sheet .../> 的拼接
        public static string WorkbookXml(string sheetsXml)
        {
            return XmlHead +
                "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                "<sheets>" + sheetsXml + "</sheets></workbook>";
        }

        // workbook.xml.rels：relsXml 为各 <Relationship .../> 的拼接
        public static string WorkbookRels(string relsXml)
        {
            return XmlHead +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                relsXml + "</Relationships>";
        }

        // [Content_Types].xml：sheetCount 决定 sheet Override 条目数；withStyles 追加 styles.xml 条目
        public static string ContentTypes(int sheetCount, bool withStyles)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(XmlHead);
            sb.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
            sb.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
            sb.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
            sb.Append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
            if (withStyles)
                sb.Append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
            for (int i = 1; i <= sheetCount; i++)
            {
                sb.Append("<Override PartName=\"/xl/worksheets/sheet").Append(i)
                  .Append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
            }
            sb.Append("</Types>");
            return sb.ToString();
        }

        public static string ContentTypes(int sheetCount)
        {
            return ContentTypes(sheetCount, false);
        }

        // 多 sheet 的 workbook.xml.rels（RelId 从 1 起，按 sheet 顺序）
        public static string SheetsRels(int sheetCount)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 1; i <= sheetCount; i++)
            {
                sb.Append("<Relationship Id=\"rId").Append(i)
                  .Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet")
                  .Append(i).Append(".xml\"/>");
            }
            return WorkbookRels(sb.ToString());
        }
    }
}
