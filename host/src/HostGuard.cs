// HostGuard —— 出网请求安全守卫（Mimosa 约束落地，形态沿用 WorkBuddyRepair HostGuard）
//   * 仅允许 http/https；
//   * 请求前校验 host：解析出的每个 IP 都必须为公网；
//   * 默认拒绝 localhost/环回/私有/保留地址（IP 直连与域名解析双查）；
//   * 企业内网网关场景由用户在设置中显式勾选"允许内网端点"后才放行私有段；
//   * 不自动跟随重定向（由调用方手动复检，防 302 跳内网）。
using System;
using System.Net;
using System.Net.Sockets;

namespace OfficeAgent.Host
{
    public static class HostGuard
    {
        // 返回 null = 放行；否则为拒绝原因
        public static string Check(string url, bool allowPrivate)
        {
            if (url == null || url.Trim().Length == 0) return "URL 为空";
            url = url.Trim();
            Uri uri;
            try { uri = new Uri(url); }
            catch { return "URL 无法解析"; }
            if (uri.Scheme != "http" && uri.Scheme != "https") return "仅允许 http/https";
            string host = uri.Host;
            if (host == null || host.Length == 0) return "URL 缺少主机名";
            if (host.StartsWith("[")) host = host.Trim('[', ']');   // IPv6 字面量去方括号
            if (host.ToLowerInvariant() == "localhost" || host.ToLowerInvariant().EndsWith(".localhost"))
                return allowPrivate ? null : "拒绝 localhost（如为企业网关请勾选允许内网端点）";

            IPAddress[] addresses;
            IPAddress ipLiteral;
            if (IPAddress.TryParse(host, out ipLiteral))
            {
                addresses = new IPAddress[] { ipLiteral };
            }
            else
            {
                try { addresses = Dns.GetHostAddresses(host); }
                catch { return "域名无法解析: " + host; }
                if (addresses == null || addresses.Length == 0) return "域名无解析记录: " + host;
            }
            foreach (IPAddress ip in addresses)
            {
                string reason;
                if (IsPrivateOrReserved(ip, out reason))
                {
                    if (!allowPrivate) return "拒绝" + reason + "地址 " + ip + "（如为企业网关请勾选允许内网端点）";
                }
            }
            return null;
        }

        public static bool IsPrivateOrReserved(IPAddress ip, out string reason)
        {
            reason = "";
            if (ip.AddressFamily == AddressFamily.InterNetwork)
            {
                byte[] b = ip.GetAddressBytes();
                if (b[0] == 127) { reason = " 环回"; return true; }
                if (b[0] == 10) { reason = " 私有"; return true; }
                if (b[0] == 172 && b[1] >= 16 && b[1] <= 31) { reason = " 私有"; return true; }
                if (b[0] == 192 && b[1] == 168) { reason = " 私有"; return true; }
                if (b[0] == 169 && b[1] == 254) { reason = " 链路本地"; return true; }
                if (b[0] == 0) { reason = " 保留"; return true; }
                if (b[0] == 100 && b[1] >= 64 && b[1] <= 127) { reason = " CGNAT"; return true; }
                if (b[0] == 192 && b[1] == 0 && b[2] == 0) { reason = " 保留"; return true; }
                if (b[0] == 198 && (b[1] == 18 || b[1] == 19)) { reason = " 保留"; return true; }
                if (b[0] >= 224) { reason = " 组播/保留"; return true; }
                return false;
            }
            if (ip.AddressFamily == AddressFamily.InterNetworkV6)
            {
                if (ip.IsIPv6LinkLocal) { reason = " 链路本地"; return true; }
                if (ip.IsIPv6SiteLocal) { reason = " 站点本地"; return true; }
                byte[] b = ip.GetAddressBytes();
                bool allZero = true;
                for (int i = 0; i < 16; i++) { if (b[i] != 0) { allZero = false; break; } }
                if (allZero) { reason = " 未指定"; return true; }
                bool onlyLast = true;                       // ::1 环回
                for (int i = 0; i < 15; i++) { if (b[i] != 0) { onlyLast = false; break; } }
                if (onlyLast && b[15] == 1) { reason = " 环回"; return true; }
                if (b[0] == 0xfe && (b[1] & 0xc0) == 0x80) { reason = " 链路本地"; return true; }
                if ((b[0] & 0xfe) == 0xfc) { reason = " ULA 私有"; return true; }
                if (b[0] == 0xff) { reason = " 组播"; return true; }
                // IPv4-mapped 回查
                bool mapped = true;
                for (int i = 0; i < 10; i++) { if (b[i] != 0) { mapped = false; break; } }
                if (mapped && b[10] == 0xff && b[11] == 0xff)
                {
                    IPAddress v4 = new IPAddress(new byte[] { b[12], b[13], b[14], b[15] });
                    string r;
                    bool priv = IsPrivateOrReserved(v4, out r);
                    if (priv) { reason = r; return true; }
                }
                return false;
            }
            reason = " 未知协议族";
            return true;
        }
    }
}
