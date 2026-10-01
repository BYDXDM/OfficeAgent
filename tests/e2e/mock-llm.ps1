# Mock LLM server (OpenAI-compatible subset) for E2E — raw TcpListener, no admin/ACL needed.
# GET  /v1/models            -> {"data":[{"id":"mock-a"},{"id":"mock-b"}]}
# POST /v1/chat/completions  -> {"choices":[{"message":{"role":"assistant","content":"pong from mock"}}]}
# Logs every request to mock-server.log. Usage: powershell -File mock-llm.ps1
$ErrorActionPreference = "Stop"
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$log = Join-Path $dir "mock-server.log"
function Log($s) { Add-Content -LiteralPath $log -Value ("[" + (Get-Date -Format "HH:mm:ss.fff") + "] " + $s) -Encoding UTF8 }

$listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, 18080)
$listener.Start()
Log "mock server started on 127.0.0.1:18080"

while ($true) {
  $client = $listener.AcceptTcpClient()
  Log "ACCEPT from $($client.Client.RemoteEndPoint)"
  try {
    $stream = $client.GetStream()
    $stream.ReadTimeout = 8000
    $buf = New-Object byte[] 65536
    $sb = New-Object System.Text.StringBuilder
    # 读到响应头结束，再按 Content-Length 读全请求体（按字节计数——body 含中文时字符数 != 字节数）
    $headerEnd = -1
    $raw = ""
    $totalBytes = 0
    while ($true) {
      $n = $stream.Read($buf, 0, $buf.Length)
      Log ("READ chunk=" + $n + " totalBytes=" + ($totalBytes + $n))
      if ($n -le 0) { break }
      $totalBytes += $n
      $sb.Append([Text.Encoding]::UTF8.GetString($buf, 0, $n)) | Out-Null
      $raw = $sb.ToString()
      $headerEnd = $raw.IndexOf("`r`n`r`n")
      if ($headerEnd -ge 0) {
        $cl = 0
        if ($raw -match "(?i)content-length:\s*(\d+)") { $cl = [int]$Matches[1] }
        if (($totalBytes - $headerEnd - 4) -ge $cl) { break }
      }
    }
    Log ("READ-END rawLen=" + $raw.Length)
    if ($raw) {
      $lines = $raw -split "`r`n"
      $reqLine = $lines[0]
      Log "REQ $reqLine"
      $parts = $reqLine -split " "
      $method = $parts[0]; $path = $parts[1]
      $body = ""
      if ($headerEnd -ge 0) {
        $body = $raw.Substring($headerEnd + 4)
        if ($body.Length -gt 0) { Log ("BODY " + $body.Substring(0, [Math]::Min(8000, $body.Length))) }
      }
      $json = ""
      if ($path -like "*\/models" -or $path -like "*/models") {
        $json = '{"data":[{"id":"mock-a"},{"id":"mock-b"}],"object":"list"}'
      } elseif ($path -like "*chat/completions*") {
        if ($body -like '*"stream":true*') {
          Log "STREAM=true (SSE mode)"
          $chunks = @(
            '{"choices":[{"delta":{"content":"pong "}}]}',
            '{"choices":[{"delta":{"content":"from mock (stream)"}}]}',
            '[DONE]'
          )
          $sse = ""
          foreach ($c in $chunks) { $sse += "data: $c`n`n" }
          $bodyBytes = [Text.Encoding]::UTF8.GetBytes($sse)
          $hdr = "HTTP/1.1 200 OK`r`nContent-Type: text/event-stream`r`nConnection: close`r`n`r`n"
          $hdrBytes = [Text.Encoding]::ASCII.GetBytes($hdr)
          $stream.Write($hdrBytes, 0, $hdrBytes.Length)
          $stream.Write($bodyBytes, 0, $bodyBytes.Length)
          $stream.Flush()
          Log "RSP SSE 3 chunks"
          try { $client.Close() } catch {}
          continue
        }
        if ($body -like '*HISTTEST*') {
          # 上下文回归测试：数请求里 user 消息条数；>=2 说明模型真的收到了历史
          try { $parsed = $body | ConvertFrom-Json; $userMsgs = @($parsed.messages | Where-Object { $_.role -eq 'user' }) } catch { $userMsgs = @() }
          Log ("HISTTEST user-msgs=" + $userMsgs.Count)
          if ($userMsgs.Count -ge 2) {
            $first = [string]$userMsgs[0].content
            if ($first.Length -gt 30) { $first = $first.Substring(0, 30) }
            $reply = "HIST-OK 收到历史 " + $userMsgs.Count + " 条，首句=" + $first
            $json = '{"id":"mock-4","choices":[{"index":0,"message":{"role":"assistant","content":"' + $reply.Replace('"', '\"') + '"},"finish_reason":"stop"}]}'
          } else {
            $json = '{"id":"mock-4","choices":[{"index":0,"message":{"role":"assistant","content":"HIST-FAIL no history received"},"finish_reason":"stop"}]}'
          }
        } elseif ($body -like '*TOOLTEST*') {
          if ($body -like '*"role":"tool"*') {
            # 第二段：工具结果已回传 → 给最终答复（引用结果内容证明链路通了）
            $json = '{"id":"mock-3","choices":[{"index":0,"message":{"role":"assistant","content":"TOOL-OK agent loop finished, tool result received."},"finish_reason":"stop"}]}'
          } else {
            # 第一段：要求调用 read_text_file 读一个已知文件
            Log "TOOLTEST stage-1 -> tool_call"
            $json = '{"id":"mock-2","choices":[{"index":0,"message":{"role":"assistant","content":"","tool_calls":[{"id":"call_1","type":"function","function":{"name":"read_text_file","arguments":"{\\"path\":\\"D:\\ai\\工作\\OfficeAgent\\store\\components.ini\\"}"}}]},"finish_reason":"tool_calls"}]}'
          }
        } else {
          $json = '{"id":"mock-1","choices":[{"index":0,"message":{"role":"assistant","content":"pong from mock"},"finish_reason":"stop"}],"usage":{"total_tokens":7}}'
        }
      } else {
        $json = '{"error":{"message":"not found"}}'
      }
      $bodyBytes = [Text.Encoding]::UTF8.GetBytes($json)
      $hdr = "HTTP/1.1 200 OK`r`nContent-Type: application/json`r`nContent-Length: $($bodyBytes.Length)`r`nConnection: close`r`n`r`n"
      $hdrBytes = [Text.Encoding]::ASCII.GetBytes($hdr)
      $stream.Write($hdrBytes, 0, $hdrBytes.Length)
      $stream.Write($bodyBytes, 0, $bodyBytes.Length)
      $stream.Flush()
      Log ("RSP " + $json.Substring(0, [Math]::Min(120, $json.Length)))
    }
  } catch { Log ("ERR " + $_.Exception.Message) }
  finally { try { $client.Close() } catch {} }
}
