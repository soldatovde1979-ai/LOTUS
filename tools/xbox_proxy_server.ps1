# xbox_proxy_server.ps1 - HTTPS CONNECT proxy with DoH (RFC 8484) via xbox-dns.
# Run:  powershell -NoProfile -ExecutionPolicy Bypass -File xbox_proxy_server.ps1 [port]
# No admin rights needed (port above 1024). ASCII-only output.
param([int]$Port = 18080)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$DohUrl = 'https://xbox-dns.ru/dns-query'
$script:Cache = @{}

function Get-DohIPv4([string]$Name) {
    if ($script:Cache.ContainsKey($Name) -and $script:Cache[$Name].Exp -gt (Get-Date)) {
        return $script:Cache[$Name].Ip
    }
    $q = New-Object 'System.Collections.Generic.List[byte]'
    foreach ($b in [byte[]](0x12, 0x34, 0x01, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00)) { $q.Add($b) }
    foreach ($label in $Name.Split('.')) {
        $bytes = [Text.Encoding]::ASCII.GetBytes($label)
        $q.Add([byte]$bytes.Length)
        foreach ($b in $bytes) { $q.Add($b) }
    }
    $q.Add([byte]0)
    foreach ($b in [byte[]](0x00, 0x01, 0x00, 0x01)) { $q.Add($b) }
    $b64 = [Convert]::ToBase64String($q.ToArray()).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    $url = "$DohUrl`?dns=$b64"
    try {
        $resp = Invoke-WebRequest -UseBasicParsing -Uri $url -Headers @{ Accept = 'application/dns-json' } -TimeoutSec 10
    } catch {
        Write-Host "[doh] $Name : DoH error: $($_.Exception.Message)"
        return $null
    }
    $data = $resp.Content | ConvertFrom-Json
    $a = $data.Answer | Where-Object { $_.type -eq 1 } | Select-Object -First 1
    if (-not $a) {
        Write-Host "[doh] $Name : no A record"
        return $null
    }
    $ip = [string]$a.data
    $ttl = 300
    try { $ttl = [int]$a.TTL } catch { $ttl = 300 }
    if ($ttl -lt 10) { $ttl = 10 }
    if ($ttl -gt 3600) { $ttl = 3600 }
    $script:Cache[$Name] = @{ Ip = $ip; Exp = (Get-Date).AddSeconds($ttl) }
    Write-Host "[doh] $Name -> $ip"
    return $ip
}

function Write-ProxyError($Stream, [string]$Status) {
    $resp = "HTTP/1.1 $Status`r`nContent-Length: 0`r`n`r`n"
    $bytes = [Text.Encoding]::ASCII.GetBytes($resp)
    try { $Stream.Write($bytes, 0, $bytes.Length); $Stream.Flush() } catch { }
}

$script:ClientHandler = {
    param($Client)
    try {
        $Stream = $Client.GetStream()
        $Buf = New-Object byte[] 8192
        $Req = ''
        while ($Req -notmatch "`r`n`r`n") {
            $n = $Stream.Read($Buf, 0, $Buf.Length)
            if ($n -le 0) { break }
            $Req += [Text.Encoding]::ASCII.GetString($Buf, 0, $n)
            if ($Req.Length -gt 8192) { break }
        }
        $Line = ($Req -split "`r`n")[0]
        $M = [regex]::Match($Line, '^CONNECT\s+([^\s:]+):(\d+)')
        if (-not $M.Success) {
            Write-ProxyError $Stream '400 Bad Request'
            $Client.Close()
            return
        }
        $HostName = $M.Groups[1].Value
        $HostPort = [int]$M.Groups[2].Value
        $Ip = Get-DohIPv4 $HostName
        if (-not $Ip) {
            Write-ProxyError $Stream '502 Bad Gateway'
            $Client.Close()
            return
        }
        $Remote = New-Object Net.Sockets.TcpClient
        try {
            $Iar = $Remote.BeginConnect($Ip, $HostPort, $null, $null)
            if (-not $Iar.AsyncWaitHandle.WaitOne(10000)) {
                Write-ProxyError $Stream '502 Bad Gateway'
                $Remote.Close(); $Client.Close()
                return
            }
            $Remote.EndConnect($Iar)
        } catch {
            Write-Host "[proxy] $HostName`:$HostPort -> $Ip : connect failed: $($_.Exception.Message)"
            Write-ProxyError $Stream '502 Bad Gateway'
            $Remote.Close(); $Client.Close()
            return
        }
        $Remote.ReceiveTimeout = 300000
        $Remote.SendTimeout = 300000
        $Client.ReceiveTimeout = 300000
        $Client.SendTimeout = 300000
        Write-Host "[proxy] CONNECT $HostName`:$HostPort -> $Ip"
        $resp = "HTTP/1.1 200 Connection Established`r`n`r`n"
        $bytes = [Text.Encoding]::ASCII.GetBytes($resp)
        try { $Stream.Write($bytes, 0, $bytes.Length); $Stream.Flush() } catch { $Remote.Close(); $Client.Close(); return }
        try {
            $RS = $Remote.GetStream()
            $t1 = $Stream.CopyToAsync($RS)
            $t2 = $RS.CopyToAsync($Stream)
            [Threading.Tasks.Task]::WaitAll(@($t1, $t2))
        } catch { }
        $Remote.Close()
    } catch {
        Write-Host "[proxy] handler error: $($_.Exception.Message)"
    } finally {
        try { $Client.Close() } catch { }
    }
}

$Listener = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Any, $Port)
$Listener.Start()
Write-Host "[proxy] listening 0.0.0.0:$Port, DoH: $DohUrl"
while ($true) {
    $Client = $Listener.AcceptTcpClient()
    $pts = [System.Threading.ParameterizedThreadStart]$script:ClientHandler
    $th = New-Object System.Threading.Thread($pts)
    $th.IsBackground = $true
    $th.Start($Client)
}
