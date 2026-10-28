# E2E-тест отправки отчёта: локальный TCP-сервер, config -> zip -> multipart
# -> ретрай (первая попытка 500, вторая 200) -> send_result.txt.
# Затем негативный сценарий: сервер недоступен, ожидаем FAILED после retry_count.
# Запуск: powershell -File test_send.ps1 [-XlamPath <путь>]
param([string]$XlamPath = (Join-Path $env:APPDATA 'Microsoft\AddIns\MicShot.xlam'))
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$cfgDir = Join-Path $env:APPDATA 'MicShot'
$cfgPath = Join-Path $cfgDir 'send_config.txt'
$cfgBak = $null
if (Test-Path $cfgPath) {
    $cfgBak = $cfgPath + '.test_bak'
    Move-Item $cfgPath $cfgBak -Force
}

$okAll = $true
try {
    # ---------- фейковая папка отчёта ----------
    $repDir = Join-Path $env:TEMP ('MicShot_sendtest_' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $repDir | Out-Null
    Set-Content (Join-Path $repDir 'README.txt') 'test report' -Encoding ASCII
    Set-Content (Join-Path $repDir 'actions_log.bas') "Attribute VB_Name = ""x""" -Encoding ASCII
    [IO.File]::WriteAllBytes((Join-Path $repDir 'narration.wav'), (New-Object byte[] 4096))
    # комментарий по-настоящему UTF-8 (проверяем чтение в BuildCaseBody)
    $st = [System.Text.UTF8Encoding]::new($false)
    [IO.File]::WriteAllText((Join-Path $repDir 'comment.txt'), "Комментарий: не работает фильтр", $st)

    # ---------- сервер: первая попытка 500, вторая 200 ----------
    $port = 18000 + (Get-Random -Maximum 2000)
    $logDir = Join-Path $env:TEMP ('MicShot_sendlog_' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $logDir | Out-Null
    $job = Start-Job -ArgumentList $port, $logDir -ScriptBlock {
        param($port, $logDir)
        $listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, $port)
        $listener.Start()
        try {
            for ($n = 1; $n -le 3; $n++) {
                $client = $listener.AcceptTcpClient()
                $stream = $client.GetStream()
                # читаем сырые байты до CRLFCRLF (заголовки), остальное - тело
                $buf = [System.Collections.Generic.List[byte]]::new()
                $data = New-Object byte[] 65536
                $hdrEnd = -1
                while ($hdrEnd -lt 0) {
                    $len = $stream.Read($data, 0, $data.Length)
                    if ($len -le 0) { break }
                    for ($i = 0; $i -lt $len; $i++) { $buf.Add($data[$i]) }
                    for ($i = 3; $i -lt $buf.Count; $i++) {
                        if ($buf[$i-3] -eq 13 -and $buf[$i-2] -eq 10 -and $buf[$i-1] -eq 13 -and $buf[$i] -eq 10) { $hdrEnd = $i + 1; break }
                    }
                }
                $hdrText = [System.Text.Encoding]::ASCII.GetString($buf.GetRange(0, $hdrEnd).ToArray())
                $clen = 0
                foreach ($h in ($hdrText -split "`r`n")) {
                    if ($h -like 'Content-Length:*') { $clen = [int]($h.Split(':')[1].Trim()) }
                }
                $body = New-Object byte[] $clen
                $already = [Math]::Min($buf.Count - $hdrEnd, $clen)
                if ($already -gt 0) { $buf.CopyTo($hdrEnd, $body, 0, $already) }
                $r = $already
                while ($r -lt $clen) {
                    $len = $stream.Read($body, $r, $clen - $r)
                    if ($len -le 0) { break }
                    $r += $len
                }
                # ищем сигнатуру zip (PK\x03\x04) в первых байтах тела
                $zipOff = -1
                for ($i = 0; $i -lt [Math]::Min($clen - 4, 3000); $i++) {
                    if ($body[$i] -eq 80 -and $body[$i+1] -eq 75 -and $body[$i+2] -eq 3 -and $body[$i+3] -eq 4) { $zipOff = $i; break }
                }
                $auth = ''
                $ctype = ''
                foreach ($h in ($hdrText -split "`r`n")) {
                    if ($h -like 'X-Test-Key:*') { $auth = $h.Substring(11).Trim() }
                    if ($h -like 'Content-Type:*') { $ctype = $h.Substring(13).Trim() }
                }
                @{ n = $n; auth = $auth; bodyLen = $clen; zipOff = $zipOff; ctype = $ctype } |
                    ConvertTo-Json | Set-Content (Join-Path $logDir ("req{0}.json" -f $n))
                $code = 500; $respBody = '{"error":"try again"}'
                if ($n -ge 2) { $code = 200; $respBody = '{"ok":true,"id":"case-42"}' }
                $respBytes = [System.Text.Encoding]::UTF8.GetBytes($respBody)
                $head = "HTTP/1.1 $code Test`r`nContent-Type: application/json`r`nContent-Length: $($respBytes.Length)`r`nConnection: close`r`n`r`n"
                $headB = [System.Text.Encoding]::ASCII.GetBytes($head)
                $stream.Write($headB, 0, $headB.Length)
                $stream.Write($respBytes, 0, $respBytes.Length)
                $stream.Close()
                $client.Close()
            }
        } finally { $listener.Stop() }
    }

    # ---------- конфиг на локальный сервер ----------
    New-Item -ItemType Directory -Force -Path $cfgDir | Out-Null
    @"
endpoint=http://127.0.0.1:$port/upload
auth=header
header_name=X-Test-Key
token=secret123
file_field=file
timeout_sec=30
retry_count=3
retry_delay_sec=1
retry_backoff=1
"@ | Set-Content $cfgPath -Encoding ASCII

    # ---------- запуск отправки ----------
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try {
        [void]$excel.Workbooks.Open($XlamPath)
        Start-Sleep -Milliseconds 800
        $excel.Run('CompileCheck') | Out-Null
        Write-Host "SENDTEST: отправляю на 127.0.0.1:$port (ожидание: 500 -> повтор -> 200)..."
        $result = $excel.Run('SendReportToServer', $repDir)
        if ($result -ne $true) { throw "ожидался True после ретрая, получено: $result" }
        Write-Host 'PASS: SendReportToServer -> True (после ретрая)'

        Wait-Job $job -Timeout 10 | Out-Null
        $r1 = Get-Content (Join-Path $logDir 'req1.json') -Raw | ConvertFrom-Json
        $r2 = Get-Content (Join-Path $logDir 'req2.json') -Raw | ConvertFrom-Json
        if ($r1.auth -ne 'secret123') { throw "заголовок X-Test-Key не дошёл: '$($r1.auth)'" }
        if ($r2.auth -ne 'secret123') { throw "заголовок X-Test-Key не дошёл на повторе" }
        if ($r1.zipOff -le 0 -or $r2.zipOff -le 0) { throw "zip-сигнатура не найдена в теле (req1=$($r1.zipOff), req2=$($r2.zipOff))" }
        if ($r1.bodyLen -lt 2000) { throw "тело подозрительно мало: $($r1.bodyLen)" }
        if ($r2.ctype -notlike 'multipart/form-data; boundary=*') { throw "Content-Type не multipart: $($r2.ctype)" }
        Write-Host ("PASS: обе попытки дошли (auth=X-Test-Key, zip в теле с offset {0}, {1} байт)" -f $r2.zipOff, $r2.bodyLen)

        $resFile = Join-Path $repDir 'send_result.txt'
        if (-not (Test-Path $resFile)) { throw 'нет send_result.txt' }
        $resText = [IO.File]::ReadAllText($resFile, [System.Text.Encoding]::UTF8)
        if ($resText -notlike 'SENT*') { throw "send_result.txt не SENT: $resText" }
        if ($resText -notlike '*case-42*') { throw "в send_result.txt нет id от сервера: $resText" }
        Write-Host "PASS: send_result.txt = $($resText.Trim())"
    }
    finally {
        try { $excel.Quit() } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        Stop-Job $job -ErrorAction SilentlyContinue
        Remove-Job $job -Force -ErrorAction SilentlyContinue
    }

    # ---------- негативный: сервер недоступен ----------
    $deadPort = $port + 1   # никто не слушает
    @"
endpoint=http://127.0.0.1:$deadPort/upload
auth=none
timeout_sec=5
retry_count=2
retry_delay_sec=1
retry_backoff=1
"@ | Set-Content $cfgPath -Encoding ASCII

    $excel2 = New-Object -ComObject Excel.Application
    $excel2.Visible = $false
    $excel2.DisplayAlerts = $false
    try {
        [void]$excel2.Workbooks.Open($XlamPath)
        Start-Sleep -Milliseconds 800
        $excel2.Run('CompileCheck') | Out-Null
        Write-Host "SENDTEST: негативный сценарий (порт $deadPort закрыт, 2 попытки)..."
        $result2 = $excel2.Run('SendReportToServer', $repDir)
        if ($result2 -ne $false) { throw "недоступный сервер должен давать False, получено: $result2" }
        $resText2 = [IO.File]::ReadAllText((Join-Path $repDir 'send_result.txt'), [System.Text.Encoding]::UTF8)
        if ($resText2 -notlike 'FAILED*') { throw "send_result.txt не FAILED: $resText2" }
        $logTail = (Get-Content (Join-Path $cfgDir 'send_log.txt') -Tail 4) -join ' | '
        Write-Host "PASS: недоступный сервер -> False, send_result.txt FAILED"
        Write-Host ("  send_log хвост: {0}" -f $logTail)
    }
    finally {
        try { $excel2.Quit() } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel2)
    }
    Write-Host 'SENDTEST_OK'
}
catch {
    Write-Host ("SENDTEST_FAIL: " + $_.Exception.Message)
    $okAll = $false
}
finally {
    if ($cfgBak) { Move-Item $cfgBak $cfgPath -Force }
    if (Test-Path $repDir) { Remove-Item $repDir -Recurse -Force -ErrorAction SilentlyContinue }
    Remove-Item $logDir -Recurse -Force -ErrorAction SilentlyContinue
}
if (-not $okAll) { exit 1 }
