# Синхронизация MicShot_source.bas с heredoc $vbaCode из build_micshot.ps1
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$raw = [System.IO.File]::ReadAllText((Join-Path $root "build_micshot.ps1"))
$m = [regex]::Match($raw, "(?s)\$vbaCode = @'\r?\n(.*?)\r?\n'@")
if (-not $m.Success) { throw "heredoc `$vbaCode не найден" }
$content = $m.Groups[1].Value
$sb = New-Object System.Text.StringBuilder
[void]$sb.Append('Attribute VB_Name = "MicShot"')
[void]$sb.Append("`r`n")
[void]$sb.Append($content)
[void]$sb.Append("`r`n")
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText((Join-Path $root "MicShot_source.bas"), $sb.ToString(), $utf8Bom)
Write-Host ("OK: MicShot_source.bas обновлён, " + $content.Length + " символов кода")
