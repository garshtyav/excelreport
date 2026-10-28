# ============================================================
#  install_micshot.ps1 - пользовательская установка MicShot.
#  Права администратора НЕ нужны: всё пишется в профиль
#  пользователя (%APPDATA%) и в HKCU.
# ============================================================
$ErrorActionPreference = "Continue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$src = Join-Path $PSScriptRoot "MicShot.xlam"
if (-not (Test-Path $src)) {
    Write-Host "ОШИБКА: MicShot.xlam не найден рядом с установщиком."
    Read-Host "Enter для выхода"
    exit 1
}

if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
    Write-Host "ВНИМАНИЕ: Excel сейчас запущен."
    Write-Host "Закройте Excel и запустите установщик заново."
    Read-Host "Enter для выхода"
    exit 1
}

# 1. Копируем надстройку в пользовательскую папку надстроек
$addIns = Join-Path $env:APPDATA "Microsoft\AddIns"
New-Item -ItemType Directory -Force -Path $addIns | Out-Null
try {
    Copy-Item $src (Join-Path $addIns "MicShot.xlam") -Force
    Unblock-File (Join-Path $addIns "MicShot.xlam") -ErrorAction SilentlyContinue
    Write-Host "[ok] Надстройка скопирована: $addIns\MicShot.xlam"
} catch {
    Write-Host "ОШИБКА копирования: $($_.Exception.Message)"
    Read-Host "Enter для выхода"
    exit 1
}

# 2. Папка надстроек -> доверенное расположение Excel
$sec = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"
$tl  = "$sec\Trusted Locations"
New-Item -Path $tl -Force | Out-Null
$trusted = $false
Get-ChildItem $tl -ErrorAction SilentlyContinue | ForEach-Object {
    $p = (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).Path
    if ($p -and $p.TrimEnd('\') -ieq $addIns) { $trusted = $true }
}
if (-not $trusted) {
    $i = 0
    while (Test-Path (Join-Path $tl "Location$i")) { $i++ }
    $loc = Join-Path $tl "Location$i"
    New-Item -Path $loc -Force | Out-Null
    Set-ItemProperty -Path $loc -Name Path -Value "$addIns\"
    Set-ItemProperty -Path $loc -Name AllowSubFolders -Value 1 -Type DWord
    Write-Host "[ok] Доверенное расположение добавлено"
} else {
    Write-Host "[ok] Доверенное расположение уже есть"
}

# 3. Регистрируем автозагрузку надстройки (OPEN-запись, HKCU)
$opt = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"
New-Item -Path $opt -Force | Out-Null
$op = Get-Item $opt
$keep = @()
foreach ($n in @($op.GetValueNames() | Where-Object { $_ -cmatch "^OPEN\d*$" })) {
    $v = $op.GetValue($n)
    if ($v -notmatch "MicShot\.xlam") { $keep += $v }
}
$keep += ('"' + (Join-Path $addIns "MicShot.xlam") + '"')
foreach ($n in @($op.GetValueNames() | Where-Object { $_ -cmatch "^OPEN\d*$" })) {
    Remove-ItemProperty -Path $opt -Name $n
}
$j = 0
foreach ($v in $keep) {
    $nm = $(if ($j -eq 0) { "OPEN" } else { "OPEN$j" })
    Set-ItemProperty -Path $opt -Name $nm -Value $v
    $j++
}
Write-Host "[ok] Надстройка зарегистрирована на автозагрузку"

# 4. Конфиг распознавания речи (LAN-Whisper, без ключа)
$cfgDir = Join-Path $env:APPDATA "MicShot"
New-Item -ItemType Directory -Force -Path $cfgDir | Out-Null
$cfg = Join-Path $cfgDir "asr_config.txt"
if (-not (Test-Path $cfg)) {
    $tmpl = @(
        "# Распознавание речи MicShot (LAN-Whisper). Подробности у владельца надстройки.",
        "endpoint=http://192.168.1.77:8000/transcribe",
        "model=",
        "language=ru",
        "vad_filter=true",
        "chunk_sec=60",
        "key="
    )
    $tmpl | Set-Content -Path $cfg -Encoding Default
    Write-Host "[ok] Конфиг распознавания: $cfg"
}

Write-Host ""
Write-Host "=============================================="
Write-Host " Готово! Откройте Excel - на вкладке"
Write-Host " 'НАДСТРОЙКИ' появятся кнопки MicShot."
Write-Host " Подробности - в файле ИНСТРУКЦИЯ.txt"
Write-Host "=============================================="
Read-Host "Enter для выхода"
