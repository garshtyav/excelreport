# ============================================================
#  uninstall_micshot.ps1 - полное удаление MicShot из профиля
#  пользователя. Права администратора НЕ нужны.
#  Excel должен быть закрыт.
# ============================================================
$ErrorActionPreference = "Continue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
    Write-Host "ВНИМАНИЕ: Excel сейчас запущен."
    Write-Host "Закройте Excel и запустите удалитель заново."
    Read-Host "Enter для выхода"
    exit 1
}

$sec    = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"
$tl     = "$sec\Trusted Locations"
$opt    = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"
$mgr    = "HKCU:\Software\Microsoft\Office\16.0\Excel\Add-in Manager"
$addIns = Join-Path $env:APPDATA "Microsoft\AddIns"
$xlam   = Join-Path $addIns "MicShot.xlam"

# 1. Убираем автозагрузку (OPEN-записи в Options), другие надстройки не трогаем
if (Test-Path $opt) {
    $op = Get-Item $opt
    $keep = @()
    foreach ($n in @($op.GetValueNames() | Where-Object { $_ -cmatch "^OPEN\d*$" })) {
        $v = $op.GetValue($n)
        if ($v -notmatch "MicShot\.xlam") { $keep += $v }
    }
    foreach ($n in @($op.GetValueNames() | Where-Object { $_ -cmatch "^OPEN\d*$" })) {
        Remove-ItemProperty -Path $opt -Name $n -ErrorAction SilentlyContinue
    }
    $j = 0
    foreach ($v in $keep) {
        $nm = $(if ($j -eq 0) { "OPEN" } else { "OPEN$j" })
        Set-ItemProperty -Path $opt -Name $nm -Value $v
        $j++
    }
    Write-Host "[ok] Автозагрузка MicShot убрана"
}

# 2. Убираем из списка надстроек Excel (Add-in Manager)
if (Test-Path $mgr) {
    $mg = Get-Item $mgr
    foreach ($n in @($mg.GetValueNames())) {
        if ("$($mg.GetValue($n))" -match "MicShot\.xlam") {
            Remove-ItemProperty -Path $mgr -Name $n -ErrorAction SilentlyContinue
        }
    }
    Write-Host "[ok] Убрана из списка надстроек Excel"
}

# 3. Удаляем файл надстройки
if (Test-Path $xlam) {
    Remove-Item $xlam -Force
    Write-Host "[ok] Удалён файл: $xlam"
} else {
    Write-Host "[i] Файл надстройки не найден (уже удалён?)"
}

# 4. На случай ручной установки в XLSTART
$xlstart = Join-Path $env:APPDATA "Microsoft\Excel\XLSTART"
$f = Join-Path $xlstart "MicShot.xlam"
if (Test-Path $f) {
    Remove-Item $f -Force
    Write-Host "[ok] Удалён из XLSTART: $f"
}

# 5. Доверенное расположение папки AddIns убираем, только если там
#    не осталось ДРУГИХ надстроек (чтобы не сломать чужие)
$otherAddIns = @(Get-ChildItem $addIns -ErrorAction SilentlyContinue |
    Where-Object { $_.Extension -ieq ".xlam" -or $_.Extension -ieq ".xla" })
if (Test-Path $tl) {
    if ($otherAddIns.Count -eq 0) {
        Get-ChildItem $tl -ErrorAction SilentlyContinue | ForEach-Object {
            $p = (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).Path
            if ($p -and $p.TrimEnd('\') -ieq $addIns) {
                Remove-Item $_.PSPath -Force
                Write-Host "[ok] Доверенное расположение папки надстроек убрано"
            }
        }
    } else {
        Write-Host "[i] В AddIns есть другие надстройки - доверенное расположение оставлено"
    }
}

# 6. Конфиг распознавания - по вопросу
$cfgDir = Join-Path $env:APPDATA "MicShot"
if (Test-Path $cfgDir) {
    $a = Read-Host "Удалить настройки распознавания ($cfgDir)? [y/N]"
    if ($a -match "^[yYдД]") {
        Remove-Item $cfgDir -Recurse -Force
        Write-Host "[ok] Настройки распознавания удалены"
    } else {
        Write-Host "[i] Настройки распознавания оставлены"
    }
}

# 7. Отчёты - по вопросу (по умолчанию оставляем)
$docs = [Environment]::GetFolderPath("MyDocuments")
$reports = @(Get-ChildItem $docs -Directory -Filter "MicShot_Report_*" -ErrorAction SilentlyContinue)
if ($reports.Count -gt 0) {
    Write-Host ("Найдено отчётов MicShot_Report_*: " + $reports.Count + " (в " + $docs + ")")
    $a = Read-Host "Удалить их тоже? [y/N]"
    if ($a -match "^[yYдД]") {
        $reports | Remove-Item -Recurse -Force
        Write-Host "[ok] Отчёты удалены"
    } else {
        Write-Host "[i] Отчёты оставлены"
    }
}

Write-Host ""
Write-Host "=============================================="
Write-Host " MicShot полностью удалён."
Write-Host " Если Excel был открыт - перезапустите его,"
Write-Host " чтобы кнопки исчезли."
Write-Host "=============================================="
Read-Host "Enter для выхода"
