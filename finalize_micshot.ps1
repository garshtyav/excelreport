# ============================================================
#  finalize_micshot.ps1
#  Чистая регистрация MicShot.xlam для текущего пользователя.
#  Запускается, когда Excel ЗАКРЫТ (иначе работающий Excel
#  перезапишет регистрацию при выходе).
# ============================================================
$ErrorActionPreference = "Stop"

$officeVer = "16.0"
$secKey    = "HKCU:\Software\Microsoft\Office\$officeVer\Excel\Security"
$tlKey     = "$secKey\Trusted Locations"
$optKey    = "HKCU:\Software\Microsoft\Office\$officeVer\Excel\Options"
$addInsDir = Join-Path $env:APPDATA "Microsoft\AddIns"
$addInsTrusted = "$addInsDir\"
$xlamPath  = Join-Path $addInsDir "MicShot.xlam"

if (-not (Test-Path $xlamPath)) { throw "Не найден файл надстройки: $xlamPath" }
if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) { throw "Excel ещё запущен - registration отложена" }

# --- 1. Папка AddIns -> доверенное расположение
New-Item -Path $tlKey -Force | Out-Null
$trusted = $false
Get-ChildItem $tlKey -ErrorAction SilentlyContinue | ForEach-Object {
    $tp = (Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue).Path
    if ($tp -and $tp.TrimEnd('\') -ieq $addInsDir) { $trusted = $true }
}
if (-not $trusted) {
    $i = 0
    while (Test-Path (Join-Path $tlKey "Location$i")) { $i++ }
    $locPath = Join-Path $tlKey "Location$i"
    New-Item -Path $locPath -Force | Out-Null
    Set-ItemProperty -Path $locPath -Name Path -Value $addInsTrusted
    Set-ItemProperty -Path $locPath -Name AllowSubFolders -Value 1 -Type DWord
    Write-Host "[ok] Доверенное расположение: $addInsDir"
} else {
    Write-Host "[ok] Доверенное расположение уже есть"
}

# --- 2. Убираем возможные устаревшие записи MicShot (OPEN* без "дыр")
if (Test-Path $optKey) {
    $op = Get-Item $optKey
    $keep = @()
    foreach ($n in @($op.GetValueNames() | Where-Object { $_ -cmatch "^OPEN\d*$" })) {
        $v = $op.GetValue($n)
        if ($v -and ($v -match "MicShot\.xlam")) {
            Remove-ItemProperty -Path $optKey -Name $n
            Write-Host "[ok] Убрана устаревшая запись ($n)"
        } else {
            $keep += $v
        }
    }
    foreach ($n in @($op.GetValueNames() | Where-Object { $_ -cmatch "^OPEN\d*$" })) {
        Remove-ItemProperty -Path $optKey -Name $n
    }
    $j = 0
    foreach ($v in $keep) {
        $newName = $(if ($j -eq 0) { "OPEN" } else { "OPEN$j" })
        Set-ItemProperty -Path $optKey -Name $newName -Value $v
        $j++
    }
}

# --- 3. Регистрируем надстройку и включаем
# (во скрытом экземпляре без открытой книги коллекция AddIns недоступна,
#  поэтому сначала создаём пустую книгу)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$toolbarOk = $false
try {
    [void]$excel.Workbooks.Add()
    $ai = $null
    foreach ($a in $excel.AddIns) {
        if ($a.Name -ieq "MicShot.xlam") { $ai = $a; break }
    }
    if (-not $ai) { $ai = $excel.AddIns.Add($xlamPath, $false) }
    $ai.Installed = $true
    Start-Sleep -Milliseconds 500
    try { [void]$excel.CommandBars.Item("MicShot"); $toolbarOk = $true } catch {}
    Write-Host ("[ok] Надстройка зарегистрирована, Installed={0}" -f $ai.Installed)
}
finally {
    try { $excel.Quit() } catch {}
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}

# --- 4. Проверяем, что регистрация сохранилась после выхода Excel
Start-Sleep -Milliseconds 800
$openOk = $false
if (Test-Path $optKey) {
    $op = Get-Item $optKey
    foreach ($n in $op.GetValueNames()) {
        if ($n -cmatch "^OPEN\d*$" -and ($op.GetValue($n) -match "MicShot\.xlam")) { $openOk = $true }
    }
}
Write-Host ("FINALIZE_RESULT: toolbar={0} registry={1}" -f $toolbarOk, $openOk)
if ($toolbarOk -and $openOk) { Write-Host "FINALIZE_OK" }
elseif ($toolbarOk) { Write-Host "FINALIZE_PARTIAL" }
else { Write-Host "FINALIZE_FAILED" }
