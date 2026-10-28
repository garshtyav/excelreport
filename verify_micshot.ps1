# Проверка скриншота на ВИДИМОМ экземпляре Excel с реальными данными.
# Запускается после пересборки; требует, чтобы надстройка была уже зарегистрирована.
$ErrorActionPreference = "Stop"
$xlamPath = Join-Path $env:APPDATA "Microsoft\AddIns\MicShot.xlam"
$testBmp  = Join-Path $env:TEMP "MicShot_verify.bmp"
$testPng  = Join-Path $env:TEMP "MicShot_verify.png"
Remove-Item $testBmp, $testPng -ErrorAction SilentlyContinue

if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) { Write-Host "VERIFY: Excel уже запущен, проверка отложена"; exit 3 }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true          # важно: захват должен идти с реального видимого окна
$excel.DisplayAlerts = $false
$wb = $null
try {
    $wb = $excel.Workbooks.Add()
    $ws = $excel.ActiveSheet
    $ws.Range("B2").Value2 = "Тест скриншота MicShot"
    $ws.Range("B3").Value2 = "Если этот текст виден на PNG - PrintWindow работает"
    for ($i = 1; $i -le 10; $i++) {
        $ws.Cells.Item($i + 4, 2).Value2 = $i
        $ws.Cells.Item($i + 4, 3).Value2 = $i * $i
    }
    $ws.Range("B:C").EntireColumn.AutoFit() | Out-Null
    Start-Sleep -Seconds 2

    $toolbarOk = $false
    try { [void]$excel.CommandBars.Item("MicShot"); $toolbarOk = $true } catch {}
    if (-not $toolbarOk) {
        # COM-экземпляры не резолвят голое имя надстройки из OPEN - открываем файл напрямую
        try {
            [void]$excel.Workbooks.Open($xlamPath)
            Start-Sleep -Milliseconds 500
            try { [void]$excel.CommandBars.Item("MicShot"); $toolbarOk = $true } catch {}
        } catch { Write-Host ("VERIFY: прямое открытие не удалось: {0}" -f $_.Exception.Message) }
        Write-Host ("VERIFY: надстройка открыта напрямую (панель={0})" -f $toolbarOk)
    } else {
        Write-Host "VERIFY: надстройка загружена автоматически (панель=True)"
    }

    $excel.Run("SelfTestBmp", $testBmp) | Out-Null
    if ((Test-Path $testBmp) -and ((Get-Item $testBmp).Length -gt 10000)) {
        Add-Type -AssemblyName System.Drawing
        $img = [System.Drawing.Image]::FromFile($testBmp)
        $img.Save($testPng, [System.Drawing.Imaging.ImageFormat]::Png)
        $img.Dispose()
        Write-Host ("VERIFY: PNG создан, {0} байт" -f (Get-Item $testPng).Length)
    } else {
        Write-Host "VERIFY: BMP не создан"
    }
}
finally {
    try { if ($wb) { $wb.Close($false) } } catch {}
    try { $excel.Quit() } catch {}
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
if (Test-Path $testPng) { Write-Host "VERIFY_OK" } else { Write-Host "VERIFY_FAILED" }
