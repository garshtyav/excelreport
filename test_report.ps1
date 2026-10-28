# Автотест полного цикла "отчёта о проблеме":
# старт -> имитация действий -> периодический скриншот -> стоп -> проверка файлов.
# Можно проверить сборку из другого места: test_report.ps1 -XlamPath <путь>
param([string]$XlamPath = (Join-Path $env:APPDATA "Microsoft\AddIns\MicShot.xlam"))
$ErrorActionPreference = "Continue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$xlamPath = $XlamPath
$docs = [Environment]::GetFolderPath("MyDocuments")
$before = Get-ChildItem $docs -Directory -Filter "MicShot_Report_*" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false
$script:tbStart = $false   # кнопка-тумблер в состоянии "старт"
$script:tbRec = $false     # кнопка-тумблер в состоянии "стоп" во время записи
$script:tbBack = $false    # после остановки вернулась в "старт"
$script:panelUp = $false   # окно панели показано при записи
$script:panelDown = $false # после стопа панель закрыта
$script:liveTextOk = $false
$script:wavOk = $false
$script:sendStubOk = $true
function Get-ToggleButton {
    try { $cb = $excel.CommandBars.Item("MicShot") } catch { return $null }
    if (-not $cb) { return $null }
    foreach ($c in $cb.Controls) {
        if ($c.Tag -eq "MicShotToggle") { return $c }
    }
    return $null
}
try {
    $wb = $excel.Workbooks.Add()
    [void]$excel.Workbooks.Open($xlamPath)   # прямое открытие надстройки
    Start-Sleep -Milliseconds 800

    Write-Host "TEST: форс-компиляция классов..."
    $excel.Run("CompileCheck") | Out-Null

    Write-Host "TEST: панель кнопок (Auto_Open)..."
    $excel.Run("Auto_Open") | Out-Null
    Start-Sleep -Milliseconds 300
    $tb = Get-ToggleButton
    if ($tb -and $tb.Caption -like "*старт*" -and $tb.State -eq 0) {
        Write-Host "  LOG_OK: кнопка-тумблер в состоянии 'старт'"
        $script:tbStart = $true
    } else {
        Write-Host "  !! кнопка-тумблер не найдена или не в состоянии 'старт'"
    }

    Write-Host "TEST: старт записи отчёта..."
    $excel.Run("MicReportStart") | Out-Null
    Start-Sleep -Seconds 2
    $tb = Get-ToggleButton
    if ($tb -and $tb.Caption -like "*стоп*" -and $tb.State -eq -1) {
        Write-Host "  LOG_OK: во время записи кнопка - нажатая 'стоп'"
        $script:tbRec = $true
    } else {
        Write-Host "  !! во время записи кнопка не переключилась на 'стоп'"
    }

    Write-Host "TEST: добавляю текст в буфер расшифровки..."
    $excel.Run("LiveSpeechCommit", "тест распознавания речи из автотеста") | Out-Null
    $liveText = [string]$excel.Run("MicShot_LiveText")
    if ($liveText -like "*тест распознавания*") {
        Write-Host "  LOG_OK: живой текст попадает в буфер и панель"
        $script:liveTextOk = $true
    } else {
        Write-Host "  !! MicShot_LiveText не содержит тестовую фразу"
    }
    if ($excel.Run("MicShot_PanelShown")) {
        Write-Host "  LOG_OK: окно живой расшифровки показано"
        $script:panelUp = $true
    } else {
        Write-Host "  !! окно панели не показано при записи"
    }

    Write-Host "TEST: живой тик сервера (пауза/сохранение/хвост)..."
    $excel.Run("MicShot_LiveTick") | Out-Null

    # имитация действий пользователя
    $ws1 = $wb.Worksheets.Item(1)
    $ws1.Activate()
    $excel.ActiveSheet.Range("B2").Select() | Out-Null
    Start-Sleep -Milliseconds 300
    $excel.ActiveCell.Value2 = "тестMicShot"
    Start-Sleep -Milliseconds 300
    $excel.ActiveSheet.Range("B5:C9").Select() | Out-Null
    Start-Sleep -Milliseconds 300
    $excel.ActiveCell.Value2 = 123
    $ws2 = $wb.Worksheets.Add()
    $ws2.Activate()
    Start-Sleep -Milliseconds 300
    $ws1.Activate()

    # форматирование (событий у Excel нет - ловит цикл опроса раз в 2 с).
    # Каждая операция в своём try/catch: сбой одной не должен ронять тест.
    function Try-Com([scriptblock]$sb) {
        try { & $sb } catch { Write-Host ("  (шаг пропущен: " + $_.Exception.Message.Split("`n")[0] + ")") }
    }
    Write-Host "TEST: меняю форматирование выделения..."
    $excel.ActiveSheet.Range("B5:C9").Select() | Out-Null
    Start-Sleep -Milliseconds 500
    Try-Com { $excel.Selection.Font.Bold = $true }
    Start-Sleep -Milliseconds 500
    Try-Com { $excel.Selection.Interior.Color = 65535 }          # жёлтая заливка
    Start-Sleep -Milliseconds 500
    Try-Com { $excel.Selection.NumberFormatLocal = "0,00" }   # NumberFormat из COM в русской локали не принимает "0.00"
    Start-Sleep -Milliseconds 500
    Try-Com { $excel.ActiveSheet.Columns.Item("B").ColumnWidth = 25 }
    Start-Sleep -Milliseconds 500
    Try-Com { $excel.ActiveWindow.Zoom = 85 }
    # форматирование ВНЕ выделения (выделение остаётся B5:C9) - ловит слой 2.
    # Сначала значение: ячейка попадает в отслеживаемую область, опрос её поглощает.
    Try-Com { $ws1.Range("D3").Value2 = "x" }
    Start-Sleep -Milliseconds 1600
    Try-Com { $ws1.Range("D3:E4").Font.Bold = $true }
    Try-Com { $ws1.Range("D3:E4").Interior.Color = 65280 }
    # автофильтр
    Try-Com { $ws1.Range("A1").Value2 = "Колонка" }
    Try-Com { [void]$ws1.Range("A1").AutoFilter() }
    Write-Host "TEST: жду 6 с (цикл опроса форматирования - 2 c)..."
    Start-Sleep -Seconds 6

    Write-Host "TEST: жду 12 с (периодический скриншот на 10-й секунде)..."
    Start-Sleep -Seconds 12

    Write-Host "TEST: стоп записи..."
    $excel.Run("MicReportStop", $true) | Out-Null   # тихий режим, без MsgBox и Проводника
    Start-Sleep -Seconds 1
    $tb = Get-ToggleButton
    if ($tb -and $tb.Caption -like "*старт*" -and $tb.State -eq 0) {
        Write-Host "  LOG_OK: после остановки кнопка вернулась в 'старт'"
        $script:tbBack = $true
    } else {
        Write-Host "  !! после остановки кнопка не вернулась в 'старт'"
    }
    if (-not $excel.Run("MicShot_PanelShown")) {
        Write-Host "  LOG_OK: после стопа окно панели закрыто"
        $script:panelDown = $true
    } else {
        Write-Host "  !! после стопа окно панели осталось открытым"
    }

    Write-Host "TEST: полная расшифровка (RunCloudTranscription)..."
    # именно здесь раньше вылезала компил-ошибка дубля объявлений:
    # вызываем принудительно, чтобы путь компилировался в каждом прогоне
    $excel.Run("RunCloudTranscription") | Out-Null
    Write-Host "  LOG_OK: RunCloudTranscription отработала без ошибок"

    Write-Host "TEST: диалог отправки (форма + не-настроенный endpoint)..."
    $excel.Run("PrepareSendDialog") | Out-Null    # загрузка формы без показа (компиляция)
    Write-Host "  LOG_OK: MicShotSendForm загружается"
    # изолируем пользовательский send_config: проверяем именно ветку "не настроено"
    $sendCfg = Join-Path $env:APPDATA "MicShot\send_config.txt"
    $sendCfgBak = $null
    if (Test-Path $sendCfg) {
        $sendCfgBak = $sendCfg + ".test_bak"
        Move-Item $sendCfg $sendCfgBak -Force
    }
    $stub = $excel.Run("SendReportToServer", $env:TEMP)
    if ($sendCfgBak) { Move-Item $sendCfgBak $sendCfg -Force }
    if ($stub -eq $false) { Write-Host "  LOG_OK: без endpoint SendReportToServer -> False" }
    else { Write-Host "  !! без endpoint вернулось не False"; $script:sendStubOk = $false }
    # лист "MicShot Отчёт" по умолчанию выключен (CREATE_REPORT_SHEET=False) - не проверяем
}
finally {
    try { if ($wb) { $wb.Close($false) } } catch {}
    try { $excel.Quit() } catch {}
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}

# проверка результатов
$rep = Get-ChildItem $docs -Directory -Filter "MicShot_Report_*" -ErrorAction SilentlyContinue |
    Where-Object { $before -notcontains $_.FullName } |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $rep) { Write-Host "TEST_FAIL: папка отчёта не создана"; exit 1 }

Write-Host ("TEST: отчёт в {0}" -f $rep.FullName)
# проверки полноты журнала действий
$failures = New-Object System.Collections.Generic.List[string]
if (-not $script:tbStart) { $failures.Add("кнопка не в 'старт' до записи") }
if (-not $script:tbRec) { $failures.Add("кнопка не 'стоп' при записи") }
if (-not $script:tbBack) { $failures.Add("кнопка не вернулась в 'старт'") }
if (-not $script:panelUp) { $failures.Add("панель не показана при записи") }
if (-not $script:panelDown) { $failures.Add("панель не закрыта после стопа") }
if (-not $script:liveTextOk) { $failures.Add("живой текст не дошёл до буфера") }
if (-not (Get-ChildItem $rep.FullName -Filter "narration.wav" -ErrorAction SilentlyContinue)) { $failures.Add("нет narration.wav") }
if (-not (Get-ChildItem $rep.FullName -Filter "shot_*.png" -ErrorAction SilentlyContinue)) { $failures.Add("нет скриншотов") }

# живой тик (кусок stop+save+перезапуск) не должен был оборвать запись WAV:
# при ~25 c записи файл обязан быть > 1,5 МБ (иначе склейка кусков сломана)
$wav = Get-ChildItem $rep.FullName -Filter "narration.wav" -ErrorAction SilentlyContinue | Select-Object -First 1
if ($wav -and $wav.Length -gt 1500000) {
    Write-Host ("  LOG_OK: narration.wav непрерывный ({0:N0} байт)" -f $wav.Length)
    $script:wavOk = $true
} elseif ($wav) {
    Write-Host ("  !! narration.wav слишком мал ({0} байт)" -f $wav.Length)
}
if (-not $script:wavOk) { $failures.Add("narration.wav прерван") }
if (-not $script:sendStubOk) { $failures.Add("заглушка отправки не False") }

# копия книги в папке отчёта (папка самодостаточна для отправки)
$wbCopy = Get-ChildItem $rep.FullName -Include "*.xlsx", "*.xlsm", "*.xls", "*.xlsb" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
if ($wbCopy) {
    Write-Host ("  LOG_OK: копия книги в отчёте ({0}, {1:N0} байт)" -f $wbCopy.Name, $wbCopy.Length)
} else {
    Write-Host "  !! копия книги не приложена к отчёту"
    $failures.Add("нет копии книги")
}

$logPath = Join-Path $rep.FullName "actions_log.bas"
if (Test-Path $logPath) {
    Write-Host "=== actions_log.bas (все строки) ==="
    Get-Content $logPath -Encoding Default | ForEach-Object { Write-Host $_ }
    Copy-Item (Get-ChildItem $rep.FullName -Filter "shot_*.png" | Select-Object -First 1).FullName (Join-Path $env:TEMP "MicShot_test_shot.png") -Force

    # проверки полноты журнала действий
    $raw = Get-Content $logPath -Encoding Default -Raw
    foreach ($needle in @(
        ".Select",                        # выделение ячеек
        ".Value =",                       # ввод значений
        "было:",                          # старое значение перед изменением
        "изменено форматирование",        # детектор форматирования сработал
        "полужирный", "заливка", "числовой формат", "ширина столбца",
        "вне выделения",                  # слой 2: кисть/вставка вне выделения
        "ИТОГОВЫЙ ДИФФ ФОРМАТИРОВАНИЯ",   # слой 3: снимок старта -> дифф на стопе
        "копия книги",                    # книга приложена к папке отчёта
        "масштаб", "автофильтр",
        "добавлен лист", "активирован лист"
    )) {
        if ($raw -like "*$needle*") { Write-Host ("  LOG_OK: есть '" + $needle + "'") }
        else { Write-Host ("  !! в журнале нет: " + $needle); $failures.Add("журнал: нет '" + $needle + "'") }
    }
} else {
    Write-Host "  !! нет actions_log.bas"; $failures.Add("нет actions_log.bas")
}

$trPath = Join-Path $rep.FullName "transcript.txt"
if (Test-Path $trPath) {
    Write-Host ("=== transcript.txt: " + ((Get-Content $trPath -Encoding UTF8 -Raw).Trim()) + " ===")
} else {
    Write-Host "  ВНИМАНИЕ: transcript.txt нет - расшифровка не сохранилась"; $failures.Add("нет transcript.txt")
}

Get-ChildItem $rep.FullName | ForEach-Object {
    Write-Host ("  FILE: {0}  {1} байт" -f $_.Name, $_.Length)
}

# папку тестового отчёта оставляем на месте для просмотра
if ($failures.Count -gt 0) {
    Write-Host ("FAILURES: " + ($failures -join "; "))
    Write-Host "TEST_FAIL"
    exit 1
} else {
    Write-Host "TEST_OK"
}
