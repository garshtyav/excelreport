# Пробная сборка MicShot.xlam во временную папку - НЕ трогает установленную
# надстройку в AddIns (её нельзя перезаписать, пока Excel запущен).
# Проверяет: код вставляется, компилируется, панель создаётся.
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$raw = [System.IO.File]::ReadAllText((Join-Path $root "build_micshot.ps1"))
function Get-Heredoc([string]$name) {
    $pattern = "(?s)\$" + $name + " = @'\r?\n(.*?)\r?\n'@"
    $m = [regex]::Match($raw, $pattern)
    if (-not $m.Success) { throw "heredoc $name не найден" }
    $m.Groups[1].Value
}
$vbaCode       = Get-Heredoc "vbaCode"
$vbaClass      = Get-Heredoc "vbaClass"
$vbaClassSapi  = Get-Heredoc "vbaClassSapi"
$vbaForm       = Get-Heredoc "vbaForm"
$vbaSendForm   = Get-Heredoc "vbaSendForm"

$outDir = Join-Path $env:TEMP "MicShot_build"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
# имя НЕ MicShot.xlam: Excel путает staging-файл с зарегистрированной
# установленной надстройкой того же имени и Workbooks.Open падает
$xlam = Join-Path $outDir "MicShot_staging.xlam"
Remove-Item $xlam -Force -ErrorAction SilentlyContinue

$secKey = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"
# staging-папка должна быть доверенной, иначе Excel заблокирует макросы .xlam
$tlStage = "$secKey\Trusted Locations\LocationMicShotStage"
New-Item -Path $tlStage -Force | Out-Null
Set-ItemProperty -Path $tlStage -Name Path -Value "$outDir\"
Set-ItemProperty -Path $tlStage -Name AllowSubFolders -Value 1 -Type DWord
Write-Host "[ok] Доверенное расположение (staging): $outDir"
$oldVbom = $null
$prop = Get-ItemProperty -Path $secKey -ErrorAction SilentlyContinue
if ($prop -and ($prop.PSObject.Properties.Name -contains "AccessVBOM")) { $oldVbom = [int]$prop.AccessVBOM }
$excel = $null
try {
New-Item -Path $secKey -Force | Out-Null
Set-ItemProperty -Path $secKey -Name AccessVBOM -Value 1 -Type DWord

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Add()
    $comp = $wb.VBProject.VBComponents.Add(1)    # 1 = стандартный модуль
    $comp.Name = "MicShot"
    [void]$comp.CodeModule.AddFromString($vbaCode)
    if ($comp.CodeModule.CountOfLines -lt 10) { throw "Код VBA не вставлен в модуль" }

    $cls = $wb.VBProject.VBComponents.Add(2)     # 2 = модуль класса
    $cls.Name = "AppEvents"
    [void]$cls.CodeModule.AddFromString($vbaClass)
    if ($cls.CodeModule.CountOfLines -lt 5) { throw "Код класса не вставлен" }

    try { [void]$wb.VBProject.References.AddFromGuid("{C866CA3A-32F7-11D2-9602-00C04F8EE628}", 5, 4) }
    catch { Write-Warning "SpeechLib: $($_.Exception.Message)" }

    $cls2 = $wb.VBProject.VBComponents.Add(2)
    $cls2.Name = "CSapiEvents"
    [void]$cls2.CodeModule.AddFromString($vbaClassSapi)
    if ($cls2.CodeModule.CountOfLines -lt 5) { throw "Код CSapiEvents не вставлен" }

    # UserForm - та же, что в build_micshot.ps1 (копия блока создания)
    $frm = $wb.VBProject.VBComponents.Add(3)     # 3 = UserForm
    $frm.Name = "MicShotForm"
    $p = $frm.Properties
    $p.Item("Caption").Value = "MicShot - субтитры"
    $p.Item("Width").Value = 700
    $p.Item("Height").Value = 64
    $p.Item("ShowModal").Value = $false
    if ($p.Item("ShowModal").Value -ne $false) { throw "ShowModal не применился" }
    $d = $frm.Designer
    $li = $d.Controls.Add("Forms.Label.1"); $li.Name = "lblInterim"
    $li.Left = 10; $li.Top = 4; $li.Width = 680; $li.Height = 20
    $li.Caption = ""
    $li.ForeColor = [uint32]8421504
    $lb = $d.Controls.Add("Forms.Label.1"); $lb.Name = "lblLive"
    $lb.Left = 10; $lb.Top = 24; $lb.Width = 680; $lb.Height = 36
    $lb.Caption = ""
    [void]$frm.CodeModule.AddFromString($vbaForm)
    if ($frm.CodeModule.CountOfLines -lt 10) { throw "Код формы не вставлен" }

    # MicShotSendForm - диалог отправки (копия блока build_micshot.ps1)
    $frm2 = $wb.VBProject.VBComponents.Add(3)
    $frm2.Name = "MicShotSendForm"
    $p2 = $frm2.Properties
    $p2.Item("Caption").Value = "MicShot - отправка отчёта"
    $p2.Item("Width").Value = 520
    $p2.Item("Height").Value = 420
    $d2 = $frm2.Designer
    $h = $d2.Controls.Add("Forms.Label.1"); $h.Name = "lblHeader"
    $h.Left = 10; $h.Top = 6; $h.Width = 500; $h.Height = 16
    $h.Caption = "В службу поддержки будет отправлено (папка целиком):"
    $lb = $d2.Controls.Add("Forms.ListBox.1"); $lb.Name = "lbxFiles"
    $lb.Left = 10; $lb.Top = 26; $lb.Width = 500; $lb.Height = 130
    $fc = $d2.Controls.Add("Forms.Label.1"); $fc.Name = "lblCount"
    $fc.Left = 10; $fc.Top = 160; $fc.Width = 500; $fc.Height = 14
    $fc.Caption = ""
    $cl = $d2.Controls.Add("Forms.Label.1"); $cl.Name = "lblComment"
    $cl.Left = 10; $cl.Top = 182; $cl.Width = 500; $cl.Height = 16
    $cl.Caption = "Комментарий (текстом - если микрофона не было или хотите дополнить):"
    $tc = $d2.Controls.Add("Forms.TextBox.1"); $tc.Name = "tbxComment"
    $tc.Left = 10; $tc.Top = 202; $tc.Width = 500; $tc.Height = 120
    $tc.MultiLine = $true
    $bs = $d2.Controls.Add("Forms.CommandButton.1"); $bs.Name = "btnSend"
    $bs.Caption = "Отправить"
    $bs.Left = 250; $bs.Top = 334; $bs.Width = 125; $bs.Height = 30
    $bk = $d2.Controls.Add("Forms.CommandButton.1"); $bk.Name = "btnSkip"
    $bk.Caption = "Не отправлять"
    $bk.Left = 385; $bk.Top = 334; $bk.Width = 125; $bk.Height = 30
    [void]$frm2.CodeModule.AddFromString($vbaSendForm)
    if ($frm2.CodeModule.CountOfLines -lt 10) { throw "Код MicShotSendForm не вставлен" }

    $excel.Run("CompileCheck") | Out-Null         # форс-компиляция классов
    $wb.SaveAs($xlam, 55)                        # 55 = xlOpenXMLAddIn
    $wb.Close($false)
    Write-Host "[ok] Пробная сборка: $xlam"

    # повторное открытие: Auto_Open должен создать панель, код - выполняться
    $wb2 = $excel.Workbooks.Open($xlam)
    Start-Sleep -Milliseconds 800
    $excel.Run("CompileCheck") | Out-Null

    # выгружаем ФАКТИЧЕСКИЙ код модулей из собранного .xlam для линтера
    foreach ($cn in @("MicShot", "AppEvents", "CSapiEvents", "MicShotForm", "MicShotSendForm")) {
        try {
            $cm = $wb2.VBProject.VBComponents.Item($cn).CodeModule
            $out = Join-Path $outDir ($cn + "_extracted.txt")
            [System.IO.File]::WriteAllText($out, $cm.Lines(1, $cm.CountOfLines), [System.Text.Encoding]::Unicode)
            Write-Host "[ok] Выгружен код: $cn ($($cm.CountOfLines) строк)"
        } catch { Write-Warning "Не выгрузить $cn : $($_.Exception.Message)" }
    }

    $tb = $false
    try { [void]$excel.CommandBars.Item("MicShot"); $tb = $true } catch {}
    if ($tb) { Write-Host "[ok] Панель MicShot создана - код выполняется" }
    else { Write-Warning "Панель MicShot не найдена" }
    $wb2.Close($false)
    Write-Host "STAGING_OK"
}
finally {
    try { $excel.Quit() } catch {}
    if ($null -ne $excel) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    if ($null -ne $oldVbom) { Set-ItemProperty -Path $secKey -Name AccessVBOM -Value $oldVbom -Type DWord }
    else { Remove-ItemProperty -Path $secKey -Name AccessVBOM -ErrorAction SilentlyContinue }
}
