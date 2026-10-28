$ErrorActionPreference = "Continue"
$xlamPath = Join-Path $env:APPDATA "Microsoft\AddIns\MicShot.xlam"
$xlStart  = Join-Path $env:APPDATA "Microsoft\Excel\XLSTART"

if (-not (Test-Path $xlamPath)) { throw "Не найден $xlamPath" }
New-Item -ItemType Directory -Force -Path $xlStart | Out-Null

# 1. Убираем OPEN-записи MicShot (чтобы не было двойной загрузки)
$opt = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"
if (Test-Path $opt) {
    $op = Get-Item $opt
    $keep = @()
    foreach ($n in @($op.GetValueNames() | Where-Object { $_ -cmatch "^OPEN\d*$" })) {
        $v = $op.GetValue($n)
        if ($v -match "MicShot") { Remove-ItemProperty -Path $opt -Name $n }
        else { $keep += $v }
    }
    foreach ($n in @($op.GetValueNames() | Where-Object { $_ -cmatch "^OPEN\d*$" })) {
        Remove-ItemProperty -Path $opt -Name $n
    }
    $j = 0
    foreach ($v in $keep) {
        $newName = $(if ($j -eq 0) { "OPEN" } else { "OPEN$j" })
        Set-ItemProperty -Path $opt -Name $newName -Value $v
        $j++
    }
    Write-Host "[ok] OPEN-записи MicShot удалены"
}

# 2. Кладём надстройку в XLSTART
Copy-Item $xlamPath (Join-Path $xlStart "MicShot.xlam") -Force
Write-Host "[ok] Скопировано в $xlStart"

if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) { Write-Host "EXCEL_RUNNING - живой тест пропущен"; exit 0 }

# 3. Живой тест: загружается ли в свежем экземпляре
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    [void]$excel.Workbooks.Add()
    $loaded = $false
    foreach ($wbo in $excel.Workbooks) {
        if ($wbo.Name -ieq "MicShot.xlam") { $loaded = $true }
    }
    Write-Host ("WORKBOOK MicShot.xlam загружен: {0}" -f $loaded)
    $tb = $false
    try { [void]$excel.CommandBars.Item("MicShot"); $tb = $true } catch {}
    Write-Host ("Toolbar MicShot: {0}" -f $tb)
}
finally {
    try { $excel.Quit() } catch {}
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
