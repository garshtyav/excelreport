$ErrorActionPreference = "Continue"
$opt = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"
$xlamPath = Join-Path $env:APPDATA "Microsoft\AddIns\MicShot.xlam"

if (Test-Path $opt) {
    $op = Get-Item $opt
    foreach ($n in $op.GetValueNames()) {
        if ($n -cmatch "^OPEN\d*$" -and ($op.GetValue($n) -match "MicShot")) {
            Set-ItemProperty -Path $opt -Name $n -Value ('"' + $xlamPath + '"')
            Write-Host "[ok] OPEN $n теперь содержит полный путь"
        }
    }
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    [void]$excel.Workbooks.Add()
    foreach ($a in $excel.AddIns) {
        if ($a.Name -ieq "MicShot.xlam") {
            $o = "?"
            try { $o = $a.IsOpen } catch {}
            Write-Host ("ADDIN: Installed={0} IsOpen={1}" -f $a.Installed, $o)
        }
    }
    $tb = $false
    try { [void]$excel.CommandBars.Item("MicShot"); $tb = $true } catch {}
    Write-Host ("Toolbar MicShot: {0}" -f $tb)
}
finally {
    try { $excel.Quit() } catch {}
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
Start-Sleep -Milliseconds 800
$op = Get-Item $opt
foreach ($n in $op.GetValueNames()) {
    if ($n -cmatch "^OPEN\d*$") { Write-Host ("REG ПОСЛЕ ВЫХОДА: {0} = {1}" -f $n, $op.GetValue($n)) }
}
