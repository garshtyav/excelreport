$ErrorActionPreference = "Continue"
Write-Host "=== REGISTRY ==="
$opt = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"
if (Test-Path $opt) {
    $op = Get-Item $opt
    foreach ($n in $op.GetValueNames()) {
        if ($n -cmatch "^OPEN\d*$") { Write-Host ("REG {0} = {1}" -f $n, $op.GetValue($n)) }
    }
} else { Write-Host "нет ключа Options" }
$tl = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations"
Get-ChildItem $tl -ErrorAction SilentlyContinue | ForEach-Object {
    $p = Get-ItemProperty $_.PSPath
    Write-Host ("TRUST {0}: Path={1} AllowSubFolders={2}" -f $_.PSChildName, $p.Path, $p.AllowSubFolders)
}

if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) { Write-Host "EXCEL_RUNNING - живой тест пропущен"; exit 0 }

Write-Host "=== LIVE INSTANCE ==="
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    [void]$excel.Workbooks.Add()
    Write-Host ("AutomationSecurity={0}" -f $excel.AutomationSecurity)
    foreach ($a in $excel.AddIns) {
        $inst = "?"; $openSt = "?"
        try { $inst = $a.Installed } catch {}
        try { $openSt = $a.IsOpen } catch {}
        Write-Host ("ADDIN: {0} | Installed={1} | IsOpen={2}" -f $a.Name, $inst, $openSt)
    }
    $tb = $false
    try { [void]$excel.CommandBars.Item("MicShot"); $tb = $true } catch {}
    Write-Host ("Toolbar MicShot: {0}" -f $tb)
}
finally {
    try { $excel.Quit() } catch {}
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
