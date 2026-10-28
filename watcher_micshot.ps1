# Ждёт закрытия Excel (максимум 4 часа), затем запускает finalize_micshot.ps1
$finalize = 'C:\Users\Cartman\.zcode\workspace\default\MicShot\finalize_micshot.ps1'
$deadline = (Get-Date).AddHours(4)
while ((Get-Date) -lt $deadline) {
    if (-not (Get-Process -Name EXCEL -ErrorAction SilentlyContinue)) {
        Start-Sleep -Seconds 3
        if (-not (Get-Process -Name EXCEL -ErrorAction SilentlyContinue)) {
            Write-Host "Excel closed - registering add-in..."
            & $finalize
            exit 0
        }
    }
    Start-Sleep -Seconds 15
}
Write-Host "EXCEL_STILL_RUNNING_AFTER_4H"
exit 3
