# Конвейер: ждать закрытия Excel -> пересобрать MicShot.xlam (фикс PrintWindow) -> проверить скриншот
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$deadline = (Get-Date).AddHours(4)
while ((Get-Date) -lt $deadline) {
    if (-not (Get-Process -Name EXCEL -ErrorAction SilentlyContinue)) {
        Start-Sleep -Seconds 3
        if (-not (Get-Process -Name EXCEL -ErrorAction SilentlyContinue)) { break }
    }
    Start-Sleep -Seconds 15
}
if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
    Write-Host "ORCH: Excel так и не закрыли за 4 часа - отмена"
    exit 3
}
Write-Host "=== BUILD ==="
& (Join-Path $dir "build_micshot.ps1")
Write-Host "=== REPORT TEST ==="
& (Join-Path $dir "test_report.ps1")
