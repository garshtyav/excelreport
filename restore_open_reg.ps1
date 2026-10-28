$ErrorActionPreference = "Continue"
$xlamPath = Join-Path $env:APPDATA "Microsoft\AddIns\MicShot.xlam"
$xlStart  = Join-Path $env:APPDATA "Microsoft\Excel\XLSTART"

# 1. Убираем XLSTART-копию (возвращаемся на стандартную регистрацию через OPEN)
Remove-Item (Join-Path $xlStart "MicShot.xlam") -ErrorAction SilentlyContinue
Write-Host "[ok] XLSTART-копия удалена"

# 2. Восстанавливаем регистрацию через finalize (работает при закрытом Excel)
& 'C:\Users\Cartman\.zcode\workspace\default\MicShot\finalize_micshot.ps1'
