@echo off
rem Uninstall MicShot (Excel add-in), no admin rights needed
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall_micshot.ps1"
