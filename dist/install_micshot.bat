@echo off
rem Установка MicShot (Excel): без прав администратора
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install_micshot.ps1"
