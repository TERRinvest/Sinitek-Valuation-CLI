@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0sinitek-cli.ps1" %*
exit /b %ERRORLEVEL%
