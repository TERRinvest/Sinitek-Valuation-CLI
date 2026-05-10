@echo off
setlocal
REM Convenience wrapper: bypasses PowerShell execution policy without requiring
REM `Set-ExecutionPolicy`. For the recommended entry point, call sinitek.ps1
REM directly from PowerShell. See README for details.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0sinitek.ps1" %*
exit /b %ERRORLEVEL%
