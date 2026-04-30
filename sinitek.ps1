& powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'sinitek-cli.ps1') @args
exit $LASTEXITCODE
