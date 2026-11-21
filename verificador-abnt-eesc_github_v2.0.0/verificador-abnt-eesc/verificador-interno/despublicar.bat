@echo off
REM Tira do ar o site publico do Verificador ABNT (o app interno continua funcionando).
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0publicar.ps1" -Desativar
echo.
pause
