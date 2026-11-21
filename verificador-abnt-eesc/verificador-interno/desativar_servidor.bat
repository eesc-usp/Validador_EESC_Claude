@echo off
REM Para o servidor da rede e tira da inicializacao do Windows (o app continua funcionando neste PC).
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0instalar.ps1" -DesativarServidor
echo.
pause
