@echo off
REM Transforma este computador no servidor do Verificador ABNT para a rede da biblioteca.
REM Todos acessam pelo navegador; atualizacoes feitas aqui valem para todos.
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0instalar.ps1" -Servidor
echo.
pause
