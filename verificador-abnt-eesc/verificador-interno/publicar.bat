@echo off
REM Coloca o Verificador ABNT na internet (site com link para os alunos), rodando neste PC.
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0publicar.ps1" %*
echo.
pause
