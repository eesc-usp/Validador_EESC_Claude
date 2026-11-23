@echo off
REM Envia o projeto para o GitHub (cria o repositorio na primeira vez). Aceita os mesmos parametros do .ps1.
cd /d "%~dp0.."
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0enviar_para_github.ps1" %*
echo.
pause
