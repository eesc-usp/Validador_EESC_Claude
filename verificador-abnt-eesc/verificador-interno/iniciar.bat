@echo off
REM Abre o Verificador ABNT. Se a versao mudou ou faltar algo, repara antes de abrir.
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0instalar.ps1" -Rapido -Iniciar
if errorlevel 1 pause
