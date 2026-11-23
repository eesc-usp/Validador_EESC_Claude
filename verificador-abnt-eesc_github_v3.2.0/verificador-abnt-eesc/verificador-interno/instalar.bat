@echo off
REM Instala ou repara o Verificador ABNT (Python, bibliotecas, Ollama e modelo de IA).
REM Pode rodar quantas vezes quiser: so instala o que falta e corrige o que estiver errado.
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0instalar.ps1" %*
echo.
pause
