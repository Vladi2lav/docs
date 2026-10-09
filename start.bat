@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Учёт времени педагога: http://localhost:8080  (закройте окно, чтобы остановить)

where python >nul 2>&1
if %ERRORLEVEL% equ 0 (
    start "" "http://localhost:8080"
    python serve.py
    exit /b
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0serve.ps1"
