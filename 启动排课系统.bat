@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "BUNDLED_PY=C:\Users\Administrator\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

if exist "%BUNDLED_PY%" (
  "%BUNDLED_PY%" app.py
) else (
  python app.py
)

if errorlevel 1 pause
