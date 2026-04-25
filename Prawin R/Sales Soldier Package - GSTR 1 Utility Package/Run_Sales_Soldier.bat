@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (
    start "" pyw "%~dp0gstr1_tool.py"
) else (
    start "" pythonw "%~dp0gstr1_tool.py"
)
