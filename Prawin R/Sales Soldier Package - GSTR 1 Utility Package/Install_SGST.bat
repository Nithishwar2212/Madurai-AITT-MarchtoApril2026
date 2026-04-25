@echo off
setlocal
cd /d "%~dp0"

echo Installing Sales Soldier requirements...
where py >nul 2>nul
if %errorlevel%==0 (
    py -m pip install --upgrade pip
    py -m pip install -r requirements.txt
) else (
    python -m pip install --upgrade pip
    python -m pip install -r requirements.txt
)

echo.
echo Installation complete.
echo You can now open Sales Soldier using Sales Soldier.lnk, Sales Soldier.vbs, or Run_Sales_Soldier.bat
pause
