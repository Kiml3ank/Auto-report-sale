@echo off
rem Double-click to run the daily sales report (Input -> adjust -> output).
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
if not exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" (
  echo Python is not installed on this PC yet. Double-click Setup.bat first, then run this again.
  echo.
  pause
  exit /b 1
)
"%LOCALAPPDATA%\Programs\Python\Python313\python.exe" -X utf8 "%~dp0sales_report.py"
echo.
pause
