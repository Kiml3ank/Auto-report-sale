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
rem 1) download today's files from KolaoERP into Input (STOP window or Esc stops it)
"%LOCALAPPDATA%\Programs\Python\Python313\python.exe" -X utf8 "%~dp0erp_export.py"
if errorlevel 1 (
  echo.
  echo ERP download stopped - the report was NOT run. See the reason above.
  echo.
  pause
  exit /b 1
)
rem 2) build the report from Input
"%LOCALAPPDATA%\Programs\Python\Python313\python.exe" -X utf8 "%~dp0sales_report.py"
echo.
pause
