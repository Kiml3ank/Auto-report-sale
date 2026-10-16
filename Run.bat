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
rem 1) erp_export.py downloads today's files from KolaoERP into Input (STOP window or Esc stops it)
rem 2) sales_report.py builds the report from Input - only when the download worked
rem With Diana installed (D:\Diana, her own add-on), she opens and runs them when you click her.
set "PY=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
if exist "D:\Diana\diana.py" (
  "%PY%" -X utf8 "D:\Diana\diana.py"
) else (
  "%PY%" -X utf8 "%~dp0erp_export.py" && "%PY%" -X utf8 "%~dp0sales_report.py"
)
echo.
pause
