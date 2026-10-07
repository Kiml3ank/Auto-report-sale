@echo off
rem Run once on a new PC: installs Python and the libraries the daily sales report needs.
rem Needs an internet connection. Microsoft Excel must already be installed (this cannot install it).
setlocal
chcp 65001 >nul
set "PYVER=3.13.15"
set "PY=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
set "INSTALLER=%TEMP%\python-%PYVER%-setup.exe"

echo ============================================
echo   Setup for the daily sales report
echo ============================================
echo.

echo [1/5] Python
if exist "%PY%" goto have_python
echo       not found - downloading Python %PYVER%, about 30 MB ...
powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol='Tls12'; $ProgressPreference='SilentlyContinue'; Invoke-WebRequest -UseBasicParsing 'https://www.python.org/ftp/python/%PYVER%/python-%PYVER%-amd64.exe' -OutFile '%INSTALLER%'"
if errorlevel 1 goto fail_download
if not exist "%INSTALLER%" goto fail_download
echo       installing, this takes a minute or two ...
"%INSTALLER%" /quiet InstallAllUsers=0 PrependPath=0 Include_launcher=0 Include_test=0
del "%INSTALLER%" >nul 2>&1
if not exist "%PY%" goto fail_python
echo       installed.
goto libraries
:have_python
echo       already installed.

:libraries
echo.
echo [2/5] Libraries (openpyxl, xlrd, pywin32)
"%PY%" -m pip install --disable-pip-version-check --no-warn-script-location -r "%~dp0requirements.txt"
if errorlevel 1 goto fail_libraries

echo.
echo [3/5] Check that the libraries load
"%PY%" -c "import openpyxl, xlrd, pythoncom, win32com.client; print('      OK')"
if errorlevel 1 goto fail_libraries

echo.
echo [4/5] Microsoft Excel
reg query "HKCR\Excel.Application" >nul 2>&1
if errorlevel 1 goto no_excel
echo       found.
goto folders
:no_excel
echo       NOT FOUND. The report needs Microsoft Excel on this PC. Install Excel, then the report will work.
set "WARN=1"

:folders
echo.
echo [5/5] Folders
for %%D in (Input adjust Main output) do if not exist "%~dp0%%D" mkdir "%~dp0%%D"
echo       Input, adjust, Main, output are ready.

echo.
echo ============================================
if defined WARN (echo   Setup finished, but Excel is missing - see above.) else (echo   Setup finished. Everything is ready.)
echo   Next: put the main workbook in Main, the exports in Input, then double-click Run.bat
echo ============================================
goto end

:fail_download
echo.
echo   FAILED: Python could not be downloaded. Check the internet connection and run Setup.bat again.
echo   Or install Python %PYVER% by hand from https://www.python.org/downloads/ with the default options.
goto end
:fail_python
echo.
echo   FAILED: Python did not install. Run Setup.bat again, or install Python %PYVER% by hand
echo   from https://www.python.org/downloads/ with the default options ("Install Now").
goto end
:fail_libraries
echo.
echo   FAILED: the libraries could not be installed. Check the internet connection and run Setup.bat again.

:end
echo.
pause
