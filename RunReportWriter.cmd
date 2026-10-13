@echo off
setlocal
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" goto install
python -m venv .venv
if errorlevel 1 goto failed
:install
".venv\Scripts\python.exe" -m pip install --disable-pip-version-check --no-input -e .
if errorlevel 1 goto failed
".venv\Scripts\python.exe" -m trw.gui
if errorlevel 1 goto failed
exit /b 0
:failed
echo.
echo Report Writer could not start. The error appears above.
echo Python 3.11 or newer is required. Automatic TOC and PDF output also need LibreOffice.
pause
exit /b 1
