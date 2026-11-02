@echo off
rem Starts the Adapter Lookup search window.
cd /d "%~dp0"
where pyw >nul 2>nul
if %errorlevel%==0 (
    start "" pyw "%~dp0adapter_lookup.py" %*
    exit /b
)
where pythonw >nul 2>nul
if %errorlevel%==0 (
    start "" pythonw "%~dp0adapter_lookup.py" %*
    exit /b
)
echo Python was not found on this computer.
echo Install it from https://www.python.org/downloads/ (tick "Add python.exe to PATH"),
echo then double-click "Install requirements.bat".
pause
