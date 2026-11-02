@echo off
rem Builds "Adapter Lookup.exe" from adapter_lookup.py using PyInstaller.
cd /d "%~dp0"
set PY=py
where py >nul 2>nul || set PY=python
%PY% -m pip install --upgrade pyinstaller pyodbc
%PY% -m PyInstaller --onefile --windowed --noconfirm --name "Adapter Lookup" --hidden-import pyodbc adapter_lookup.py
if errorlevel 1 goto fail
copy /y "dist\Adapter Lookup.exe" "Adapter Lookup.exe" >nul
echo.
echo Done. "Adapter Lookup.exe" is in this folder, next to SIGBOARDS.mdb.
echo You can delete the build and dist folders and the .spec file.
pause
exit /b 0
:fail
echo.
echo Build failed - see the messages above.
pause
exit /b 1
