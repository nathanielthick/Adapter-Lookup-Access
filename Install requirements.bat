@echo off
rem One-time setup: installs the pyodbc package and checks for the Access driver.
cd /d "%~dp0"
set PY=py
where py >nul 2>nul || set PY=python
%PY% -m pip install --upgrade pyodbc
echo.
echo Checking for the Microsoft Access ODBC driver...
%PY% -c "import pyodbc,struct; d=[x for x in pyodbc.drivers() if 'Access' in x]; b=struct.calcsize('P')*8; print('Found: '+', '.join(d) if d else 'NOT FOUND for %%d-bit Python. Install the Microsoft Access Database Engine 2016 Redistributable (%%s).' %% (b, 'x64' if b==64 else 'x86'))"
echo.
pause
