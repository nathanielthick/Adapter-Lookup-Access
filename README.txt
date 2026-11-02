ADAPTER LOOKUP - search window for the adapter Access database
================================================================

There are two ways to run it:
 A. "Adapter Lookup.exe"  - no Python needed. Easiest for most users.
 B. "Adapter Lookup.bat"  - runs adapter_lookup.py with Python (for editing the code).

OPTION A - THE EXE (no Python needed)
 1. Copy "Adapter Lookup.exe" into the same folder as the .accdb / .mdb file
    (e.g. next to SIGBOARDS.mdb).
 2. The Microsoft Access ODBC driver must be installed. Most PCs with Office
    already have it. If the exe says no driver was found, install the free
    "Microsoft Access Database Engine 2016 Redistributable" from Microsoft -
    the x64 (64-bit) version, because the exe is 64-bit.
 3. Double-click "Adapter Lookup.exe". It may take a few seconds to start
    (it unpacks itself each time). Windows SmartScreen may warn about an
    unrecognised app the first time - click "More info" > "Run anyway".
 Settings are saved in adapter_lookup_settings.json next to the exe.

OPTION B - PYTHON SCRIPT
 First-time setup:
 1. Python 3.9 or newer must be installed (python.org; tick "Add python.exe to PATH").
 2. Double-click "Install requirements.bat". It installs pyodbc and checks that
    the Microsoft Access ODBC driver is present.
    - If it says NOT FOUND: install the free "Microsoft Access Database Engine 2016
      Redistributable" from Microsoft. Pick the version (x64 / x86) that matches
      the bitness the script reports. If you have 64-bit Office, use 64-bit Python.
 3. Put your .accdb (or .mdb) file in this folder.
 Running:
 Double-click "Adapter Lookup.bat".
 Try it without a database:  py adapter_lookup.py --demo

Either way, it opens the database in its folder (or asks you to pick one, and
remembers your choice). File > Open database... to switch.

REBUILDING THE EXE (after changing adapter_lookup.py)
 Requires Python (see Option B). Double-click "Build EXE.bat". It installs
 PyInstaller and writes a new "Adapter Lookup.exe" into this folder. The build
 and dist folders it creates can be deleted. The exe matches the bitness of
 the Python used to build it (64-bit Python -> 64-bit exe -> needs the x64 driver).

SEARCHING
 - Part number / keywords box: results update as you type. Several words = all
   must match. * is a wildcard (ADP-*-1005). Use the "in" box to search one
   column only (e.g. just the part number column).
 - Filters panel (left) is built from your table's columns:
     pick-lists  for columns with up to 60 distinct values (connector type, gender...)
                 picking a value = exact match; typing = "contains"
     X to Y      for number columns (pin count, length, size...) - leave either side blank
     Yes/No      for checkbox columns
     dates       YYYY-MM-DD or MM/DD/YYYY
   Press Enter in a box (or pick from a list) to apply. "Clear all" resets everything.
 - Table box: switch between tables and saved Access queries.
 - Click a column header to sort. Double-click a row for all its details.
 - Right-click / Ctrl+C to copy rows or just part numbers.
 - Export CSV saves the current results (opens in Excel).

The database is opened read-only. Nothing in it is ever changed.
