ADAPTER LOOKUP - search window for the adapter Access database
================================================================

There are two ways to run it:
 A. "Adapter Lookup.exe"  - no Python needed. Easiest for most users.
 B. "Adapter Lookup.bat"  - runs adapter_lookup.py with Python (for editing the code).

OPTION A - THE EXE (no Python needed)
 1. Copy "Adapter Lookup.exe" anywhere (e.g. your Desktop). It opens the
    shared database on the Z: drive (see DATABASE below).
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
 Running:
 Double-click "Adapter Lookup.bat".
 Try it without a database:  py adapter_lookup.py --demo

DATABASE
 Either way, it opens the shared database:
     Z:\Engineering\Test Engineering\Test\Test Board DB\SIGBOARDS.mdb
 (only its CONNECTORS table is used). If the Z: drive isn't connected, it looks
 for a .mdb / .accdb next to the program or asks you to pick one, and remembers
 your choice. File > Open database... to switch.

REBUILDING THE EXE (after changing adapter_lookup.py)
 Requires Python (see Option B). Double-click "Build EXE.bat". It installs
 PyInstaller and writes a new "Adapter Lookup.exe" into this folder. The build
 and dist folders it creates can be deleted. The exe matches the bitness of
 the Python used to build it (64-bit Python -> 64-bit exe -> needs the x64 driver).

SEARCHING
 - Search box (top): results update as you type. Several words = all must
   match. * is a wildcard (ADP-*-1005). Use the "in" box to search one column
   only. Ctrl+F jumps to the search box, Esc clears it.
 - Filters panel (left). The most useful filters are always shown; click
   "More filters" for the rest.
     pick-lists  for columns with up to 250 distinct values (Category, Contact Type...)
                 picking a value = exact match; typing = "contains"
     min to max  for number columns (Positions, Qty) - leave either side blank
     Yes/No      for checkbox columns
   Filters apply as you type. Each active filter shows as a blue chip above the
   results - click its x to remove it. "Reset all" / "Clear all" resets everything.
 - Group by adapter: shows every connector on each matching adapter, with the
   connectors that matched highlighted. Hide scrapped: leaves out scrapped rows.
 - Details panel (right) shows every field of the selected row. Double-click a
   value there to copy it. View > Show / hide details panel to turn it off.
 - Click a column header to sort. Right-click / Ctrl+C to copy rows or part numbers.
 - Zoom the results table in or out: Ctrl + mouse wheel over the table,
   Ctrl+plus / Ctrl+minus, the - 100% + buttons at the bottom right, or the View
   menu. Ctrl+0 (or clicking the %) resets to 100%. The zoom is remembered.
 - Export CSV saves the current results (opens in Excel).

EDITING (adding, changing and deleting connectors)
 The app opens the database read-only. To make changes, click "Enable editing"
 in the top bar. It offers to save a backup copy of the database first (yes is
 recommended the first time). While editing is on an orange bar is shown, and:
 - "+ New connector" (Ctrl+N) adds a row. If an adapter or one of its connectors
   is selected, the adapter number is filled in for you. Right-click >
   "Add connector to this adapter" does the same.
 - "Edit" (F2, or double-click a row) changes the selected connector. Changed
   fields turn orange in the form. Ctrl+S or Save writes the change.
 - "Duplicate" (Ctrl+D) copies the selected connector into a new row to adjust.
 - "Delete" (Del key) deletes the selected rows after asking. Selecting an
   adapter heading selects all of its connectors.
 - Ctrl+Z / "Undo last change" undoes your changes one at a time, until you close
   the app.
 - Fields marked * are required. Drop-downs offer the existing values; if you
   type a value that only differs in capitals from an existing one (e.g.
   "d-sub"), the existing spelling is used, so the data stays consistent.
 - Every change is written to SIGBOARDS_change_log.csv next to the database
   (or next to the app if that folder can't be written). It records when, who
   (Windows user name), what was done and the before/after values.
   Edit > Open change log opens it in Excel.
 - The CONNECTORS table has no ID column, so a row is found by matching all of
   its values. If two rows are completely identical, the app won't change or
   delete either of them (it can't tell them apart) - fix those in Access.
   A change is also refused if someone else changed the row since you loaded
   it; press F5 to reload and try again.
 - Click "Editing ON - click to stop" to go back to read-only.
 File > Back up database now... saves a copy at any time.
