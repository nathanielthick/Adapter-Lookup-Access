ADAPTER LOOKUP - search window for the adapter Access database
================================================================

FIRST-TIME SETUP
 1. Python 3.9 or newer must be installed (python.org; tick "Add python.exe to PATH").
 2. Double-click "Install requirements.bat". It installs pyodbc and checks that
    the Microsoft Access ODBC driver is present.
    - If it says NOT FOUND: install the free "Microsoft Access Database Engine 2016
      Redistributable" from Microsoft. Pick the version (x64 / x86) that matches
      the bitness the script reports. If you have 64-bit Office, use 64-bit Python.
 3. Put your .accdb (or .mdb) file in this folder.

RUNNING
 Double-click "Adapter Lookup.bat". It opens the database in this folder (or asks
 you to pick one, and remembers your choice). File > Open database... to switch.
 Try it without a database:  py adapter_lookup.py --demo

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
