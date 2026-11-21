CSV import
==========

The helpers below implement
:meth:`gspread.worksheet.Worksheet.import_csv`, which loads a CSV payload
into a bounded A1 data area of an existing worksheet. The exceptions it
raises are documented under :doc:`exceptions`.

Import
------

.. autofunction:: gspread.csv_import.import_csv

Parsing
-------

.. autofunction:: gspread.csv_import.parse_csv

Result
------

.. autoclass:: gspread.csv_import.CsvImportResult
   :members:
