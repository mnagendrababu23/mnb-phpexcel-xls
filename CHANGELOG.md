# Changelog

## 2.0.0

- Replaced the external XLS adapter with an independent native BIFF8 engine.
- Added native CFB/OLE header, FAT, DIFAT, MiniFAT, directory, and stream processing.
- Added native BIFF8 workbook globals, SST/CONTINUE, worksheet, date, style, and formula processing.
- Added native BIFF8 workbook and worksheet writing.
- Added malformed-file limits, boundary validation, and cycle detection.
- Removed the old compatibility reader and writer classes.
- Added LibreOffice fixtures and native smoke tests.
