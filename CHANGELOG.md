# Changelog

## 2.0.1

- Added native OLE/BIFF8 metadata inspection using the shared metadata schema.
- Added SummaryInformation, DocumentSummaryInformation, and typed custom-property reading and writing.
- Added atomic XLS metadata updates for document, revision, application, workbook visibility/active-sheet/date-system, and calculation settings.
- Added personal-information removal, unknown-stream preservation, digital-signature safeguards, and forensic stream hashes.
- Added document properties to newly written native XLS workbooks.
- Added LibreOffice-validated metadata round-trip and preservation regressions.
- Updated the Core dependency to `^2.0.5`.

## 2.0.0

- Replaced the external XLS adapter with an independent native BIFF8 engine.
- Added native CFB/OLE header, FAT, DIFAT, MiniFAT, directory, and stream processing.
- Added native BIFF8 workbook globals, SST/CONTINUE, worksheet, date, style, and formula processing.
- Added native BIFF8 workbook and worksheet writing.
- Added malformed-file limits, boundary validation, and cycle detection.
- Removed the old compatibility reader and writer classes.
- Added LibreOffice fixtures and native smoke tests.
