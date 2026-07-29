# Native XLS roadmap

## Completed native milestone

- independent CFB/OLE reader
- FAT, DIFAT, MiniFAT, directory, and stream parsing
- independent CFB writer for ordinary workbook sizes
- BIFF8 workbook globals
- native worksheet value reader
- native worksheet writer
- Unicode SST and `CONTINUE` handling
- dates and date-times
- common formulas and cached formula values
- row iteration, projection, and safety limits
- LibreOffice-generated fixture compatibility
- LibreOffice open/convert validation of MNB-generated XLS files

## Next reader milestones

1. Shared and array formula records.
2. External sheet and defined-name resolution.
3. Complete XF/font/fill/border exposure through cell-detail APIs.
4. Hyperlinks, comments, and data validation inspection.
5. Legacy BIFF5/BIFF7 compatibility where practical.
6. FILEPASS detection and explicit encryption diagnostics.

## Next writer milestones

1. Full style table generation from `WorksheetData` styles.
2. AutoFilter and filter column records.
3. Hyperlinks and comments.
4. Data validation and conditional formatting.
5. Defined names and cross-sheet formulas.
6. Drawings and images.
7. Optional native legacy encryption only after a dedicated security review.

## Release gates

Before a stable major release:

- add Excel-generated fixtures alongside LibreOffice fixtures
- test files from Excel 97, 2003, 2007+, and current Microsoft 365
- run malformed-file and allocation-chain fuzzing
- add memory and throughput benchmarks for large BIFF8 files
- validate every writer fixture by opening and resaving in Microsoft Excel and LibreOffice
- publish an exact supported-record matrix
