# Native XLS architecture

## Layers

```text
Format\Xls
├── Reader\XlsReader
│   ├── Compound\CompoundFileReader
│   │   ├── CFB header
│   │   ├── DIFAT / FAT
│   │   ├── MiniFAT / root mini stream
│   │   └── directory entries and streams
│   ├── Biff\WorkbookGlobalsReader
│   │   ├── BOUNDSHEET
│   │   ├── SST / CONTINUE
│   │   ├── FORMAT / XF
│   │   └── date system and code page
│   └── Reader\WorksheetReader
│       ├── cell records
│       ├── formula tokens and cached results
│       ├── date conversion
│       └── row iteration and column projection
└── Writer\XlsWriter
    ├── Writer\WorkbookWriter
    │   ├── workbook globals
    │   ├── BOUNDSHEET offsets
    │   ├── fonts / XFs / styles
    │   └── SST records
    ├── Writer\WorksheetWriter
    │   ├── BIFF8 cells
    │   ├── formulas
    │   ├── dimensions / widths / heights
    │   ├── merges
    │   └── frozen panes
    └── Compound\CompoundFileWriter
        ├── CFB header
        ├── workbook stream sectors
        ├── directory sector
        └── FAT sectors
```

## Public API stability

`Mnb\PHPExcel\Format\Xls` retains the existing read/write facade. The implementation moved from third-party compatibility classes to native `Reader` and `Writer` classes.


## CFB reader behavior

The reader supports both 512-byte and 4096-byte regular sectors and the standard 64-byte mini sector. It resolves FAT sector locations from the header DIFAT and optional DIFAT chain, then parses the directory stream. Streams smaller than 4096 bytes are read from the MiniFAT and root mini stream.

Every chain is cycle-checked and boundary-checked before data is returned.

## CFB writer behavior

The writer produces a version-3 CFB file with 512-byte sectors and one `Workbook` stream. The BIFF stream is padded to the normal-stream cutoff, avoiding MiniFAT generation. FAT locations use the 109 header DIFAT slots first and then native DIFAT-chain sectors for larger files.

This writer design is deliberately deterministic while still supporting workbook streams beyond the header-only FAT range.

## BIFF record processing

All records use the standard four-byte BIFF header:

```text
record type:   uint16 little-endian
record length: uint16 little-endian
payload:       record length bytes
```

Payloads above 8224 bytes are rejected. SST data is the exception at the logical level: it is split into one `SST` record and one or more `CONTINUE` records, each within the record-size limit.

## Formula strategy

Reading and calculation are separate concerns.

- The native reader decodes common BIFF RPN tokens into an Excel-style formula string.
- Cached formula results are read independently and remain available even for an unknown token sequence.
- The writer compiles a documented subset of Excel formula syntax into BIFF8 RPN tokens.
- The package does not call an external formula engine.

## Date handling

The workbook global `DATEMODE` record selects the 1900 or 1904 date system. Cell XF records point to built-in or custom number formats. Numeric cells using date-like formats are converted according to the same read options used by the XLSX reader:

- `format_dates`
- `return_datetime`
- `date_format`
- `datetime_format`

## Unsupported features

Unsupported read records are ignored when they are not needed to recover worksheet values. Unsupported writer features are either skipped or rejected when `strict_features` is enabled. This prevents data exports from failing merely because a `WorksheetData` object contains an advanced XLSX-only presentation feature, while still allowing strict compatibility audits.
