# MNB PHPExcel — Native XLS

A fully independent, pure-PHP BIFF8 reader and writer for legacy Excel 97–2003 `.xls` files.

This package **does not require, suggest, wrap, or call PhpSpreadsheet**. Both the OLE Compound File container and the BIFF8 workbook stream are processed natively.

## Install

```bash
composer require mnb/mnb-phpexcel-xls:^2.0
```

Requirements:

- PHP 8.1+
- `ext-iconv`
- `mnb/mnb-phpexcel-core`

## Read an XLS file

```php
<?php

declare(strict_types=1);

require __DIR__ . '/vendor/autoload.php';

use Mnb\PHPExcel\Format\Xls;

$rows = Xls::read(__DIR__ . '/files/report.xls')
    ->sheet('Data')
    ->withHeaderRow()
    ->toArray([
        'skip_empty_rows' => true,
        'max_rows' => 5000,
    ]);

foreach ($rows as $row) {
    echo $row['Name'] . PHP_EOL;
}
```

You can also use the reader directly:

```php
use Mnb\PHPExcel\Reader\XlsReader;

$reader = new XlsReader();
$sheetNames = $reader->sheetNames('report.xls');
$rows = $reader->readSheet('report.xls', 1);
```

## Write an XLS file

```php
use Mnb\PHPExcel\Core\CellValue;
use Mnb\PHPExcel\Format\Xls;

Xls::write([
    ['Name' => 'Alice', 'Amount' => 1250.50, 'Approved' => true],
    ['Name' => 'Bob', 'Amount' => 875.00, 'Approved' => false],
], __DIR__ . '/files/report.xls', [
    'with_header' => true,
    'sheet_name' => 'Data',
]);
```

Typed cells and cached formulas are supported:

```php
Xls::write([
    ['Date', 'Amount', 'Double'],
    [
        CellValue::date('2026-07-29'),
        CellValue::number('1250.50'),
        CellValue::formula('B2*2', 2501.00),
    ],
], 'typed.xls');
```

## Formula read modes

```php
$session = Xls::read('report.xls');

$formulas = $session->toArray(['formula_cells' => 'formula']);
$cached   = $session->toArray(['formula_cells' => 'cached_value']);
$both     = $session->toArray(['formula_cells' => 'both']);
```

`both` returns `Mnb\PHPExcel\Reader\State\FormulaResult` values containing the formula, cached result, result type, and native BIFF token metadata.

## Native feature coverage

### Reader

- OLE/CFB version 3 and version 4 containers
- FAT, DIFAT, MiniFAT, directory entries, normal streams, and mini streams
- `Workbook` and older `Book` stream names
- BIFF8 workbook globals and worksheet substreams
- sheet names, order, hidden state, and worksheet selection
- Unicode shared strings, including SST data split across `CONTINUE` records
- numeric `NUMBER`, `RK`, and `MULRK` cells
- strings through `LABELSST` and `LABEL`
- Boolean, error, blank, and multi-blank cells
- formulas with numeric, string, Boolean, error, and blank cached values
- common BIFF8 formula token decoding
- built-in and custom date-format detection
- 1900 and 1904 date systems
- source-level row and column limits/projection
- forward row iteration through `IterableReaderInterface`

### Writer

- valid BIFF8 workbooks inside native CFB/OLE containers
- multiple worksheets
- Unicode shared strings with `CONTINUE` splitting
- numbers, text, Booleans, blanks, errors, dates, and date-times
- common formulas compiled into BIFF8 RPN tokens
- formula cached values
- merged cells
- column widths and row heights
- frozen rows and columns
- atomic file replacement
- BIFF8 worksheet limits: 65,536 rows and 256 columns

## Formula compiler coverage

The native writer currently supports:

- numeric, text, Boolean, and error literals
- A1 cell references and ranges
- `+`, `-`, `*`, `/`, `^`, `&`
- `=`, `<>`, `<`, `<=`, `>`, `>=`
- unary plus/minus and percentage
- parentheses
- common built-in functions such as `SUM`, `AVERAGE`, `MIN`, `MAX`, `IF`, `COUNT`, `COUNTA`, `ROUND`, `DATE`, `TIME`, `VLOOKUP`, and `SUMIF`

Unsupported formula tokens or functions raise `UnsupportedXlsFeatureException` when writing. On reading, unknown formula token streams remain available as a `=#BIFF_FORMULA[...]` representation while their cached value can still be returned.

## Unsupported writer features in the first native milestone

These features are not silently advertised as implemented:

- images and drawings
- charts and pivot tables
- comments and hyperlinks
- conditional formatting
- data validation
- AutoFilter records
- VBA projects and macros
- external workbook links and defined names
- legacy XLS encryption

By default, unsupported worksheet presentation features are skipped so data export remains usable. Set `strict_features` to `true` to receive an `UnsupportedXlsFeatureException` instead:

```php
Xls::write($workbook, 'report.xls', [
    'strict_features' => true,
]);
```

## Security limits

The native binary reader validates and limits:

- file and stream sizes
- FAT, DIFAT, MiniFAT, and directory counts
- cyclic or invalid allocation chains
- stream and sector boundaries
- BIFF record lengths and total record counts
- shared-string counts
- BIFF8 row and column boundaries

Limits can be reduced for untrusted uploads:

```php
$rows = Xls::read('upload.xls')->toArray([
    'max_file_size' => 64 * 1024 * 1024,
    'max_stream_size' => 32 * 1024 * 1024,
    'max_chain_sectors' => 200000,
    'max_biff_records' => 1000000,
    'max_shared_strings' => 250000,
]);
```

## Tests

With Composer development dependencies installed:

```bash
vendor/bin/phpunit tests
```

Without PHPUnit, from a workspace containing sibling `mnb-phpexcel-core` and `mnb-phpexcel-xls` repositories:

```bash
php -d assert.exception=1 tests/run-native-smoke.php
```

The fixture suite includes a LibreOffice-generated BIFF8 file, native write/read round-trips, formulas, dates, and SST continuation coverage.

See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md), [docs/SUPPORTED-RECORDS.md](docs/SUPPORTED-RECORDS.md), and [docs/ROADMAP.md](docs/ROADMAP.md) for internals, the exact record matrix, and next compatibility milestones.

## Unified metadata

```php
$report = Xls::metaInfo('report.xls', [
    'profile' => 'full',
    'max_items' => 1000,
]);

$report = Xls::read('report.xls')->metaInfo([
    'profile' => 'forensic',
    'include_hash' => true,
]);
```

The native collector reads OLE SummaryInformation, DocumentSummaryInformation, custom properties, workbook/sheet state, calculation settings, protection indicators, macros, named objects, links, hidden content, comments, embedded streams, print records, validations, pivots, and forensic stream hashes. Complex BIFF objects that are counted but not fully decoded are explicitly marked `partial`.

### Atomic metadata updates

```php
Xls::updateMetaInfo('source.xls', 'updated.xls', [
    'document' => [
        'title' => 'Annual Report',
        'creator' => 'MNB',
        'category' => 'Finance',
    ],
    'revision' => [
        'last_saved_by' => 'Release Bot',
        'revision_number' => '7',
        'total_editing_time_seconds' => 3600,
    ],
    'application' => [
        'application_name' => 'MNB PHPExcel',
        'application_version' => '2.0',
        'manager' => 'Finance Manager',
        'company' => 'MNB',
    ],
    'custom_properties' => [
        'Project ID' => 'PRJ-1001',
        'Approved' => ['type' => 'boolean', 'value' => true],
        'Budget' => ['type' => 'float', 'value' => 1250.50],
        'Revision Date' => ['type' => 'datetime', 'value' => '2026-08-02T10:00:00Z'],
    ],
    'workbook' => [
        'active_sheet' => 'Summary',
        'sheet_visibility' => ['Raw Data' => 'hidden'],
        'date1904' => false,
    ],
    'calculation' => [
        'mode' => 'automatic',
        'iterate' => false,
        'iterate_count' => 100,
        'iterate_delta' => 0.001,
        'calc_on_save' => true,
        'reference_mode' => 'a1',
    ],
]);
```

Updates are atomic, preserve unknown OLE streams and unmodified BIFF records, and support the same source and destination path. Password-encrypted BIFF workbooks are read safely as `password_required`, but native decryption/update is not implemented. Files containing digital-signature streams are rejected by default because metadata changes may invalidate the signature; proceeding requires the explicit `allow_invalidate_digital_signatures` option.

```php
Xls::removePersonalInfo('source.xls', 'clean.xls', [
    'remove_custom_properties' => true,
    'remove_descriptive_properties' => false,
]);
```
