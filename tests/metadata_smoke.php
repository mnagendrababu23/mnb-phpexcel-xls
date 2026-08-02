<?php

declare(strict_types=1);

$root = dirname(__DIR__);
$workspace = dirname($root);
spl_autoload_register(static function (string $class) use ($root, $workspace): void {
    $prefix = 'Mnb\\PHPExcel\\';
    if (!str_starts_with($class, $prefix)) return;
    $relative = str_replace('\\', '/', substr($class, strlen($prefix))) . '.php';
    foreach ([$root . '/src/', $workspace . '/mnb-phpexcel-core/src/'] as $base) {
        if (is_file($base . $relative)) { require $base . $relative; return; }
    }
});

use Mnb\PHPExcel\Core\CellValue;
use Mnb\PHPExcel\Core\WorkbookData;
use Mnb\PHPExcel\Core\WorksheetData;
use Mnb\PHPExcel\Format\Xls;
use Mnb\PHPExcel\Writer\XlsWriter;

$fixture = __DIR__ . '/fixtures/libreoffice-sample.xls';
$fixtureMeta = Xls::metaInfo($fixture, ['profile' => 'full']);
assert($fixtureMeta['format'] === 'xls');
assert($fixtureMeta['workbook']['sheet_count'] >= 1);
assert($fixtureMeta['format_details']['container'] === 'CFB/OLE2');
assert($fixtureMeta['document']['state'] === 'available' || $fixtureMeta['document']['state'] === 'partial');

$path = __DIR__ . '/runtime/metadata-native.xls';
(new XlsWriter())->write(new WorkbookData([
    new WorksheetData('Visible', [
        ['Name', 'Value', 'Formula'],
        ['Alice', 10, CellValue::formula('B2*2', 20)],
    ], mergeCells: ['A4:B4']),
    new WorksheetData('Second', [['x']]),
]), $path);
$full = Xls::metaInfo($path, ['profile' => 'full']);
assert($full['workbook']['sheet_count'] === 2);
assert($full['statistics']['formula_count'] === 1);
assert($full['statistics']['merged_range_count'] === 1);
assert($full['security']['encrypted'] === false);
assert($full['xml_metadata']['state'] === 'not_applicable');


$rich = Xls::metaInfo(__DIR__ . '/fixtures/rich-metadata.xls', ['profile' => 'full']);
assert(($rich['document']['title'] ?? null) === 'Rich XLS Metadata');
assert(($rich['document']['category'] ?? null) === 'Finance');
assert(($rich['workbook']['sheet_count'] ?? null) === 2);
assert(($rich['hidden_content']['hidden_sheet_count'] ?? null) === 1);
assert(($rich['hidden_content']['hidden_row_count'] ?? null) === 1);
assert(($rich['hidden_content']['hidden_column_count'] ?? null) === 1);
assert(($rich['comments_notes']['comment_count'] ?? null) === 1);
assert(($rich['links']['hyperlink_count'] ?? null) === 1);
assert(($rich['links']['external_book_count'] ?? null) === 0);
assert(($rich['security']['protected_sheet_count'] ?? null) === 0);
$customNames = array_column($rich['custom_properties']['items'], 'name');
assert(in_array('Project ID', $customNames, true));
assert(in_array('Approved', $customNames, true));

$session = Xls::read($path);
$viaSession = $session->metaInfo(['profile' => 'quick']);
assert($viaSession['workbook']['sheet_count'] === 2);
assert($viaSession['statistics']['state'] === 'partial');

echo json_encode([
    'status' => 'passed',
    'fixture_title' => $fixtureMeta['document']['title'] ?? null,
    'sheets' => $full['workbook']['sheet_count'],
    'formulas' => $full['statistics']['formula_count'],
], JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES) . PHP_EOL;
