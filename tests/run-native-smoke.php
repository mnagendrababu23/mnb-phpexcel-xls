<?php

declare(strict_types=1);

$root = dirname(__DIR__);
$workspace = dirname($root);

spl_autoload_register(static function (string $class) use ($root, $workspace): void {
    $prefix = 'Mnb\\PHPExcel\\';
    if (!str_starts_with($class, $prefix)) {
        return;
    }
    $relative = str_replace('\\', '/', substr($class, strlen($prefix))) . '.php';
    foreach ([$root . '/src/', $workspace . '/mnb-phpexcel-core/src/'] as $base) {
        if (is_file($base . $relative)) {
            require $base . $relative;
            return;
        }
    }
});

use Mnb\PHPExcel\Core\CellValue;
use Mnb\PHPExcel\Core\WorkbookData;
use Mnb\PHPExcel\Core\WorksheetData;
use Mnb\PHPExcel\Reader\State\FormulaResult;
use Mnb\PHPExcel\Reader\XlsReader;
use Mnb\PHPExcel\Writer\XlsWriter;

$runtime = __DIR__ . '/runtime';
if (!is_dir($runtime)) {
    mkdir($runtime, 0775, true);
}

$fixtureRows = (new XlsReader())->readSheet(__DIR__ . '/fixtures/libreoffice-sample.xls');
assert($fixtureRows[1] === ['Alice', 30, 'TRUE']);

$path = $runtime . '/native-smoke.xls';
$workbook = new WorkbookData([
    new WorksheetData('Data', [
        ['Name', 'Age', 'Active', 'Date', 'Formula'],
        ['Alice', 30, true, CellValue::date('2026-07-29'), CellValue::formula('B2*2', 60)],
        ['Bob', 25, false, CellValue::date('2026-07-30 12:30:00'), CellValue::formula('SUM(B2:B3)', 55)],
    ], freezeHeader: true, mergeCells: ['A5:B5'], columnWidths: ['A' => 20]),
]);
(new XlsWriter())->write($workbook, $path);
$rows = (new XlsReader())->readSheet($path, 1, ['formula_cells' => 'both']);
assert($rows[1][2] === true);
assert($rows[1][3] === '2026-07-29');
assert($rows[1][4] instanceof FormulaResult);
assert($rows[1][4]->cachedValue === 60);

$largePath = $runtime . '/native-large-sst.xls';
$largeRows = [];
for ($i = 0; $i < 700; $i++) {
    $largeRows[] = ['Unicode ' . $i . ' — ' . str_repeat('x', 30), $i];
}
(new XlsWriter())->write(new WorkbookData([new WorksheetData('Unicode', $largeRows)]), $largePath);
$readLarge = (new XlsReader())->readSheet($largePath);
assert(count($readLarge) === 700);
assert($readLarge[699] === $largeRows[699]);

echo json_encode([
    'status' => 'passed',
    'libreoffice_fixture_rows' => count($fixtureRows),
    'native_roundtrip_rows' => count($rows),
    'sst_continue_rows' => count($readLarge),
    'native_file' => $path,
], JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES) . PHP_EOL;
