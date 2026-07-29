<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Xls\Tests;

use Mnb\PHPExcel\Core\CellValue;
use Mnb\PHPExcel\Core\WorkbookData;
use Mnb\PHPExcel\Core\WorksheetData;
use Mnb\PHPExcel\Reader\State\FormulaResult;
use Mnb\PHPExcel\Reader\XlsReader;
use Mnb\PHPExcel\Writer\XlsWriter;
use PHPUnit\Framework\TestCase;

final class NativeXlsTest extends TestCase
{
    public function testReadsLibreOfficeGeneratedBiff8File(): void
    {
        $reader = new XlsReader();
        self::assertSame(['sample'], $reader->sheetNames(__DIR__ . '/fixtures/libreoffice-sample.xls'));
        self::assertSame([
            ['Name', 'Age', 'Active'],
            ['Alice', 30, 'TRUE'],
            ['Bob', 25, 'FALSE'],
        ], $reader->readSheet(__DIR__ . '/fixtures/libreoffice-sample.xls'));
    }

    public function testNativeRoundTripIncludesDatesBooleansAndFormulas(): void
    {
        $path = $this->temporaryPath('native-roundtrip.xls');
        $sheet = new WorksheetData('Data', [
            ['Name', 'Age', 'Active', 'Date', 'Calc'],
            ['Alice', 30, true, CellValue::date('2026-07-29'), CellValue::formula('B2*2', 60)],
            ['Bob', 25, false, CellValue::date('2026-07-30 12:30:00'), CellValue::formula('SUM(B2:B3)', 55)],
        ], freezeHeader: true, mergeCells: ['A5:B5'], columnWidths: ['A' => 20]);

        (new XlsWriter())->write(new WorkbookData([$sheet]), $path);
        $rows = (new XlsReader())->readSheet($path, 1, ['formula_cells' => 'both']);

        self::assertSame('2026-07-29', $rows[1][3]);
        self::assertSame('2026-07-30 12:30:00', $rows[2][3]);
        self::assertTrue($rows[1][2]);
        self::assertInstanceOf(FormulaResult::class, $rows[1][4]);
        self::assertSame('=B2*2', $rows[1][4]->formula);
        self::assertSame(60, $rows[1][4]->cachedValue);
        self::assertSame('=SUM(B2:B3)', $rows[2][4]->formula);
    }

    public function testSstContinueRecordsRoundTripUnicodeData(): void
    {
        $path = $this->temporaryPath('large-sst.xls');
        $source = [];
        for ($i = 0; $i < 700; $i++) {
            $source[] = ['Row ' . $i . ' — ' . str_repeat(chr(65 + ($i % 26)), 30), $i];
        }
        (new XlsWriter())->write(new WorkbookData([new WorksheetData('Unicode', $source)]), $path);
        $rows = (new XlsReader())->readSheet($path);
        self::assertCount(700, $rows);
        self::assertSame($source[699], $rows[699]);
    }


    public function testFormulaCachedResultTypesRoundTrip(): void
    {
        $path = $this->temporaryPath('formula-types.xls');
        $sheet = new WorksheetData('Formulas', [[
            CellValue::formula('IF(TRUE,"yes","no")', 'yes'),
            CellValue::formula('TRUE', true),
            CellValue::formula('NA()', '#N/A'),
            CellValue::formula('IF(FALSE,1,"")', null),
        ]]);
        (new XlsWriter())->write(new WorkbookData([$sheet]), $path);

        $cached = (new XlsReader())->readSheet($path, 1, ['formula_cells' => 'cached_value']);
        self::assertSame(['yes', true, '#N/A', null], $cached[0]);

        $formulas = (new XlsReader())->readSheet($path, 1, ['formula_cells' => 'formula']);
        self::assertSame('=IF(TRUE,"yes","no")', $formulas[0][0]);
        self::assertSame('=NA()', $formulas[0][2]);
    }

    private function temporaryPath(string $name): string
    {
        $directory = sys_get_temp_dir() . '/mnb-native-xls-tests-' . getmypid();
        if (!is_dir($directory)) {
            mkdir($directory, 0775, true);
        }
        return $directory . '/' . $name;
    }
}
