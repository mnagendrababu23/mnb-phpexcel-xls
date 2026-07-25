<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Compatibility;

use Mnb\PHPExcel\Reader\ColumnProjection;
use Mnb\PHPExcel\Reader\IterableReaderInterface;
use Mnb\PHPExcel\Reader\State\FormulaResult;
use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;

/**
 * Optional legacy .xls adapter powered by PhpSpreadsheet.
 *
 * Kept outside the lightweight core because BIFF parsing is a substantial
 * dependency and is not suitable for the native streaming path.
 */
final class XlsReader implements IterableReaderInterface
{
    /** @return list<list<mixed>> */
    public function readSheet(string $path, int|string $sheet = 1, array $options = []): array
    {
        return array_values(iterator_to_array($this->iterateSheet($path, $sheet, $options), true));
    }

    /** @return \Generator<int,list<mixed>> */
    public function iterateSheet(string $path, int|string $sheet = 1, array $options = []): iterable
    {
        $this->ensureDependency();
        if (!is_file($path)) {
            throw MnbExcelException::withCode('XLS file not found: ' . $path, ErrorCode::FILE_NOT_FOUND, ['path' => $path]);
        }

        $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReaderForFile($path);
        if (method_exists($reader, 'setReadDataOnly')) {
            $reader->setReadDataOnly((bool) ($options['read_data_only'] ?? false));
        }
        if (method_exists($reader, 'setReadEmptyCells')) {
            $reader->setReadEmptyCells((bool) ($options['include_empty_cells'] ?? true));
        }

        $sheetNames = $reader->listWorksheetNames($path);
        $sheetName = $this->resolveSheetName($sheetNames, $sheet);
        if (method_exists($reader, 'setLoadSheetsOnly')) {
            $reader->setLoadSheetsOnly([$sheetName]);
        }

        try {
            $spreadsheet = $reader->load($path);
        } catch (\Throwable $e) {
            throw MnbExcelException::withCode('Unable to read legacy XLS file: ' . $e->getMessage(), ErrorCode::FILE_READ_FAILED, ['path' => $path], $e);
        }

        $worksheet = $spreadsheet->getSheetByName($sheetName);
        if ($worksheet === null) {
            $spreadsheet->disconnectWorksheets();
            throw new MnbExcelException('XLS sheet does not exist: ' . $sheetName);
        }

        $projection = ColumnProjection::fromOptions($options);
        $startRow = max(1, (int) ($options['start_row'] ?? 1));
        $endRow = isset($options['end_row']) ? max(1, (int) $options['end_row']) : $worksheet->getHighestDataRow();
        $sourceLimit = isset($options['source_limit_rows']) ? max(0, (int) $options['source_limit_rows']) : null;
        $highestColumn = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($worksheet->getHighestDataColumn());
        $formulaMode = strtolower((string) ($options['formula_cells'] ?? 'formula'));
        $delivered = 0;

        try {
            for ($rowNumber = $startRow; $rowNumber <= $endRow; $rowNumber++) {
                if ($sourceLimit !== null && $delivered >= $sourceLimit) {
                    break;
                }
                $row = [];
                for ($column = 1; $column <= $highestColumn; $column++) {
                    if ($projection->active() && !$projection->includesIndex($column)) {
                        continue;
                    }
                    $cell = $worksheet->getCell([$column, $rowNumber]);
                    $value = $cell->getValue();
                    if (is_string($value) && str_starts_with($value, '=')) {
                        if ($formulaMode === 'cached_value') {
                            try {
                                $value = $cell->getCalculatedValue();
                            } catch (\Throwable) {
                                $value = null;
                            }
                        } elseif ($formulaMode === 'both') {
                            try {
                                $cached = $cell->getCalculatedValue();
                            } catch (\Throwable) {
                                $cached = null;
                            }
                            $value = new FormulaResult($value, $cached, get_debug_type($cached), ['source' => 'phpspreadsheet']);
                        }
                    }
                    if ($projection->active() && $projection->compact()) {
                        $row[] = $value;
                    } else {
                        $row[$column - 1] = $value;
                    }
                }
                if ($row !== [] && !array_is_list($row)) {
                    ksort($row);
                    $max = max(array_keys($row));
                    $row = array_replace(array_fill(0, $max + 1, null), $row);
                }
                yield $rowNumber - 1 => array_values($row);
                $delivered++;
            }
        } finally {
            $spreadsheet->disconnectWorksheets();
            unset($spreadsheet);
        }
    }

    /** @return list<string> */
    public function sheetNames(string $path): array
    {
        $this->ensureDependency();
        $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReaderForFile($path);
        return array_values($reader->listWorksheetNames($path));
    }

    /** @param list<string> $names */
    private function resolveSheetName(array $names, int|string $sheet): string
    {
        if (is_string($sheet) && !ctype_digit($sheet)) {
            if (!in_array($sheet, $names, true)) {
                throw new MnbExcelException('XLS sheet does not exist: ' . $sheet);
            }
            return $sheet;
        }
        $index = max(1, (int) $sheet) - 1;
        if (!isset($names[$index])) {
            throw new MnbExcelException('XLS sheet index does not exist: ' . (string) $sheet);
        }
        return $names[$index];
    }

    private function ensureDependency(): void
    {
        if (!class_exists('PhpOffice\\PhpSpreadsheet\\IOFactory')) {
            throw MnbExcelException::withCode(
                'Legacy XLS support requires the optional phpoffice/phpspreadsheet package.',
                ErrorCode::EXTENSION_MISSING,
                [],
                null,
                'Install mnb/mnb-phpexcel-xls (or phpoffice/phpspreadsheet) to read .xls files.'
            );
        }
    }
}
