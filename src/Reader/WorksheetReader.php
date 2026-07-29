<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader;

use DateTimeImmutable;
use Mnb\PHPExcel\Biff\BiffRecord;
use Mnb\PHPExcel\Biff\BiffRecordReader;
use Mnb\PHPExcel\Biff\Formula\FormulaDecoder;
use Mnb\PHPExcel\Biff\RecordType;
use Mnb\PHPExcel\Biff\String\BiffString;
use Mnb\PHPExcel\Biff\WorkbookInfo;
use Mnb\PHPExcel\Exception\InvalidBiffRecordException;
use Mnb\PHPExcel\Reader\State\FormulaResult;
use Mnb\PHPExcel\Support\Binary;

final class WorksheetReader
{
    public function __construct(
        private readonly string $stream,
        private readonly WorkbookInfo $workbook,
        private readonly array $options = [],
    ) {
    }

    /** @return \Generator<int,list<mixed>> */
    public function rows(int $sheetOffset): \Generator
    {
        $projection = ColumnProjection::fromOptions($this->options);
        $startRow = max(1, (int) ($this->options['start_row'] ?? 1));
        $endRow = isset($this->options['end_row']) ? max(1, (int) $this->options['end_row']) : null;
        $sourceLimit = isset($this->options['source_limit_rows']) ? max(0, (int) $this->options['source_limit_rows']) : null;
        $includeEmptyCells = (bool) ($this->options['include_empty_cells'] ?? true);
        $maxWorksheetRows = min(65536, (int) ($this->options['max_worksheet_rows'] ?? 65536));
        $maxWorksheetColumns = min(256, (int) ($this->options['max_worksheet_columns'] ?? 256));

        $recordIterator = (new BiffRecordReader($this->stream, $this->options))->records($sheetOffset);
        $first = true;
        $currentRowNumber = null;
        $currentRow = [];
        $highestColumn = 0;
        $dimensionLastRow = 0;
        $dimensionLastColumn = 0;
        $delivered = 0;
        $lastYieldedSourceRow = 0;
        $pendingStringFormula = null;
        $formulaDecoder = new FormulaDecoder();

        $emitRow = function (int $sourceRow, array $row) use (
            $projection,
            $includeEmptyCells,
            &$highestColumn,
            &$dimensionLastColumn,
            $startRow,
            $endRow,
            $sourceLimit,
            &$delivered,
            &$lastYieldedSourceRow,
            $maxWorksheetColumns
        ): ?array {
            if ($sourceRow < $startRow || ($endRow !== null && $sourceRow > $endRow)) {
                return null;
            }
            if ($sourceLimit !== null && $delivered >= $sourceLimit) {
                return null;
            }
            $width = min($maxWorksheetColumns, max($highestColumn, $dimensionLastColumn));
            if ($includeEmptyCells && $width > 0) {
                $row = array_replace(array_fill(0, $width, null), $row);
            } elseif ($row !== []) {
                ksort($row);
                $row = array_replace(array_fill(0, max(array_keys($row)) + 1, null), $row);
                while ($row !== [] && end($row) === null) {
                    array_pop($row);
                }
            }
            $row = array_values($projection->project($row));
            $delivered++;
            $lastYieldedSourceRow = $sourceRow;
            return $row;
        };

        foreach ($recordIterator as $record) {
            if ($first) {
                $first = false;
                if ($record->type !== RecordType::BOF || $record->length() < 4 || Binary::u16($record->payload, 2) !== 0x0010) {
                    throw InvalidBiffRecordException::because('Worksheet BOF record is missing at the BOUNDSHEET offset.', ['offset' => $sheetOffset]);
                }
                continue;
            }
            if ($record->type === RecordType::EOF) {
                break;
            }
            if ($record->type === RecordType::DIMENSIONS && $record->length() >= 12) {
                $dimensionLastRow = min($maxWorksheetRows, Binary::u32($record->payload, 4));
                $dimensionLastColumn = min($maxWorksheetColumns, Binary::u16($record->payload, 10));
                continue;
            }
            if ($record->type === RecordType::STRING && $pendingStringFormula !== null) {
                $decoded = BiffString::readUnicodeString($record->payload, 0, false)['value'];
                $pendingStringFormula['cached'] = $decoded;
                $value = $this->formulaValue(
                    $pendingStringFormula['formula'],
                    $decoded,
                    'string',
                    $pendingStringFormula['tokens'],
                    $pendingStringFormula['xf']
                );
                if ($currentRowNumber === $pendingStringFormula['row']) {
                    $currentRow[$pendingStringFormula['column']] = $value;
                }
                $pendingStringFormula = null;
                continue;
            }

            $cells = $this->decodeCells($record, $formulaDecoder, $pendingStringFormula);
            foreach ($cells as $cell) {
                $rowNumber = $cell['row'];
                $columnIndex = $cell['column'];
                if ($rowNumber < 1 || $rowNumber > $maxWorksheetRows || $columnIndex < 0 || $columnIndex >= $maxWorksheetColumns) {
                    throw InvalidBiffRecordException::because('Cell lies outside BIFF8 worksheet limits.', ['row' => $rowNumber, 'column' => $columnIndex + 1]);
                }
                if ($currentRowNumber === null) {
                    $currentRowNumber = $rowNumber;
                }
                if ($rowNumber < $currentRowNumber) {
                    throw InvalidBiffRecordException::because('Worksheet cell records are not ordered by row.', ['previous_row' => $currentRowNumber, 'row' => $rowNumber]);
                }
                if ($rowNumber > $currentRowNumber) {
                    $emitted = $emitRow($currentRowNumber, $currentRow);
                    if ($emitted !== null) {
                        yield $currentRowNumber - 1 => $emitted;
                    }
                    if (($sourceLimit !== null && $delivered >= $sourceLimit) || ($endRow !== null && $currentRowNumber >= $endRow)) {
                        return;
                    }
                    // Preserve blank source rows described by DIMENSIONS.
                    for ($gap = $currentRowNumber + 1; $gap < $rowNumber; $gap++) {
                        if ($gap > $dimensionLastRow || ($endRow !== null && $gap > $endRow)) {
                            break;
                        }
                        $blank = $emitRow($gap, []);
                        if ($blank !== null) {
                            yield $gap - 1 => $blank;
                        }
                        if ($sourceLimit !== null && $delivered >= $sourceLimit) {
                            return;
                        }
                    }
                    $currentRowNumber = $rowNumber;
                    $currentRow = [];
                }
                $currentRow[$columnIndex] = $cell['value'];
                $highestColumn = max($highestColumn, $columnIndex + 1);
            }
        }

        if ($currentRowNumber !== null) {
            $emitted = $emitRow($currentRowNumber, $currentRow);
            if ($emitted !== null) {
                yield $currentRowNumber - 1 => $emitted;
            }
        }
        if ($sourceLimit !== null && $delivered >= $sourceLimit) {
            return;
        }
        $last = $currentRowNumber ?? 0;
        $finalRow = min($maxWorksheetRows, $endRow ?? $dimensionLastRow);
        for ($row = $last + 1; $row <= $finalRow; $row++) {
            $blank = $emitRow($row, []);
            if ($blank !== null) {
                yield $row - 1 => $blank;
            }
            if ($sourceLimit !== null && $delivered >= $sourceLimit) {
                return;
            }
        }
    }

    /**
     * @param array<string,mixed>|null $pendingStringFormula
     * @return list<array{row:int,column:int,value:mixed}>
     */
    private function decodeCells(BiffRecord $record, FormulaDecoder $formulaDecoder, ?array &$pendingStringFormula): array
    {
        $payload = $record->payload;
        switch ($record->type) {
            case RecordType::NUMBER:
                Binary::requireBytes($payload, 0, 14);
                return [[
                    'row' => Binary::u16($payload, 0) + 1,
                    'column' => Binary::u16($payload, 2),
                    'value' => $this->numericValue(Binary::double($payload, 6), Binary::u16($payload, 4)),
                ]];

            case RecordType::RK:
                Binary::requireBytes($payload, 0, 10);
                return [[
                    'row' => Binary::u16($payload, 0) + 1,
                    'column' => Binary::u16($payload, 2),
                    'value' => $this->numericValue($this->decodeRk(Binary::u32($payload, 6)), Binary::u16($payload, 4)),
                ]];

            case RecordType::MULRK:
                Binary::requireBytes($payload, 0, 12);
                $row = Binary::u16($payload, 0) + 1;
                $firstColumn = Binary::u16($payload, 2);
                $lastColumn = Binary::u16($payload, strlen($payload) - 2);
                $count = $lastColumn - $firstColumn + 1;
                if ($count < 1 || 4 + ($count * 6) + 2 !== strlen($payload)) {
                    throw InvalidBiffRecordException::because('Malformed MULRK record.', ['offset' => $record->offset]);
                }
                $cells = [];
                for ($i = 0; $i < $count; $i++) {
                    $offset = 4 + ($i * 6);
                    $xf = Binary::u16($payload, $offset);
                    $cells[] = [
                        'row' => $row,
                        'column' => $firstColumn + $i,
                        'value' => $this->numericValue($this->decodeRk(Binary::u32($payload, $offset + 2)), $xf),
                    ];
                }
                return $cells;

            case RecordType::LABELSST:
                Binary::requireBytes($payload, 0, 10);
                $sst = $this->workbook->sharedStrings;
                if ($sst === null) {
                    throw InvalidBiffRecordException::because('LABELSST record appears without an SST record.', ['offset' => $record->offset]);
                }
                return [[
                    'row' => Binary::u16($payload, 0) + 1,
                    'column' => Binary::u16($payload, 2),
                    'value' => $sst->get(Binary::u32($payload, 6)),
                ]];

            case RecordType::LABEL:
                Binary::requireBytes($payload, 0, 9);
                return [[
                    'row' => Binary::u16($payload, 0) + 1,
                    'column' => Binary::u16($payload, 2),
                    'value' => BiffString::readUnicodeString($payload, 6, false)['value'],
                ]];

            case RecordType::BOOLERR:
                Binary::requireBytes($payload, 0, 8);
                $isError = Binary::u8($payload, 7) !== 0;
                return [[
                    'row' => Binary::u16($payload, 0) + 1,
                    'column' => Binary::u16($payload, 2),
                    'value' => $isError ? self::errorText(Binary::u8($payload, 6)) : Binary::u8($payload, 6) !== 0,
                ]];

            case RecordType::BLANK:
                Binary::requireBytes($payload, 0, 6);
                return [[
                    'row' => Binary::u16($payload, 0) + 1,
                    'column' => Binary::u16($payload, 2),
                    'value' => null,
                ]];

            case RecordType::MULBLANK:
                Binary::requireBytes($payload, 0, 8);
                $row = Binary::u16($payload, 0) + 1;
                $firstColumn = Binary::u16($payload, 2);
                $lastColumn = Binary::u16($payload, strlen($payload) - 2);
                $count = $lastColumn - $firstColumn + 1;
                if ($count < 1 || 4 + ($count * 2) + 2 !== strlen($payload)) {
                    throw InvalidBiffRecordException::because('Malformed MULBLANK record.', ['offset' => $record->offset]);
                }
                $cells = [];
                for ($i = 0; $i < $count; $i++) {
                    $cells[] = ['row' => $row, 'column' => $firstColumn + $i, 'value' => null];
                }
                return $cells;

            case RecordType::FORMULA:
                Binary::requireBytes($payload, 0, 22);
                $row = Binary::u16($payload, 0) + 1;
                $column = Binary::u16($payload, 2);
                $xf = Binary::u16($payload, 4);
                $tokenLength = Binary::u16($payload, 20);
                Binary::requireBytes($payload, 22, $tokenLength);
                $tokens = substr($payload, 22, $tokenLength);
                try {
                    $formula = '=' . $formulaDecoder->decode($tokens, $this->workbook->sheetNames(false));
                } catch (\Throwable) {
                    $formula = '=#BIFF_FORMULA[' . strtoupper(bin2hex($tokens)) . ']';
                }
                [$cached, $resultType] = $this->formulaCachedValue(substr($payload, 6, 8), $xf);
                if ($resultType === 'pending_string') {
                    $pendingStringFormula = [
                        'row' => $row,
                        'column' => $column,
                        'xf' => $xf,
                        'formula' => $formula,
                        'tokens' => $tokens,
                        'cached' => null,
                    ];
                }
                return [[
                    'row' => $row,
                    'column' => $column,
                    'value' => $this->formulaValue($formula, $cached, $resultType, $tokens, $xf),
                ]];
        }
        return [];
    }

    /** @return array{0:mixed,1:string} */
    private function formulaCachedValue(string $result, int $xf): array
    {
        if (strlen($result) !== 8) {
            return [null, 'blank'];
        }
        if (substr($result, 6, 2) !== "\xFF\xFF") {
            return [$this->numericValue(Binary::double($result, 0), $xf), 'number'];
        }
        $type = ord($result[0]);
        $value = ord($result[2]);
        return match ($type) {
            0 => [null, 'pending_string'],
            1 => [$value !== 0, 'boolean'],
            2 => [self::errorText($value), 'error'],
            3 => [null, 'blank'],
            default => [null, 'unknown'],
        };
    }

    private function formulaValue(string $formula, mixed $cached, string $resultType, string $tokens, int $xf): mixed
    {
        $mode = strtolower((string) ($this->options['formula_cells'] ?? 'formula'));
        if ($mode === 'cached_value') {
            return $cached;
        }
        if ($mode === 'both') {
            return new FormulaResult($formula, $cached, $resultType, [
                'source' => 'native-biff8',
                'token_hex' => strtoupper(bin2hex($tokens)),
                'xf_index' => $xf,
            ]);
        }
        return $formula;
    }

    private function numericValue(float $value, int $xf): mixed
    {
        if (($this->options['format_dates'] ?? true) && $this->workbook->isDateStyle($xf)) {
            return $this->formatExcelSerialDate($value);
        }
        if (($this->options['preserve_numeric_strings'] ?? false) === true) {
            return self::numberText($value);
        }
        return floor($value) === $value && $value >= PHP_INT_MIN && $value <= PHP_INT_MAX ? (int) $value : $value;
    }

    private function formatExcelSerialDate(float $serial): mixed
    {
        if ($serial < 0) {
            return $serial;
        }
        $wholeDays = (int) floor($serial);
        $seconds = (int) round(($serial - $wholeDays) * 86400);
        if ($this->workbook->date1904) {
            $date = (new DateTimeImmutable('1904-01-01 00:00:00'))->modify('+' . $wholeDays . ' days')->modify('+' . $seconds . ' seconds');
        } else {
            $offset = $wholeDays > 59 ? $wholeDays - 1 : $wholeDays;
            $date = (new DateTimeImmutable('1899-12-31 00:00:00'))->modify('+' . $offset . ' days')->modify('+' . $seconds . ' seconds');
        }
        if (($this->options['return_datetime'] ?? false) === true) {
            return $date;
        }
        return $date->format($seconds !== 0
            ? (string) ($this->options['datetime_format'] ?? 'Y-m-d H:i:s')
            : (string) ($this->options['date_format'] ?? 'Y-m-d'));
    }

    private function decodeRk(int $rk): float|int
    {
        $divideBy100 = ($rk & 0x01) !== 0;
        $isInteger = ($rk & 0x02) !== 0;
        if ($isInteger) {
            $signed = $rk >= 0x80000000 ? $rk - 0x100000000 : $rk;
            $value = $signed >> 2;
        } else {
            $value = unpack('e', pack('V2', 0, $rk & 0xFFFFFFFC))[1];
        }
        return $divideBy100 ? $value / 100 : $value;
    }

    private static function errorText(int $code): string
    {
        return match ($code) {
            0x00 => '#NULL!', 0x07 => '#DIV/0!', 0x0F => '#VALUE!', 0x17 => '#REF!',
            0x1D => '#NAME?', 0x24 => '#NUM!', 0x2A => '#N/A', default => '#ERROR!',
        };
    }

    private static function numberText(float $number): string
    {
        if (is_finite($number) && floor($number) === $number) {
            return sprintf('%.0F', $number);
        }
        return sprintf('%.15G', $number);
    }
}
