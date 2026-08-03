<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Writer;

use DateTimeImmutable;
use DateTimeInterface;
use Mnb\PHPExcel\Biff\BiffRecordWriter;
use Mnb\PHPExcel\Biff\Formula\FormulaEncoder;
use Mnb\PHPExcel\Biff\RecordType;
use Mnb\PHPExcel\Core\CellValue;
use Mnb\PHPExcel\Core\RichText;
use Mnb\PHPExcel\Core\WorksheetData;
use Mnb\PHPExcel\Exception\UnsupportedXlsFeatureException;
use Mnb\PHPExcel\Support\Coordinate;

final class WorksheetWriter
{
    public function __construct(
        private readonly SharedStringWriter $sharedStrings,
        private readonly XlsStyleRegistry $styles,
        private readonly bool $date1904 = false,
        private readonly array $options = [],
    ) {
    }

    public function registerStrings(WorksheetData $sheet): void
    {
        foreach ($sheet->rows as $row) {
            foreach (array_values($row) as $value) {
                $text = $this->stringValue($value);
                if ($text !== null) {
                    $this->sharedStrings->add($text);
                }
            }
        }
    }

    public function write(WorksheetData $sheet): string
    {
        $this->validateSheet($sheet);
        $rows = array_values($sheet->rows);
        $rowCount = count($rows);
        if ($rowCount > 65536) {
            throw UnsupportedXlsFeatureException::forFeature('more than 65536 rows', ['sheet' => $sheet->name, 'rows' => $rowCount]);
        }
        $columnCount = 0;
        foreach ($rows as $row) {
            $columnCount = max($columnCount, count($row));
        }
        if ($columnCount > 256) {
            throw UnsupportedXlsFeatureException::forFeature('more than 256 columns', ['sheet' => $sheet->name, 'columns' => $columnCount]);
        }

        $stream = BiffRecordWriter::bof(0x0010);
        $stream .= BiffRecordWriter::record(RecordType::CALCMODE, pack('v', 1));
        $stream .= BiffRecordWriter::record(RecordType::CALCCOUNT, pack('v', 100));
        $stream .= BiffRecordWriter::record(RecordType::REFMODE, pack('v', 1));
        $stream .= BiffRecordWriter::record(RecordType::ITERATION, pack('v', 0));
        $stream .= BiffRecordWriter::record(RecordType::DELTA, pack('e', 0.001));
        $stream .= BiffRecordWriter::record(RecordType::SAVERECALC, pack('v', 1));

        foreach ($sheet->columnWidths as $column => $width) {
            $index = is_int($column) || ctype_digit((string) $column)
                ? (int) $column - 1
                : Coordinate::columnNameToIndex((string) $column) - 1;
            if ($index < 0 || $index > 255) {
                continue;
            }
            $biffWidth = max(0, min(65535, (int) round((float) $width * 256)));
            $stream .= BiffRecordWriter::record(RecordType::COLINFO, pack('vvvvvv', $index, $index, $biffWidth, 0, 0, 0));
        }

        $stream .= BiffRecordWriter::record(
            RecordType::DIMENSIONS,
            pack('VVvvv', 0, $rowCount, 0, $columnCount, 0)
        );

        $formulaEncoder = new FormulaEncoder();
        foreach ($rows as $rowIndex => $row) {
            if ($rowIndex >= 65536) {
                break;
            }
            $values = array_values($row);
            $height = $sheet->rowHeights[$rowIndex + 1] ?? null;
            if ($height !== null) {
                $twips = max(0, min(8191, (int) round((float) $height * 20)));
                $stream .= BiffRecordWriter::record(RecordType::ROW, pack('vvvvVV', $rowIndex, 0, count($values), $twips, 0, 0x00000140));
            }
            foreach ($values as $columnIndex => $value) {
                if ($columnIndex >= 256) {
                    break;
                }
                $cellRef = Coordinate::columnIndexToName($columnIndex + 1) . ($rowIndex + 1);
                $xf = $this->styles->styleId($sheet, $rowIndex + 1, $columnIndex + 1, $cellRef, $value);
                $stream .= $this->cellRecord($rowIndex, $columnIndex, $value, $formulaEncoder, $xf);
            }
        }

        if ($sheet->mergeCells !== []) {
            $ranges = [];
            foreach ($sheet->mergeCells as $range) {
                $ranges[] = $this->parseRange($range);
            }
            foreach (array_chunk($ranges, 1027) as $chunk) {
                $payload = pack('v', count($chunk));
                foreach ($chunk as [$row1, $row2, $col1, $col2]) {
                    $payload .= pack('vvvv', $row1, $row2, $col1, $col2);
                }
                $stream .= BiffRecordWriter::record(RecordType::MERGEDCELLS, $payload);
            }
        }

        $freezeRows = $sheet->freezeTopLeftCell !== null ? Coordinate::splitCellRef($sheet->freezeTopLeftCell)[1] - 1 : ($sheet->freezeHeader ? 1 : $sheet->freezeRows);
        $freezeColumns = $sheet->freezeTopLeftCell !== null ? Coordinate::splitCellRef($sheet->freezeTopLeftCell)[0] - 1 : $sheet->freezeColumns;
        $windowFlags = ($freezeRows > 0 || $freezeColumns > 0) ? 0x06B6 : 0x00B6;
        $stream .= BiffRecordWriter::record(RecordType::WINDOW2, pack('vvvVvvV', $windowFlags, 0, 0, 64, 60, 100, 0));
        if ($freezeRows > 0 || $freezeColumns > 0) {
            $activePane = $freezeRows > 0 && $freezeColumns > 0 ? 0 : ($freezeRows > 0 ? 2 : 1);
            $stream .= BiffRecordWriter::record(RecordType::PANE, pack('vvvvC', $freezeColumns, $freezeRows, $freezeRows, $freezeColumns, $activePane));
        }

        return $stream . BiffRecordWriter::eof();
    }

    private function cellRecord(int $row, int $column, mixed $value, FormulaEncoder $formulaEncoder, int $xf): string
    {
        if ($value instanceof RichText) {
            $value = implode('', array_map(static fn ($run): string => $run->text, $value->runs));
        }
        if ($value instanceof CellValue) {
            return $this->typedCellRecord($row, $column, $value, $formulaEncoder, $xf);
        }
        if ($value instanceof DateTimeInterface) {
            $serial = $this->dateSerial($value);
            return BiffRecordWriter::record(RecordType::NUMBER, pack('vvve', $row, $column, $xf, $serial));
        }
        if ($value === null) {
            return '';
        }
        if (is_bool($value)) {
            return BiffRecordWriter::record(RecordType::BOOLERR, pack('vvvCC', $row, $column, $xf, $value ? 1 : 0, 0));
        }
        if (is_int($value) || is_float($value)) {
            return BiffRecordWriter::record(RecordType::NUMBER, pack('vvve', $row, $column, $xf, (float) $value));
        }
        if (is_string($value) && str_starts_with($value, '=')) {
            return $this->formulaRecord($row, $column, substr($value, 1), null, $formulaEncoder, $xf);
        }
        $text = is_scalar($value) ? (string) $value : json_encode($value, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES | JSON_THROW_ON_ERROR);
        return BiffRecordWriter::record(RecordType::LABELSST, pack('vvvV', $row, $column, $xf, $this->sharedStrings->index($text)));
    }

    private function typedCellRecord(int $row, int $column, CellValue $value, FormulaEncoder $formulaEncoder, int $xf): string
    {
        return match ($value->type()) {
            CellValue::TYPE_BLANK => BiffRecordWriter::record(RecordType::BLANK, pack('vvv', $row, $column, $xf)),
            CellValue::TYPE_TEXT => BiffRecordWriter::record(RecordType::LABELSST, pack('vvvV', $row, $column, $xf, $this->sharedStrings->index((string) $value->value()))),
            CellValue::TYPE_NUMBER => BiffRecordWriter::record(RecordType::NUMBER, pack('vvve', $row, $column, $xf, (float) $value->value())),
            CellValue::TYPE_BOOLEAN => BiffRecordWriter::record(RecordType::BOOLERR, pack('vvvCC', $row, $column, $xf, $value->value() ? 1 : 0, 0)),
            CellValue::TYPE_ERROR => BiffRecordWriter::record(RecordType::BOOLERR, pack('vvvCC', $row, $column, $xf, self::errorCode((string) $value->value()), 1)),
            CellValue::TYPE_DATE => $this->dateCellRecord($row, $column, $value, $xf),
            CellValue::TYPE_FORMULA => $this->formulaRecord($row, $column, (string) $value->value(), $value->cachedValue(), $formulaEncoder, $xf),
            default => throw UnsupportedXlsFeatureException::forFeature('cell value type ' . $value->type()),
        };
    }

    private function dateCellRecord(int $row, int $column, CellValue $value, int $xf): string
    {
        $date = $value->value() instanceof DateTimeInterface
            ? $value->value()
            : new DateTimeImmutable((string) $value->value());
        return BiffRecordWriter::record(RecordType::NUMBER, pack('vvve', $row, $column, $xf, $this->dateSerial($date)));
    }

    private function formulaRecord(int $row, int $column, string $formula, mixed $cached, FormulaEncoder $encoder, int $xf): string
    {
        $tokens = $encoder->encode($formula);
        [$result, $trailing] = $this->formulaCachedResult($cached);
        $payload = pack('vvv', $row, $column, $xf)
            . $result
            . pack('vVv', 0, 0, strlen($tokens))
            . $tokens;
        return BiffRecordWriter::record(RecordType::FORMULA, $payload) . $trailing;
    }

    /** @return array{0:string,1:string} */
    private function formulaCachedResult(mixed $cached): array
    {
        if (is_int($cached) || is_float($cached)) {
            return [pack('e', (float) $cached), ''];
        }
        if (is_bool($cached)) {
            return ["\x01\x00" . ($cached ? "\x01" : "\x00") . "\x00\x00\x00\xFF\xFF", ''];
        }
        if (is_string($cached) && self::errorCodeOrNull($cached) !== null) {
            return ["\x02\x00" . chr(self::errorCode($cached)) . "\x00\x00\x00\xFF\xFF", ''];
        }
        if (is_string($cached)) {
            $encoded = iconv('UTF-8', 'UTF-16LE', $cached);
            if ($encoded === false || intdiv(strlen($encoded), 2) > 65535 || strlen($encoded) + 3 > 8224) {
                throw UnsupportedXlsFeatureException::forFeature('long formula string cached result');
            }
            $stringRecord = BiffRecordWriter::record(RecordType::STRING, pack('vC', intdiv(strlen($encoded), 2), 1) . $encoded);
            return ["\x00\x00\x00\x00\x00\x00\xFF\xFF", $stringRecord];
        }
        return ["\x03\x00\x00\x00\x00\x00\xFF\xFF", ''];
    }

    private function stringValue(mixed $value): ?string
    {
        if ($value instanceof RichText) {
            return implode('', array_map(static fn ($run): string => $run->text, $value->runs));
        }
        if ($value instanceof CellValue) {
            return $value->type() === CellValue::TYPE_TEXT ? (string) $value->value() : null;
        }
        if (is_string($value) && !str_starts_with($value, '=')) {
            return $value;
        }
        if ($value !== null && !is_scalar($value) && !$value instanceof DateTimeInterface) {
            return json_encode($value, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES | JSON_THROW_ON_ERROR);
        }
        return null;
    }

    /** @return array{0:int,1:int,2:int,3:int} */
    private function parseRange(string $range): array
    {
        [$start, $end] = array_pad(explode(':', strtoupper($range), 2), 2, strtoupper($range));
        [$startColumn, $startRow] = Coordinate::splitCellRef($start);
        [$endColumn, $endRow] = Coordinate::splitCellRef($end);
        if ($startColumn > 256 || $endColumn > 256 || $startRow > 65536 || $endRow > 65536) {
            throw UnsupportedXlsFeatureException::forFeature('merge range outside BIFF8 limits', ['range' => $range]);
        }
        return [min($startRow, $endRow) - 1, max($startRow, $endRow) - 1, min($startColumn, $endColumn) - 1, max($startColumn, $endColumn) - 1];
    }

    private function dateSerial(DateTimeInterface $date): float
    {
        $utc = DateTimeImmutable::createFromInterface($date);
        $base = new DateTimeImmutable($this->date1904 ? '1904-01-01 00:00:00' : '1899-12-31 00:00:00', $utc->getTimezone());
        $days = (int) $base->diff($utc)->format('%r%a');
        if (!$this->date1904 && $days >= 60) {
            $days++; // Excel's fake 1900-02-29
        }
        $seconds = ((int) $utc->format('H') * 3600) + ((int) $utc->format('i') * 60) + (int) $utc->format('s');
        return $days + ($seconds / 86400);
    }

    private function validateSheet(WorksheetData $sheet): void
    {
        $nameBytes = iconv('UTF-8', 'UTF-16LE', $sheet->name);
        if ($nameBytes === false || intdiv(strlen($nameBytes), 2) < 1 || intdiv(strlen($nameBytes), 2) > 31) {
            throw UnsupportedXlsFeatureException::forFeature('worksheet name outside the 1-31 character BIFF8 limit', ['sheet' => $sheet->name]);
        }
        if (preg_match('~[\\/?*\[\]:]~', $sheet->name) === 1) {
            throw UnsupportedXlsFeatureException::forFeature('worksheet name containing an invalid Excel character', ['sheet' => $sheet->name]);
        }
        $unsupported = [
            'images' => $sheet->images,
            'hyperlinks' => $sheet->hyperlinks,
            'comments' => $sheet->comments,
            'conditional formats' => $sheet->conditionalFormats,
            'data validations' => $sheet->dataValidations,
            'charts' => $sheet->charts,
            'pivot tables' => $sheet->pivotTables,
            'auto filter' => $sheet->autoFilter || $sheet->autoFilterRange !== null,
        ];
        if (($this->options['strict_features'] ?? false) === true) {
            foreach ($unsupported as $feature => $value) {
                if ($value !== [] && $value !== false && $value !== null) {
                    throw UnsupportedXlsFeatureException::forFeature($feature, ['sheet' => $sheet->name]);
                }
            }
        }
    }

    private static function errorCode(string $error): int
    {
        return self::errorCodeOrNull($error) ?? throw UnsupportedXlsFeatureException::forFeature('Excel error code ' . $error);
    }

    private static function errorCodeOrNull(string $error): ?int
    {
        return match (strtoupper($error)) {
            '#NULL!' => 0x00, '#DIV/0!' => 0x07, '#VALUE!' => 0x0F, '#REF!' => 0x17,
            '#NAME?' => 0x1D, '#NUM!' => 0x24, '#N/A' => 0x2A, default => null,
        };
    }
}
