<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader;

use DateTimeInterface;
use Mnb\PHPExcel\Biff\BiffRecordReader;
use Mnb\PHPExcel\Biff\BiffStyleMap;
use Mnb\PHPExcel\Biff\RecordType;
use Mnb\PHPExcel\Biff\WorkbookGlobalsReader;
use Mnb\PHPExcel\Compound\CompoundFileReader;
use Mnb\PHPExcel\Core\StyleNormalizer;
use Mnb\PHPExcel\Reader\State\FormulaResult;
use Mnb\PHPExcel\Snapshot\VisualSnapshot;
use Mnb\PHPExcel\Snapshot\VisualSnapshotReaderInterface;
use Mnb\PHPExcel\Support\Binary;
use Mnb\PHPExcel\Support\Coordinate;
use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;

/** Native BIFF8 visual snapshot reader with style and layout preservation. */
final class XlsVisualSnapshotReader implements VisualSnapshotReaderInterface
{
    /** @param array<string,mixed> $options @return array<string,mixed> */
    public function visualSnapshot(string $path, array $options = []): array
    {
        if (!is_file($path)) {
            throw MnbExcelException::withCode('XLS file not found: ' . $path, ErrorCode::FILE_NOT_FOUND, ['path' => $path]);
        }
        $compound = new CompoundFileReader($path, $options);
        $streamName = $compound->hasStream('Workbook') ? 'Workbook' : ($compound->hasStream('Book') ? 'Book' : null);
        if ($streamName === null) {
            throw new MnbExcelException('XLS compound file has no Workbook or Book stream.');
        }
        $stream = $compound->readStream($streamName);
        $workbook = (new WorkbookGlobalsReader())->read($stream, $options);
        $styleMap = BiffStyleMap::fromWorkbookStream($stream, $options);
        $activeGlobalSheet = $this->activeSheet($stream, $options, count($workbook->sheets));
        $activeWorksheet = 1;
        $styleTable = [];
        $styleIds = [];
        $sheets = [];
        $warnings = [];
        $maximumCells = max(1, (int) ($options['max_cells'] ?? 1_000_000));
        $cellCount = 0;
        $worksheetNumber = 0;

        foreach ($workbook->sheets as $globalSheetOffset => $sheetInfo) {
            if ((int) $sheetInfo['type'] !== 0) {
                if ($activeGlobalSheet === $globalSheetOffset + 1) {
                    $warnings[] = 'The active BIFF sheet is not a worksheet; the snapshot selected the first worksheet.';
                }
                continue;
            }
            $worksheetNumber++;
            if ($activeGlobalSheet === $globalSheetOffset + 1) {
                $activeWorksheet = $worksheetNumber;
            }
            $scan = $this->scanWorksheet($stream, (int) $sheetInfo['offset'], $options);
            $cellCount += count($scan['cells']);
            $gridCells = max(1, $scan['max_row']) * max(1, $scan['max_column']);
            if ($cellCount > $maximumCells || $gridCells > $maximumCells) {
                throw new MnbExcelException('Visual snapshot exceeds max_cells.');
            }
            $worksheetReader = new WorksheetReader($stream, $workbook, array_replace($options, [
                'formula_cells' => 'both',
                'return_datetime' => true,
                'format_dates' => true,
                'include_empty_cells' => true,
                'end_row' => max(1, $scan['max_row']),
                'max_worksheet_rows' => max(1, $scan['max_row']),
                'max_worksheet_columns' => max(1, $scan['max_column']),
            ]));
            $rows = array_values(iterator_to_array($worksheetReader->rows((int) $sheetInfo['offset']), false));
            $cells = [];
            foreach ($scan['cells'] as $coordinate => $xf) {
                [$column, $row] = Coordinate::splitCellRef($coordinate);
                $value = $rows[$row - 1][$column - 1] ?? null;
                $style = StyleNormalizer::normalize($styleMap->styleForIndex($xf));
                $styleId = $this->registerStyle($style, $styleTable, $styleIds);
                $cells[$coordinate] = $this->encodeCell($value, $style, $styleId);
            }
            $dimension = 'A1:' . Coordinate::columnIndexToName(max(1, $scan['max_column'])) . max(1, $scan['max_row']);
            $state = match ((int) $sheetInfo['state']) { 1 => 'hidden', 2 => 'veryHidden', default => 'visible' };
            $sheets[] = [
                'index' => $worksheetNumber,
                'name' => (string) $sheetInfo['name'],
                'state' => $state,
                'dimension' => $dimension,
                'cells' => $cells,
                'layout' => [
                    'column_widths' => $scan['column_widths'],
                    'row_heights' => $scan['row_heights'],
                    'freeze_panes' => $scan['freeze_panes'],
                    'auto_filter_range' => null,
                ],
                'merged_cells' => $scan['merged_cells'],
                'comments' => [],
                'hyperlinks' => [],
                'images' => [],
                'conditional_formats' => [],
                'data_validations' => [],
                'capabilities' => [
                    'values' => 'available',
                    'styles' => 'available_biff8_subset',
                    'layout' => 'available_biff8_subset',
                    'conditional_formats' => 'inventory_only',
                    'data_validations' => 'inventory_only',
                    'images' => 'inventory_only',
                ],
            ];
        }

        $metadata = [];
        try {
            $report = (new XlsMetadataReader())->metaInfo($path, ['profile' => 'quick']);
            $document = is_array($report['document'] ?? null) ? $report['document'] : [];
            $revision = is_array($report['revision'] ?? null) ? $report['revision'] : [];
            $application = is_array($report['application'] ?? null) ? $report['application'] : [];
            foreach (['title','subject','creator','keywords','description','category','manager','company'] as $key) {
                foreach ([$document, $application] as $section) {
                    if (isset($section[$key]) && is_scalar($section[$key])) { $metadata[$key] = $section[$key]; break; }
                }
            }
            if (isset($application['application_name']) && is_scalar($application['application_name'])) $metadata['application'] = $application['application_name'];
            foreach (['last_saved_by'=>'last_modified_by','revision_number'=>'revision_number','total_editing_time_seconds'=>'total_editing_time_seconds','created_at'=>'created_at','modified_at'=>'modified_at','last_printed_at'=>'last_printed_at'] as $source => $target) {
                if (isset($revision[$source]) && is_scalar($revision[$source])) $metadata[$target] = $revision[$source];
            }
            $custom = is_array($report['custom_properties'] ?? null) ? $report['custom_properties'] : [];
            if (is_array($custom['items'] ?? null)) $metadata['custom_properties'] = $custom['items'];
        } catch (\Throwable $error) {
            $warnings[] = 'Document metadata could not be added to the visual snapshot: ' . $error->getMessage();
        }

        $snapshot = [
            'schema' => VisualSnapshot::SCHEMA,
            'schema_version' => VisualSnapshot::VERSION,
            'format' => 'xls',
            'source' => ['file_name' => basename($path), 'size_bytes' => filesize($path) ?: 0],
            'workbook' => [
                'active_sheet' => min(max(1, $activeWorksheet), max(1, count($sheets))),
                'active_sheet_name' => $sheets[min(max(1, $activeWorksheet), max(1, count($sheets))) - 1]['name'] ?? '',
                'date1904' => $workbook->date1904,
                'metadata' => $metadata,
            ],
            'styles' => $styleTable,
            'sheets' => $sheets,
            'capabilities' => [
                'values' => 'available',
                'styles' => 'available_biff8_subset',
                'layout' => 'available_biff8_subset',
                'round_trip' => 'supported_for_native_subset',
            ],
            'warnings' => $warnings,
        ];
        VisualSnapshot::assertValid($snapshot);
        return $snapshot;
    }

    private function activeSheet(string $stream, array $options, int $sheetCount): int
    {
        foreach ((new BiffRecordReader($stream, $options))->records(0) as $record) {
            if ($record->type === RecordType::WINDOW1 && $record->length() >= 12) {
                return min(max(1, Binary::u16($record->payload, 10) + 1), max(1, $sheetCount));
            }
            if ($record->type === RecordType::EOF) break;
        }
        return 1;
    }

    /** @return array{cells:array<string,int>,max_row:int,max_column:int,column_widths:array<string,float>,row_heights:array<int,float>,freeze_panes:array<string,mixed>,merged_cells:list<string>} */
    private function scanWorksheet(string $stream, int $offset, array $options): array
    {
        $cells = [];
        $maxRow = 1;
        $maxColumn = 1;
        $widths = [];
        $heights = [];
        $freeze = ['rows' => 0, 'columns' => 0, 'top_left_cell' => null];
        $merged = [];
        foreach ((new BiffRecordReader($stream, $options))->records($offset) as $record) {
            if ($record->type === RecordType::EOF) break;
            $payload = $record->payload;
            if ($record->type === RecordType::DIMENSIONS && $record->length() >= 12) {
                // DIMENSIONS is only a producer hint and can be stale or maliciously huge.
                // Actual cell and merge records below define the restorable grid.
            } elseif (in_array($record->type, [RecordType::NUMBER,RecordType::RK,RecordType::LABEL,RecordType::LABELSST,RecordType::BOOLERR,RecordType::BLANK,RecordType::FORMULA], true) && $record->length() >= 6) {
                $row = Binary::u16($payload, 0) + 1;
                $column = Binary::u16($payload, 2) + 1;
                $cells[Coordinate::columnIndexToName($column) . $row] = Binary::u16($payload, 4);
                $maxRow = max($maxRow, $row); $maxColumn = max($maxColumn, $column);
            } elseif ($record->type === RecordType::MULRK && $record->length() >= 12) {
                $row = Binary::u16($payload, 0) + 1;
                $first = Binary::u16($payload, 2);
                $last = Binary::u16($payload, strlen($payload) - 2);
                for ($column = $first; $column <= $last; $column++) {
                    $xf = Binary::u16($payload, 4 + (($column - $first) * 6));
                    $cells[Coordinate::columnIndexToName($column + 1) . $row] = $xf;
                }
                $maxRow = max($maxRow, $row); $maxColumn = max($maxColumn, $last + 1);
            } elseif ($record->type === RecordType::MULBLANK && $record->length() >= 8) {
                $row = Binary::u16($payload, 0) + 1;
                $first = Binary::u16($payload, 2);
                $last = Binary::u16($payload, strlen($payload) - 2);
                for ($column = $first; $column <= $last; $column++) {
                    $xf = Binary::u16($payload, 4 + (($column - $first) * 2));
                    $cells[Coordinate::columnIndexToName($column + 1) . $row] = $xf;
                }
                $maxRow = max($maxRow, $row); $maxColumn = max($maxColumn, $last + 1);
            } elseif ($record->type === RecordType::COLINFO && $record->length() >= 12) {
                $first = Binary::u16($payload, 0); $last = Binary::u16($payload, 2); $width = Binary::u16($payload, 4) / 256;
                for ($column = $first; $column <= $last; $column++) $widths[Coordinate::columnIndexToName($column + 1)] = $width;
            } elseif ($record->type === RecordType::ROW && $record->length() >= 16) {
                $row = Binary::u16($payload, 0) + 1;
                $height = Binary::u16($payload, 6) & 0x7FFF;
                if ($height > 0) $heights[$row] = $height / 20;
            } elseif ($record->type === RecordType::PANE && $record->length() >= 8) {
                $columns = Binary::u16($payload, 0); $rows = Binary::u16($payload, 2);
                $freeze = ['rows' => $rows, 'columns' => $columns, 'top_left_cell' => Coordinate::columnIndexToName($columns + 1) . ($rows + 1)];
            } elseif ($record->type === RecordType::MERGEDCELLS && $record->length() >= 2) {
                $count = Binary::u16($payload, 0);
                for ($i = 0; $i < $count && 2 + ($i * 8) + 7 < strlen($payload); $i++) {
                    $position = 2 + ($i * 8);
                    $r1 = Binary::u16($payload, $position) + 1; $r2 = Binary::u16($payload, $position + 2) + 1;
                    $c1 = Binary::u16($payload, $position + 4) + 1; $c2 = Binary::u16($payload, $position + 6) + 1;
                    $merged[] = Coordinate::columnIndexToName($c1) . $r1 . ':' . Coordinate::columnIndexToName($c2) . $r2;
                    $maxRow = max($maxRow, $r2);
                    $maxColumn = max($maxColumn, $c2);
                }
            }
        }
        return ['cells'=>$cells,'max_row'=>$maxRow,'max_column'=>$maxColumn,'column_widths'=>$widths,'row_heights'=>$heights,'freeze_panes'=>$freeze,'merged_cells'=>$merged];
    }

    /** @param array<string,mixed> $style @param array<string,array<string,mixed>> $table @param array<string,string> $ids */
    private function registerStyle(array $style, array &$table, array &$ids): ?string
    {
        if ($style === []) return null;
        $key = json_encode($style, JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE);
        if (!is_string($key)) return null;
        if (!isset($ids[$key])) { $id = 's' . (count($table) + 1); $ids[$key] = $id; $table[$id] = $style; }
        return $ids[$key];
    }

    /** @param array<string,mixed> $style @return array<string,mixed> */
    private function encodeCell(mixed $value, array $style, ?string $styleId): array
    {
        $cell = [];
        if ($styleId !== null) $cell['style_id'] = $styleId;
        if ($value instanceof FormulaResult) {
            $cell['type'] = 'formula'; $cell['formula'] = ltrim($value->formula, '='); $cell['cached_value'] = $value->cachedValue; return $cell;
        }
        if ($value instanceof DateTimeInterface) {
            $cell['type'] = 'date'; $cell['value'] = $value->format('Y-m-d\TH:i:sP'); $cell['format'] = (string) ($style['format'] ?? 'm/d/yy'); return $cell;
        }
        if ($value === null) { $cell['type'] = 'blank'; $cell['value'] = null; }
        elseif (is_bool($value)) { $cell['type'] = 'boolean'; $cell['value'] = $value; }
        elseif (is_int($value) || is_float($value)) { $cell['type'] = 'number'; $cell['value'] = $value; }
        elseif (is_string($value) && str_starts_with($value, '#') && in_array($value, ['#NULL!','#DIV/0!','#VALUE!','#REF!','#NAME?','#NUM!','#N/A'], true)) { $cell['type'] = 'error'; $cell['value'] = $value; }
        else { $cell['type'] = 'text'; $cell['value'] = (string) $value; }
        return $cell;
    }
}
