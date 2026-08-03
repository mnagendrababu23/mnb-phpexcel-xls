<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Writer;

use Mnb\PHPExcel\Biff\BiffRecordWriter;
use Mnb\PHPExcel\Biff\RecordType;
use Mnb\PHPExcel\Biff\String\BiffString;
use Mnb\PHPExcel\Core\WorkbookData;
use Mnb\PHPExcel\Exception\UnsupportedXlsFeatureException;

final class WorkbookWriter
{
    /** @param array<string,mixed> $options */
    public function write(WorkbookData $workbook, array $options = []): string
    {
        $sheets = $workbook->sheets;
        if ($sheets === []) {
            throw UnsupportedXlsFeatureException::forFeature('workbook without worksheets');
        }
        if (count($sheets) > 255) {
            throw UnsupportedXlsFeatureException::forFeature('more than 255 worksheets', ['count' => count($sheets)]);
        }
        $names = [];
        foreach ($sheets as $sheet) {
            $key = strtolower($sheet->name);
            if (isset($names[$key])) {
                throw UnsupportedXlsFeatureException::forFeature('duplicate worksheet name', ['sheet' => $sheet->name]);
            }
            $names[$key] = true;
        }

        $sharedStrings = new SharedStringWriter();
        $styles = new XlsStyleRegistry();
        $styles->registerWorkbook($workbook);
        $worksheetWriter = new WorksheetWriter($sharedStrings, $styles, (bool) ($workbook->metadata['date1904'] ?? false), $options);
        foreach ($sheets as $sheet) {
            $worksheetWriter->registerStrings($sheet);
        }

        $worksheetStreams = [];
        foreach ($sheets as $sheet) {
            $worksheetStreams[] = $worksheetWriter->write($sheet);
        }

        $placeholderOffsets = array_fill(0, count($sheets), 0);
        $globals = $this->workbookGlobals($workbook, $sharedStrings, $styles, $placeholderOffsets);
        $offsets = [];
        $offset = strlen($globals);
        foreach ($worksheetStreams as $sheetStream) {
            $offsets[] = $offset;
            $offset += strlen($sheetStream);
        }
        $globalsWithOffsets = $this->workbookGlobals($workbook, $sharedStrings, $styles, $offsets);
        if (strlen($globalsWithOffsets) !== strlen($globals)) {
            throw new \LogicException('BOUNDSHEET offset patching changed workbook global size.');
        }

        return $globalsWithOffsets . implode('', $worksheetStreams);
    }

    /** @param list<int> $sheetOffsets */
    private function workbookGlobals(WorkbookData $workbook, SharedStringWriter $sharedStrings, XlsStyleRegistry $styles, array $sheetOffsets): string
    {
        $stream = BiffRecordWriter::bof(0x0005);
        $stream .= BiffRecordWriter::record(RecordType::CODEPAGE, pack('v', 1200));
        $activeSheet = $workbook->metadata['_mnb_active_sheet'] ?? 1;
        if (is_string($activeSheet) && !ctype_digit($activeSheet)) {
            foreach ($workbook->sheets as $index => $candidate) {
                if ($candidate->name === $activeSheet) { $activeSheet = $index + 1; break; }
            }
        }
        $activeTab = max(0, min(count($workbook->sheets) - 1, ((int) $activeSheet) - 1));
        $stream .= BiffRecordWriter::record(RecordType::WINDOW1, pack('vvvvvvvvv', 0, 0, 0x4000, 0x2000, 0x0038, $activeTab, 0, 1, 600));
        $stream .= BiffRecordWriter::record(RecordType::DATEMODE, pack('v', (int) (($workbook->metadata['date1904'] ?? false) === true)));
        $stream .= BiffRecordWriter::record(RecordType::CALCMODE, pack('v', 1));
        $stream .= BiffRecordWriter::record(RecordType::CALCCOUNT, pack('v', 100));
        $stream .= BiffRecordWriter::record(RecordType::REFMODE, pack('v', 1));
        $stream .= BiffRecordWriter::record(RecordType::ITERATION, pack('v', 0));
        $stream .= BiffRecordWriter::record(RecordType::DELTA, pack('e', 0.001));
        $stream .= BiffRecordWriter::record(RecordType::SAVERECALC, pack('v', 1));
        $stream .= $styles->globalsRecords();
        $stream .= BiffRecordWriter::record(RecordType::STYLE, pack('vCC', 0x8000, 0, 0xFF));

        $states = is_array($workbook->metadata['_mnb_sheet_states'] ?? null) ? $workbook->metadata['_mnb_sheet_states'] : [];
        foreach ($workbook->sheets as $index => $sheet) {
            $stateName = (string) ($states[$sheet->name] ?? $states[$index + 1] ?? 'visible');
            $state = match ($stateName) { 'hidden' => 1, 'veryHidden' => 2, default => 0 };
            $stream .= BiffRecordWriter::record(
                RecordType::BOUNDSHEET,
                pack('VCC', $sheetOffsets[$index] ?? 0, $state, 0) . BiffString::writeUnicodeString($sheet->name, true)
            );
        }
        $stream .= $sharedStrings->records();
        return $stream . BiffRecordWriter::eof();
    }

}
