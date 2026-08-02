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
        $worksheetWriter = new WorksheetWriter($sharedStrings, $options);
        foreach ($sheets as $sheet) {
            $worksheetWriter->registerStrings($sheet);
        }

        $worksheetStreams = [];
        foreach ($sheets as $sheet) {
            $worksheetStreams[] = $worksheetWriter->write($sheet);
        }

        $placeholderOffsets = array_fill(0, count($sheets), 0);
        $globals = $this->workbookGlobals($workbook, $sharedStrings, $placeholderOffsets);
        $offsets = [];
        $offset = strlen($globals);
        foreach ($worksheetStreams as $sheetStream) {
            $offsets[] = $offset;
            $offset += strlen($sheetStream);
        }
        $globalsWithOffsets = $this->workbookGlobals($workbook, $sharedStrings, $offsets);
        if (strlen($globalsWithOffsets) !== strlen($globals)) {
            throw new \LogicException('BOUNDSHEET offset patching changed workbook global size.');
        }

        return $globalsWithOffsets . implode('', $worksheetStreams);
    }

    /** @param list<int> $sheetOffsets */
    private function workbookGlobals(WorkbookData $workbook, SharedStringWriter $sharedStrings, array $sheetOffsets): string
    {
        $stream = BiffRecordWriter::bof(0x0005);
        $stream .= BiffRecordWriter::record(RecordType::CODEPAGE, pack('v', 1200));
        $stream .= BiffRecordWriter::record(RecordType::WINDOW1, hex2bin('000000000040002038000000000001005802'));
        $stream .= BiffRecordWriter::record(RecordType::DATEMODE, pack('v', (int) (($workbook->metadata['date1904'] ?? false) === true)));
        $stream .= BiffRecordWriter::record(RecordType::CALCMODE, pack('v', 1));
        $stream .= BiffRecordWriter::record(RecordType::CALCCOUNT, pack('v', 100));
        $stream .= BiffRecordWriter::record(RecordType::REFMODE, pack('v', 1));
        $stream .= BiffRecordWriter::record(RecordType::ITERATION, pack('v', 0));
        $stream .= BiffRecordWriter::record(RecordType::DELTA, pack('e', 0.001));
        $stream .= BiffRecordWriter::record(RecordType::SAVERECALC, pack('v', 1));
        $stream .= BiffRecordWriter::record(RecordType::FONT, $this->fontRecord('Arial', 10));
        $stream .= BiffRecordWriter::record(RecordType::XF, $this->xfRecord(0));
        $stream .= BiffRecordWriter::record(RecordType::XF, $this->xfRecord(14));
        $stream .= BiffRecordWriter::record(RecordType::XF, $this->xfRecord(22));
        $stream .= BiffRecordWriter::record(RecordType::STYLE, pack('vCC', 0x8000, 0, 0xFF));

        foreach ($workbook->sheets as $index => $sheet) {
            $stream .= BiffRecordWriter::record(
                RecordType::BOUNDSHEET,
                pack('VCC', $sheetOffsets[$index] ?? 0, 0, 0) . BiffString::writeUnicodeString($sheet->name, true)
            );
        }
        $stream .= $sharedStrings->records();
        return $stream . BiffRecordWriter::eof();
    }

    private function fontRecord(string $name, int $points): string
    {
        return pack('vvvvvCCCC', $points * 20, 0, 0x7FFF, 400, 0, 0, 0, 0, 0)
            . BiffString::writeUnicodeString($name, true);
    }

    private function xfRecord(int $formatId): string
    {
        $base = hex2bin('000000000100200000000000000000000000C020');
        return pack('vv', 0, $formatId) . substr($base, 4);
    }
}
