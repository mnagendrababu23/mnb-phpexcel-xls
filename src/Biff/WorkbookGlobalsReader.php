<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Biff;

use Mnb\PHPExcel\Biff\String\BiffString;
use Mnb\PHPExcel\Biff\String\SharedStringTable;
use Mnb\PHPExcel\Exception\InvalidBiffRecordException;
use Mnb\PHPExcel\Support\Binary;

final class WorkbookGlobalsReader
{
    /** @param array<string,mixed> $options */
    public function read(string $stream, array $options = []): WorkbookInfo
    {
        $records = [];
        foreach ((new BiffRecordReader($stream, $options))->records(0) as $record) {
            $records[] = $record;
            if ($record->type === RecordType::EOF) {
                break;
            }
        }
        if ($records === [] || $records[0]->type !== RecordType::BOF) {
            throw InvalidBiffRecordException::because('Workbook globals BOF record is missing.');
        }
        if (strlen($records[0]->payload) < 4 || Binary::u16($records[0]->payload, 2) !== 0x0005) {
            throw InvalidBiffRecordException::because('First BIFF substream is not workbook globals.');
        }

        $sheets = [];
        $sharedStrings = null;
        $xfNumberFormats = [];
        $customFormats = [];
        $date1904 = false;
        $codePage = 1200;

        for ($i = 0, $count = count($records); $i < $count; $i++) {
            $record = $records[$i];
            switch ($record->type) {
                case RecordType::BOUNDSHEET:
                    if ($record->length() < 8) {
                        throw InvalidBiffRecordException::because('BOUNDSHEET record is too short.', ['offset' => $record->offset]);
                    }
                    $decoded = BiffString::readUnicodeString($record->payload, 6, true);
                    $sheets[] = [
                        'name' => $decoded['value'],
                        'offset' => Binary::u32($record->payload, 0),
                        'state' => Binary::u8($record->payload, 4),
                        'type' => Binary::u8($record->payload, 5),
                    ];
                    break;

                case RecordType::CODEPAGE:
                    if ($record->length() >= 2) {
                        $codePage = Binary::u16($record->payload, 0);
                    }
                    break;

                case RecordType::DATEMODE:
                    if ($record->length() >= 2) {
                        $date1904 = Binary::u16($record->payload, 0) === 1;
                    }
                    break;

                case RecordType::SST:
                    $segments = [$record->payload];
                    while (isset($records[$i + 1]) && $records[$i + 1]->type === RecordType::CONTINUE) {
                        $segments[] = $records[++$i]->payload;
                    }
                    $sharedStrings = SharedStringTable::fromSegments($segments, (int) ($options['max_shared_strings'] ?? 2_000_000));
                    break;

                case RecordType::FORMAT:
                    if ($record->length() >= 5) {
                        $formatId = Binary::u16($record->payload, 0);
                        $decoded = BiffString::readUnicodeString($record->payload, 2, false);
                        $customFormats[$formatId] = $decoded['value'];
                    }
                    break;

                case RecordType::XF:
                    if ($record->length() >= 4) {
                        $xfNumberFormats[] = Binary::u16($record->payload, 2);
                    }
                    break;
            }
        }

        if ($sheets === []) {
            throw InvalidBiffRecordException::because('Workbook contains no BOUNDSHEET records.');
        }

        return new WorkbookInfo($sheets, $sharedStrings, $xfNumberFormats, $customFormats, $date1904, $codePage);
    }
}
