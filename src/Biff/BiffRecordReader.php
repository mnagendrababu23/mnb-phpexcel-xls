<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Biff;

use Mnb\PHPExcel\Exception\InvalidBiffRecordException;
use Mnb\PHPExcel\Support\Binary;

final class BiffRecordReader
{
    /** @param array<string,mixed> $options */
    public function __construct(private readonly string $stream, private readonly array $options = [])
    {
    }

    /** @return \Generator<int,BiffRecord> */
    public function records(int $offset = 0): \Generator
    {
        $length = strlen($this->stream);
        $count = 0;
        $maxRecords = (int) ($this->options['max_biff_records'] ?? 10_000_000);
        while ($offset + 4 <= $length) {
            $type = Binary::u16($this->stream, $offset);
            $recordLength = Binary::u16($this->stream, $offset + 2);
            if ($type === 0 && $recordLength === 0) {
                break; // padded CFB stream after workbook EOF
            }
            if ($recordLength > 8224) {
                throw InvalidBiffRecordException::because('BIFF record exceeds the 8224-byte limit.', [
                    'offset' => $offset,
                    'type' => sprintf('0x%04X', $type),
                    'length' => $recordLength,
                ]);
            }
            if ($offset + 4 + $recordLength > $length) {
                throw InvalidBiffRecordException::because('BIFF record is truncated.', [
                    'offset' => $offset,
                    'type' => sprintf('0x%04X', $type),
                    'declared_length' => $recordLength,
                    'remaining' => $length - $offset - 4,
                ]);
            }
            if (++$count > $maxRecords) {
                throw InvalidBiffRecordException::because('BIFF record count exceeds the configured limit.', ['limit' => $maxRecords]);
            }
            yield $offset => new BiffRecord($type, substr($this->stream, $offset + 4, $recordLength), $offset);
            $offset += 4 + $recordLength;
        }
    }
}
