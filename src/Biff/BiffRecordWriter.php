<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Biff;

final class BiffRecordWriter
{
    private function __construct()
    {
    }

    public static function record(int $type, string $payload = ''): string
    {
        $length = strlen($payload);
        if ($length > 8224) {
            throw new \InvalidArgumentException(sprintf('BIFF record 0x%04X exceeds 8224 bytes.', $type));
        }
        return pack('vv', $type, $length) . $payload;
    }

    public static function bof(int $substreamType): string
    {
        return self::record(RecordType::BOF, pack('vvvvVV', 0x0600, $substreamType, 0x0DBB, 0x07CC, 0x00000041, 0x00000006));
    }

    public static function eof(): string
    {
        return self::record(RecordType::EOF);
    }
}
