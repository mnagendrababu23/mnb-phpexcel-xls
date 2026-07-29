<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Biff\String;

use Mnb\PHPExcel\Exception\InvalidBiffRecordException;
use Mnb\PHPExcel\Support\Binary;

final class BiffString
{
    private function __construct()
    {
    }

    /** @return array{value:string,next:int} */
    public static function readUnicodeString(string $data, int $offset = 0, bool $shortLength = false): array
    {
        $characterCount = $shortLength ? Binary::u8($data, $offset) : Binary::u16($data, $offset);
        $offset += $shortLength ? 1 : 2;
        $flags = Binary::u8($data, $offset++);
        $is16Bit = ($flags & 0x01) !== 0;
        $hasExtended = ($flags & 0x04) !== 0;
        $hasRich = ($flags & 0x08) !== 0;
        $runCount = $hasRich ? Binary::u16($data, $offset) : 0;
        if ($hasRich) {
            $offset += 2;
        }
        $extendedSize = $hasExtended ? Binary::u32($data, $offset) : 0;
        if ($hasExtended) {
            $offset += 4;
        }
        $byteLength = $characterCount * ($is16Bit ? 2 : 1);
        Binary::requireBytes($data, $offset, $byteLength + ($runCount * 4) + $extendedSize);
        $raw = substr($data, $offset, $byteLength);
        $value = self::decodeCharacters($raw, $is16Bit);
        $offset += $byteLength + ($runCount * 4) + $extendedSize;
        return ['value' => $value, 'next' => $offset];
    }

    public static function writeUnicodeString(string $value, bool $shortLength = false): string
    {
        $encoded = iconv('UTF-8', 'UTF-16LE', $value);
        if ($encoded === false) {
            throw new \InvalidArgumentException('Unable to encode BIFF Unicode string.');
        }
        $characters = intdiv(strlen($encoded), 2);
        if ($shortLength && $characters > 255) {
            throw new \InvalidArgumentException('BIFF short Unicode string cannot exceed 255 characters.');
        }
        if (!$shortLength && $characters > 65535) {
            throw new \InvalidArgumentException('BIFF Unicode string cannot exceed 65535 characters.');
        }
        return ($shortLength ? chr($characters) : pack('v', $characters)) . "\x01" . $encoded;
    }

    public static function decodeCharacters(string $raw, bool $is16Bit): string
    {
        $value = iconv($is16Bit ? 'UTF-16LE' : 'ISO-8859-1', 'UTF-8//IGNORE', $raw);
        if ($value === false) {
            throw InvalidBiffRecordException::because('Unable to decode BIFF string characters.');
        }
        return $value;
    }
}
