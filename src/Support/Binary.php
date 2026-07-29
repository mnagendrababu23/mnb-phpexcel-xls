<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Support;

use Mnb\PHPExcel\Exception\InvalidBiffRecordException;

final class Binary
{
    private function __construct()
    {
    }

    public static function u8(string $data, int $offset): int
    {
        self::requireBytes($data, $offset, 1);
        return ord($data[$offset]);
    }

    public static function u16(string $data, int $offset): int
    {
        self::requireBytes($data, $offset, 2);
        return unpack('v', substr($data, $offset, 2))[1];
    }

    public static function u32(string $data, int $offset): int
    {
        self::requireBytes($data, $offset, 4);
        return unpack('V', substr($data, $offset, 4))[1];
    }

    public static function i32(string $data, int $offset): int
    {
        $value = self::u32($data, $offset);
        return $value >= 0x80000000 ? $value - 0x100000000 : $value;
    }

    public static function u64(string $data, int $offset): int
    {
        self::requireBytes($data, $offset, 8);
        $parts = unpack('Vlow/Vhigh', substr($data, $offset, 8));
        if ($parts['high'] > 0x7FFFFFFF || ($parts['high'] !== 0 && PHP_INT_SIZE < 8)) {
            throw InvalidBiffRecordException::because('64-bit binary value exceeds the supported PHP integer range.', ['offset' => $offset]);
        }
        return (int) ($parts['low'] + ($parts['high'] * 4294967296));
    }

    public static function double(string $data, int $offset): float
    {
        self::requireBytes($data, $offset, 8);
        return unpack('e', substr($data, $offset, 8))[1];
    }

    public static function packU64(int $value): string
    {
        if ($value < 0) {
            throw new \InvalidArgumentException('Unsigned 64-bit value cannot be negative.');
        }
        return pack('V2', $value & 0xFFFFFFFF, intdiv($value, 4294967296));
    }

    public static function requireBytes(string $data, int $offset, int $length): void
    {
        if ($offset < 0 || $length < 0 || $offset + $length > strlen($data)) {
            throw InvalidBiffRecordException::because('Unexpected end of binary data.', [
                'offset' => $offset,
                'required_bytes' => $length,
                'available_bytes' => max(0, strlen($data) - $offset),
            ]);
        }
    }
}
