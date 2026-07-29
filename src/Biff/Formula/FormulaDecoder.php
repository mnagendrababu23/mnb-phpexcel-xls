<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Biff\Formula;

use Mnb\PHPExcel\Exception\InvalidBiffRecordException;
use Mnb\PHPExcel\Support\Binary;
use Mnb\PHPExcel\Support\Coordinate;

/** Best-effort BIFF8 RPN token decoder for common worksheet formulas. */
final class FormulaDecoder
{
    /** @param list<string> $sheetNames */
    public function decode(string $tokens, array $sheetNames = []): string
    {
        $stack = [];
        $offset = 0;
        $length = strlen($tokens);
        while ($offset < $length) {
            $rawToken = ord($tokens[$offset++]);
            $token = $rawToken >= 0x20 ? (($rawToken & 0x1F) | 0x20) : $rawToken;
            switch ($token) {
                case 0x03: $this->binary($stack, '+'); break;
                case 0x04: $this->binary($stack, '-'); break;
                case 0x05: $this->binary($stack, '*'); break;
                case 0x06: $this->binary($stack, '/'); break;
                case 0x07: $this->binary($stack, '^'); break;
                case 0x08: $this->binary($stack, '&'); break;
                case 0x09: $this->binary($stack, '<'); break;
                case 0x0A: $this->binary($stack, '<='); break;
                case 0x0B: $this->binary($stack, '='); break;
                case 0x0C: $this->binary($stack, '>='); break;
                case 0x0D: $this->binary($stack, '>'); break;
                case 0x0E: $this->binary($stack, '<>'); break;
                case 0x0F: $this->binary($stack, ' '); break;
                case 0x10: $this->binary($stack, ','); break;
                case 0x11: $this->binary($stack, ':'); break;
                case 0x12: $stack[] = '+' . $this->pop($stack); break;
                case 0x13: $stack[] = '-' . $this->pop($stack); break;
                case 0x14: $stack[] = $this->pop($stack) . '%'; break;
                case 0x15: $stack[] = '(' . $this->pop($stack) . ')'; break;
                case 0x16: $stack[] = ''; break;
                case 0x17:
                    Binary::requireBytes($tokens, $offset, 2);
                    $count = ord($tokens[$offset++]);
                    $flags = ord($tokens[$offset++]);
                    $is16 = ($flags & 1) !== 0;
                    $bytes = $count * ($is16 ? 2 : 1);
                    Binary::requireBytes($tokens, $offset, $bytes);
                    $raw = substr($tokens, $offset, $bytes);
                    $offset += $bytes;
                    $value = iconv($is16 ? 'UTF-16LE' : 'ISO-8859-1', 'UTF-8//IGNORE', $raw);
                    $stack[] = '"' . str_replace('"', '""', $value === false ? '' : $value) . '"';
                    break;
                case 0x19: // ptgAttr
                    Binary::requireBytes($tokens, $offset, 3);
                    $flags = ord($tokens[$offset]);
                    $offset += 3;
                    if (($flags & 0x10) !== 0 && $stack !== []) {
                        $stack[] = 'SUM(' . $this->pop($stack) . ')';
                    }
                    break;
                case 0x1C:
                    Binary::requireBytes($tokens, $offset, 1);
                    $stack[] = self::errorText(ord($tokens[$offset++]));
                    break;
                case 0x1D:
                    Binary::requireBytes($tokens, $offset, 1);
                    $stack[] = ord($tokens[$offset++]) === 0 ? 'FALSE' : 'TRUE';
                    break;
                case 0x1E:
                    $stack[] = (string) Binary::u16($tokens, $offset);
                    $offset += 2;
                    break;
                case 0x1F:
                    $stack[] = self::numberText(Binary::double($tokens, $offset));
                    $offset += 8;
                    break;
                case 0x21:
                    $functionId = Binary::u16($tokens, $offset) & 0x7FFF;
                    $offset += 2;
                    $definition = FunctionRegistry::byId($functionId);
                    $argCount = $definition['args'] ?? 0;
                    $this->function($stack, $definition['name'] ?? ('FUNC' . $functionId), $argCount);
                    break;
                case 0x22:
                    Binary::requireBytes($tokens, $offset, 3);
                    $argCount = ord($tokens[$offset++]);
                    $functionId = Binary::u16($tokens, $offset) & 0x7FFF;
                    $offset += 2;
                    $definition = FunctionRegistry::byId($functionId);
                    $this->function($stack, $definition['name'] ?? ('FUNC' . $functionId), $argCount);
                    break;
                case 0x23: // name
                    Binary::requireBytes($tokens, $offset, 4);
                    $nameIndex = Binary::u32($tokens, $offset);
                    $offset += 4;
                    $stack[] = '#NAME[' . $nameIndex . ']';
                    break;
                case 0x24:
                    Binary::requireBytes($tokens, $offset, 4);
                    $row = Binary::u16($tokens, $offset);
                    $columnFlags = Binary::u16($tokens, $offset + 2);
                    $offset += 4;
                    $stack[] = $this->cellReference($row, $columnFlags);
                    break;
                case 0x25:
                    Binary::requireBytes($tokens, $offset, 8);
                    $row1 = Binary::u16($tokens, $offset);
                    $row2 = Binary::u16($tokens, $offset + 2);
                    $col1 = Binary::u16($tokens, $offset + 4);
                    $col2 = Binary::u16($tokens, $offset + 6);
                    $offset += 8;
                    $stack[] = $this->cellReference($row1, $col1) . ':' . $this->cellReference($row2, $col2);
                    break;
                case 0x26: // MemArea
                case 0x27: // MemErr
                case 0x28: // MemNoMem
                    Binary::requireBytes($tokens, $offset, 6);
                    $offset += 6;
                    break;
                case 0x29: // MemFunc
                    Binary::requireBytes($tokens, $offset, 2);
                    $offset += 2;
                    break;
                case 0x2A:
                    Binary::requireBytes($tokens, $offset, 4);
                    $offset += 4;
                    $stack[] = '#REF!';
                    break;
                case 0x2B:
                    Binary::requireBytes($tokens, $offset, 8);
                    $offset += 8;
                    $stack[] = '#REF!:#REF!';
                    break;
                case 0x2C:
                    Binary::requireBytes($tokens, $offset, 4);
                    $row = Binary::u16($tokens, $offset);
                    $columnFlags = Binary::u16($tokens, $offset + 2);
                    $offset += 4;
                    $stack[] = $this->cellReference($row, $columnFlags);
                    break;
                case 0x2D:
                    Binary::requireBytes($tokens, $offset, 8);
                    $row1 = Binary::u16($tokens, $offset);
                    $row2 = Binary::u16($tokens, $offset + 2);
                    $col1 = Binary::u16($tokens, $offset + 4);
                    $col2 = Binary::u16($tokens, $offset + 6);
                    $offset += 8;
                    $stack[] = $this->cellReference($row1, $col1) . ':' . $this->cellReference($row2, $col2);
                    break;
                case 0x39: // NameX
                    Binary::requireBytes($tokens, $offset, 6);
                    $sheetIndex = Binary::u16($tokens, $offset);
                    $nameIndex = Binary::u32($tokens, $offset + 2);
                    $offset += 6;
                    $stack[] = '#EXTERNAL_NAME[' . $sheetIndex . ',' . $nameIndex . ']';
                    break;
                case 0x3A:
                    Binary::requireBytes($tokens, $offset, 6);
                    $sheetIndex = Binary::u16($tokens, $offset);
                    $row = Binary::u16($tokens, $offset + 2);
                    $col = Binary::u16($tokens, $offset + 4);
                    $offset += 6;
                    $stack[] = $this->sheetPrefix($sheetIndex, $sheetNames) . $this->cellReference($row, $col);
                    break;
                case 0x3B:
                    Binary::requireBytes($tokens, $offset, 10);
                    $sheetIndex = Binary::u16($tokens, $offset);
                    $row1 = Binary::u16($tokens, $offset + 2);
                    $row2 = Binary::u16($tokens, $offset + 4);
                    $col1 = Binary::u16($tokens, $offset + 6);
                    $col2 = Binary::u16($tokens, $offset + 8);
                    $offset += 10;
                    $prefix = $this->sheetPrefix($sheetIndex, $sheetNames);
                    $stack[] = $prefix . $this->cellReference($row1, $col1) . ':' . $this->cellReference($row2, $col2);
                    break;
                default:
                    throw InvalidBiffRecordException::because('Unsupported BIFF formula token.', [
                        'token' => sprintf('0x%02X', $rawToken),
                        'normalized_token' => sprintf('0x%02X', $token),
                        'offset' => $offset - 1,
                    ]);
            }
        }
        return $stack === [] ? '' : (string) end($stack);
    }

    /** @param list<string> $stack */
    private function binary(array &$stack, string $operator): void
    {
        $right = $this->pop($stack);
        $left = $this->pop($stack);
        $stack[] = $left . $operator . $right;
    }

    /** @param list<string> $stack */
    private function function(array &$stack, string $name, int $argCount): void
    {
        if ($argCount < 0 || $argCount > count($stack)) {
            throw InvalidBiffRecordException::because('Formula function argument count is invalid.', ['function' => $name, 'arguments' => $argCount, 'stack' => count($stack)]);
        }
        $args = $argCount === 0 ? [] : array_splice($stack, -$argCount);
        $stack[] = $name . '(' . implode(',', $args) . ')';
    }

    /** @param list<string> $stack */
    private function pop(array &$stack): string
    {
        if ($stack === []) {
            throw InvalidBiffRecordException::because('Formula token stack underflow.');
        }
        return (string) array_pop($stack);
    }

    private function cellReference(int $row, int $columnFlags): string
    {
        $column = ($columnFlags & 0x00FF) + 1;
        $columnRelative = ($columnFlags & 0x4000) !== 0;
        $rowRelative = ($columnFlags & 0x8000) !== 0;
        return ($columnRelative ? '' : '$') . Coordinate::columnIndexToName($column) . ($rowRelative ? '' : '$') . ($row + 1);
    }

    /** @param list<string> $sheetNames */
    private function sheetPrefix(int $index, array $sheetNames): string
    {
        $name = $sheetNames[$index] ?? ('Sheet' . ($index + 1));
        if (preg_match('/^[A-Za-z_][A-Za-z0-9_.]*$/', $name) !== 1) {
            $name = "'" . str_replace("'", "''", $name) . "'";
        }
        return $name . '!';
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
        return rtrim(rtrim(sprintf('%.15G', $number), '0'), '.');
    }
}
