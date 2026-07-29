<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Biff\Formula;

use Mnb\PHPExcel\Exception\UnsupportedXlsFeatureException;
use Mnb\PHPExcel\Support\Coordinate;

/**
 * Small native formula compiler for common BIFF8 formulas.
 *
 * Supported: literals, A1 references/ranges, arithmetic/comparison operators,
 * concatenation, percent, parentheses, and functions in FunctionRegistry.
 */
final class FormulaEncoder
{
    /** @var list<array{type:string,value:string}> */
    private array $tokens = [];
    private int $position = 0;

    public function encode(string $formula): string
    {
        $formula = trim($formula);
        if (str_starts_with($formula, '=')) {
            $formula = substr($formula, 1);
        }
        if ($formula === '') {
            throw UnsupportedXlsFeatureException::forFeature('empty formula');
        }
        $this->tokens = $this->tokenize($formula);
        $this->position = 0;
        $result = $this->parseExpression(0);
        if ($this->current() !== null) {
            throw UnsupportedXlsFeatureException::forFeature('formula syntax near ' . $this->current()['value']);
        }
        return $result;
    }

    /** @return list<array{type:string,value:string}> */
    private function tokenize(string $formula): array
    {
        $tokens = [];
        $offset = 0;
        $length = strlen($formula);
        while ($offset < $length) {
            if (preg_match('/\G\s+/A', $formula, $match, 0, $offset) === 1) {
                $offset += strlen($match[0]);
                continue;
            }
            if (preg_match('/\G"(?:[^"]|"")*"/A', $formula, $match, 0, $offset) === 1) {
                $tokens[] = ['type' => 'string', 'value' => $match[0]];
                $offset += strlen($match[0]);
                continue;
            }
            if (preg_match('/\G(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?/A', $formula, $match, 0, $offset) === 1) {
                $tokens[] = ['type' => 'number', 'value' => $match[0]];
                $offset += strlen($match[0]);
                continue;
            }
            if (preg_match('/\G\$?[A-Za-z]{1,3}\$?\d+/A', $formula, $match, 0, $offset) === 1) {
                $tokens[] = ['type' => 'cell', 'value' => strtoupper($match[0])];
                $offset += strlen($match[0]);
                continue;
            }
            if (preg_match('/\G(?:#[A-Za-z0-9\/]+[!?]?|[A-Za-z_][A-Za-z0-9_.]*)/A', $formula, $match, 0, $offset) === 1) {
                $tokens[] = ['type' => 'identifier', 'value' => strtoupper($match[0])];
                $offset += strlen($match[0]);
                continue;
            }
            $two = substr($formula, $offset, 2);
            if (in_array($two, ['<=', '>=', '<>'], true)) {
                $tokens[] = ['type' => 'operator', 'value' => $two];
                $offset += 2;
                continue;
            }
            $one = $formula[$offset];
            if (str_contains('+-*/^&=<>%():,', $one)) {
                $tokens[] = ['type' => in_array($one, ['(', ')', ',', ':'], true) ? $one : 'operator', 'value' => $one];
                $offset++;
                continue;
            }
            throw UnsupportedXlsFeatureException::forFeature('formula token ' . $one, ['offset' => $offset]);
        }
        return $tokens;
    }

    private function parseExpression(int $minimumPrecedence): string
    {
        $left = $this->parsePrefix();
        while (($token = $this->current()) !== null) {
            if ($token['type'] === 'operator' && $token['value'] === '%') {
                $this->position++;
                $left .= "\x14";
                continue;
            }
            if ($token['type'] === ':') {
                $precedence = 90;
                if ($precedence < $minimumPrecedence) {
                    break;
                }
                $this->position++;
                $right = $this->parseExpression($precedence + 1);
                // A range is emitted directly when both sides are references.
                $leftRef = $this->decodeReferenceToken($left);
                $rightRef = $this->decodeReferenceToken($right);
                if ($leftRef === null || $rightRef === null) {
                    $left .= $right . "\x11";
                } else {
                    $left = "\x25" . pack('vvvv', $leftRef['row'], $rightRef['row'], $leftRef['columnFlags'], $rightRef['columnFlags']);
                }
                continue;
            }
            if ($token['type'] !== 'operator' || !isset(self::PRECEDENCE[$token['value']])) {
                break;
            }
            $precedence = self::PRECEDENCE[$token['value']];
            if ($precedence < $minimumPrecedence) {
                break;
            }
            $operator = $token['value'];
            $this->position++;
            $nextMinimum = $operator === '^' ? $precedence : $precedence + 1;
            $right = $this->parseExpression($nextMinimum);
            $left .= $right . self::operatorToken($operator);
        }
        return $left;
    }

    private function parsePrefix(): string
    {
        $token = $this->current();
        if ($token === null) {
            throw UnsupportedXlsFeatureException::forFeature('incomplete formula');
        }
        if ($token['type'] === 'operator' && in_array($token['value'], ['+', '-'], true)) {
            $this->position++;
            return $this->parseExpression(80) . ($token['value'] === '+' ? "\x12" : "\x13");
        }
        if ($token['type'] === '(') {
            $this->position++;
            $value = $this->parseExpression(0);
            $this->expect(')');
            return $value . "\x15";
        }
        if ($token['type'] === 'number') {
            $this->position++;
            $number = (float) $token['value'];
            if (ctype_digit($token['value']) && $number >= 0 && $number <= 65535) {
                return "\x1E" . pack('v', (int) $number);
            }
            return "\x1F" . pack('e', $number);
        }
        if ($token['type'] === 'string') {
            $this->position++;
            $value = str_replace('""', '"', substr($token['value'], 1, -1));
            $encoded = iconv('UTF-8', 'UTF-16LE', $value);
            if ($encoded === false || intdiv(strlen($encoded), 2) > 255) {
                throw UnsupportedXlsFeatureException::forFeature('formula string longer than 255 characters');
            }
            return "\x17" . chr(intdiv(strlen($encoded), 2)) . "\x01" . $encoded;
        }
        if ($token['type'] === 'cell') {
            $this->position++;
            return $this->referenceToken($token['value']);
        }
        if ($token['type'] === 'identifier') {
            $this->position++;
            $identifier = $token['value'];
            if ($identifier === 'TRUE' || $identifier === 'FALSE') {
                return "\x1D" . ($identifier === 'TRUE' ? "\x01" : "\x00");
            }
            $error = self::errorCode($identifier);
            if ($error !== null) {
                return "\x1C" . chr($error);
            }
            if ($this->current()['type'] ?? null === '(') {
                $this->position++;
                $args = [];
                if (($this->current()['type'] ?? null) !== ')') {
                    do {
                        $args[] = $this->parseExpression(0);
                        if (($this->current()['type'] ?? null) !== ',') {
                            break;
                        }
                        $this->position++;
                    } while (true);
                }
                $this->expect(')');
                $definition = FunctionRegistry::byName($identifier);
                if ($definition === null) {
                    throw UnsupportedXlsFeatureException::forFeature('formula function ' . $identifier);
                }
                if ($definition['args'] !== null && $definition['args'] !== count($args)) {
                    throw UnsupportedXlsFeatureException::forFeature(sprintf('%s with %d arguments', $identifier, count($args)));
                }
                if (count($args) > 255) {
                    throw UnsupportedXlsFeatureException::forFeature('formula function with more than 255 arguments');
                }
                return implode('', $args) . "\x22" . chr(count($args)) . pack('v', $definition['id']);
            }
            throw UnsupportedXlsFeatureException::forFeature('defined name ' . $identifier);
        }
        throw UnsupportedXlsFeatureException::forFeature('formula syntax near ' . $token['value']);
    }

    private function referenceToken(string $reference): string
    {
        if (preg_match('/^(\$?)([A-Z]{1,3})(\$?)(\d+)$/', $reference, $match) !== 1) {
            throw UnsupportedXlsFeatureException::forFeature('cell reference ' . $reference);
        }
        $column = Coordinate::columnNameToIndex($match[2]);
        $row = (int) $match[4];
        if ($column > 256 || $row < 1 || $row > 65536) {
            throw UnsupportedXlsFeatureException::forFeature('XLS cell outside BIFF8 limits', ['cell' => $reference]);
        }
        $columnFlags = ($column - 1)
            | ($match[1] === '$' ? 0 : 0x4000)
            | ($match[3] === '$' ? 0 : 0x8000);
        return "\x24" . pack('vv', $row - 1, $columnFlags);
    }

    /** @return array{row:int,columnFlags:int}|null */
    private function decodeReferenceToken(string $bytes): ?array
    {
        if (strlen($bytes) !== 5 || ord($bytes[0]) !== 0x24) {
            return null;
        }
        $values = unpack('vrow/vcolumnFlags', substr($bytes, 1));
        return ['row' => $values['row'], 'columnFlags' => $values['columnFlags']];
    }

    private function expect(string $type): void
    {
        $token = $this->current();
        if ($token === null || $token['type'] !== $type) {
            throw UnsupportedXlsFeatureException::forFeature('formula syntax; expected ' . $type);
        }
        $this->position++;
    }

    /** @return array{type:string,value:string}|null */
    private function current(): ?array
    {
        return $this->tokens[$this->position] ?? null;
    }

    private const PRECEDENCE = [
        '=' => 10, '<>' => 10, '<' => 10, '<=' => 10, '>' => 10, '>=' => 10,
        '&' => 20, '+' => 30, '-' => 30, '*' => 40, '/' => 40, '^' => 50,
    ];

    private static function operatorToken(string $operator): string
    {
        return match ($operator) {
            '+' => "\x03", '-' => "\x04", '*' => "\x05", '/' => "\x06", '^' => "\x07",
            '&' => "\x08", '<' => "\x09", '<=' => "\x0A", '=' => "\x0B", '>=' => "\x0C",
            '>' => "\x0D", '<>' => "\x0E",
            default => throw UnsupportedXlsFeatureException::forFeature('formula operator ' . $operator),
        };
    }

    private static function errorCode(string $error): ?int
    {
        return match ($error) {
            '#NULL!' => 0x00, '#DIV/0!' => 0x07, '#VALUE!' => 0x0F, '#REF!' => 0x17,
            '#NAME?' => 0x1D, '#NUM!' => 0x24, '#N/A' => 0x2A, default => null,
        };
    }
}
