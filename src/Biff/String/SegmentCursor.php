<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Biff\String;

use Mnb\PHPExcel\Exception\InvalidBiffRecordException;
use Mnb\PHPExcel\Support\Binary;

/** Cursor over an SST record followed by CONTINUE record payloads. */
final class SegmentCursor
{
    private int $segmentIndex = 0;
    private int $offset = 0;

    /** @param list<string> $segments */
    public function __construct(private readonly array $segments)
    {
        if ($segments === []) {
            throw new \InvalidArgumentException('Segment cursor requires at least one segment.');
        }
    }

    public function u8(): int
    {
        return ord($this->read(1));
    }

    public function u16(): int
    {
        return Binary::u16($this->read(2), 0);
    }

    public function u32(): int
    {
        return Binary::u32($this->read(4), 0);
    }

    public function read(int $length): string
    {
        $result = '';
        while (strlen($result) < $length) {
            if (!isset($this->segments[$this->segmentIndex])) {
                throw InvalidBiffRecordException::because('SST continuation stream ended unexpectedly.', ['required' => $length, 'read' => strlen($result)]);
            }
            $segment = $this->segments[$this->segmentIndex];
            $available = strlen($segment) - $this->offset;
            if ($available <= 0) {
                $this->segmentIndex++;
                $this->offset = 0;
                continue;
            }
            $take = min($length - strlen($result), $available);
            $result .= substr($segment, $this->offset, $take);
            $this->offset += $take;
        }
        return $result;
    }

    public function readCharacters(int $characterCount, bool $is16Bit): string
    {
        $result = '';
        $remaining = $characterCount;
        while ($remaining > 0) {
            if (!isset($this->segments[$this->segmentIndex])) {
                throw InvalidBiffRecordException::because('SST string data ended unexpectedly.', ['remaining_characters' => $remaining]);
            }
            $segment = $this->segments[$this->segmentIndex];
            $bytesPerCharacter = $is16Bit ? 2 : 1;
            $availableBytes = strlen($segment) - $this->offset;
            $availableCharacters = intdiv($availableBytes, $bytesPerCharacter);
            if ($availableCharacters === 0) {
                $this->segmentIndex++;
                $this->offset = 0;
                if (!isset($this->segments[$this->segmentIndex]) || $this->segments[$this->segmentIndex] === '') {
                    throw InvalidBiffRecordException::because('SST CONTINUE record is missing its string option byte.');
                }
                $continuationFlags = ord($this->segments[$this->segmentIndex][0]);
                $this->offset = 1;
                $is16Bit = ($continuationFlags & 0x01) !== 0;
                continue;
            }
            $takeCharacters = min($remaining, $availableCharacters);
            $takeBytes = $takeCharacters * $bytesPerCharacter;
            $raw = substr($segment, $this->offset, $takeBytes);
            $result .= BiffString::decodeCharacters($raw, $is16Bit);
            $this->offset += $takeBytes;
            $remaining -= $takeCharacters;
        }
        return $result;
    }
}
