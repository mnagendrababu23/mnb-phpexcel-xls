<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Biff\String;

use Mnb\PHPExcel\Exception\InvalidBiffRecordException;

final class SharedStringTable
{
    /** @param list<string> $strings */
    private function __construct(private readonly array $strings)
    {
    }

    /** @param list<string> $segments */
    public static function fromSegments(array $segments, int $maxStrings = 2_000_000): self
    {
        $cursor = new SegmentCursor($segments);
        $total = $cursor->u32();
        $unique = $cursor->u32();
        if ($unique > $maxStrings || $total > max($maxStrings * 100, $maxStrings)) {
            throw InvalidBiffRecordException::because('SST string count exceeds the configured limit.', ['total' => $total, 'unique' => $unique, 'limit' => $maxStrings]);
        }
        $strings = [];
        for ($i = 0; $i < $unique; $i++) {
            $characterCount = $cursor->u16();
            $flags = $cursor->u8();
            $is16Bit = ($flags & 0x01) !== 0;
            $hasExtended = ($flags & 0x04) !== 0;
            $hasRich = ($flags & 0x08) !== 0;
            $runCount = $hasRich ? $cursor->u16() : 0;
            $extendedSize = $hasExtended ? $cursor->u32() : 0;
            $strings[] = $cursor->readCharacters($characterCount, $is16Bit);
            if ($runCount > 0) {
                $cursor->read($runCount * 4);
            }
            if ($extendedSize > 0) {
                $cursor->read($extendedSize);
            }
        }
        return new self($strings);
    }

    public function get(int $index): string
    {
        if (!array_key_exists($index, $this->strings)) {
            throw InvalidBiffRecordException::because('LABELSST references a missing shared string.', ['index' => $index, 'count' => count($this->strings)]);
        }
        return $this->strings[$index];
    }

    /** @return list<string> */
    public function all(): array
    {
        return $this->strings;
    }
}
