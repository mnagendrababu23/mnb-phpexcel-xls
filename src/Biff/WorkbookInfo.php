<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Biff;

use Mnb\PHPExcel\Biff\String\SharedStringTable;

final class WorkbookInfo
{
    /**
     * @param list<array{name:string,offset:int,state:int,type:int}> $sheets
     * @param array<int,int> $xfNumberFormats
     * @param array<int,string> $customNumberFormats
     */
    public function __construct(
        public readonly array $sheets,
        public readonly ?SharedStringTable $sharedStrings,
        public readonly array $xfNumberFormats,
        public readonly array $customNumberFormats,
        public readonly bool $date1904,
        public readonly int $codePage,
    ) {
    }

    /** @return list<string> */
    public function sheetNames(bool $worksheetsOnly = true): array
    {
        $names = [];
        foreach ($this->sheets as $sheet) {
            if (!$worksheetsOnly || $sheet['type'] === 0) {
                $names[] = $sheet['name'];
            }
        }
        return $names;
    }

    public function isDateStyle(int $xfIndex): bool
    {
        $formatId = $this->xfNumberFormats[$xfIndex] ?? 0;
        if (in_array($formatId, [14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47], true)) {
            return true;
        }
        $format = $this->customNumberFormats[$formatId] ?? '';
        if ($format === '') {
            return false;
        }
        // Remove quoted literals, escaped characters, colors, conditions, and
        // bracketed elapsed-time tokens before looking for date/time symbols.
        $normalized = preg_replace('/"(?:[^"]|"")*"|\\.|\[(?:[^\]]+)\]/', '', strtolower($format)) ?? '';
        return preg_match('/(^|[^a-z])[dmyhs]+([^a-z]|$)/', $normalized) === 1;
    }
}
