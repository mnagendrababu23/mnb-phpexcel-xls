<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Compound;

final class DirectoryEntry
{
    public const TYPE_EMPTY = 0;
    public const TYPE_STORAGE = 1;
    public const TYPE_STREAM = 2;
    public const TYPE_ROOT = 5;

    public function __construct(
        public readonly int $id,
        public readonly string $name,
        public readonly int $type,
        public readonly int $leftSiblingId,
        public readonly int $rightSiblingId,
        public readonly int $childId,
        public readonly int $startSector,
        public readonly int $streamSize,
        public readonly string $rawBytes = '',
        public readonly int $color = 1,
    ) {
    }
}
