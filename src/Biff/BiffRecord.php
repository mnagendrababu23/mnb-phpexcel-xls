<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Biff;

final class BiffRecord
{
    public function __construct(
        public readonly int $type,
        public readonly string $payload,
        public readonly int $offset,
    ) {
    }

    public function length(): int
    {
        return strlen($this->payload);
    }
}
