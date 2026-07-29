<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Exception;

use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;

final class InvalidCompoundFileException extends MnbExcelException
{
    /** @param array<string,mixed> $context */
    public static function because(string $message, array $context = []): self
    {
        return new self($message, ErrorCode::FILE_READ_FAILED, 'xls', $context, null, 'The XLS compound file is invalid or corrupted.');
    }
}
