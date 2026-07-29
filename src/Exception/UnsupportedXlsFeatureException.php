<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Exception;

use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;

final class UnsupportedXlsFeatureException extends MnbExcelException
{
    /** @param array<string,mixed> $context */
    public static function forFeature(string $feature, array $context = []): self
    {
        return new self(
            'Native XLS feature is not supported yet: ' . $feature,
            ErrorCode::UNSUPPORTED_FORMAT,
            'xls',
            ['feature' => $feature] + $context,
            null,
            'This XLS feature is not supported by the native BIFF8 engine yet.'
        );
    }
}
