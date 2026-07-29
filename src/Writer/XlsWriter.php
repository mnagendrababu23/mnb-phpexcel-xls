<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Writer;

use Mnb\PHPExcel\Compound\CompoundFileWriter;
use Mnb\PHPExcel\Core\WorkbookData;
use Mnb\PHPExcel\Support\AtomicFileWriter;
use Mnb\PHPExcel\Support\ErrorCode;

/** Fully independent native BIFF8/XLS writer. */
final class XlsWriter
{
    /** @param array<string,mixed> $options */
    public function write(WorkbookData $workbook, string $path, array $options = []): void
    {
        $workbookStream = (new WorkbookWriter())->write($workbook, $options);
        $compoundFile = CompoundFileWriter::build('Workbook', $workbookStream);
        AtomicFileWriter::writeString($path, $compoundFile, ErrorCode::FILE_WRITE_FAILED);
    }
}
