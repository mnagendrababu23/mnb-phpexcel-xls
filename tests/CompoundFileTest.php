<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Xls\Tests;

use Mnb\PHPExcel\Compound\CompoundFileReader;
use Mnb\PHPExcel\Compound\CompoundFileWriter;
use Mnb\PHPExcel\Exception\InvalidCompoundFileException;
use PHPUnit\Framework\TestCase;

final class CompoundFileTest extends TestCase
{
    public function testLargeStreamUsesDifatAndRoundTrips(): void
    {
        $payload = str_repeat('DIFAT', 2_000_000); // 10 MB
        $path = sys_get_temp_dir() . '/mnb-native-xls-difat-' . getmypid() . '.cfb';
        file_put_contents($path, CompoundFileWriter::build('Workbook', $payload));
        $reader = new CompoundFileReader($path, ['max_stream_size' => 16 * 1024 * 1024]);
        self::assertSame($payload, $reader->readStream('Workbook'));
        @unlink($path);
    }

    public function testInvalidSignatureIsRejected(): void
    {
        $path = sys_get_temp_dir() . '/mnb-native-xls-invalid-' . getmypid() . '.xls';
        file_put_contents($path, str_repeat("\0", 512));
        $this->expectException(InvalidCompoundFileException::class);
        try {
            new CompoundFileReader($path);
        } finally {
            @unlink($path);
        }
    }
}
