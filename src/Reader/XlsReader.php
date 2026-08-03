<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader;

use Mnb\PHPExcel\Biff\WorkbookGlobalsReader;
use Mnb\PHPExcel\Biff\WorkbookInfo;
use Mnb\PHPExcel\Compound\CompoundFileReader;
use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;
use Mnb\PHPExcel\Snapshot\VisualSnapshotReaderInterface;

/** Fully independent native BIFF8/XLS reader. */
final class XlsReader implements XlsReaderInterface, MetadataReaderInterface, VisualSnapshotReaderInterface
{
    public function format(): string
    {
        return 'xls';
    }

    /** @param array<string,mixed> $options @return array<string,mixed> */
    public function metaInfo(string $path, array $options = []): array
    {
        return (new XlsMetadataReader())->metaInfo($path, $options);
    }

    /** @param array<string,mixed> $options @return array<string,mixed> */
    public function visualSnapshot(string $path, array $options = []): array
    {
        return (new XlsVisualSnapshotReader())->visualSnapshot($path, $options);
    }

    /** @return list<list<mixed>> */
    public function readSheet(string $path, int|string $sheet = 1, array $options = []): array
    {
        return array_values(iterator_to_array($this->iterateSheet($path, $sheet, $options), true));
    }

    /** @return \Generator<int,list<mixed>> */
    public function iterateSheet(string $path, int|string $sheet = 1, array $options = []): iterable
    {
        [$stream, $workbook] = $this->openWorkbook($path, $options);
        $sheetInfo = $this->resolveSheet($workbook, $sheet);
        yield from (new WorksheetReader($stream, $workbook, $options))->rows($sheetInfo['offset']);
    }

    /** @return list<string> */
    public function sheetNames(string $path, array $options = []): array
    {
        [, $workbook] = $this->openWorkbook($path, $options);
        return $workbook->sheetNames();
    }

    /** @param array<string,mixed> $options @return array{0:string,1:WorkbookInfo} */
    private function openWorkbook(string $path, array $options): array
    {
        if (!is_file($path)) {
            throw MnbExcelException::withCode('XLS file not found: ' . $path, ErrorCode::FILE_NOT_FOUND, ['path' => $path]);
        }
        try {
            $compound = new CompoundFileReader($path, $options);
            $streamName = $compound->hasStream('Workbook') ? 'Workbook' : ($compound->hasStream('Book') ? 'Book' : null);
            if ($streamName === null) {
                throw MnbExcelException::withCode('XLS compound file has no Workbook or Book stream.', ErrorCode::FILE_READ_FAILED, ['path' => $path]);
            }
            $stream = $compound->readStream($streamName);
            $workbook = (new WorkbookGlobalsReader())->read($stream, $options);
            return [$stream, $workbook];
        } catch (MnbExcelException $e) {
            throw $e;
        } catch (\Throwable $e) {
            throw MnbExcelException::withCode('Unable to read native XLS file: ' . $e->getMessage(), ErrorCode::FILE_READ_FAILED, ['path' => $path], $e);
        }
    }

    /** @return array{name:string,offset:int,state:int,type:int} */
    private function resolveSheet(WorkbookInfo $workbook, int|string $sheet): array
    {
        $worksheets = array_values(array_filter($workbook->sheets, static fn (array $item): bool => $item['type'] === 0));
        if (is_string($sheet) && !ctype_digit($sheet)) {
            foreach ($worksheets as $item) {
                if ($item['name'] === $sheet) {
                    return $item;
                }
            }
            throw new MnbExcelException('XLS sheet does not exist: ' . $sheet);
        }
        $index = max(1, (int) $sheet) - 1;
        if (!isset($worksheets[$index])) {
            throw new MnbExcelException('XLS sheet index does not exist: ' . (string) $sheet);
        }
        return $worksheets[$index];
    }
}
