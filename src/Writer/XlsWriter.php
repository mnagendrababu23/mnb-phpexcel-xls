<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Writer;

use Mnb\PHPExcel\Compound\CompoundFileWriter;
use Mnb\PHPExcel\Core\WorkbookData;
use Mnb\PHPExcel\Metadata\OlePropertySetWriter;
use Mnb\PHPExcel\Support\AtomicFileWriter;
use Mnb\PHPExcel\Support\ErrorCode;

/** Fully independent native BIFF8/XLS writer. */
final class XlsWriter
{
    /** @param array<string,mixed> $options */
    public function write(WorkbookData $workbook, string $path, array $options = []): void
    {
        $workbookStream = (new WorkbookWriter())->write($workbook, $options);
        $propertyWriter = new OlePropertySetWriter();
        $metadata = $workbook->metadata;
        $now = gmdate('Y-m-d\TH:i:s\Z');

        $summary = [
            'creator' => (string) ($metadata['creator'] ?? 'MNB PHPExcel'),
            'last_saved_by' => (string) ($metadata['last_modified_by'] ?? $metadata['modified_by'] ?? $metadata['creator'] ?? 'MNB PHPExcel'),
            'application_name' => (string) ($metadata['application'] ?? 'MNB PHPExcel'),
            'created_at' => $metadata['created_at'] ?? $metadata['document_created_at'] ?? $now,
            'modified_at' => $metadata['modified_at'] ?? $metadata['document_modified_at'] ?? $now,
        ];
        foreach ([
            'title' => 'title',
            'subject' => 'subject',
            'keywords' => 'keywords',
            'description' => 'comments',
            'comments' => 'comments',
            'revision_number' => 'revision_number',
            'total_editing_time_seconds' => 'total_editing_time_seconds',
            'last_printed_at' => 'last_printed_at',
        ] as $source => $target) {
            if (array_key_exists($source, $metadata) && $metadata[$source] !== null && $metadata[$source] !== '') {
                $summary[$target] = $metadata[$source];
            }
        }

        $documentSummary = [];
        foreach ([
            'category' => 'category',
            'manager' => 'manager',
            'company' => 'company',
            'application_version' => 'application_version',
            'content_type' => 'content_type',
            'content_status' => 'content_status',
            'language' => 'language',
            'document_version' => 'document_version',
        ] as $source => $target) {
            if (array_key_exists($source, $metadata) && $metadata[$source] !== null && $metadata[$source] !== '') {
                $documentSummary[$target] = $metadata[$source];
            }
        }

        $compoundFile = CompoundFileWriter::buildStreams([
            'Workbook' => $workbookStream,
            "\x05SummaryInformation" => $propertyWriter->newSummary($summary),
            "\x05DocumentSummaryInformation" => $propertyWriter->newDocumentSummary(
                $documentSummary,
                $metadata['custom_properties'] ?? []
            ),
        ]);
        AtomicFileWriter::writeString($path, $compoundFile, ErrorCode::FILE_WRITE_FAILED);
    }
}
