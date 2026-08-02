<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Metadata;

use Mnb\PHPExcel\Biff\BiffRecord;
use Mnb\PHPExcel\Biff\BiffRecordReader;
use Mnb\PHPExcel\Biff\RecordType;
use Mnb\PHPExcel\Biff\WorkbookGlobalsReader;
use Mnb\PHPExcel\Compound\CompoundFileReader;
use Mnb\PHPExcel\Compound\CompoundFileRewriter;
use Mnb\PHPExcel\Compound\DirectoryEntry;
use Mnb\PHPExcel\Reader\XlsMetadataReader;
use Mnb\PHPExcel\Support\AtomicFileWriter;
use Mnb\PHPExcel\Support\Binary;
use Mnb\PHPExcel\Support\MnbExcelException;

/** Atomic metadata updater for native OLE2/BIFF8 XLS workbooks. */
final class XlsMetadataWriter implements MetadataWriterInterface
{
    private const FILEPASS = 0x002F;

    /** @var array<string,list<string>> */
    private const ALLOWED_FIELDS = [
        'document' => ['title', 'subject', 'creator', 'author', 'keywords', 'comments', 'description', 'category'],
        'revision' => [
            'last_saved_by', 'last_modified_by', 'revision_number', 'total_editing_time_seconds',
            'last_printed_at', 'document_created_at', 'document_modified_at',
        ],
        'application' => [
            'application', 'application_name', 'application_version', 'manager', 'company',
            'scale_crop', 'links_dirty', 'shared_document', 'hyperlinks_changed',
            'content_type', 'content_status', 'language', 'document_version',
        ],
        'custom_properties' => [],
        'workbook' => ['active_sheet', 'sheet_visibility', 'date1904'],
        'calculation' => [
            'mode', 'iterate', 'iteration_enabled', 'iterate_count', 'maximum_iterations',
            'iterate_delta', 'maximum_change', 'calc_on_save', 'save_recalculation', 'reference_mode',
        ],
    ];

    /** @param array<string,mixed> $changes @param array<string,mixed> $options */
    public function updateMetaInfo(string $source, string $destination, array $changes, array $options = []): void
    {
        $this->assertChanges($changes, (bool) ($options['strict'] ?? true));
        $sourcePath = realpath($source);
        if ($sourcePath === false || !is_file($sourcePath)) {
            throw new MnbExcelException('Invalid XLS source path: ' . $source);
        }
        if (trim($destination) === '') {
            throw new MnbExcelException('XLS metadata destination path cannot be empty.');
        }

        $compound = new CompoundFileReader($sourcePath, $options);
        $entries = $compound->directoryEntries();
        $streams = [];
        $streamIds = [];
        foreach ($entries as $entry) {
            if ($entry->type !== DirectoryEntry::TYPE_STREAM) {
                continue;
            }
            $streams[$entry->id] = $compound->readStreamById($entry->id);
            $streamIds[strtolower($entry->name)] = $entry->id;
        }
        $workbookId = $streamIds['workbook'] ?? $streamIds['book'] ?? null;
        if ($workbookId === null) {
            throw new MnbExcelException('XLS compound file has no Workbook or Book stream.');
        }
        $signatureNames = [];
        foreach ($entries as $entry) {
            if ($entry->type !== DirectoryEntry::TYPE_STREAM) {
                continue;
            }
            $normalizedName = strtolower(str_replace(["\0", "\x05"], '', $entry->name));
            if (str_contains($normalizedName, 'digitalsignature') || str_contains($normalizedName, 'digital signature') || str_contains($normalizedName, 'msodigitalsignature')) {
                $signatureNames[] = $entry->name;
            }
        }
        if ($signatureNames !== [] && !(bool) ($options['allow_invalidate_digital_signatures'] ?? false)) {
            throw new MnbExcelException('The XLS file contains digital-signature streams. Metadata changes could invalidate them; set allow_invalidate_digital_signatures=true to proceed explicitly.');
        }
        if ($this->containsFilePass($streams[$workbookId], $options)) {
            throw new MnbExcelException('Native metadata updates for password-encrypted BIFF workbooks are not supported.');
        }

        if (isset($changes['workbook']) || isset($changes['calculation'])) {
            $streams[$workbookId] = $this->updateWorkbookStream(
                $streams[$workbookId],
                (array) ($changes['workbook'] ?? []),
                (array) ($changes['calculation'] ?? []),
                $options
            );
        }

        $propertyWriter = new OlePropertySetWriter();
        $rootStreams = [];
        $summaryChanges = $this->summaryChanges(
            (array) ($changes['document'] ?? []),
            (array) ($changes['revision'] ?? []),
            (array) ($changes['application'] ?? [])
        );
        if ($summaryChanges !== []) {
            $name = "\x05SummaryInformation";
            $id = $streamIds[strtolower($name)] ?? null;
            $existing = $id === null ? $propertyWriter->newSummary() : $streams[$id];
            $rootStreams[$name] = $propertyWriter->updateSummary($existing, $summaryChanges);
        }

        $documentSummaryChanges = $this->documentSummaryChanges(
            (array) ($changes['document'] ?? []),
            (array) ($changes['application'] ?? [])
        );
        $customWasRequested = array_key_exists('custom_properties', $changes);
        if ($documentSummaryChanges !== [] || $customWasRequested) {
            $name = "\x05DocumentSummaryInformation";
            $id = $streamIds[strtolower($name)] ?? null;
            $existing = $id === null ? $propertyWriter->newDocumentSummary() : $streams[$id];
            $rootStreams[$name] = $propertyWriter->updateDocumentSummary(
                $existing,
                $documentSummaryChanges,
                $customWasRequested ? $changes['custom_properties'] : null,
                (bool) ($options['replace_custom_properties'] ?? false)
            );
        }

        if ($rootStreams !== []) {
            [$entries, $streams] = CompoundFileRewriter::upsertRootStreams($entries, $streams, $rootStreams);
        }
        $output = CompoundFileRewriter::build($entries, $streams);

        AtomicFileWriter::writeViaTemp(
            $destination,
            static function (string $temporary) use ($output): void {
                $written = @file_put_contents($temporary, $output, LOCK_EX);
                if ($written === false || $written !== strlen($output)) {
                    throw new MnbExcelException('Unable to write updated XLS metadata file.');
                }
            },
            (bool) ($options['validate_output'] ?? true)
                ? static function (string $temporary) use ($options): void {
                    $report = (new XlsMetadataReader())->metaInfo($temporary, ['profile' => 'quick'] + $options);
                    if (($report['status'] ?? 'error') === 'error') {
                        throw new MnbExcelException('Updated XLS metadata file failed validation.');
                    }
                }
                : null
        );
    }

    /** @param array<string,mixed> $options */
    public function removePersonalInfo(string $source, string $destination, array $options = []): void
    {
        $document = ['creator' => null];
        if ((bool) ($options['remove_descriptive_properties'] ?? false)) {
            $document += [
                'title' => null,
                'subject' => null,
                'keywords' => null,
                'comments' => null,
                'category' => null,
            ];
        }
        $this->updateMetaInfo($source, $destination, [
            'document' => $document,
            'revision' => ['last_saved_by' => null],
            'application' => ['manager' => null, 'company' => null],
            'custom_properties' => [],
        ], ['replace_custom_properties' => (bool) ($options['remove_custom_properties'] ?? true)] + $options);
    }

    /** @param array<string,mixed> $changes */
    private function assertChanges(array $changes, bool $strict): void
    {
        foreach ($changes as $section => $fields) {
            if (!array_key_exists($section, self::ALLOWED_FIELDS)) {
                if ($strict) {
                    throw new MnbExcelException('Unsupported XLS metadata section: ' . $section);
                }
                continue;
            }
            if ($section === 'custom_properties') {
                if (!is_array($fields)) {
                    throw new MnbExcelException('custom_properties changes must be an array.');
                }
                continue;
            }
            if (!is_array($fields)) {
                throw new MnbExcelException('XLS metadata section must be an array: ' . $section);
            }
            foreach ($fields as $field => $_value) {
                if (!in_array((string) $field, self::ALLOWED_FIELDS[$section], true) && $strict) {
                    throw new MnbExcelException('Unsupported XLS metadata field: ' . $section . '.' . $field);
                }
            }
        }
    }

    /** @param array<string,mixed> $document @param array<string,mixed> $revision @param array<string,mixed> $application @return array<string,mixed> */
    private function summaryChanges(array $document, array $revision, array $application): array
    {
        $out = [];
        foreach (['title', 'subject', 'creator', 'author', 'keywords', 'comments'] as $key) {
            if (array_key_exists($key, $document)) {
                $out[$key] = $document[$key];
            }
        }
        if (array_key_exists('description', $document)) {
            $out['comments'] = $document['description'];
        }
        foreach ([
            'last_saved_by', 'last_modified_by', 'revision_number', 'total_editing_time_seconds',
            'last_printed_at', 'document_created_at', 'document_modified_at',
        ] as $key) {
            if (array_key_exists($key, $revision)) {
                $out[$key] = $revision[$key];
            }
        }
        foreach (['application', 'application_name'] as $key) {
            if (array_key_exists($key, $application)) {
                $out['application_name'] = $application[$key];
            }
        }
        return $out;
    }

    /** @param array<string,mixed> $document @param array<string,mixed> $application @return array<string,mixed> */
    private function documentSummaryChanges(array $document, array $application): array
    {
        $out = [];
        if (array_key_exists('category', $document)) {
            $out['category'] = $document['category'];
        }
        foreach ([
            'application_version', 'manager', 'company', 'scale_crop', 'links_dirty',
            'shared_document', 'hyperlinks_changed', 'content_type', 'content_status',
            'language', 'document_version',
        ] as $key) {
            if (array_key_exists($key, $application)) {
                $out[$key] = $application[$key];
            }
        }
        return $out;
    }

    /** @param array<string,mixed> $workbook @param array<string,mixed> $calculation @param array<string,mixed> $options */
    private function updateWorkbookStream(string $stream, array $workbook, array $calculation, array $options): string
    {
        $info = (new WorkbookGlobalsReader())->read($stream, $options);
        $sheetNames = array_map(static fn(array $sheet): string => (string) $sheet['name'], $info->sheets);
        $states = array_map(static fn(array $sheet): int => (int) $sheet['state'], $info->sheets);
        $records = [];
        $bounds = [];
        foreach ((new BiffRecordReader($stream, $options))->records() as $record) {
            $records[$record->type][] = $record;
            if ($record->type === RecordType::BOUNDSHEET) {
                $bounds[] = $record;
            }
            if ($record->type === RecordType::EOF) {
                break;
            }
        }

        if (array_key_exists('sheet_visibility', $workbook)) {
            if (!is_array($workbook['sheet_visibility'])) {
                throw new MnbExcelException('Workbook sheet_visibility must be an array.');
            }
            foreach ($workbook['sheet_visibility'] as $selector => $stateValue) {
                $index = $this->resolveSheetIndex($sheetNames, is_int($selector) ? $selector : (string) $selector);
                $state = strtolower((string) $stateValue);
                $states[$index - 1] = match ($state) {
                    'visible' => 0,
                    'hidden' => 1,
                    'veryhidden', 'very_hidden' => 2,
                    default => throw new MnbExcelException('Worksheet visibility must be visible, hidden, or veryHidden.'),
                };
            }
            if (!in_array(0, $states, true)) {
                throw new MnbExcelException('At least one worksheet must remain visible.');
            }
            if (count($bounds) !== count($sheetNames)) {
                throw new MnbExcelException('Unable to match XLS BOUNDSHEET records for visibility update.');
            }
            foreach ($bounds as $index => $record) {
                if ($record->length() < 6) {
                    throw new MnbExcelException('Invalid XLS BOUNDSHEET record.');
                }
                $payload = $record->payload;
                $payload[4] = chr($states[$index]);
                $stream = $this->replacePayload($stream, $record, $payload);
            }
        }

        $active = $this->currentActiveIndex($records, count($sheetNames));
        if (array_key_exists('active_sheet', $workbook)) {
            $active = $this->resolveSheetIndex($sheetNames, $workbook['active_sheet']);
            if (($states[$active - 1] ?? 0) !== 0) {
                throw new MnbExcelException('The active worksheet must be visible.');
            }
        } elseif (($states[$active - 1] ?? 0) !== 0) {
            $firstVisible = array_search(0, $states, true);
            $active = is_int($firstVisible) ? $firstVisible + 1 : 1;
        }
        if (array_key_exists('active_sheet', $workbook) || array_key_exists('sheet_visibility', $workbook)) {
            $record = $records[RecordType::WINDOW1][0] ?? null;
            if (!$record instanceof BiffRecord || $record->length() < 14) {
                throw new MnbExcelException('XLS WINDOW1 record is missing; active sheet cannot be updated safely.');
            }
            $payload = substr_replace($record->payload, pack('v', $active - 1), 10, 2);
            $payload = substr_replace($payload, pack('v', $active - 1), 12, 2);
            $stream = $this->replacePayload($stream, $record, $payload);
        }

        if (array_key_exists('date1904', $workbook)) {
            if (!is_bool($workbook['date1904'])) {
                throw new MnbExcelException('Workbook date1904 must be boolean.');
            }
            $record = $records[RecordType::DATEMODE][0] ?? null;
            if (!$record instanceof BiffRecord || $record->length() < 2) {
                throw new MnbExcelException('XLS DATEMODE record is missing.');
            }
            $stream = $this->replacePayload($stream, $record, pack('v', $workbook['date1904'] ? 1 : 0) . substr($record->payload, 2));
        }

        $calcMap = [
            'mode' => RecordType::CALCMODE,
            'iterate_count' => RecordType::CALCCOUNT,
            'maximum_iterations' => RecordType::CALCCOUNT,
            'reference_mode' => RecordType::REFMODE,
            'iterate' => RecordType::ITERATION,
            'iteration_enabled' => RecordType::ITERATION,
            'iterate_delta' => RecordType::DELTA,
            'maximum_change' => RecordType::DELTA,
            'calc_on_save' => RecordType::SAVERECALC,
            'save_recalculation' => RecordType::SAVERECALC,
        ];
        foreach ($calculation as $field => $value) {
            if (!isset($calcMap[$field])) {
                continue;
            }
            $type = $calcMap[$field];
            $record = $records[$type][0] ?? null;
            if (!$record instanceof BiffRecord) {
                throw new MnbExcelException(sprintf('XLS calculation record 0x%04X is missing.', $type));
            }
            $payload = match ($field) {
                'mode' => pack('v', match (strtolower((string) $value)) {
                    'manual' => 0,
                    'automatic', 'auto' => 1,
                    'automatic_except_tables', 'auto_except_tables' => 0xFFFF,
                    default => throw new MnbExcelException('Invalid XLS calculation mode.'),
                }),
                'iterate_count', 'maximum_iterations' => pack('v', $this->unsigned16($value, 'maximum iterations')),
                'reference_mode' => pack('v', match (strtolower((string) $value)) {
                    'a1' => 1,
                    'r1c1' => 0,
                    default => throw new MnbExcelException('XLS reference mode must be a1 or r1c1.'),
                }),
                'iterate', 'iteration_enabled', 'calc_on_save', 'save_recalculation' => pack('v', $this->boolean($value) ? 1 : 0),
                'iterate_delta', 'maximum_change' => pack('e', $this->nonNegativeFloat($value, 'maximum change')),
                default => $record->payload,
            };
            if (strlen($payload) !== $record->length()) {
                $payload = $payload . substr($record->payload, strlen($payload));
            }
            $stream = $this->replacePayload($stream, $record, $payload);
        }
        return $stream;
    }

    /** @param array<int,list<BiffRecord>> $records */
    private function currentActiveIndex(array $records, int $sheetCount): int
    {
        $record = $records[RecordType::WINDOW1][0] ?? null;
        if (!$record instanceof BiffRecord || $record->length() < 12) {
            return 1;
        }
        return min(max(1, Binary::u16($record->payload, 10) + 1), max(1, $sheetCount));
    }

    /** @param list<string> $sheetNames */
    private function resolveSheetIndex(array $sheetNames, mixed $selector): int
    {
        if (is_int($selector) || (is_string($selector) && ctype_digit(trim($selector)))) {
            $index = (int) $selector;
            if ($index < 1 || !isset($sheetNames[$index - 1])) {
                throw new MnbExcelException('Worksheet index is out of range: ' . $index);
            }
            return $index;
        }
        $name = trim((string) $selector);
        foreach ($sheetNames as $index => $sheetName) {
            if ($sheetName === $name) {
                return $index + 1;
            }
        }
        $matches = [];
        foreach ($sheetNames as $index => $sheetName) {
            if (strcasecmp($sheetName, $name) === 0) {
                $matches[] = $index + 1;
            }
        }
        if (count($matches) === 1) {
            return $matches[0];
        }
        throw new MnbExcelException('Worksheet not found: ' . $name);
    }

    private function replacePayload(string $stream, BiffRecord $record, string $payload): string
    {
        if (strlen($payload) !== $record->length()) {
            throw new MnbExcelException('BIFF metadata patch attempted to change a record length.');
        }
        return substr_replace($stream, $payload, $record->offset + 4, $record->length());
    }

    /** @param array<string,mixed> $options */
    private function containsFilePass(string $stream, array $options): bool
    {
        foreach ((new BiffRecordReader($stream, $options))->records() as $record) {
            if ($record->type === self::FILEPASS) {
                return true;
            }
            if ($record->type === RecordType::EOF) {
                break;
            }
        }
        return false;
    }

    private function unsigned16(mixed $value, string $name): int
    {
        if (is_string($value) && preg_match('/^\d+$/', trim($value)) !== 1) {
            throw new MnbExcelException('XLS ' . $name . ' is invalid.');
        }
        $integer = filter_var($value, FILTER_VALIDATE_INT);
        if ($integer === false || $integer < 0 || $integer > 65535) {
            throw new MnbExcelException('XLS ' . $name . ' must be between 0 and 65535.');
        }
        return $integer;
    }

    private function boolean(mixed $value): bool
    {
        if (is_bool($value)) {
            return $value;
        }
        if (is_int($value) && in_array($value, [0, 1], true)) {
            return $value === 1;
        }
        if (is_string($value)) {
            return match (strtolower(trim($value))) {
                'true', '1', 'yes' => true,
                'false', '0', 'no' => false,
                default => throw new MnbExcelException('XLS metadata boolean value is invalid.'),
            };
        }
        throw new MnbExcelException('XLS metadata boolean value is invalid.');
    }

    private function nonNegativeFloat(mixed $value, string $name): float
    {
        if (!is_numeric($value) || !is_finite((float) $value) || (float) $value < 0) {
            throw new MnbExcelException('XLS ' . $name . ' must be a non-negative finite number.');
        }
        return (float) $value;
    }
}
