<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader;

use Mnb\PHPExcel\Biff\BiffRecord;
use Mnb\PHPExcel\Biff\BiffRecordReader;
use Mnb\PHPExcel\Biff\RecordType;
use Mnb\PHPExcel\Biff\String\BiffString;
use Mnb\PHPExcel\Biff\WorkbookGlobalsReader;
use Mnb\PHPExcel\Biff\WorkbookInfo;
use Mnb\PHPExcel\Compound\CompoundFileReader;
use Mnb\PHPExcel\Metadata\MetadataCapabilities;
use Mnb\PHPExcel\Metadata\MetadataOptions;
use Mnb\PHPExcel\Metadata\MetadataProfile;
use Mnb\PHPExcel\Metadata\MetadataReport;
use Mnb\PHPExcel\Metadata\MetadataSectionState;
use Mnb\PHPExcel\Metadata\OlePropertySetReader;
use Mnb\PHPExcel\Support\Binary;
use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;

/** Native OLE2/BIFF8 metadata collector for legacy .xls workbooks. */
final class XlsMetadataReader
{
    private const NAME = 0x0018;
    private const EXTERNSHEET = 0x0017;
    private const SUPBOOK = 0x01AE;
    private const FILEPASS = 0x002F;
    private const PROTECT = 0x0012;
    private const PASSWORD = 0x0013;
    private const WINDOWPROTECT = 0x0019;
    private const WRITEPROT = 0x0086;
    private const WRITEACCESS = 0x005C;
    private const CODENAME = 0x01BA;
    private const NOTE = 0x001C;
    private const HLINK = 0x01B8;
    private const OBJ = 0x005D;
    private const MSODRAWING = 0x00EC;
    private const DVAL = 0x01B2;
    private const DV = 0x01BE;
    private const CONDFMT = 0x01B0;
    private const CF = 0x01B1;
    private const SXVIEW = 0x00B0;
    private const SXDB = 0x00C6;
    private const HEADER = 0x0014;
    private const FOOTER = 0x0015;
    private const HORIZONTAL_PAGE_BREAKS = 0x001B;
    private const VERTICAL_PAGE_BREAKS = 0x001A;
    private const LEFT_MARGIN = 0x0026;
    private const RIGHT_MARGIN = 0x0027;
    private const TOP_MARGIN = 0x0028;
    private const BOTTOM_MARGIN = 0x0029;
    private const PRINT_HEADERS = 0x002A;
    private const PRINT_GRIDLINES = 0x002B;
    private const SETUP = 0x00A1;
    private const SHEET_PROTECTION = 0x0867;
    private const LIST12 = 0x0877;

    /** @param array<string,mixed> $options @return array<string,mixed> */
    public function metaInfo(string $path, array $options = []): array
    {
        if (!is_file($path)) {
            throw MnbExcelException::withCode('XLS file not found: ' . $path, ErrorCode::FILE_NOT_FOUND, ['path' => $path]);
        }
        $metadataOptions = MetadataOptions::fromArray($options);
        try {
            $compound = new CompoundFileReader($path, $options);
        } catch (\Throwable $e) {
            throw MnbExcelException::withCode('Unable to open XLS compound file: ' . $e->getMessage(), ErrorCode::FILE_READ_FAILED, ['path' => $path], $e);
        }
        $streamName = $compound->hasStream('Workbook') ? 'Workbook' : ($compound->hasStream('Book') ? 'Book' : null);
        if ($streamName === null) {
            throw MnbExcelException::withCode('XLS compound file has no Workbook or Book stream.', ErrorCode::FILE_READ_FAILED, ['path' => $path]);
        }
        $stream = $compound->readStream($streamName);
        $encrypted = $this->containsRecord($stream, self::FILEPASS, 256);

        $report = new MetadataReport('xls', 'biff8', 'application/vnd.ms-excel', $metadataOptions->profile());
        $stat = @stat($path) ?: [];
        $report->setSection('file', MetadataSectionState::AVAILABLE, [
            'path' => $path,
            'resolved_path' => realpath($path) ?: $path,
            'name' => basename($path),
            'extension' => strtolower((string) pathinfo($path, PATHINFO_EXTENSION)),
            'size_bytes' => isset($stat['size']) ? (int) $stat['size'] : null,
            'filesystem_created_at' => isset($stat['ctime']) ? date(DATE_ATOM, (int) $stat['ctime']) : null,
            'filesystem_modified_at' => isset($stat['mtime']) ? date(DATE_ATOM, (int) $stat['mtime']) : null,
            'filesystem_accessed_at' => isset($stat['atime']) ? date(DATE_ATOM, (int) $stat['atime']) : null,
            'readable' => is_readable($path),
            'writable' => is_writable($path),
            'sha256' => $metadataOptions->includeHash() ? hash_file('sha256', $path) : null,
        ]);

        $streamInfo = $compound->streamInfo();
        $report->setSection('format_details', MetadataSectionState::AVAILABLE, [
            'container' => 'CFB/OLE2',
            'workbook_stream' => $streamName,
            'workbook_stream_size_bytes' => strlen($stream),
            'stream_count' => count($streamInfo),
            'biff_version' => 8,
            'maximum_rows' => 65536,
            'maximum_columns' => 256,
        ]);

        $properties = $this->propertyMetadata($compound);
        $document = array_filter([
            'title' => $properties['summary']['title'] ?? null,
            'subject' => $properties['summary']['subject'] ?? null,
            'creator' => $properties['summary']['creator'] ?? null,
            'keywords' => $properties['summary']['keywords'] ?? null,
            'comments' => $properties['summary']['comments'] ?? null,
            'category' => $properties['document_summary']['category'] ?? ($properties['custom_properties']['OOXMLCorePropertyCategory']['value'] ?? null),
        ], static fn(mixed $v): bool => $v !== null && $v !== '');
        $report->setSection('document', $document === [] ? MetadataSectionState::PARTIAL : MetadataSectionState::AVAILABLE, $document + [
            'warnings' => $properties['warnings'],
        ]);
        $revision = array_filter([
            'last_saved_by' => $properties['summary']['last_saved_by'] ?? null,
            'revision_number' => $properties['summary']['revision_number'] ?? null,
            'total_editing_time' => $properties['summary']['total_editing_time'] ?? null,
            'last_printed_at' => $properties['summary']['last_printed_at'] ?? null,
            'document_created_at' => $properties['summary']['created_at'] ?? null,
            'document_modified_at' => $properties['summary']['modified_at'] ?? null,
        ], static fn(mixed $v): bool => $v !== null && $v !== '');
        $report->setSection('revision', $revision === [] ? MetadataSectionState::PARTIAL : MetadataSectionState::AVAILABLE, $revision);
        $application = array_filter([
            'application_name' => $properties['summary']['application_name'] ?? null,
            'application_version' => $properties['document_summary']['application_version'] ?? ($properties['custom_properties']['AppVersion']['value'] ?? null),
            'manager' => $properties['document_summary']['manager'] ?? null,
            'company' => $properties['document_summary']['company'] ?? null,
        ], static fn(mixed $v): bool => $v !== null && $v !== '');
        $report->setSection('application', $application === [] ? MetadataSectionState::PARTIAL : MetadataSectionState::AVAILABLE, $application);
        $customItems = [];
        foreach ($properties['custom_properties'] as $name => $property) {
            $customItems[] = ['name' => $name] + $property;
        }
        $report->setSection('custom_properties', MetadataSectionState::AVAILABLE, [
            'count' => count($customItems),
            'items' => array_slice($customItems, 0, $metadataOptions->maxItems()),
            'truncated' => count($customItems) > $metadataOptions->maxItems(),
        ]);

        if ($encrypted) {
            $report->status('password_required')
                ->setSection('security', MetadataSectionState::PASSWORD_REQUIRED, [
                    'encrypted' => true,
                    'encryption_type' => 'BIFF FILEPASS',
                    'warnings' => ['Native BIFF decryption is not implemented; workbook-record metadata could not be scanned.'],
                ]);
            foreach (['workbook','macros','named_objects','links','hidden_content','comments_notes','tracked_changes','embedded_objects','calculation','print_settings','validation','pivot_metadata','statistics'] as $section) {
                $report->setSection($section, MetadataSectionState::ENCRYPTED);
            }
            $report->setSection('xml_metadata', MetadataSectionState::NOT_APPLICABLE);
            $array = $report->toArray();
            $report->capabilities(MetadataCapabilities::fromReport($array));
            return $report->toArray();
        }

        try {
            $workbook = (new WorkbookGlobalsReader())->read($stream, $options);
        } catch (\Throwable $e) {
            $report->error('Unable to parse BIFF workbook globals: ' . $e->getMessage());
            foreach (['workbook','security','macros','named_objects','links','hidden_content','comments_notes','tracked_changes','embedded_objects','calculation','print_settings','validation','pivot_metadata','statistics'] as $section) {
                $report->setSection($section, MetadataSectionState::ERROR);
            }
            $report->setSection('xml_metadata', MetadataSectionState::NOT_APPLICABLE);
            $array = $report->toArray();
            $report->capabilities(MetadataCapabilities::fromReport($array));
            return $report->toArray();
        }

        $global = $this->scanGlobals($stream, $workbook, $metadataOptions);
        $sheets = [];
        $sheetScans = [];
        foreach ($workbook->sheets as $index => $sheet) {
            $item = [
                'index' => $index + 1,
                'name' => $sheet['name'],
                'state' => $this->sheetState($sheet['state']),
                'type' => $this->sheetType($sheet['type']),
                'offset' => $sheet['offset'],
            ];
            if ($metadataOptions->atLeast(MetadataProfile::STANDARD) && $sheet['type'] === 0) {
                $scan = $this->scanSheet($stream, $sheet['offset'], $metadataOptions);
                $sheetScans[$index] = $scan;
                $item += [
                    'dimension' => $scan['dimension'],
                    'row_count' => $scan['row_count'],
                    'column_count' => $scan['column_count'],
                    'cell_count' => $scan['cell_count'],
                    'formula_count' => $scan['formula_count'],
                ];
            }
            $sheets[] = $item;
        }
        $activeIndex = min(max(0, $global['active_sheet_zero_based']), max(0, count($sheets) - 1));
        $report->setSection('workbook', MetadataSectionState::AVAILABLE, [
            'name' => pathinfo($path, PATHINFO_FILENAME),
            'sheet_count' => count($sheets),
            'active_sheet_index' => $activeIndex + 1,
            'active_sheet_name' => $sheets[$activeIndex]['name'] ?? null,
            'date_system' => $workbook->date1904 ? 1904 : 1900,
            'code_page' => $workbook->codePage,
            'code_name' => $global['code_name'],
            'sheets' => $sheets,
            'count' => count($sheets),
            'items' => $sheets,
        ]);

        $hiddenItems = array_values(array_filter($sheets, static fn(array $sheet): bool => $sheet['state'] !== 'visible'));
        $hiddenRows = array_sum(array_column($sheetScans, 'hidden_row_count'));
        $hiddenColumns = array_sum(array_column($sheetScans, 'hidden_column_count'));
        $report->setSection('hidden_content', $metadataOptions->profile() === MetadataProfile::QUICK ? MetadataSectionState::PARTIAL : MetadataSectionState::AVAILABLE, [
            'hidden_sheet_count' => count($hiddenItems),
            'hidden_row_count' => $metadataOptions->profile() === MetadataProfile::QUICK ? null : $hiddenRows,
            'hidden_column_count' => $metadataOptions->profile() === MetadataProfile::QUICK ? null : $hiddenColumns,
            'count' => count($hiddenItems),
            'items' => $hiddenItems,
        ]);

        $macroStreams = $this->macroStreams($streamInfo);
        $signatureStreams = $this->matchingStreams($streamInfo, ['digitalsignature', 'digital signature', 'msodigitalsignature']);
        $report->setSection('macros', MetadataSectionState::AVAILABLE, [
            'has_vba_project' => $macroStreams !== [],
            'count' => count($macroStreams),
            'items' => $macroStreams,
        ]);
        $report->setSection('security', MetadataSectionState::AVAILABLE, [
            'encrypted' => false,
            'workbook_protected' => $global['workbook_protected'],
            'window_protected' => $global['window_protected'],
            'write_protected' => $global['write_protected'],
            'password_hash_present' => $global['password_hash_present'],
            'protected_sheet_count' => array_sum(array_column($sheetScans, 'protected') ?: [0]),
            'document_security_flags' => $properties['summary']['security_flags'] ?? null,
            'digital_signature_present' => $signatureStreams !== [],
            'digital_signature_stream_count' => count($signatureStreams),
            'digital_signature_streams' => $signatureStreams,
            'digital_signature_validation' => $signatureStreams === [] ? 'not_applicable' : 'not_validated',
        ]);
        $report->setSection('named_objects', MetadataSectionState::PARTIAL, [
            'defined_name_count' => count($global['defined_names']),
            'table_count' => array_sum(array_column($sheetScans, 'table_count') ?: [0]),
            'count' => count($global['defined_names']) + array_sum(array_column($sheetScans, 'table_count') ?: [0]),
            'items' => array_slice($global['defined_names'], 0, $metadataOptions->maxItems()),
            'truncated' => count($global['defined_names']) > $metadataOptions->maxItems(),
            'warnings' => ['Defined names and table records are inventoried; embedded chart objects are not fully decoded.'],
        ]);
        $report->setSection('links', MetadataSectionState::PARTIAL, [
            'external_book_count' => $global['external_book_count'],
            'external_sheet_reference_count' => $global['external_sheet_reference_count'],
            'hyperlink_count' => array_sum(array_column($sheetScans, 'hyperlink_count') ?: [0]),
            'count' => $global['external_book_count'] + array_sum(array_column($sheetScans, 'hyperlink_count') ?: [0]),
            'warnings' => ['External books and hyperlinks are counted; query-table and legacy data-connection definitions are not fully decoded.'],
        ]);
        $report->setSection('comments_notes', $metadataOptions->profile() === MetadataProfile::QUICK ? MetadataSectionState::NOT_SCANNED : MetadataSectionState::PARTIAL, [
            'comment_count' => array_sum(array_column($sheetScans, 'comment_count') ?: [0]),
            'count' => array_sum(array_column($sheetScans, 'comment_count') ?: [0]),
        ]);
        $tracked = $this->matchingStreams($streamInfo, ['revision', 'history', 'user names']);
        $report->setSection('tracked_changes', MetadataSectionState::PARTIAL, [
            'detected' => $tracked !== [],
            'count' => count($tracked),
            'items' => $tracked,
            'warnings' => ['Legacy shared-workbook history is inventoried by stream/record presence; individual changes are not decoded.'],
        ]);
        $embedded = $this->embeddedStreams($streamInfo);
        $report->setSection('embedded_objects', MetadataSectionState::PARTIAL, [
            'ole_stream_count' => count($embedded),
            'drawing_record_count' => array_sum(array_column($sheetScans, 'drawing_count') ?: [0]),
            'object_record_count' => array_sum(array_column($sheetScans, 'object_count') ?: [0]),
            'count' => count($embedded),
            'items' => array_slice($embedded, 0, $metadataOptions->maxItems()),
            'truncated' => count($embedded) > $metadataOptions->maxItems(),
            'warnings' => ['OLE streams and drawing/object records are inventoried; embedded payload semantics are not executed or fully decoded.'],
        ]);
        $report->setSection('calculation', MetadataSectionState::AVAILABLE, [
            'mode' => $global['calculation']['mode'],
            'iteration_enabled' => $global['calculation']['iteration_enabled'],
            'maximum_iterations' => $global['calculation']['maximum_iterations'],
            'maximum_change' => $global['calculation']['maximum_change'],
            'reference_mode' => $global['calculation']['reference_mode'],
            'save_recalculation' => $global['calculation']['save_recalculation'],
            'formula_count' => array_sum(array_column($sheetScans, 'formula_count') ?: [0]),
        ]);
        $report->setSection('print_settings', $metadataOptions->profile() === MetadataProfile::QUICK ? MetadataSectionState::NOT_SCANNED : MetadataSectionState::PARTIAL, [
            'sheet_count_with_print_settings' => count(array_filter($sheetScans, static fn(array $s): bool => $s['print_record_count'] > 0)),
            'record_count' => array_sum(array_column($sheetScans, 'print_record_count') ?: [0]),
            'count' => array_sum(array_column($sheetScans, 'print_record_count') ?: [0]),
        ]);
        $validationCount = array_sum(array_column($sheetScans, 'validation_count') ?: [0]);
        $conditionalCount = array_sum(array_column($sheetScans, 'conditional_format_count') ?: [0]);
        $report->setSection('validation', $metadataOptions->profile() === MetadataProfile::QUICK ? MetadataSectionState::NOT_SCANNED : MetadataSectionState::PARTIAL, [
            'data_validation_count' => $validationCount,
            'conditional_format_count' => $conditionalCount,
            'count' => $validationCount + $conditionalCount,
        ]);
        $pivotCount = array_sum(array_column($sheetScans, 'pivot_count') ?: [0]) + $global['pivot_cache_count'];
        $report->setSection('pivot_metadata', $metadataOptions->profile() === MetadataProfile::QUICK ? MetadataSectionState::NOT_SCANNED : MetadataSectionState::PARTIAL, [
            'pivot_table_count' => array_sum(array_column($sheetScans, 'pivot_count') ?: [0]),
            'pivot_cache_record_count' => $global['pivot_cache_count'],
            'count' => $pivotCount,
            'warnings' => $pivotCount > 0 ? ['Pivot structures are inventoried; cache fields and source definitions are not fully decoded.'] : [],
        ]);
        $report->setSection('xml_metadata', MetadataSectionState::NOT_APPLICABLE);
        $report->setSection('statistics', $metadataOptions->profile() === MetadataProfile::QUICK ? MetadataSectionState::PARTIAL : MetadataSectionState::AVAILABLE, [
            'worksheet_count' => count(array_filter($sheets, static fn(array $s): bool => $s['type'] === 'worksheet')),
            'row_count' => $metadataOptions->profile() === MetadataProfile::QUICK ? null : array_sum(array_column($sheetScans, 'row_count') ?: [0]),
            'cell_count' => $metadataOptions->profile() === MetadataProfile::QUICK ? null : array_sum(array_column($sheetScans, 'cell_count') ?: [0]),
            'formula_count' => $metadataOptions->profile() === MetadataProfile::QUICK ? null : array_sum(array_column($sheetScans, 'formula_count') ?: [0]),
            'merged_range_count' => $metadataOptions->profile() === MetadataProfile::QUICK ? null : array_sum(array_column($sheetScans, 'merged_range_count') ?: [0]),
            'shared_string_count' => $workbook->sharedStrings === null ? 0 : count($workbook->sharedStrings->all()),
        ]);

        if ($metadataOptions->profile() === MetadataProfile::FORENSIC) {
            $parts = [];
            foreach ($streamInfo as $item) {
                $entry = $item;
                try {
                    $data = $compound->readStream($item['name']);
                    $entry['sha256'] = hash('sha256', $data);
                } catch (\Throwable $e) {
                    $entry['error'] = $e->getMessage();
                }
                $entry['name_display'] = $this->displayStreamName($item['name']);
                $parts[] = $entry;
            }
            $report->setSection('format_details', MetadataSectionState::AVAILABLE, [
                'container' => 'CFB/OLE2', 'workbook_stream' => $streamName,
                'workbook_stream_size_bytes' => strlen($stream), 'stream_count' => count($streamInfo),
                'biff_version' => 8, 'maximum_rows' => 65536, 'maximum_columns' => 256,
                'streams' => array_slice($parts, 0, $metadataOptions->maxItems()),
                'truncated' => count($parts) > $metadataOptions->maxItems(),
                'record_counts' => $global['record_counts'],
            ]);
        }

        $array = $report->toArray();
        $report->capabilities(MetadataCapabilities::fromReport($array));
        return $report->toArray();
    }

    /** @return array<string,mixed> */
    private function propertyMetadata(CompoundFileReader $compound): array
    {
        $result = ['summary' => [], 'document_summary' => [], 'custom_properties' => [], 'warnings' => []];
        $reader = new OlePropertySetReader();
        foreach (["\x05SummaryInformation", "\x05DocumentSummaryInformation"] as $name) {
            if (!$compound->hasStream($name)) {
                continue;
            }
            try {
                $parsed = $reader->read($compound->readStream($name));
                foreach (['summary', 'document_summary', 'custom_properties'] as $key) {
                    $result[$key] = array_replace($result[$key], $parsed[$key]);
                }
                $result['warnings'] = array_merge($result['warnings'], $parsed['warnings']);
            } catch (\Throwable $e) {
                $result['warnings'][] = 'Unable to read ' . $this->displayStreamName($name) . ': ' . $e->getMessage();
            }
        }
        return $result;
    }

    /** @return array<string,mixed> */
    private function scanGlobals(string $stream, WorkbookInfo $workbook, MetadataOptions $options): array
    {
        $active = 0; $codeName = null; $protected = false; $windowProtected = false; $writeProtected = false; $password = false;
        $definedNames = []; $externalBooks = 0; $externalSheets = 0; $supbooks = []; $pivotCache = 0; $recordCounts = [];
        $calc = ['mode' => null, 'maximum_iterations' => null, 'reference_mode' => null, 'iteration_enabled' => null, 'maximum_change' => null, 'save_recalculation' => null];
        foreach ((new BiffRecordReader($stream))->records(0) as $record) {
            $recordCounts[sprintf('0x%04X', $record->type)] = ($recordCounts[sprintf('0x%04X', $record->type)] ?? 0) + 1;
            switch ($record->type) {
                case RecordType::WINDOW1: if ($record->length() >= 12) $active = Binary::u16($record->payload, 10); break;
                case self::CODENAME: $codeName = $this->readRecordString($record); break;
                case self::PROTECT: $protected = $this->recordBool($record); break;
                case self::WINDOWPROTECT: $windowProtected = $this->recordBool($record); break;
                case self::WRITEPROT: $writeProtected = true; break;
                case self::PASSWORD: $password = $record->length() >= 2 && Binary::u16($record->payload, 0) !== 0; break;
                case self::NAME: $definedNames[] = $this->parseDefinedName($record); break;
                case self::SUPBOOK:
                    $isExternal = $this->isExternalSupbook($record);
                    $supbooks[] = $isExternal;
                    if ($isExternal) $externalBooks++;
                    break;
                case self::EXTERNSHEET:
                    if ($record->length() >= 2) {
                        $referenceCount = Binary::u16($record->payload, 0);
                        for ($reference = 0; $reference < $referenceCount; $reference++) {
                            $entryOffset = 2 + ($reference * 6);
                            if ($entryOffset + 6 > $record->length()) break;
                            $supbookIndex = Binary::u16($record->payload, $entryOffset);
                            if (($supbooks[$supbookIndex] ?? false) === true) $externalSheets++;
                        }
                    }
                    break;
                case self::SXDB: $pivotCache++; break;
                case RecordType::CALCMODE: if ($record->length() >= 2) $calc['mode'] = match (Binary::u16($record->payload, 0)) {0 => 'manual', 1 => 'automatic', -1, 0xFFFF => 'automatic_except_tables', default => 'unknown'}; break;
                case RecordType::CALCCOUNT: if ($record->length() >= 2) $calc['maximum_iterations'] = Binary::u16($record->payload, 0); break;
                case RecordType::REFMODE: if ($record->length() >= 2) $calc['reference_mode'] = Binary::u16($record->payload, 0) === 1 ? 'A1' : 'R1C1'; break;
                case RecordType::ITERATION: if ($record->length() >= 2) $calc['iteration_enabled'] = Binary::u16($record->payload, 0) !== 0; break;
                case RecordType::DELTA: if ($record->length() >= 8) $calc['maximum_change'] = Binary::double($record->payload, 0); break;
                case RecordType::SAVERECALC: if ($record->length() >= 2) $calc['save_recalculation'] = Binary::u16($record->payload, 0) !== 0; break;
            }
            if ($record->type === RecordType::EOF) break;
        }
        return [
            'active_sheet_zero_based' => $active, 'code_name' => $codeName,
            'workbook_protected' => $protected, 'window_protected' => $windowProtected,
            'write_protected' => $writeProtected, 'password_hash_present' => $password,
            'defined_names' => array_values(array_filter($definedNames)),
            'external_book_count' => $externalBooks, 'external_sheet_reference_count' => $externalSheets,
            'pivot_cache_count' => $pivotCache, 'calculation' => $calc, 'record_counts' => $recordCounts,
        ];
    }

    /** @return array<string,mixed> */
    private function scanSheet(string $stream, int $offset, MetadataOptions $options): array
    {
        $dimension = null; $rows = 0; $columns = 0; $cells = 0; $formulas = 0; $hiddenRows = 0; $hiddenColumns = 0;
        $comments = 0; $hyperlinks = 0; $objects = 0; $drawings = 0; $validations = 0; $conditional = 0;
        $pivots = 0; $merged = 0; $print = 0; $protected = false; $tables = 0;
        foreach ((new BiffRecordReader($stream))->records($offset) as $record) {
            switch ($record->type) {
                case RecordType::DIMENSIONS:
                    if ($record->length() >= 14) {
                        $firstRow = Binary::u32($record->payload, 0); $lastRow = Binary::u32($record->payload, 4);
                        $firstCol = Binary::u16($record->payload, 8); $lastCol = Binary::u16($record->payload, 10);
                        $rows = max(0, $lastRow - $firstRow); $columns = max(0, $lastCol - $firstCol);
                        $dimension = ['first_row' => $firstRow + 1, 'last_row' => $lastRow, 'first_column' => $firstCol + 1, 'last_column' => $lastCol];
                    } break;
                case RecordType::ROW: if ($record->length() >= 14 && (Binary::u16($record->payload, 12) & 0x0020) !== 0) $hiddenRows++; break;
                case RecordType::COLINFO: if ($record->length() >= 10 && (Binary::u16($record->payload, 8) & 0x0001) !== 0) $hiddenColumns += Binary::u16($record->payload, 2) - Binary::u16($record->payload, 0) + 1; break;
                case RecordType::NUMBER: case RecordType::RK: case RecordType::LABEL: case RecordType::LABELSST: case RecordType::BOOLERR: case RecordType::BLANK: $cells++; break;
                case RecordType::MULRK: if ($record->length() >= 6) $cells += max(0, Binary::u16($record->payload, $record->length() - 2) - Binary::u16($record->payload, 2) + 1); break;
                case RecordType::MULBLANK: if ($record->length() >= 6) $cells += max(0, Binary::u16($record->payload, $record->length() - 2) - Binary::u16($record->payload, 2) + 1); break;
                case RecordType::FORMULA: $cells++; $formulas++; break;
                case self::NOTE: $comments++; break;
                case self::HLINK: $hyperlinks++; break;
                case self::OBJ: $objects++; break;
                case self::MSODRAWING: $drawings++; break;
                case self::DV: $validations++; break;
                case self::CONDFMT: $conditional++; break;
                case self::SXVIEW: $pivots++; break;
                case self::LIST12: $tables++; break;
                case RecordType::MERGEDCELLS: if ($record->length() >= 2) $merged += Binary::u16($record->payload, 0); break;
                case self::PROTECT: $protected = $this->recordBool($record); break;
                case self::HEADER: case self::FOOTER: case self::HORIZONTAL_PAGE_BREAKS: case self::VERTICAL_PAGE_BREAKS:
                case self::LEFT_MARGIN: case self::RIGHT_MARGIN: case self::TOP_MARGIN: case self::BOTTOM_MARGIN:
                case self::PRINT_HEADERS: case self::PRINT_GRIDLINES: case self::SETUP: $print++; break;
            }
            if ($record->type === RecordType::EOF) break;
        }
        return [
            'dimension' => $dimension, 'row_count' => $rows, 'column_count' => $columns,
            'cell_count' => $cells, 'formula_count' => $formulas, 'hidden_row_count' => $hiddenRows,
            'hidden_column_count' => $hiddenColumns, 'comment_count' => $comments, 'hyperlink_count' => $hyperlinks,
            'object_count' => $objects, 'drawing_count' => $drawings, 'validation_count' => $validations,
            'conditional_format_count' => $conditional, 'pivot_count' => $pivots, 'merged_range_count' => $merged,
            'print_record_count' => $print, 'protected' => $protected ? 1 : 0, 'table_count' => $tables,
        ];
    }

    /** @return array<string,mixed>|null */
    private function parseDefinedName(BiffRecord $record): ?array
    {
        if ($record->length() < 15) return null;
        $flags = Binary::u16($record->payload, 0); $length = Binary::u8($record->payload, 3); $formulaLength = Binary::u16($record->payload, 4);
        $sheetIndex = Binary::u16($record->payload, 8); $cursor = 14; if ($cursor >= $record->length()) return null;
        $option = Binary::u8($record->payload, $cursor++); $unicode = ($option & 0x01) !== 0; $byteLength = $length * ($unicode ? 2 : 1);
        if ($cursor + $byteLength > $record->length()) return null;
        $raw = substr($record->payload, $cursor, $byteLength);
        $builtIn = ($flags & 0x0020) !== 0;
        $name = $builtIn && $length === 1 ? $this->builtInName(ord($raw[0])) : BiffString::decodeCharacters($raw, $unicode);
        $cursor += $byteLength;
        return [
            'name' => $name, 'built_in' => $builtIn, 'hidden' => ($flags & 0x0001) !== 0,
            'function' => ($flags & 0x0002) !== 0, 'sheet_index' => $sheetIndex,
            'formula_token_bytes' => $formulaLength,
            'formula_token_hex' => ($formulaLength > 0 && $cursor + $formulaLength <= $record->length()) ? strtoupper(bin2hex(substr($record->payload, $cursor, $formulaLength))) : null,
        ];
    }

    private function isExternalSupbook(BiffRecord $record): bool
    {
        if ($record->length() < 4) {
            return false;
        }
        $marker = substr($record->payload, 2, 2);
        // 0x0401 is the current workbook; 0x3A01 is the add-in function table.
        return $marker !== "\x01\x04" && $marker !== "\x01\x3A";
    }

    private function containsRecord(string $stream, int $type, int $maxRecords): bool
    {
        $count = 0;
        try {
            foreach ((new BiffRecordReader($stream, ['max_biff_records' => max(1, $maxRecords)]))->records(0) as $record) {
                if ($record->type === $type) return true;
                if (++$count >= $maxRecords || $record->type === RecordType::EOF) break;
            }
        } catch (\Throwable) {
        }
        return false;
    }

    private function readRecordString(BiffRecord $record): ?string
    {
        try { return BiffString::readUnicodeString($record->payload, 0, false)['value']; } catch (\Throwable) { return null; }
    }
    private function recordBool(BiffRecord $record): bool { return $record->length() >= 2 && Binary::u16($record->payload, 0) !== 0; }
    private function sheetState(int $state): string { return match ($state) {1 => 'hidden', 2 => 'very_hidden', default => 'visible'}; }
    private function sheetType(int $type): string { return match ($type) {0 => 'worksheet', 1 => 'macro_sheet', 2 => 'chart', 6 => 'vb_module', default => 'unknown_' . $type}; }
    private function builtInName(int $id): string { return match ($id) {0x06 => '_xlnm.Print_Area', 0x07 => '_xlnm.Print_Titles', 0x0D => '_xlnm.FilterDatabase', 0x00 => 'Consolidate_Area', 0x01 => 'Auto_Open', 0x02 => 'Auto_Close', default => '_xlnm.BuiltIn_' . sprintf('%02X', $id)}; }

    /** @param list<array{name:string,size_bytes:int,type:int}> $streams @return list<array<string,mixed>> */
    private function macroStreams(array $streams): array
    {
        $items = [];
        foreach ($streams as $stream) {
            $lower = strtolower($stream['name']);
            if (str_contains($lower, 'vba') || in_array($lower, ['project', 'projectwm', 'dir', '_vba_project_cur'], true)) {
                $items[] = ['name' => $this->displayStreamName($stream['name']), 'size_bytes' => $stream['size_bytes']];
            }
        }
        return $items;
    }
    /** @param list<array{name:string,size_bytes:int,type:int}> $streams @param list<string> $needles @return list<array<string,mixed>> */
    private function matchingStreams(array $streams, array $needles): array
    {
        $items = [];
        foreach ($streams as $stream) {
            $lower = strtolower($stream['name']);
            foreach ($needles as $needle) {
                if (str_contains($lower, $needle)) { $items[] = ['name' => $this->displayStreamName($stream['name']), 'size_bytes' => $stream['size_bytes']]; break; }
            }
        }
        return $items;
    }
    /** @param list<array{name:string,size_bytes:int,type:int}> $streams @return list<array<string,mixed>> */
    private function embeddedStreams(array $streams): array
    {
        $items = [];
        foreach ($streams as $stream) {
            $lower = strtolower($stream['name']);
            if (str_contains($lower, 'ole10native') || str_contains($lower, 'package') || str_contains($lower, 'objectpool')) {
                $items[] = ['name' => $this->displayStreamName($stream['name']), 'size_bytes' => $stream['size_bytes']];
            }
        }
        return $items;
    }
    private function displayStreamName(string $name): string { return preg_replace_callback('/[\x00-\x1F]/', static fn(array $m): string => sprintf('\\x%02X', ord($m[0])), $name) ?? $name; }
}
