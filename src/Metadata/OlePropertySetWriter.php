<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Metadata;

use DateTimeImmutable;
use DateTimeInterface;
use DateTimeZone;
use Mnb\PHPExcel\Support\Binary;
use Mnb\PHPExcel\Support\MnbExcelException;

/** Lossless OLE Property Set updater for SummaryInformation and DocumentSummaryInformation. */
final class OlePropertySetWriter
{
    private const FMTID_SUMMARY = 'F29F85E0-4FF9-1068-AB91-08002B27B3D9';
    private const FMTID_DOCUMENT_SUMMARY = 'D5CDD502-2E9C-101B-9397-08002B2CF9AE';
    private const FMTID_USER_DEFINED = 'D5CDD505-2E9C-101B-9397-08002B2CF9AE';

    /** @param array<string,mixed> $changes */
    public function updateSummary(string $stream, array $changes): string
    {
        $parsed = $this->parse($stream);
        $index = $this->sectionIndex($parsed['sections'], self::FMTID_SUMMARY);
        if ($index === null) {
            $parsed['sections'][] = $this->emptySection(self::FMTID_SUMMARY);
            $index = count($parsed['sections']) - 1;
        }
        $map = [
            'title' => 2, 'subject' => 3, 'creator' => 4, 'author' => 4, 'keywords' => 5,
            'comments' => 6, 'description' => 6, 'template' => 7, 'last_saved_by' => 8,
            'last_modified_by' => 8, 'revision_number' => 9, 'total_editing_time' => 10,
            'total_editing_time_seconds' => 10, 'last_printed_at' => 11, 'created_at' => 12,
            'document_created_at' => 12, 'modified_at' => 13, 'document_modified_at' => 13,
            'page_count' => 14, 'word_count' => 15, 'character_count' => 16,
            'application_name' => 18, 'security_flags' => 19,
        ];
        foreach ($changes as $key => $value) {
            $normalized = strtolower((string) $key);
            if (!isset($map[$normalized])) {
                throw new MnbExcelException('Unsupported XLS SummaryInformation property: ' . $key);
            }
            $id = $map[$normalized];
            if ($value === null) {
                unset($parsed['sections'][$index]['properties'][$id]);
                continue;
            }
            $parsed['sections'][$index]['properties'][$id] = match ($id) {
                10 => $this->fileTimeDuration($value),
                11, 12, 13 => $this->fileTime($value),
                14, 15, 16, 19 => $this->integer($value),
                default => $this->string((string) $value),
            };
        }
        return $this->build($parsed);
    }

    /** @param array<string,mixed> $documentChanges @param mixed $customChanges */
    public function updateDocumentSummary(string $stream, array $documentChanges, mixed $customChanges = null, bool $replaceCustom = false): string
    {
        $parsed = $this->parse($stream);
        $docIndex = $this->sectionIndex($parsed['sections'], self::FMTID_DOCUMENT_SUMMARY);
        if ($docIndex === null) {
            $parsed['sections'][] = $this->emptySection(self::FMTID_DOCUMENT_SUMMARY);
            $docIndex = count($parsed['sections']) - 1;
        }
        $map = [
            'category' => 2, 'presentation_format' => 3, 'byte_count' => 4, 'line_count' => 5,
            'paragraph_count' => 6, 'slide_count' => 7, 'note_count' => 8, 'hidden_slide_count' => 9,
            'multimedia_clip_count' => 10, 'scale_crop' => 11, 'manager' => 14, 'company' => 15,
            'links_dirty' => 16, 'character_count_with_spaces' => 17, 'shared_document' => 19,
            'hyperlinks_changed' => 22, 'application_version' => 23, 'content_type' => 26,
            'content_status' => 27, 'language' => 28, 'document_version' => 29,
        ];
        foreach ($documentChanges as $key => $value) {
            $normalized = strtolower((string) $key);
            if (!isset($map[$normalized])) {
                throw new MnbExcelException('Unsupported XLS DocumentSummaryInformation property: ' . $key);
            }
            $id = $map[$normalized];
            if ($value === null) {
                unset($parsed['sections'][$docIndex]['properties'][$id]);
                continue;
            }
            $parsed['sections'][$docIndex]['properties'][$id] = match ($id) {
                4, 5, 6, 7, 8, 9, 10, 17 => $this->integer($value),
                23 => $this->applicationVersion($value),
                11, 16, 19, 22 => $this->boolean($value),
                default => $this->string((string) $value),
            };
        }

        if ($customChanges !== null) {
            $customIndex = $this->sectionIndex($parsed['sections'], self::FMTID_USER_DEFINED);
            if ($customIndex === null) {
                $parsed['sections'][] = $this->emptySection(self::FMTID_USER_DEFINED, 1200);
                $customIndex = count($parsed['sections']) - 1;
            }
            $section =& $parsed['sections'][$customIndex];
            if ($replaceCustom) {
                $section['properties'] = [1 => $this->codepage((int) $section['codepage'])];
                $section['dictionary'] = [];
            }
            $nameToId = [];
            foreach ($section['dictionary'] as $id => $name) {
                $nameToId[strtolower($name)] = $id;
            }
            $nextId = max([1, ...array_keys($section['properties']), ...array_keys($section['dictionary'])]) + 1;
            foreach ($this->normalizeCustomChanges($customChanges) as $change) {
                $name = trim((string) ($change['name'] ?? ''));
                if ($name === '') {
                    throw new MnbExcelException('Custom property name cannot be empty.');
                }
                $lookup = strtolower($name);
                $id = $nameToId[$lookup] ?? $nextId++;
                if (!array_key_exists('value', $change) || $change['value'] === null) {
                    unset($section['properties'][$id], $section['dictionary'][$id], $nameToId[$lookup]);
                    continue;
                }
                $type = strtolower((string) ($change['type'] ?? $this->inferType($change['value'])));
                $section['properties'][$id] = match ($type) {
                    'string' => $this->string((string) $change['value']),
                    'integer', 'int' => $this->integer($change['value']),
                    'float', 'double', 'number' => $this->float($change['value']),
                    'boolean', 'bool' => $this->boolean($change['value']),
                    'datetime', 'date' => $this->fileTime($change['value']),
                    default => throw new MnbExcelException('Unsupported custom property type: ' . $type),
                };
                $section['dictionary'][$id] = $name;
                $nameToId[$lookup] = $id;
            }
            $section['properties'][1] = $this->codepage((int) $section['codepage']);
            ksort($section['properties']);
            ksort($section['dictionary']);
        }
        return $this->build($parsed);
    }

    public function newSummary(array $changes = []): string
    {
        $stream = $this->build([
            'header' => $this->defaultHeader(),
            'sections' => [$this->emptySection(self::FMTID_SUMMARY)],
        ]);
        return $changes === [] ? $stream : $this->updateSummary($stream, $changes);
    }

    public function newDocumentSummary(array $changes = [], mixed $customChanges = null): string
    {
        $stream = $this->build([
            'header' => $this->defaultHeader(),
            'sections' => [
                $this->emptySection(self::FMTID_DOCUMENT_SUMMARY),
                $this->emptySection(self::FMTID_USER_DEFINED, 1200),
            ],
        ]);
        return $this->updateDocumentSummary($stream, $changes, $customChanges);
    }

    /** @return array{header:string,sections:list<array{fmtid:string,fmtid_bytes:string,codepage:int,properties:array<int,string>,dictionary:array<int,string>}>} */
    private function parse(string $stream): array
    {
        if (strlen($stream) < 28 || Binary::u16($stream, 0) !== 0xFFFE) {
            throw new MnbExcelException('Invalid OLE property-set stream.');
        }
        $count = Binary::u32($stream, 24);
        if ($count < 1 || $count > 32 || 28 + ($count * 20) > strlen($stream)) {
            throw new MnbExcelException('Invalid OLE property-set section table.');
        }
        $decoded = (new OlePropertySetReader())->read($stream)['sections'];
        $decodedByFmtid = [];
        foreach ($decoded as $section) {
            $decodedByFmtid[$section['fmtid']] = $section;
        }
        $sections = [];
        for ($i = 0; $i < $count; $i++) {
            $descriptor = 28 + ($i * 20);
            $fmtidBytes = substr($stream, $descriptor, 16);
            $fmtid = $this->guid($fmtidBytes);
            $offset = Binary::u32($stream, $descriptor + 16);
            $size = Binary::u32($stream, $offset);
            $propertyCount = Binary::u32($stream, $offset + 4);
            if ($size < 8 || $offset + $size > strlen($stream) || $offset + 8 + ($propertyCount * 8) > $offset + $size) {
                throw new MnbExcelException('Invalid OLE property section.');
            }
            $entries = [];
            for ($p = 0; $p < $propertyCount; $p++) {
                $id = Binary::u32($stream, $offset + 8 + ($p * 8));
                $relative = Binary::u32($stream, $offset + 12 + ($p * 8));
                $entries[] = ['id' => $id, 'relative' => $relative];
            }
            usort($entries, static fn(array $a, array $b): int => $a['relative'] <=> $b['relative']);
            $properties = [];
            foreach ($entries as $position => $entry) {
                $start = $offset + $entry['relative'];
                $end = isset($entries[$position + 1]) ? $offset + $entries[$position + 1]['relative'] : $offset + $size;
                $properties[$entry['id']] = substr($stream, $start, $end - $start);
            }
            $decodedSection = $decodedByFmtid[$fmtid] ?? [];
            $sections[] = [
                'fmtid' => $fmtid,
                'fmtid_bytes' => $fmtidBytes,
                'codepage' => (int) ($decodedSection['codepage'] ?? 1252),
                'properties' => $properties,
                'dictionary' => (array) ($decodedSection['dictionary'] ?? []),
            ];
        }
        return ['header' => substr($stream, 0, 24), 'sections' => $sections];
    }

    /** @param array{header:string,sections:list<array<string,mixed>>} $parsed */
    private function build(array $parsed): string
    {
        $sectionBytes = [];
        foreach ($parsed['sections'] as $section) {
            $properties = $section['properties'];
            if ($section['fmtid'] === self::FMTID_USER_DEFINED) {
                if ($section['dictionary'] !== []) {
                    $properties[0] = $this->dictionary($section['dictionary'], (int) $section['codepage']);
                } else {
                    unset($properties[0]);
                }
            }
            ksort($properties);
            $tableSize = 8 + (count($properties) * 8);
            $cursor = $tableSize;
            $table = '';
            $values = '';
            foreach ($properties as $id => $raw) {
                $raw = $this->pad4($raw);
                $table .= pack('V2', (int) $id, $cursor);
                $values .= $raw;
                $cursor += strlen($raw);
            }
            $size = $tableSize + strlen($values);
            $sectionBytes[] = pack('V2', $size, count($properties)) . $table . $values;
        }
        $header = strlen($parsed['header']) === 24 ? $parsed['header'] : $this->defaultHeader();
        $out = $header . pack('V', count($parsed['sections']));
        $offset = 28 + (count($parsed['sections']) * 20);
        foreach ($parsed['sections'] as $index => $section) {
            $fmtidBytes = $section['fmtid_bytes'] ?? $this->guidBytes((string) $section['fmtid']);
            $out .= $fmtidBytes . pack('V', $offset);
            $offset += strlen($sectionBytes[$index]);
        }
        return $out . implode('', $sectionBytes);
    }

    /** @return array{fmtid:string,fmtid_bytes:string,codepage:int,properties:array<int,string>,dictionary:array<int,string>} */
    private function emptySection(string $fmtid, int $codepage = 1252): array
    {
        return [
            'fmtid' => $fmtid,
            'fmtid_bytes' => $this->guidBytes($fmtid),
            'codepage' => $codepage,
            'properties' => [1 => $this->codepage($codepage)],
            'dictionary' => [],
        ];
    }

    /** @param list<array<string,mixed>> $sections */
    private function sectionIndex(array $sections, string $fmtid): ?int
    {
        foreach ($sections as $index => $section) {
            if (($section['fmtid'] ?? null) === $fmtid) {
                return $index;
            }
        }
        return null;
    }

    /** @return list<array<string,mixed>> */
    private function normalizeCustomChanges(mixed $changes): array
    {
        if (!is_array($changes)) {
            throw new MnbExcelException('custom_properties changes must be an array.');
        }
        $items = [];
        foreach ($changes as $key => $value) {
            if (is_int($key)) {
                if (!is_array($value)) {
                    throw new MnbExcelException('Each custom property list item must be an array.');
                }
                $items[] = $value;
            } elseif (is_array($value) && (array_key_exists('value', $value) || array_key_exists('type', $value))) {
                $items[] = ['name' => (string) $key] + $value;
            } else {
                $items[] = ['name' => (string) $key, 'value' => $value];
            }
        }
        return $items;
    }

    private function inferType(mixed $value): string
    {
        return match (true) {
            is_bool($value) => 'boolean',
            is_int($value) => 'integer',
            is_float($value) => 'float',
            $value instanceof DateTimeInterface => 'datetime',
            default => 'string',
        };
    }

    private function string(string $value): string
    {
        $encoded = iconv('UTF-8', 'UTF-16LE', $value);
        if (!is_string($encoded)) {
            throw new MnbExcelException('Unable to encode XLS metadata string.');
        }
        return pack('V2', 31, intdiv(strlen($encoded), 2) + 1) . $encoded . "\0\0";
    }

    private function integer(mixed $value): string
    {
        if (is_string($value) && preg_match('/^[+-]?\d+$/', trim($value)) !== 1) {
            throw new MnbExcelException('XLS metadata integer value is invalid.');
        }
        $integer = filter_var($value, FILTER_VALIDATE_INT);
        if ($integer === false || $integer < -2147483648 || $integer > 2147483647) {
            throw new MnbExcelException('XLS metadata integer must fit in signed 32 bits.');
        }
        return pack('V2', 3, $integer & 0xFFFFFFFF);
    }

    private function applicationVersion(mixed $value): string
    {
        if (is_int($value)) {
            $major = $value;
            $minor = 0;
        } elseif (is_string($value) && preg_match('/^(\d{1,5})(?:\.(\d{1,5}))?$/', trim($value), $match) === 1) {
            $major = (int) $match[1];
            $minor = isset($match[2]) ? (int) $match[2] : 0;
        } else {
            throw new MnbExcelException('XLS application version must be an integer or major.minor string.');
        }
        if ($major > 65535 || $minor > 65535) {
            throw new MnbExcelException('XLS application version components must fit in unsigned 16 bits.');
        }
        return pack('V2', 3, (($major & 0xFFFF) << 16) | ($minor & 0xFFFF));
    }

    private function float(mixed $value): string
    {
        if (!is_numeric($value) || !is_finite((float) $value)) {
            throw new MnbExcelException('XLS metadata float value is invalid.');
        }
        return pack('V', 5) . pack('e', (float) $value);
    }

    private function boolean(mixed $value): string
    {
        if (is_string($value)) {
            $normalized = strtolower(trim($value));
            if (!in_array($normalized, ['true', 'false', '1', '0', 'yes', 'no'], true)) {
                throw new MnbExcelException('XLS metadata boolean value is invalid.');
            }
            $value = in_array($normalized, ['true', '1', 'yes'], true);
        } elseif (is_int($value)) {
            if (!in_array($value, [0, 1], true)) {
                throw new MnbExcelException('XLS metadata boolean integer must be 0 or 1.');
            }
            $value = $value === 1;
        } elseif (!is_bool($value)) {
            throw new MnbExcelException('XLS metadata boolean value is invalid.');
        }
        return pack('Vv', 11, $value ? 0xFFFF : 0) . "\0\0";
    }

    private function fileTime(mixed $value): string
    {
        try {
            $date = $value instanceof DateTimeInterface
                ? DateTimeImmutable::createFromInterface($value)
                : new DateTimeImmutable((string) $value);
        } catch (\Throwable $e) {
            throw new MnbExcelException('XLS metadata date-time value is invalid: ' . $e->getMessage(), 0, $e);
        }
        $seconds = $date->setTimezone(new DateTimeZone('UTC'))->getTimestamp();
        if ($seconds < -11_644_473_600) {
            throw new MnbExcelException('XLS metadata date-time predates the FILETIME epoch.');
        }
        $ticks = ($seconds + 11_644_473_600) * 10_000_000;
        return pack('V', 64) . Binary::packU64($ticks);
    }

    private function fileTimeDuration(mixed $value): string
    {
        if (is_array($value) && isset($value['seconds'])) {
            $value = $value['seconds'];
        }
        if (!is_numeric($value) || (float) $value < 0) {
            throw new MnbExcelException('XLS total editing time must be non-negative seconds.');
        }
        $ticks = (int) round((float) $value * 10_000_000);
        return pack('V', 64) . Binary::packU64($ticks);
    }

    private function codepage(int $codepage): string
    {
        if ($codepage < 1 || $codepage > 65535) {
            $codepage = 1252;
        }
        return pack('Vv', 2, $codepage) . "\0\0";
    }

    /** @param array<int,string> $dictionary */
    private function dictionary(array $dictionary, int $codepage): string
    {
        ksort($dictionary);
        $out = pack('V', count($dictionary));
        foreach ($dictionary as $id => $name) {
            if ($codepage === 1200) {
                $encoded = iconv('UTF-8', 'UTF-16LE', $name . "\0");
                if (!is_string($encoded)) {
                    throw new MnbExcelException('Unable to encode XLS custom property name.');
                }
                $out .= pack('V2', $id, intdiv(strlen($encoded), 2)) . $encoded;
            } else {
                $encoding = $codepage === 65001 ? 'UTF-8' : ('CP' . $codepage);
                $encoded = $encoding === 'UTF-8' ? $name . "\0" : @iconv('UTF-8', $encoding . '//TRANSLIT', $name . "\0");
                if (!is_string($encoded)) {
                    throw new MnbExcelException('Unable to encode XLS custom property name for code page ' . $codepage . '.');
                }
                $out .= pack('V2', $id, strlen($encoded)) . $encoded;
            }
        }
        return $out;
    }

    private function pad4(string $value): string
    {
        return str_pad($value, (strlen($value) + 3) & ~3, "\0");
    }

    private function defaultHeader(): string
    {
        return pack('v2V', 0xFFFE, 0, 0x00020006) . str_repeat("\0", 16);
    }

    private function guid(string $bytes): string
    {
        return strtoupper(sprintf('%08X-%04X-%04X-%s-%s',
            unpack('V', substr($bytes, 0, 4))[1], unpack('v', substr($bytes, 4, 2))[1],
            unpack('v', substr($bytes, 6, 2))[1], bin2hex(substr($bytes, 8, 2)), bin2hex(substr($bytes, 10, 6))));
    }

    private function guidBytes(string $guid): string
    {
        $normalized = strtoupper(trim($guid, '{}'));
        if (preg_match('/^[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}$/', $normalized) !== 1) {
            throw new MnbExcelException('Invalid OLE property-set FMTID.');
        }
        $parts = explode('-', $normalized);
        $tail = hex2bin($parts[3] . $parts[4]);
        if (!is_string($tail)) {
            throw new MnbExcelException('Invalid OLE property-set FMTID bytes.');
        }
        return pack('Vvv', hexdec($parts[0]), hexdec($parts[1]), hexdec($parts[2])) . $tail;
    }
}
