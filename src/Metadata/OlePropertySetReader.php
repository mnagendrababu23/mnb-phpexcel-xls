<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Metadata;

use DateTimeImmutable;
use DateTimeZone;
use Mnb\PHPExcel\Support\Binary;

/** Reads OLE Property Set streams used by legacy XLS document properties. */
final class OlePropertySetReader
{
    private const FMTID_SUMMARY = 'F29F85E0-4FF9-1068-AB91-08002B27B3D9';
    private const FMTID_DOCUMENT_SUMMARY = 'D5CDD502-2E9C-101B-9397-08002B2CF9AE';
    private const FMTID_USER_DEFINED = 'D5CDD505-2E9C-101B-9397-08002B2CF9AE';

    private const VT_I2 = 2;
    private const VT_I4 = 3;
    private const VT_R4 = 4;
    private const VT_R8 = 5;
    private const VT_DATE = 7;
    private const VT_BSTR = 8;
    private const VT_BOOL = 11;
    private const VT_VARIANT = 12;
    private const VT_I1 = 16;
    private const VT_UI1 = 17;
    private const VT_UI2 = 18;
    private const VT_UI4 = 19;
    private const VT_I8 = 20;
    private const VT_UI8 = 21;
    private const VT_INT = 22;
    private const VT_UINT = 23;
    private const VT_LPSTR = 30;
    private const VT_LPWSTR = 31;
    private const VT_FILETIME = 64;
    private const VT_BLOB = 65;
    private const VT_CF = 71;
    private const VT_CLSID = 72;
    private const VT_VECTOR = 0x1000;

    /**
     * @return array{
     *   summary:array<string,mixed>,document_summary:array<string,mixed>,custom_properties:array<string,mixed>,
     *   sections:list<array<string,mixed>>,warnings:list<string>
     * }
     */
    public function read(string $stream): array
    {
        $warnings = [];
        if (strlen($stream) < 28) {
            return ['summary' => [], 'document_summary' => [], 'custom_properties' => [], 'sections' => [], 'warnings' => ['OLE property stream is too short.']];
        }
        if (Binary::u16($stream, 0) !== 0xFFFE) {
            return ['summary' => [], 'document_summary' => [], 'custom_properties' => [], 'sections' => [], 'warnings' => ['Unsupported OLE property-set byte order.']];
        }

        $sectionCount = Binary::u32($stream, 24);
        if ($sectionCount < 1 || $sectionCount > 32 || 28 + ($sectionCount * 20) > strlen($stream)) {
            return ['summary' => [], 'document_summary' => [], 'custom_properties' => [], 'sections' => [], 'warnings' => ['Invalid OLE property-set section table.']];
        }

        $summary = [];
        $documentSummary = [];
        $custom = [];
        $sections = [];
        for ($i = 0; $i < $sectionCount; $i++) {
            $descriptor = 28 + ($i * 20);
            $fmtid = $this->guid(substr($stream, $descriptor, 16));
            $offset = Binary::u32($stream, $descriptor + 16);
            try {
                $section = $this->readSection($stream, $offset, $fmtid);
                $sections[] = $section;
                if ($fmtid === self::FMTID_SUMMARY) {
                    $summary = $this->mapProperties($section['properties'], self::summaryNames(), true);
                } elseif ($fmtid === self::FMTID_DOCUMENT_SUMMARY) {
                    $documentSummary = $this->mapProperties($section['properties'], self::documentSummaryNames(), false);
                } elseif ($fmtid === self::FMTID_USER_DEFINED) {
                    foreach ($section['properties'] as $id => $property) {
                        if ($id < 2) {
                            continue;
                        }
                        $name = $section['dictionary'][$id] ?? ('Property ' . $id);
                        $custom[$name] = [
                            'id' => $id,
                            'type' => $property['type_name'],
                            'value' => $property['value'],
                        ];
                    }
                }
            } catch (\Throwable $e) {
                $warnings[] = 'Unable to parse OLE property section ' . ($i + 1) . ': ' . $e->getMessage();
            }
        }

        return [
            'summary' => $summary,
            'document_summary' => $documentSummary,
            'custom_properties' => $custom,
            'sections' => $sections,
            'warnings' => $warnings,
        ];
    }

    /** @return array{fmtid:string,codepage:int,properties:array<int,array{type:int,type_name:string,value:mixed}>,dictionary:array<int,string>} */
    private function readSection(string $stream, int $offset, string $fmtid): array
    {
        if ($offset < 0 || $offset + 8 > strlen($stream)) {
            throw new \RuntimeException('Section offset is outside the stream.');
        }
        $size = Binary::u32($stream, $offset);
        $count = Binary::u32($stream, $offset + 4);
        if ($size < 8 || $offset + $size > strlen($stream) || $count > 100000 || $offset + 8 + ($count * 8) > $offset + $size) {
            throw new \RuntimeException('Section header is invalid.');
        }

        $offsets = [];
        for ($i = 0; $i < $count; $i++) {
            $entryOffset = $offset + 8 + ($i * 8);
            $id = Binary::u32($stream, $entryOffset);
            $relative = Binary::u32($stream, $entryOffset + 4);
            if ($relative >= $size) {
                throw new \RuntimeException('Property offset is outside its section.');
            }
            $offsets[$id] = $offset + $relative;
        }

        $codepage = 1252;
        if (isset($offsets[1])) {
            $parsed = $this->readTypedValue($stream, $offsets[1], $offset + $size, $codepage);
            if (is_int($parsed['value'])) {
                $candidate = $parsed['value'] < 0 ? $parsed['value'] + 65536 : $parsed['value'];
                if ($candidate > 0) {
                    $codepage = $candidate;
                }
            }
        }

        $dictionary = [];
        if ($fmtid === self::FMTID_USER_DEFINED && isset($offsets[0])) {
            $dictionary = $this->readDictionary($stream, $offsets[0], $offset + $size, $codepage);
        }

        $properties = [];
        foreach ($offsets as $id => $propertyOffset) {
            if ($id === 0) {
                continue;
            }
            $properties[$id] = $this->readTypedValue($stream, $propertyOffset, $offset + $size, $codepage);
        }
        ksort($properties);

        return [
            'fmtid' => $fmtid,
            'codepage' => $codepage,
            'properties' => $properties,
            'dictionary' => $dictionary,
        ];
    }

    /** @return array<int,string> */
    private function readDictionary(string $stream, int $offset, int $limit, int $codepage): array
    {
        if ($offset + 4 > $limit) {
            return [];
        }
        $count = Binary::u32($stream, $offset);
        $cursor = $offset + 4;
        $dictionary = [];
        for ($i = 0; $i < $count && $cursor + 8 <= $limit; $i++) {
            $id = Binary::u32($stream, $cursor);
            $length = Binary::u32($stream, $cursor + 4);
            $cursor += 8;
            if ($length > 1_000_000) {
                break;
            }
            if ($codepage === 1200) {
                $byteLength = $length * 2;
                if ($cursor + $byteLength > $limit) {
                    break;
                }
                $raw = substr($stream, $cursor, $byteLength);
                $name = $this->convertString($this->trimUnicodeNull($raw), 1200);
                $cursor += $byteLength;
            } else {
                if ($cursor + $length > $limit) {
                    break;
                }
                $raw = substr($stream, $cursor, $length);
                $name = $this->convertString(rtrim($raw, "\0"), $codepage);
                $cursor += $length;
            }
            $dictionary[$id] = $name;
        }
        return $dictionary;
    }

    /** @return array{type:int,type_name:string,value:mixed,bytes:int} */
    private function readTypedValue(string $stream, int $offset, int $limit, int $codepage): array
    {
        if ($offset + 4 > $limit) {
            throw new \RuntimeException('Typed property value is truncated.');
        }
        $type = Binary::u32($stream, $offset);
        $cursor = $offset + 4;
        $baseType = $type & 0x0FFF;
        if (($type & self::VT_VECTOR) !== 0) {
            if ($cursor + 4 > $limit) {
                throw new \RuntimeException('Vector property is truncated.');
            }
            $count = Binary::u32($stream, $cursor);
            $cursor += 4;
            $values = [];
            for ($i = 0; $i < $count && $i < 100000; $i++) {
                if ($baseType === self::VT_VARIANT) {
                    $item = $this->readTypedValue($stream, $cursor, $limit, $codepage);
                } else {
                    $item = $this->readScalar($stream, $cursor, $limit, $baseType, $codepage, false);
                }
                $values[] = $item['value'];
                $cursor += $item['bytes'];
                $cursor = $this->align4($cursor);
            }
            return ['type' => $type, 'type_name' => 'vector<' . $this->typeName($baseType) . '>', 'value' => $values, 'bytes' => $cursor - $offset];
        }

        $scalar = $this->readScalar($stream, $cursor, $limit, $baseType, $codepage, true);
        $result = ['type' => $type, 'type_name' => $this->typeName($baseType), 'value' => $scalar['value'], 'bytes' => 4 + $scalar['bytes']];
        if (array_key_exists('raw_value', $scalar)) {
            $result['raw_value'] = $scalar['raw_value'];
        }
        return $result;
    }

    /** @return array{value:mixed,bytes:int} */
    private function readScalar(string $stream, int $offset, int $limit, int $type, int $codepage, bool $align): array
    {
        $need = static function (int $bytes) use ($offset, $limit): void {
            if ($offset + $bytes > $limit) {
                throw new \RuntimeException('Property scalar is truncated.');
            }
        };
        $bytes = 0;
        $value = null;
        $rawValue = null;
        switch ($type) {
            case self::VT_I1:
                $need(1); $v = ord($stream[$offset]); $value = $v >= 128 ? $v - 256 : $v; $bytes = 1; break;
            case self::VT_UI1:
                $need(1); $value = ord($stream[$offset]); $bytes = 1; break;
            case self::VT_I2:
                $need(2); $v = Binary::u16($stream, $offset); $value = $v >= 0x8000 ? $v - 0x10000 : $v; $bytes = 2; break;
            case self::VT_UI2:
                $need(2); $value = Binary::u16($stream, $offset); $bytes = 2; break;
            case self::VT_I4:
            case self::VT_INT:
                $need(4); $value = Binary::i32($stream, $offset); $bytes = 4; break;
            case self::VT_UI4:
            case self::VT_UINT:
                $need(4); $value = Binary::u32($stream, $offset); $bytes = 4; break;
            case self::VT_I8:
            case self::VT_UI8:
                $need(8); $value = $this->u64($stream, $offset); $bytes = 8; break;
            case self::VT_R4:
                $need(4); $value = unpack('g', substr($stream, $offset, 4))[1]; $bytes = 4; break;
            case self::VT_R8:
                $need(8); $value = Binary::double($stream, $offset); $bytes = 8; break;
            case self::VT_DATE:
                $need(8); $days = Binary::double($stream, $offset); $value = $this->oleAutomationDate($days); $bytes = 8; break;
            case self::VT_BOOL:
                $need(2); $value = Binary::u16($stream, $offset) !== 0; $bytes = 2; break;
            case self::VT_LPSTR:
            case self::VT_BSTR:
                $need(4); $length = Binary::u32($stream, $offset); $need(4 + $length);
                $value = $this->convertString(rtrim(substr($stream, $offset + 4, $length), "\0"), $codepage);
                $bytes = 4 + $length; break;
            case self::VT_LPWSTR:
                $need(4); $characters = Binary::u32($stream, $offset); $length = $characters * 2; $need(4 + $length);
                $value = $this->convertString($this->trimUnicodeNull(substr($stream, $offset + 4, $length)), 1200);
                $bytes = 4 + $length; break;
            case self::VT_FILETIME:
                $need(8); $ticks = $this->u64($stream, $offset); $rawValue = $ticks; $value = $this->fileTime($ticks); $bytes = 8; break;
            case self::VT_CLSID:
                $need(16); $value = $this->guid(substr($stream, $offset, 16)); $bytes = 16; break;
            case self::VT_BLOB:
            case self::VT_CF:
                $need(4); $length = Binary::u32($stream, $offset); $need(4 + $length);
                $value = ['size_bytes' => $length, 'sha256' => hash('sha256', substr($stream, $offset + 4, $length))];
                $bytes = 4 + $length; break;
            default:
                $value = ['unsupported_type' => $type];
                $bytes = 0;
        }
        if ($align) {
            $bytes = $this->align4($offset + $bytes) - $offset;
        }
        $result = ['value' => $value, 'bytes' => $bytes];
        if ($rawValue !== null) {
            $result['raw_value'] = $rawValue;
        }
        return $result;
    }

    /** @param array<int,array{type:int,type_name:string,value:mixed}> $properties @param array<int,string> $names @return array<string,mixed> */
    private function mapProperties(array $properties, array $names, bool $summary): array
    {
        $mapped = [];
        foreach ($properties as $id => $property) {
            if ($id === 1) {
                continue;
            }
            $name = $names[$id] ?? ('property_' . $id);
            $value = $property['value'];
            if ($summary && $id === 10 && isset($property['raw_value']) && is_int($property['raw_value'])) {
                // SummaryInformation total editing time uses FILETIME ticks as a duration.
                $value = intdiv($property['raw_value'], 10_000_000);
            } elseif (!$summary && $id === 23 && is_int($value)) {
                // PIDDSI_VERSION stores major/minor components in the high/low words.
                $value = (($value >> 16) & 0xFFFF) . '.' . ($value & 0xFFFF);
            }
            $mapped[$name] = $value;
        }
        return $mapped;
    }

    /** @return array<int,string> */
    private static function summaryNames(): array
    {
        return [
            2 => 'title', 3 => 'subject', 4 => 'creator', 5 => 'keywords', 6 => 'comments',
            7 => 'template', 8 => 'last_saved_by', 9 => 'revision_number', 10 => 'total_editing_time',
            11 => 'last_printed_at', 12 => 'created_at', 13 => 'modified_at', 14 => 'page_count',
            15 => 'word_count', 16 => 'character_count', 18 => 'application_name', 19 => 'security_flags',
        ];
    }

    /** @return array<int,string> */
    private static function documentSummaryNames(): array
    {
        return [
            2 => 'category', 3 => 'presentation_format', 4 => 'byte_count', 5 => 'line_count',
            6 => 'paragraph_count', 7 => 'slide_count', 8 => 'note_count', 9 => 'hidden_slide_count',
            10 => 'multimedia_clip_count', 11 => 'scale_crop', 12 => 'heading_pairs', 13 => 'titles_of_parts',
            14 => 'manager', 15 => 'company', 16 => 'links_dirty', 17 => 'character_count_with_spaces',
            19 => 'shared_document', 22 => 'hyperlinks_changed', 23 => 'application_version',
            26 => 'content_type', 27 => 'content_status', 28 => 'language', 29 => 'document_version',
        ];
    }

    private function trimUnicodeNull(string $value): string
    {
        return str_ends_with($value, "\0\0") ? substr($value, 0, -2) : $value;
    }

    private function convertString(string $value, int $codepage): string
    {
        if ($value === '') {
            return '';
        }
        $encoding = match ($codepage) {
            1200 => 'UTF-16LE', 65001 => 'UTF-8', 1250 => 'Windows-1250', 1251 => 'Windows-1251',
            1252 => 'Windows-1252', 1253 => 'Windows-1253', 1254 => 'Windows-1254', 1255 => 'Windows-1255',
            1256 => 'Windows-1256', 1257 => 'Windows-1257', 1258 => 'Windows-1258', 932 => 'CP932',
            936 => 'CP936', 949 => 'CP949', 950 => 'CP950', default => 'Windows-1252',
        };
        if ($encoding === 'UTF-8') {
            return $value;
        }
        $converted = @iconv($encoding, 'UTF-8//IGNORE', $value);
        return is_string($converted) ? $converted : $value;
    }

    private function typeName(int $type): string
    {
        return match ($type) {
            self::VT_I1 => 'int8', self::VT_UI1 => 'uint8', self::VT_I2 => 'int16', self::VT_UI2 => 'uint16',
            self::VT_I4, self::VT_INT => 'int32', self::VT_UI4, self::VT_UINT => 'uint32', self::VT_I8 => 'int64',
            self::VT_UI8 => 'uint64', self::VT_R4 => 'float32', self::VT_R8 => 'float64', self::VT_DATE => 'date',
            self::VT_BSTR => 'bstr', self::VT_BOOL => 'boolean', self::VT_VARIANT => 'variant', self::VT_LPSTR => 'string',
            self::VT_LPWSTR => 'unicode_string', self::VT_FILETIME => 'filetime', self::VT_BLOB => 'blob',
            self::VT_CF => 'clipboard_data', self::VT_CLSID => 'clsid', default => 'unknown_' . $type,
        };
    }

    private function fileTime(int $ticks): ?string
    {
        if ($ticks === 0) {
            return null;
        }
        $seconds = intdiv($ticks, 10_000_000) - 11_644_473_600;
        if ($seconds < -62135596800 || $seconds > 253402300799) {
            return null;
        }
        return (new DateTimeImmutable('@' . $seconds))->setTimezone(new DateTimeZone('UTC'))->format('Y-m-d\TH:i:s\Z');
    }

    private function oleAutomationDate(float $days): ?string
    {
        if (!is_finite($days)) {
            return null;
        }
        $seconds = (int) round(($days - 25569.0) * 86400.0);
        return (new DateTimeImmutable('@' . $seconds))->setTimezone(new DateTimeZone('UTC'))->format('Y-m-d\TH:i:s\Z');
    }

    private function u64(string $data, int $offset): int
    {
        $parts = unpack('Vlow/Vhigh', substr($data, $offset, 8));
        if (!is_array($parts)) {
            throw new \RuntimeException('Unable to decode 64-bit property value.');
        }
        $high = (int) $parts['high'];
        if ($high > 0x7FFFFFFF) {
            throw new \RuntimeException('Unsigned 64-bit property exceeds the supported PHP integer range.');
        }
        return (int) ((int) $parts['low'] + ($high * 4294967296));
    }

    private function align4(int $offset): int
    {
        return ($offset + 3) & ~3;
    }

    private function guid(string $bytes): string
    {
        if (strlen($bytes) !== 16) {
            return strtoupper(bin2hex($bytes));
        }
        return strtoupper(sprintf(
            '%08X-%04X-%04X-%s-%s',
            unpack('V', substr($bytes, 0, 4))[1],
            unpack('v', substr($bytes, 4, 2))[1],
            unpack('v', substr($bytes, 6, 2))[1],
            bin2hex(substr($bytes, 8, 2)),
            bin2hex(substr($bytes, 10, 6))
        ));
    }
}
