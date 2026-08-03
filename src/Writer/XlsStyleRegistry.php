<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Writer;

use Mnb\PHPExcel\Biff\BiffRecordWriter;
use Mnb\PHPExcel\Biff\RecordType;
use Mnb\PHPExcel\Biff\String\BiffString;
use Mnb\PHPExcel\Core\CellValue;
use Mnb\PHPExcel\Core\StyleNormalizer;
use Mnb\PHPExcel\Core\WorkbookData;
use Mnb\PHPExcel\Core\WorksheetData;
use Mnb\PHPExcel\Exception\UnsupportedXlsFeatureException;
use Mnb\PHPExcel\Support\Coordinate;

/** Builds a BIFF8 font/palette/format/XF table from canonical styles. */
final class XlsStyleRegistry
{
    /** @var array<int,array<string,mixed>> */
    private array $styles = [];
    /** @var array<string,int> */
    private array $styleIds = [];
    /** @var array<int,array<string,mixed>> */
    private array $fonts = [];
    /** @var array<string,int> */
    private array $fontIds = [];
    /** @var array<string,int> */
    private array $customFormatIds = [];
    /** @var array<string,int> RGB => palette index */
    private array $palette = [];

    public function __construct()
    {
        $this->registerFont(['name' => 'Arial', 'size' => 10.0]);
        $this->registerStyle([]); // XF 0
        $this->registerStyle(['format' => 'm/d/yy']); // XF 1
        $this->registerStyle(['format' => 'm/d/yy h:mm']); // XF 2
    }

    public function registerWorkbook(WorkbookData $workbook): void
    {
        foreach ($workbook->sheets as $sheet) {
            foreach ($sheet->rows as $rowIndex => $row) {
                foreach (array_values($row) as $columnIndex => $value) {
                    $cell = Coordinate::columnIndexToName($columnIndex + 1) . ($rowIndex + 1);
                    $style = $this->effectiveStyle($sheet, $rowIndex + 1, $columnIndex + 1, $cell);
                    if ($value instanceof CellValue && $value->type() === CellValue::TYPE_DATE && !isset($style['format'])) {
                        $style['format'] = (string) ($value->options()['format'] ?? 'm/d/yy');
                    }
                    $this->registerStyle($style);
                }
            }
        }
    }

    public function styleId(WorksheetData $sheet, int $row, int $column, string $cell, mixed $value): int
    {
        $style = $this->effectiveStyle($sheet, $row, $column, $cell);
        if ($value instanceof CellValue && $value->type() === CellValue::TYPE_DATE && !isset($style['format'])) {
            $style['format'] = (string) ($value->options()['format'] ?? 'm/d/yy');
        }
        if ($style === [] && $value instanceof \DateTimeInterface) {
            return $value->format('His') === '000000' ? 1 : 2;
        }
        return $this->registerStyle($style);
    }

    public function globalsRecords(): string
    {
        $stream = '';
        foreach ($this->fonts as $font) {
            $stream .= BiffRecordWriter::record(RecordType::FONT, $this->fontRecord($font));
        }
        if ($this->palette !== []) {
            $colors = array_map(static fn (string $key): string => ltrim($key, '#'), array_keys($this->palette));
            $payload = pack('v', count($colors));
            foreach ($colors as $rgb) {
                $payload .= pack('C4', hexdec(substr($rgb, 0, 2)), hexdec(substr($rgb, 2, 2)), hexdec(substr($rgb, 4, 2)), 0);
            }
            $stream .= BiffRecordWriter::record(RecordType::PALETTE, $payload);
        }
        foreach ($this->customFormatIds as $format => $id) {
            $stream .= BiffRecordWriter::record(RecordType::FORMAT, pack('v', $id) . BiffString::writeUnicodeString($format, false));
        }
        foreach ($this->styles as $style) {
            $stream .= BiffRecordWriter::record(RecordType::XF, $this->xfRecord($style));
        }
        return $stream;
    }

    /** @param array<string,mixed> $style */
    private function registerStyle(array $style): int
    {
        $style = StyleNormalizer::normalize($style);
        $key = json_encode($style, JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE) ?: serialize($style);
        if (isset($this->styleIds[$key])) {
            return $this->styleIds[$key];
        }
        if (count($this->styles) >= 4000) {
            throw UnsupportedXlsFeatureException::forFeature('more than 4000 cell formats');
        }
        $this->collectStyleResources($style);
        $id = count($this->styles);
        $this->styleIds[$key] = $id;
        $this->styles[$id] = $style;
        return $id;
    }

    /** @param array<string,mixed> $style */
    private function collectStyleResources(array $style): void
    {
        if (is_array($style['font'] ?? null)) {
            $this->registerFont($style['font']);
            $this->collectColor($style['font']['color'] ?? null);
        }
        if (is_array($style['fill'] ?? null)) {
            $this->collectColor($style['fill']['foreground'] ?? null);
            $this->collectColor($style['fill']['background'] ?? null);
        }
        if (is_array($style['border'] ?? null)) {
            foreach (['left','right','top','bottom','diagonal'] as $side) {
                $this->collectColor($style['border'][$side]['color'] ?? null);
            }
        }
        $format = $this->formatCode($style['format'] ?? null);
        if ($format !== null && $this->builtinFormatId($format) === null && !isset($this->customFormatIds[$format])) {
            $this->customFormatIds[$format] = 164 + count($this->customFormatIds);
        }
    }

    /** @param array<string,mixed> $font */
    private function registerFont(array $font): int
    {
        $font = StyleNormalizer::normalize(['font' => $font])['font'] ?? [];
        $font += ['name' => 'Arial', 'size' => 10.0];
        $key = json_encode($font, JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE) ?: serialize($font);
        if (isset($this->fontIds[$key])) {
            return $this->fontIds[$key];
        }
        if (count($this->fonts) >= 1024) {
            throw UnsupportedXlsFeatureException::forFeature('too many fonts');
        }
        $logicalIndex = count($this->fonts);
        $biffIndex = $logicalIndex >= 4 ? $logicalIndex + 1 : $logicalIndex;
        $this->fontIds[$key] = $biffIndex;
        $this->fonts[] = $font;
        return $biffIndex;
    }

    private function collectColor(mixed $color): void
    {
        $rgb = $this->rgb($color);
        if ($rgb === null || isset($this->palette['#' . $rgb])) {
            return;
        }
        if (count($this->palette) >= 56) {
            throw UnsupportedXlsFeatureException::forFeature('more than 56 custom palette colors');
        }
        $this->palette['#' . $rgb] = 8 + count($this->palette);
    }

    /** @param array<string,mixed> $font */
    private function fontRecord(array $font): string
    {
        $height = max(20, min(8191, (int) round((float) ($font['size'] ?? 10) * 20)));
        $options = 0;
        if ((bool) ($font['italic'] ?? false)) $options |= 0x0002;
        if ((bool) ($font['strike'] ?? false)) $options |= 0x0008;
        if ((bool) ($font['outline'] ?? false)) $options |= 0x0010;
        if ((bool) ($font['shadow'] ?? false)) $options |= 0x0020;
        if ((bool) ($font['condense'] ?? false)) $options |= 0x0040;
        if ((bool) ($font['extend'] ?? false)) $options |= 0x0080;
        $color = $this->colorIndex($font['color'] ?? null, 0x7FFF);
        $weight = (bool) ($font['bold'] ?? false) ? 700 : 400;
        $vertical = match ((string) ($font['vertical_align'] ?? '')) {
            'superscript', 'super' => 1,
            'subscript', 'sub' => 2,
            default => 0,
        };
        $underline = match ((string) ($font['underline'] ?? '')) {
            'single', 'true', '1' => 1,
            'double' => 2,
            'singleAccounting', 'single_accounting' => 0x21,
            'doubleAccounting', 'double_accounting' => 0x22,
            default => 0,
        };
        return pack('vvvvvCCCC', $height, $options, $color, $weight, $vertical, $underline, (int) ($font['family'] ?? 0), (int) ($font['charset'] ?? 0), 0)
            . BiffString::writeUnicodeString((string) ($font['name'] ?? 'Arial'), true);
    }

    /** @param array<string,mixed> $style */
    private function xfRecord(array $style): string
    {
        $font = is_array($style['font'] ?? null) ? $style['font'] : ['name' => 'Arial', 'size' => 10.0];
        $fontIndex = $this->registerFont($font);
        $format = $this->formatCode($style['format'] ?? null);
        $formatId = $format === null ? 0 : ($this->builtinFormatId($format) ?? $this->customFormatIds[$format] ?? 0);
        $protection = is_array($style['protection'] ?? null) ? $style['protection'] : [];
        $typeProtection = ((bool) ($protection['locked'] ?? true) ? 0x0001 : 0)
            | ((bool) ($protection['hidden'] ?? false) ? 0x0002 : 0);

        $alignment = is_array($style['alignment'] ?? null) ? $style['alignment'] : [];
        $horizontal = match ((string) ($alignment['horizontal'] ?? 'general')) {
            'left' => 1, 'center' => 2, 'right' => 3, 'fill' => 4,
            'justify' => 5, 'centerContinuous', 'center_continuous' => 6,
            'distributed' => 7, default => 0,
        };
        $vertical = match ((string) ($alignment['vertical'] ?? 'bottom')) {
            'top' => 0, 'center' => 1, 'justify' => 3, 'distributed' => 4, default => 2,
        };
        $alignmentByte = $horizontal | ((bool) ($alignment['wrap_text'] ?? false) ? 0x08 : 0) | ($vertical << 4);
        $rotation = max(0, min(180, (int) ($alignment['text_rotation'] ?? 0)));
        $text = max(0, min(15, (int) ($alignment['indent'] ?? 0)))
            | ((bool) ($alignment['shrink_to_fit'] ?? false) ? 0x10 : 0)
            | ((max(0, min(2, (int) ($alignment['reading_order'] ?? 0)))) << 6);

        $border = is_array($style['border'] ?? null) ? $style['border'] : [];
        $left = $this->borderStyle($border['left']['style'] ?? null);
        $right = $this->borderStyle($border['right']['style'] ?? null);
        $top = $this->borderStyle($border['top']['style'] ?? null);
        $bottom = $this->borderStyle($border['bottom']['style'] ?? null);
        $leftColor = $this->colorIndex($border['left']['color'] ?? null, 0x40) & 0x7F;
        $rightColor = $this->colorIndex($border['right']['color'] ?? null, 0x40) & 0x7F;
        $border1 = $left | ($right << 4) | ($top << 8) | ($bottom << 12) | ($leftColor << 16) | ($rightColor << 23);
        if ((bool) ($border['diagonal_down'] ?? false)) $border1 |= 0x40000000;
        if ((bool) ($border['diagonal_up'] ?? false)) $border1 |= 0x80000000;
        $topColor = $this->colorIndex($border['top']['color'] ?? null, 0x40) & 0x7F;
        $bottomColor = $this->colorIndex($border['bottom']['color'] ?? null, 0x40) & 0x7F;
        $diagonalColor = $this->colorIndex($border['diagonal']['color'] ?? null, 0x40) & 0x7F;
        $diagonalStyle = $this->borderStyle($border['diagonal']['style'] ?? null);

        $fill = is_array($style['fill'] ?? null) ? $style['fill'] : [];
        $pattern = $this->fillPattern((string) ($fill['pattern'] ?? 'none'));
        $border2 = $topColor | ($bottomColor << 7) | ($diagonalColor << 14) | ($diagonalStyle << 21) | ($pattern << 26);
        $foreground = $this->colorIndex($fill['foreground'] ?? null, 0x40) & 0x7F;
        $background = $this->colorIndex($fill['background'] ?? null, 0x41) & 0x7F;
        $patternColors = $foreground | ($background << 7);

        return pack('vvvCCCCVVv', $fontIndex, $formatId, $typeProtection, $alignmentByte, $rotation, $text, 0xFC, $border1, $border2, $patternColors);
    }

    /** @return array<string,mixed> */
    private function effectiveStyle(WorksheetData $sheet, int $row, int $column, string $cell): array
    {
        $style = [];
        if ($sheet->headerRowIndex !== null && $row === $sheet->headerRowIndex && $sheet->hasHeader) {
            return StyleNormalizer::normalize($this->resolveStyle($sheet, $sheet->headerStyle));
        }
        foreach ([
            $sheet->columnStyles[$column] ?? null,
            $sheet->rowStyles[$row] ?? null,
        ] as $candidate) {
            $style = $this->merge($style, $this->resolveStyle($sheet, $candidate));
        }
        foreach ($sheet->rangeStyles as $range => $candidate) {
            if ($this->cellInRange($cell, (string) $range)) {
                $style = $this->merge($style, $this->resolveStyle($sheet, $candidate));
            }
        }
        $style = $this->merge($style, $this->resolveStyle($sheet, $sheet->cellStyles[$cell] ?? null));
        return StyleNormalizer::normalize($style);
    }

    /** @return array<string,mixed> */
    private function resolveStyle(WorksheetData $sheet, mixed $style): array
    {
        if (is_string($style)) return is_array($sheet->namedStyles[$style] ?? null) ? $sheet->namedStyles[$style] : [];
        return is_array($style) ? $style : [];
    }

    /** @param array<string,mixed> $base @param array<string,mixed> $override @return array<string,mixed> */
    private function merge(array $base, array $override): array
    {
        foreach ($override as $key => $value) {
            $base[$key] = isset($base[$key]) && is_array($base[$key]) && is_array($value)
                ? $this->merge($base[$key], $value)
                : $value;
        }
        return $base;
    }

    private function cellInRange(string $cell, string $range): bool
    {
        [$start, $end] = array_pad(explode(':', strtoupper($range), 2), 2, strtoupper($range));
        try {
            [$column, $row] = Coordinate::splitCellRef($cell);
            [$c1, $r1] = Coordinate::splitCellRef($start);
            [$c2, $r2] = Coordinate::splitCellRef($end);
            return $column >= min($c1, $c2) && $column <= max($c1, $c2) && $row >= min($r1, $r2) && $row <= max($r1, $r2);
        } catch (\Throwable) {
            return false;
        }
    }

    private function formatCode(mixed $format): ?string
    {
        if ($format === null || $format === '' || strtolower((string) $format) === 'general') return null;
        return match (strtolower((string) $format)) {
            'text' => '@', 'integer', 'int' => '#,##0', 'number', 'decimal' => '#,##0.00',
            'date' => 'm/d/yy', 'datetime' => 'm/d/yy h:mm', 'percent' => '0.00%',
            default => (string) $format,
        };
    }

    private function builtinFormatId(string $format): ?int
    {
        return match ($format) {
            'General' => 0, '0' => 1, '0.00' => 2, '#,##0' => 3, '#,##0.00' => 4,
            '0%' => 9, '0.00%' => 10, '0.00E+00' => 11, '# ?/?' => 12, '# ??/??' => 13,
            'm/d/yy' => 14, 'd-mmm-yy' => 15, 'd-mmm' => 16, 'mmm-yy' => 17,
            'h:mm AM/PM' => 18, 'h:mm:ss AM/PM' => 19, 'h:mm' => 20, 'h:mm:ss' => 21,
            'm/d/yy h:mm' => 22, 'mm:ss' => 45, '[h]:mm:ss' => 46, 'mmss.0' => 47, '@' => 49,
            default => null,
        };
    }

    private function colorIndex(mixed $color, int $fallback): int
    {
        $rgb = $this->rgb($color);
        return $rgb !== null ? ($this->palette['#' . $rgb] ?? $fallback) : $fallback;
    }

    private function rgb(mixed $color): ?string
    {
        if (is_array($color)) $color = $color['rgb'] ?? null;
        if (!is_string($color)) return null;
        $value = strtoupper(ltrim(trim($color), '#'));
        if (strlen($value) === 8) $value = substr($value, 2);
        return preg_match('/^[0-9A-F]{6}$/', $value) === 1 ? $value : null;
    }

    private function borderStyle(mixed $style): int
    {
        return match ((string) $style) {
            'thin' => 1, 'medium' => 2, 'dashed' => 3, 'dotted' => 4, 'thick' => 5,
            'double' => 6, 'hair' => 7, 'mediumDashed', 'medium_dashed' => 8,
            'dashDot', 'dash_dot' => 9, 'mediumDashDot', 'medium_dash_dot' => 10,
            'dashDotDot', 'dash_dot_dot' => 11, 'mediumDashDotDot', 'medium_dash_dot_dot' => 12,
            'slantDashDot', 'slant_dash_dot' => 13, default => 0,
        };
    }

    private function fillPattern(string $pattern): int
    {
        return match ($pattern) {
            'solid' => 1, 'mediumGray', 'medium_gray' => 2, 'darkGray', 'dark_gray' => 3,
            'lightGray', 'light_gray' => 4, 'darkHorizontal', 'dark_horizontal' => 5,
            'darkVertical', 'dark_vertical' => 6, 'darkDown', 'dark_down' => 7,
            'darkUp', 'dark_up' => 8, 'darkGrid', 'dark_grid' => 9, 'darkTrellis', 'dark_trellis' => 10,
            'lightHorizontal', 'light_horizontal' => 11, 'lightVertical', 'light_vertical' => 12,
            'lightDown', 'light_down' => 13, 'lightUp', 'light_up' => 14,
            'lightGrid', 'light_grid' => 15, 'lightTrellis', 'light_trellis' => 16,
            'gray125' => 17, 'gray0625' => 18, default => 0,
        };
    }
}
