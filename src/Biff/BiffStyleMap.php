<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Biff;

use Mnb\PHPExcel\Biff\String\BiffString;
use Mnb\PHPExcel\Core\StyleNormalizer;
use Mnb\PHPExcel\Support\Binary;

/** Decodes BIFF8 FONT, PALETTE, FORMAT, and XF records into canonical styles. */
final class BiffStyleMap
{
    /** @param array<int,array<string,mixed>> $styles */
    private function __construct(private readonly array $styles)
    {
    }

    public static function fromWorkbookStream(string $stream, array $options = []): self
    {
        $fonts = [];
        $palette = self::defaultPalette();
        $formats = [];
        $xfPayloads = [];
        $fontSequence = 0;
        foreach ((new BiffRecordReader($stream, $options))->records(0) as $record) {
            if ($record->type === RecordType::EOF) break;
            if ($record->type === RecordType::FONT && $record->length() >= 15) {
                $fontIndex = $fontSequence >= 4 ? $fontSequence + 1 : $fontSequence;
                $fonts[$fontIndex] = self::parseFont($record->payload, $palette);
                $fontSequence++;
            } elseif ($record->type === RecordType::PALETTE && $record->length() >= 2) {
                $count = min(56, Binary::u16($record->payload, 0));
                for ($i = 0; $i < $count && 2 + ($i * 4) + 3 < strlen($record->payload); $i++) {
                    $offset = 2 + ($i * 4);
                    $palette[8 + $i] = sprintf('%02X%02X%02X', Binary::u8($record->payload, $offset), Binary::u8($record->payload, $offset + 1), Binary::u8($record->payload, $offset + 2));
                }
            } elseif ($record->type === RecordType::FORMAT && $record->length() >= 5) {
                $id = Binary::u16($record->payload, 0);
                $formats[$id] = BiffString::readUnicodeString($record->payload, 2, false)['value'];
            } elseif ($record->type === RecordType::XF && $record->length() >= 20) {
                $xfPayloads[] = $record->payload;
            }
        }
        // Palette can occur after fonts; resolve font colors again from retained indices.
        $fontSequence = 0;
        $fonts = [];
        foreach ((new BiffRecordReader($stream, $options))->records(0) as $record) {
            if ($record->type === RecordType::EOF) break;
            if ($record->type === RecordType::FONT && $record->length() >= 15) {
                $fontIndex = $fontSequence >= 4 ? $fontSequence + 1 : $fontSequence;
                $fonts[$fontIndex] = self::parseFont($record->payload, $palette);
                $fontSequence++;
            }
        }
        $styles = [];
        foreach ($xfPayloads as $index => $payload) {
            $styles[$index] = self::parseXf($payload, $fonts, $palette, $formats);
        }
        return new self($styles);
    }

    /** @return array<string,mixed> */
    public function styleForIndex(?int $index): array
    {
        return $index === null ? [] : ($this->styles[$index] ?? []);
    }

    /** @return array<int,array<string,mixed>> */
    public function allStyles(): array
    {
        return $this->styles;
    }

    /** @param array<int,string> $palette @return array<string,mixed> */
    private static function parseFont(string $payload, array $palette): array
    {
        $options = Binary::u16($payload, 2);
        $colorIndex = Binary::u16($payload, 4);
        $weight = Binary::u16($payload, 6);
        $vertical = Binary::u16($payload, 8);
        $underline = Binary::u8($payload, 10);
        $name = BiffString::readUnicodeString($payload, 14, true)['value'];
        $font = [
            'name' => $name,
            'size' => Binary::u16($payload, 0) / 20,
        ];
        if ($weight >= 700) $font['bold'] = true;
        if (($options & 0x0002) !== 0) $font['italic'] = true;
        if (($options & 0x0008) !== 0) $font['strike'] = true;
        if (($options & 0x0010) !== 0) $font['outline'] = true;
        if (($options & 0x0020) !== 0) $font['shadow'] = true;
        if ($vertical === 1) $font['vertical_align'] = 'superscript';
        if ($vertical === 2) $font['vertical_align'] = 'subscript';
        if ($underline !== 0) {
            $font['underline'] = match ($underline) { 2 => 'double', 0x21 => 'single_accounting', 0x22 => 'double_accounting', default => 'single' };
        }
        $color = self::color($colorIndex, $palette);
        if ($color !== null) $font['color'] = ['rgb' => 'FF' . $color];
        return $font;
    }

    /** @param array<int,array<string,mixed>> $fonts @param array<int,string> $palette @param array<int,string> $formats @return array<string,mixed> */
    private static function parseXf(string $payload, array $fonts, array $palette, array $formats): array
    {
        $fontIndex = Binary::u16($payload, 0);
        $formatId = Binary::u16($payload, 2);
        $protection = Binary::u16($payload, 4);
        $align = Binary::u8($payload, 6);
        $rotation = Binary::u8($payload, 7);
        $text = Binary::u8($payload, 8);
        $border1 = Binary::u32($payload, 10);
        $border2 = Binary::u32($payload, 14);
        $patternColors = Binary::u16($payload, 18);
        $style = [];
        if (isset($fonts[$fontIndex])) $style['font'] = $fonts[$fontIndex];
        $format = $formats[$formatId] ?? self::builtinFormat($formatId);
        if ($format !== null && $format !== 'General') $style['format'] = $format;

        $horizontal = $align & 0x07;
        $vertical = ($align >> 4) & 0x07;
        $alignment = [];
        $horizontalName = [1=>'left',2=>'center',3=>'right',4=>'fill',5=>'justify',6=>'center_continuous',7=>'distributed'][$horizontal] ?? null;
        if ($horizontalName !== null) $alignment['horizontal'] = $horizontalName;
        $verticalName = [0=>'top',1=>'center',2=>'bottom',3=>'justify',4=>'distributed'][$vertical] ?? null;
        if ($verticalName !== null && $verticalName !== 'bottom') $alignment['vertical'] = $verticalName;
        if (($align & 0x08) !== 0) $alignment['wrap_text'] = true;
        if ($rotation !== 0) $alignment['text_rotation'] = $rotation;
        if (($text & 0x0F) !== 0) $alignment['indent'] = $text & 0x0F;
        if (($text & 0x10) !== 0) $alignment['shrink_to_fit'] = true;
        if (($text >> 6) !== 0) $alignment['reading_order'] = $text >> 6;
        if ($alignment !== []) $style['alignment'] = $alignment;

        $border = [];
        $styles = [
            'left' => $border1 & 0x0F,
            'right' => ($border1 >> 4) & 0x0F,
            'top' => ($border1 >> 8) & 0x0F,
            'bottom' => ($border1 >> 12) & 0x0F,
            'diagonal' => ($border2 >> 21) & 0x0F,
        ];
        $colors = [
            'left' => ($border1 >> 16) & 0x7F,
            'right' => ($border1 >> 23) & 0x7F,
            'top' => $border2 & 0x7F,
            'bottom' => ($border2 >> 7) & 0x7F,
            'diagonal' => ($border2 >> 14) & 0x7F,
        ];
        foreach ($styles as $side => $lineStyle) {
            if ($lineStyle === 0) continue;
            $item = ['style' => self::borderStyle($lineStyle)];
            $color = self::color($colors[$side], $palette);
            if ($color !== null) $item['color'] = ['rgb' => 'FF' . $color];
            $border[$side] = $item;
        }
        if (($border1 & 0x40000000) !== 0) $border['diagonal_down'] = true;
        if (($border1 & 0x80000000) !== 0) $border['diagonal_up'] = true;
        if ($border !== []) $style['border'] = $border;

        $pattern = ($border2 >> 26) & 0x3F;
        if ($pattern !== 0) {
            $fill = ['type' => 'pattern', 'pattern' => self::fillPattern($pattern)];
            $foreground = self::color($patternColors & 0x7F, $palette);
            $background = self::color(($patternColors >> 7) & 0x7F, $palette);
            if ($foreground !== null) $fill['foreground'] = ['rgb' => 'FF' . $foreground];
            if ($background !== null) $fill['background'] = ['rgb' => 'FF' . $background];
            $style['fill'] = $fill;
        }
        $locked = ($protection & 0x0001) !== 0;
        $hidden = ($protection & 0x0002) !== 0;
        if (!$locked || $hidden) $style['protection'] = ['locked' => $locked, 'hidden' => $hidden];
        return StyleNormalizer::normalize($style);
    }

    private static function color(int $index, array $palette): ?string
    {
        if ($index === 0x7FFF || $index === 0x40 || $index === 0x41) return null;
        return $palette[$index] ?? null;
    }

    /** @return array<int,string> */
    private static function defaultPalette(): array
    {
        return [
            8=>'000000',9=>'FFFFFF',10=>'FF0000',11=>'00FF00',12=>'0000FF',13=>'FFFF00',14=>'FF00FF',15=>'00FFFF',
            16=>'800000',17=>'008000',18=>'000080',19=>'808000',20=>'800080',21=>'008080',22=>'C0C0C0',23=>'808080',
            24=>'9999FF',25=>'993366',26=>'FFFFCC',27=>'CCFFFF',28=>'660066',29=>'FF8080',30=>'0066CC',31=>'CCCCFF',
            32=>'000080',33=>'FF00FF',34=>'FFFF00',35=>'00FFFF',36=>'800080',37=>'800000',38=>'008080',39=>'0000FF',
            40=>'00CCFF',41=>'CCFFFF',42=>'CCFFCC',43=>'FFFF99',44=>'99CCFF',45=>'FF99CC',46=>'CC99FF',47=>'FFCC99',
            48=>'3366FF',49=>'33CCCC',50=>'99CC00',51=>'FFCC00',52=>'FF9900',53=>'FF6600',54=>'666699',55=>'969696',
            56=>'003366',57=>'339966',58=>'003300',59=>'333300',60=>'993300',61=>'993366',62=>'333399',63=>'333333',
        ];
    }

    private static function builtinFormat(int $id): ?string
    {
        return [0=>'General',1=>'0',2=>'0.00',3=>'#,##0',4=>'#,##0.00',9=>'0%',10=>'0.00%',11=>'0.00E+00',12=>'# ?/?',13=>'# ??/??',14=>'m/d/yy',15=>'d-mmm-yy',16=>'d-mmm',17=>'mmm-yy',18=>'h:mm AM/PM',19=>'h:mm:ss AM/PM',20=>'h:mm',21=>'h:mm:ss',22=>'m/d/yy h:mm',45=>'mm:ss',46=>'[h]:mm:ss',47=>'mmss.0',49=>'@'][$id] ?? null;
    }

    private static function borderStyle(int $id): string
    {
        return [1=>'thin',2=>'medium',3=>'dashed',4=>'dotted',5=>'thick',6=>'double',7=>'hair',8=>'medium_dashed',9=>'dash_dot',10=>'medium_dash_dot',11=>'dash_dot_dot',12=>'medium_dash_dot_dot',13=>'slant_dash_dot'][$id] ?? 'none';
    }

    private static function fillPattern(int $id): string
    {
        return [1=>'solid',2=>'medium_gray',3=>'dark_gray',4=>'light_gray',5=>'dark_horizontal',6=>'dark_vertical',7=>'dark_down',8=>'dark_up',9=>'dark_grid',10=>'dark_trellis',11=>'light_horizontal',12=>'light_vertical',13=>'light_down',14=>'light_up',15=>'light_grid',16=>'light_trellis',17=>'gray125',18=>'gray0625'][$id] ?? 'none';
    }
}
