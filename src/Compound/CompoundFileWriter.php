<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Compound;

use Mnb\PHPExcel\Support\Binary;

/** Builds a version-3 CFB container with one normal-size workbook stream. */
final class CompoundFileWriter
{
    private const SIGNATURE = "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1";
    private const FREESECT = 0xFFFFFFFF;
    private const ENDOFCHAIN = 0xFFFFFFFE;
    private const FATSECT = 0xFFFFFFFD;
    private const DIFSECT = 0xFFFFFFFC;
    private const NOSTREAM = 0xFFFFFFFF;
    private const SECTOR_SIZE = 512;
    private const FAT_ENTRIES_PER_SECTOR = 128;
    private const HEADER_DIFAT_ENTRIES = 109;
    private const DIFAT_ENTRIES_PER_SECTOR = 127;

    private function __construct()
    {
    }

    public static function build(string $streamName, string $streamData): string
    {
        return self::buildStreams([$streamName => $streamData]);
    }

    /** @param array<string,string> $streams */
    public static function buildStreams(array $streams): string
    {
        return CompoundFileRewriter::buildRootStreams($streams);
    }

    /** @return array{0:int,1:int} */
    private static function allocationSectorCounts(int $dataAndDirectorySectors): array
    {
        $fatSectors = 1;
        $difatSectors = 0;
        do {
            $previousFat = $fatSectors;
            $previousDifat = $difatSectors;
            $fatSectors = (int) ceil(($dataAndDirectorySectors + $fatSectors + $difatSectors) / self::FAT_ENTRIES_PER_SECTOR);
            $difatSectors = $fatSectors > self::HEADER_DIFAT_ENTRIES
                ? (int) ceil(($fatSectors - self::HEADER_DIFAT_ENTRIES) / self::DIFAT_ENTRIES_PER_SECTOR)
                : 0;
        } while ($fatSectors !== $previousFat || $difatSectors !== $previousDifat);

        return [$fatSectors, $difatSectors];
    }

    /** @param list<int> $fatSectorIds */
    private static function header(array $fatSectorIds, int $directorySectorId, int $firstDifatSectorId, int $difatSectors): string
    {
        $header = self::SIGNATURE;
        $header .= str_repeat("\0", 16); // CLSID
        $header .= pack('v', 0x003E); // minor version
        $header .= pack('v', 0x0003); // major version (512-byte sectors)
        $header .= pack('v', 0xFFFE); // little endian
        $header .= pack('v', 9); // sector shift
        $header .= pack('v', 6); // mini sector shift
        $header .= str_repeat("\0", 6);
        $header .= pack('V', 0); // directory sectors for v3
        $header .= pack('V', count($fatSectorIds));
        $header .= pack('V', $directorySectorId);
        $header .= pack('V', 0); // transaction signature
        $header .= pack('V', 4096);
        $header .= pack('V', self::ENDOFCHAIN);
        $header .= pack('V', 0);
        $header .= pack('V', $firstDifatSectorId);
        $header .= pack('V', $difatSectors);
        for ($i = 0; $i < self::HEADER_DIFAT_ENTRIES; $i++) {
            $header .= pack('V', $fatSectorIds[$i] ?? self::FREESECT);
        }
        return str_pad($header, 512, "\0");
    }

    /** @param list<int> $fatSectorIds */
    private static function difatSectors(array $fatSectorIds, int $firstDifatSectorId, int $difatSectorCount): string
    {
        if ($difatSectorCount === 0) {
            return '';
        }
        $remaining = array_slice($fatSectorIds, self::HEADER_DIFAT_ENTRIES);
        $bytes = '';
        for ($i = 0; $i < $difatSectorCount; $i++) {
            $entries = array_splice($remaining, 0, self::DIFAT_ENTRIES_PER_SECTOR);
            while (count($entries) < self::DIFAT_ENTRIES_PER_SECTOR) {
                $entries[] = self::FREESECT;
            }
            foreach ($entries as $entry) {
                $bytes .= pack('V', $entry);
            }
            $bytes .= pack('V', $i === $difatSectorCount - 1 ? self::ENDOFCHAIN : $firstDifatSectorId + $i + 1);
        }
        if ($remaining !== []) {
            throw new \LogicException('DIFAT sector count did not cover all FAT sector IDs.');
        }
        return $bytes;
    }

    private static function directory(string $streamName, int $streamStartSector, int $streamSize): string
    {
        $root = self::directoryEntry('Root Entry', DirectoryEntry::TYPE_ROOT, self::NOSTREAM, self::NOSTREAM, 1, self::ENDOFCHAIN, 0);
        $workbook = self::directoryEntry($streamName, DirectoryEntry::TYPE_STREAM, self::NOSTREAM, self::NOSTREAM, self::NOSTREAM, $streamStartSector, $streamSize);
        return str_pad($root . $workbook, self::SECTOR_SIZE, "\0");
    }

    private static function directoryEntry(string $name, int $type, int $left, int $right, int $child, int $startSector, int $streamSize): string
    {
        $encodedName = iconv('UTF-8', 'UTF-16LE', $name);
        if ($encodedName === false) {
            throw new \RuntimeException('Unable to encode CFB directory name.');
        }
        $encodedName .= "\0\0";
        $entry = str_pad($encodedName, 64, "\0");
        $entry .= pack('v', strlen($encodedName));
        $entry .= chr($type);
        $entry .= chr(1); // black node
        $entry .= pack('V3', $left, $right, $child);
        $entry .= str_repeat("\0", 16); // CLSID
        $entry .= pack('V', 0); // state bits
        $entry .= str_repeat("\0", 16); // timestamps
        $entry .= pack('V', $startSector);
        $entry .= Binary::packU64($streamSize);
        return $entry;
    }

    private static function align(int $value, int $alignment): int
    {
        return (int) (ceil($value / $alignment) * $alignment);
    }
}
