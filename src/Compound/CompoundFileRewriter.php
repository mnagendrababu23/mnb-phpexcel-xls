<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Compound;

use Mnb\PHPExcel\Support\Binary;

/** Rebuilds a CFB v3 container while preserving its directory hierarchy and streams. */
final class CompoundFileRewriter
{
    private const SIGNATURE = "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1";
    private const FREESECT = 0xFFFFFFFF;
    private const ENDOFCHAIN = 0xFFFFFFFE;
    private const FATSECT = 0xFFFFFFFD;
    private const DIFSECT = 0xFFFFFFFC;
    private const NOSTREAM = 0xFFFFFFFF;
    private const SECTOR_SIZE = 512;
    private const MINI_SECTOR_SIZE = 64;
    private const MINI_CUTOFF = 4096;
    private const FAT_ENTRIES = 128;
    private const HEADER_DIFAT = 109;
    private const DIFAT_ENTRIES = 127;

    private function __construct()
    {
    }

    /**
     * Creates a new compound file whose streams are direct children of the root.
     *
     * @param array<string,string> $streams
     */
    public static function buildRootStreams(array $streams): string
    {
        if ($streams === []) {
            throw new \InvalidArgumentException('At least one CFB stream is required.');
        }
        $root = new DirectoryEntry(
            0,
            'Root Entry',
            DirectoryEntry::TYPE_ROOT,
            self::NOSTREAM,
            self::NOSTREAM,
            self::NOSTREAM,
            self::ENDOFCHAIN,
            0,
        );
        [$entries, $data] = self::upsertRootStreams([$root], [], $streams);
        return self::build($entries, $data);
    }

    /**
     * Adds or replaces root-level streams while preserving every existing
     * directory entry and unrelated stream.
     *
     * @param list<DirectoryEntry> $entries
     * @param array<int,string> $streamDataByEntryId
     * @param array<string,string> $streams
     * @return array{0:list<DirectoryEntry>,1:array<int,string>}
     */
    public static function upsertRootStreams(array $entries, array $streamDataByEntryId, array $streams): array
    {
        if ($entries === []) {
            throw new \InvalidArgumentException('CFB directory entries cannot be empty.');
        }
        $byId = [];
        $rootId = null;
        $maxId = -1;
        foreach ($entries as $entry) {
            $byId[$entry->id] = $entry;
            $maxId = max($maxId, $entry->id);
            if ($entry->type === DirectoryEntry::TYPE_ROOT) {
                $rootId = $entry->id;
            }
        }
        if ($rootId === null) {
            throw new \InvalidArgumentException('CFB root directory entry is missing.');
        }

        $rootIds = [];
        self::collectSiblingTreeIds($byId[$rootId]->childId, $byId, $rootIds, []);
        $rootSet = array_fill_keys($rootIds, true);
        $streamByName = [];
        foreach ($byId as $id => $entry) {
            if ($entry->type === DirectoryEntry::TYPE_STREAM) {
                $streamByName[self::directoryNameKey($entry->name)] = $id;
            }
        }

        $addedRootStream = false;
        foreach ($streams as $name => $data) {
            if (!is_string($name) || $name === '' || !is_string($data)) {
                throw new \InvalidArgumentException('CFB stream names and values must be non-empty strings.');
            }
            $encoded = iconv('UTF-8', 'UTF-16LE', $name);
            if (!is_string($encoded) || strlen($encoded) > 62) {
                throw new \InvalidArgumentException('CFB stream name must fit in 31 UTF-16 code units.');
            }
            $key = self::directoryNameKey($name);
            $id = $streamByName[$key] ?? null;
            if ($id === null) {
                $id = ++$maxId;
                $byId[$id] = new DirectoryEntry(
                    $id,
                    $name,
                    DirectoryEntry::TYPE_STREAM,
                    self::NOSTREAM,
                    self::NOSTREAM,
                    self::NOSTREAM,
                    self::ENDOFCHAIN,
                    strlen($data),
                    '',
                    1,
                );
                $rootSet[$id] = true;
                $streamByName[$key] = $id;
                $addedRootStream = true;
            }
            $streamDataByEntryId[$id] = $data;
        }

        if ($addedRootStream) {
            $rootIds = array_map('intval', array_keys($rootSet));
            usort($rootIds, static function (int $left, int $right) use ($byId): int {
                $comparison = self::directoryNameKey($byId[$left]->name) <=> self::directoryNameKey($byId[$right]->name);
                return $comparison !== 0 ? $comparison : ($left <=> $right);
            });
            $redLevel = self::redLevel(count($rootIds));
            $rootChild = self::buildSiblingTree($rootIds, 0, count($rootIds) - 1, 0, $redLevel, $byId);
            $root = $byId[$rootId];
            $byId[$rootId] = self::copyEntry($root, child: $rootChild, color: 1);
        }
        ksort($byId);
        return [array_values($byId), $streamDataByEntryId];
    }

    /**
     * @param list<DirectoryEntry> $entries
     * @param array<int,string> $streamDataByEntryId
     */
    public static function build(array $entries, array $streamDataByEntryId): string
    {
        if ($entries === []) {
            throw new \InvalidArgumentException('CFB directory entries cannot be empty.');
        }
        $byId = [];
        $rootId = null;
        $maxId = 0;
        foreach ($entries as $entry) {
            $byId[$entry->id] = $entry;
            $maxId = max($maxId, $entry->id);
            if ($entry->type === DirectoryEntry::TYPE_ROOT) {
                $rootId = $entry->id;
            }
        }
        if ($rootId === null) {
            throw new \InvalidArgumentException('CFB root directory entry is missing.');
        }

        $miniFat = [];
        $miniStream = '';
        $miniStarts = [];
        $largeStreams = [];
        foreach ($entries as $entry) {
            if ($entry->type !== DirectoryEntry::TYPE_STREAM) {
                continue;
            }
            $data = $streamDataByEntryId[$entry->id] ?? '';
            $length = strlen($data);
            if ($length > 0 && $length < self::MINI_CUTOFF) {
                $count = (int) ceil($length / self::MINI_SECTOR_SIZE);
                $start = count($miniFat);
                $miniStarts[$entry->id] = $start;
                for ($i = 0; $i < $count; $i++) {
                    $miniFat[] = $i === $count - 1 ? self::ENDOFCHAIN : $start + $i + 1;
                }
                $miniStream .= str_pad($data, $count * self::MINI_SECTOR_SIZE, "\0");
            } else {
                $largeStreams[$entry->id] = $data;
            }
        }

        $largeSectorCounts = [];
        $dataSectorCount = 0;
        foreach ($largeStreams as $id => $data) {
            $count = $data === '' ? 0 : (int) ceil(strlen($data) / self::SECTOR_SIZE);
            $largeSectorCounts[$id] = $count;
            $dataSectorCount += $count;
        }
        $rootMiniSectors = $miniStream === '' ? 0 : (int) ceil(strlen($miniStream) / self::SECTOR_SIZE);
        $miniFatBytes = '';
        foreach ($miniFat as $value) {
            $miniFatBytes .= pack('V', $value);
        }
        $miniFatSectors = $miniFatBytes === '' ? 0 : (int) ceil(strlen($miniFatBytes) / self::SECTOR_SIZE);
        $directorySectors = (int) ceil((($maxId + 1) * 128) / self::SECTOR_SIZE);
        $baseSectors = $dataSectorCount + $rootMiniSectors + $miniFatSectors + $directorySectors;
        [$fatSectors, $difatSectors] = self::allocationCounts($baseSectors);

        $nextSector = 0;
        $largeStarts = [];
        foreach ($largeStreams as $id => $data) {
            $count = $largeSectorCounts[$id];
            $largeStarts[$id] = $count > 0 ? $nextSector : self::ENDOFCHAIN;
            $nextSector += $count;
        }
        $rootMiniStart = $rootMiniSectors > 0 ? $nextSector : self::ENDOFCHAIN;
        $nextSector += $rootMiniSectors;
        $firstMiniFat = $miniFatSectors > 0 ? $nextSector : self::ENDOFCHAIN;
        $nextSector += $miniFatSectors;
        $firstDirectory = $nextSector;
        $nextSector += $directorySectors;
        $firstFat = $nextSector;
        $nextSector += $fatSectors;
        $firstDifat = $difatSectors > 0 ? $nextSector : self::ENDOFCHAIN;
        $nextSector += $difatSectors;
        $totalSectors = $nextSector;

        $fat = array_fill(0, $fatSectors * self::FAT_ENTRIES, self::FREESECT);
        foreach ($largeStreams as $id => $data) {
            self::chain($fat, $largeStarts[$id], $largeSectorCounts[$id]);
        }
        self::chain($fat, $rootMiniStart, $rootMiniSectors);
        self::chain($fat, $firstMiniFat, $miniFatSectors);
        self::chain($fat, $firstDirectory, $directorySectors);
        for ($i = 0; $i < $fatSectors; $i++) {
            $fat[$firstFat + $i] = self::FATSECT;
        }
        for ($i = 0; $i < $difatSectors; $i++) {
            $fat[$firstDifat + $i] = self::DIFSECT;
        }

        $directory = '';
        for ($id = 0; $id <= $maxId; $id++) {
            if (!isset($byId[$id])) {
                $directory .= str_repeat("\0", 128);
                continue;
            }
            $entry = $byId[$id];
            if ($entry->type === DirectoryEntry::TYPE_ROOT) {
                $start = $rootMiniStart;
                $size = strlen($miniStream);
            } elseif ($entry->type === DirectoryEntry::TYPE_STREAM) {
                $data = $streamDataByEntryId[$id] ?? '';
                $size = strlen($data);
                $start = $size === 0 ? self::ENDOFCHAIN : ($size < self::MINI_CUTOFF ? $miniStarts[$id] : $largeStarts[$id]);
            } else {
                $start = self::ENDOFCHAIN;
                $size = 0;
            }
            $directory .= self::serializeDirectoryEntry($entry, $start, $size);
        }
        $directory = str_pad($directory, $directorySectors * self::SECTOR_SIZE, "\0");

        $fatSectorIds = [];
        for ($i = 0; $i < $fatSectors; $i++) {
            $fatSectorIds[] = $firstFat + $i;
        }
        $header = self::header($fatSectorIds, $firstDirectory, $firstMiniFat, $miniFatSectors, $firstDifat, $difatSectors);
        $body = '';
        foreach ($largeStreams as $id => $data) {
            $body .= str_pad($data, $largeSectorCounts[$id] * self::SECTOR_SIZE, "\0");
        }
        $body .= str_pad($miniStream, $rootMiniSectors * self::SECTOR_SIZE, "\0");
        $body .= str_pad($miniFatBytes, $miniFatSectors * self::SECTOR_SIZE, pack('V', self::FREESECT));
        $body .= $directory;
        foreach ($fat as $value) {
            $body .= pack('V', $value);
        }
        $body .= self::difat($fatSectorIds, $firstDifat, $difatSectors);

        $file = $header . $body;
        $expected = 512 + ($totalSectors * self::SECTOR_SIZE);
        if (strlen($file) !== $expected) {
            throw new \RuntimeException('Internal CFB rewrite size calculation failed.');
        }
        return $file;
    }

    /** @param list<int> $fat */
    private static function chain(array &$fat, int $start, int $count): void
    {
        if ($count === 0 || $start === self::ENDOFCHAIN) {
            return;
        }
        for ($i = 0; $i < $count; $i++) {
            $fat[$start + $i] = $i === $count - 1 ? self::ENDOFCHAIN : $start + $i + 1;
        }
    }

    /** @return array{0:int,1:int} */
    private static function allocationCounts(int $baseSectors): array
    {
        $fat = 1;
        $difat = 0;
        do {
            $oldFat = $fat;
            $oldDifat = $difat;
            $fat = (int) ceil(($baseSectors + $fat + $difat) / self::FAT_ENTRIES);
            $difat = $fat > self::HEADER_DIFAT ? (int) ceil(($fat - self::HEADER_DIFAT) / self::DIFAT_ENTRIES) : 0;
        } while ($fat !== $oldFat || $difat !== $oldDifat);
        return [$fat, $difat];
    }

    /** @param list<int> $fatSectorIds */
    private static function header(array $fatSectorIds, int $directory, int $miniFat, int $miniFatCount, int $difat, int $difatCount): string
    {
        $h = self::SIGNATURE . str_repeat("\0", 16);
        $h .= pack('v', 0x003E) . pack('v', 3) . pack('v', 0xFFFE) . pack('v', 9) . pack('v', 6);
        $h .= str_repeat("\0", 6) . pack('V', 0) . pack('V', count($fatSectorIds)) . pack('V', $directory);
        $h .= pack('V', 0) . pack('V', self::MINI_CUTOFF) . pack('V', $miniFat) . pack('V', $miniFatCount);
        $h .= pack('V', $difat) . pack('V', $difatCount);
        for ($i = 0; $i < self::HEADER_DIFAT; $i++) {
            $h .= pack('V', $fatSectorIds[$i] ?? self::FREESECT);
        }
        return str_pad($h, 512, "\0");
    }

    /** @param list<int> $fatSectorIds */
    private static function difat(array $fatSectorIds, int $firstDifat, int $count): string
    {
        if ($count === 0) {
            return '';
        }
        $remaining = array_slice($fatSectorIds, self::HEADER_DIFAT);
        $out = '';
        for ($i = 0; $i < $count; $i++) {
            $chunk = array_splice($remaining, 0, self::DIFAT_ENTRIES);
            while (count($chunk) < self::DIFAT_ENTRIES) {
                $chunk[] = self::FREESECT;
            }
            foreach ($chunk as $value) {
                $out .= pack('V', $value);
            }
            $out .= pack('V', $i === $count - 1 ? self::ENDOFCHAIN : $firstDifat + $i + 1);
        }
        return $out;
    }

    /** @param array<int,DirectoryEntry> $byId @param array<int,true> $seen @param list<int> $ids */
    private static function collectSiblingTreeIds(int $id, array $byId, array &$ids, array $seen): void
    {
        $stack = [$id];
        while ($stack !== []) {
            $current = array_pop($stack);
            if ($current === self::NOSTREAM || !isset($byId[$current]) || isset($seen[$current])) {
                continue;
            }
            $seen[$current] = true;
            $entry = $byId[$current];
            $ids[] = $current;
            $stack[] = $entry->leftSiblingId;
            $stack[] = $entry->rightSiblingId;
        }
    }

    /** @param list<int> $ids @param array<int,DirectoryEntry> $byId */
    private static function buildSiblingTree(array $ids, int $low, int $high, int $level, int $redLevel, array &$byId): int
    {
        if ($high < $low) {
            return self::NOSTREAM;
        }
        $middle = intdiv($low + $high, 2);
        $id = $ids[$middle];
        $left = self::buildSiblingTree($ids, $low, $middle - 1, $level + 1, $redLevel, $byId);
        $right = self::buildSiblingTree($ids, $middle + 1, $high, $level + 1, $redLevel, $byId);
        $entry = $byId[$id];
        $byId[$id] = self::copyEntry($entry, left: $left, right: $right, color: $level === $redLevel ? 0 : 1);
        return $id;
    }

    private static function redLevel(int $size): int
    {
        $level = 0;
        for ($value = $size - 1; $value >= 0; $value = intdiv($value, 2) - 1) {
            $level++;
        }
        return $level;
    }

    private static function directoryNameKey(string $name): string
    {
        return strtoupper($name);
    }

    private static function copyEntry(
        DirectoryEntry $entry,
        ?int $left = null,
        ?int $right = null,
        ?int $child = null,
        ?int $color = null,
    ): DirectoryEntry {
        return new DirectoryEntry(
            $entry->id,
            $entry->name,
            $entry->type,
            $left ?? $entry->leftSiblingId,
            $right ?? $entry->rightSiblingId,
            $child ?? $entry->childId,
            $entry->startSector,
            $entry->streamSize,
            $entry->rawBytes,
            $color ?? $entry->color,
        );
    }

    private static function serializeDirectoryEntry(DirectoryEntry $entry, int $start, int $size): string
    {
        if (strlen($entry->rawBytes) === 128) {
            $raw = $entry->rawBytes;
            $raw[66] = chr($entry->type);
            $raw[67] = chr($entry->color);
            $raw = substr_replace($raw, pack('V3', $entry->leftSiblingId, $entry->rightSiblingId, $entry->childId), 68, 12);
            $raw = substr_replace($raw, pack('V', $start), 116, 4);
            return substr_replace($raw, Binary::packU64($size), 120, 8);
        }
        $name = iconv('UTF-8', 'UTF-16LE', $entry->name);
        if (!is_string($name)) {
            throw new \RuntimeException('Unable to encode CFB directory name.');
        }
        $name .= "\0\0";
        $raw = str_pad($name, 64, "\0") . pack('v', strlen($name));
        $raw .= chr($entry->type) . chr($entry->color);
        $raw .= pack('V3', $entry->leftSiblingId, $entry->rightSiblingId, $entry->childId);
        $raw .= str_repeat("\0", 16) . pack('V', 0) . str_repeat("\0", 16);
        $raw .= pack('V', $start) . Binary::packU64($size);
        return $raw;
    }
}
