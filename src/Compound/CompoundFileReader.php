<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Compound;

use Mnb\PHPExcel\Exception\InvalidCompoundFileException;
use Mnb\PHPExcel\Support\Binary;

/**
 * Native reader for Microsoft Compound File Binary (CFB/OLE2) containers.
 *
 * It intentionally exposes streams only; BIFF parsing belongs to the XLS layer.
 */
final class CompoundFileReader
{
    private const SIGNATURE = "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1";
    private const FREESECT = 0xFFFFFFFF;
    private const ENDOFCHAIN = 0xFFFFFFFE;
    private const FATSECT = 0xFFFFFFFD;
    private const DIFSECT = 0xFFFFFFFC;
    private const NOSTREAM = 0xFFFFFFFF;

    /** @var resource */
    private $handle;
    private int $fileSize;
    private int $sectorSize;
    private int $sectorBaseOffset;
    private int $miniSectorSize;
    private int $miniStreamCutoff;
    private int $firstDirectorySector;
    private int $firstMiniFatSector;
    private int $miniFatSectorCount;
    private int $firstDifatSector;
    private int $difatSectorCount;
    /** @var list<int> */
    private array $fat = [];
    /** @var list<int> */
    private array $miniFat = [];
    /** @var list<DirectoryEntry> */
    private array $directoryEntries = [];
    /** @var array<int,string> */
    private array $directoryEntryBytes = [];
    private ?DirectoryEntry $rootEntry = null;
    private ?string $rootMiniStream = null;

    /** @param array<string,mixed> $options */
    public function __construct(private readonly string $path, private readonly array $options = [])
    {
        if (!is_file($path)) {
            throw InvalidCompoundFileException::because('Compound file does not exist.', ['path' => $path]);
        }
        $size = filesize($path);
        if ($size === false || $size < 512) {
            throw InvalidCompoundFileException::because('Compound file is too small.', ['path' => $path, 'size' => $size]);
        }
        $maxFileSize = (int) ($options['max_file_size'] ?? 512 * 1024 * 1024);
        if ($size > $maxFileSize) {
            throw InvalidCompoundFileException::because('Compound file exceeds the configured size limit.', ['size' => $size, 'limit' => $maxFileSize]);
        }
        $handle = @fopen($path, 'rb');
        if (!is_resource($handle)) {
            throw InvalidCompoundFileException::because('Unable to open compound file.', ['path' => $path]);
        }
        $this->handle = $handle;
        $this->fileSize = (int) $size;

        try {
            $this->parseHeaderAndAllocationTables();
            $this->parseDirectory();
        } catch (\Throwable $e) {
            fclose($this->handle);
            if ($e instanceof InvalidCompoundFileException) {
                throw $e;
            }
            throw InvalidCompoundFileException::because('Unable to parse compound file: ' . $e->getMessage(), ['path' => $path]);
        }
    }

    public function __destruct()
    {
        if (isset($this->handle) && is_resource($this->handle)) {
            fclose($this->handle);
        }
    }

    /** @return list<DirectoryEntry> */
    public function directoryEntries(): array
    {
        return $this->directoryEntries;
    }

    public function readStreamById(int $entryId): string
    {
        foreach ($this->directoryEntries as $entry) {
            if ($entry->id === $entryId && $entry->type === DirectoryEntry::TYPE_STREAM) {
                return $this->readEntryStream($entry);
            }
        }
        throw InvalidCompoundFileException::because('Compound stream entry was not found.', ['entry_id' => $entryId]);
    }

    /** @return list<array{name:string,size_bytes:int,type:int}> */
    public function streamInfo(): array
    {
        $items = [];
        foreach ($this->directoryEntries as $entry) {
            if ($entry->type === DirectoryEntry::TYPE_STREAM) {
                $items[] = [
                    'name' => $entry->name,
                    'size_bytes' => $entry->streamSize,
                    'type' => $entry->type,
                ];
            }
        }
        return $items;
    }

    /** @return list<string> */
    public function streamNames(): array
    {
        $names = [];
        foreach ($this->directoryEntries as $entry) {
            if ($entry->type === DirectoryEntry::TYPE_STREAM) {
                $names[] = $entry->name;
            }
        }
        return $names;
    }

    public function hasStream(string $name): bool
    {
        return $this->findStream($name) !== null;
    }

    public function readStream(string $name): string
    {
        $entry = $this->findStream($name);
        if ($entry === null) {
            throw InvalidCompoundFileException::because('Compound stream was not found.', ['stream' => $name, 'available' => $this->streamNames()]);
        }
        return $this->readEntryStream($entry);
    }

    private function parseHeaderAndAllocationTables(): void
    {
        $header = $this->readAt(0, 512);
        if (substr($header, 0, 8) !== self::SIGNATURE) {
            throw InvalidCompoundFileException::because('Invalid CFB signature.', ['path' => $this->path]);
        }
        $byteOrder = Binary::u16($header, 28);
        if ($byteOrder !== 0xFFFE) {
            throw InvalidCompoundFileException::because('Unsupported CFB byte order.', ['byte_order' => $byteOrder]);
        }
        $majorVersion = Binary::u16($header, 26);
        $sectorShift = Binary::u16($header, 30);
        $miniSectorShift = Binary::u16($header, 32);
        if (!in_array($majorVersion, [3, 4], true)) {
            throw InvalidCompoundFileException::because('Unsupported CFB major version.', ['version' => $majorVersion]);
        }
        $this->sectorSize = 1 << $sectorShift;
        $this->miniSectorSize = 1 << $miniSectorShift;
        if (!in_array($this->sectorSize, [512, 4096], true) || $this->miniSectorSize !== 64) {
            throw InvalidCompoundFileException::because('Unsupported CFB sector geometry.', [
                'sector_size' => $this->sectorSize,
                'mini_sector_size' => $this->miniSectorSize,
            ]);
        }
        $this->sectorBaseOffset = $this->sectorSize;
        if ($this->fileSize < $this->sectorBaseOffset || ($this->fileSize - $this->sectorBaseOffset) % $this->sectorSize !== 0) {
            throw InvalidCompoundFileException::because('CFB file length is not aligned to its sector size.', ['size' => $this->fileSize, 'sector_size' => $this->sectorSize]);
        }

        $fatSectorCount = Binary::u32($header, 44);
        $this->firstDirectorySector = Binary::u32($header, 48);
        $this->miniStreamCutoff = Binary::u32($header, 56);
        $this->firstMiniFatSector = Binary::u32($header, 60);
        $this->miniFatSectorCount = Binary::u32($header, 64);
        $this->firstDifatSector = Binary::u32($header, 68);
        $this->difatSectorCount = Binary::u32($header, 72);

        if ($this->miniStreamCutoff !== 4096) {
            throw InvalidCompoundFileException::because('Unsupported CFB mini stream cutoff.', ['cutoff' => $this->miniStreamCutoff]);
        }

        $fatSectorIds = [];
        for ($i = 0; $i < 109; $i++) {
            $sector = Binary::u32($header, 76 + ($i * 4));
            if ($sector !== self::FREESECT) {
                $fatSectorIds[] = $sector;
            }
        }

        $entriesPerDifatSector = intdiv($this->sectorSize, 4) - 1;
        $nextDifat = $this->firstDifatSector;
        $seenDifat = [];
        for ($i = 0; $i < $this->difatSectorCount; $i++) {
            if ($nextDifat === self::ENDOFCHAIN || $nextDifat === self::FREESECT) {
                throw InvalidCompoundFileException::because('DIFAT chain ended before the declared sector count.', ['index' => $i]);
            }
            $this->assertSectorId($nextDifat, 'DIFAT');
            if (isset($seenDifat[$nextDifat])) {
                throw InvalidCompoundFileException::because('Cyclic DIFAT chain detected.', ['sector' => $nextDifat]);
            }
            $seenDifat[$nextDifat] = true;
            $sectorData = $this->readSector($nextDifat);
            for ($j = 0; $j < $entriesPerDifatSector; $j++) {
                $fatSector = Binary::u32($sectorData, $j * 4);
                if ($fatSector !== self::FREESECT) {
                    $fatSectorIds[] = $fatSector;
                }
            }
            $nextDifat = Binary::u32($sectorData, $entriesPerDifatSector * 4);
        }

        if (count($fatSectorIds) < $fatSectorCount) {
            throw InvalidCompoundFileException::because('CFB header declares more FAT sectors than the DIFAT provides.', [
                'declared' => $fatSectorCount,
                'found' => count($fatSectorIds),
            ]);
        }
        $fatSectorIds = array_slice($fatSectorIds, 0, $fatSectorCount);
        $maxFatSectors = (int) ($this->options['max_fat_sectors'] ?? 65536);
        if (count($fatSectorIds) > $maxFatSectors) {
            throw InvalidCompoundFileException::because('FAT sector count exceeds the configured limit.', ['count' => count($fatSectorIds), 'limit' => $maxFatSectors]);
        }

        foreach ($fatSectorIds as $sectorId) {
            $this->assertSectorId($sectorId, 'FAT');
            $entries = array_values(unpack('V*', $this->readSector($sectorId)));
            foreach ($entries as $entry) {
                $this->fat[] = (int) $entry;
            }
        }

        if ($this->miniFatSectorCount > 0) {
            $miniFatBytes = $this->readSectorChain(
                $this->firstMiniFatSector,
                $this->fat,
                $this->sectorSize,
                'MiniFAT',
                $this->miniFatSectorCount
            );
            $this->miniFat = array_values(unpack('V*', substr($miniFatBytes, 0, $this->miniFatSectorCount * $this->sectorSize)));
        }
    }

    private function parseDirectory(): void
    {
        $directoryBytes = $this->readSectorChain($this->firstDirectorySector, $this->fat, $this->sectorSize, 'directory');
        $entryCount = intdiv(strlen($directoryBytes), 128);
        $maxEntries = (int) ($this->options['max_directory_entries'] ?? 65536);
        if ($entryCount > $maxEntries) {
            throw InvalidCompoundFileException::because('Directory entry count exceeds the configured limit.', ['count' => $entryCount, 'limit' => $maxEntries]);
        }

        for ($id = 0; $id < $entryCount; $id++) {
            $entry = substr($directoryBytes, $id * 128, 128);
            $this->directoryEntryBytes[$id] = $entry;
            $type = Binary::u8($entry, 66);
            if ($type === DirectoryEntry::TYPE_EMPTY) {
                continue;
            }
            $nameLength = Binary::u16($entry, 64);
            if ($nameLength < 2 || $nameLength > 64 || ($nameLength % 2) !== 0) {
                throw InvalidCompoundFileException::because('Invalid CFB directory name length.', ['entry_id' => $id, 'name_length' => $nameLength]);
            }
            $nameBytes = substr($entry, 0, $nameLength - 2);
            $name = iconv('UTF-16LE', 'UTF-8//IGNORE', $nameBytes);
            if ($name === false) {
                throw InvalidCompoundFileException::because('Unable to decode CFB directory name.', ['entry_id' => $id]);
            }
            $directoryEntry = new DirectoryEntry(
                $id,
                $name,
                $type,
                Binary::u32($entry, 68),
                Binary::u32($entry, 72),
                Binary::u32($entry, 76),
                Binary::u32($entry, 116),
                Binary::u64($entry, 120),
                $entry,
                Binary::u8($entry, 67),
            );
            $this->directoryEntries[] = $directoryEntry;
            if ($type === DirectoryEntry::TYPE_ROOT) {
                $this->rootEntry = $directoryEntry;
            }
        }
        if ($this->rootEntry === null) {
            throw InvalidCompoundFileException::because('CFB root directory entry is missing.');
        }
    }

    private function readEntryStream(DirectoryEntry $entry): string
    {
        $maxStreamSize = (int) ($this->options['max_stream_size'] ?? 256 * 1024 * 1024);
        if ($entry->streamSize > $maxStreamSize) {
            throw InvalidCompoundFileException::because('Compound stream exceeds the configured size limit.', [
                'stream' => $entry->name, 'size' => $entry->streamSize, 'limit' => $maxStreamSize,
            ]);
        }
        if ($entry->streamSize === 0) {
            return '';
        }
        if ($entry->streamSize < $this->miniStreamCutoff) {
            return $this->readMiniStream($entry);
        }
        return substr($this->readSectorChain($entry->startSector, $this->fat, $this->sectorSize, 'stream:' . $entry->name), 0, $entry->streamSize);
    }

    private function readMiniStream(DirectoryEntry $entry): string
    {
        if ($this->miniFat === []) {
            throw InvalidCompoundFileException::because('Small stream references a missing MiniFAT.', ['stream' => $entry->name]);
        }
        if ($this->rootMiniStream === null) {
            $root = $this->rootEntry;
            if ($root === null || $root->streamSize === 0 || $root->startSector === self::ENDOFCHAIN) {
                throw InvalidCompoundFileException::because('CFB root mini stream is missing.');
            }
            $this->rootMiniStream = substr(
                $this->readSectorChain($root->startSector, $this->fat, $this->sectorSize, 'root-mini-stream'),
                0,
                $root->streamSize
            );
        }

        $result = '';
        $sector = $entry->startSector;
        $seen = [];
        $maxMiniSectors = (int) ($this->options['max_chain_sectors'] ?? 2_000_000);
        while ($sector !== self::ENDOFCHAIN) {
            if ($sector === self::FREESECT || $sector >= count($this->miniFat)) {
                throw InvalidCompoundFileException::because('Invalid MiniFAT sector reference.', ['stream' => $entry->name, 'sector' => $sector]);
            }
            if (isset($seen[$sector])) {
                throw InvalidCompoundFileException::because('Cyclic MiniFAT chain detected.', ['stream' => $entry->name, 'sector' => $sector]);
            }
            if (count($seen) >= $maxMiniSectors) {
                throw InvalidCompoundFileException::because('MiniFAT chain exceeds the configured limit.', ['stream' => $entry->name, 'limit' => $maxMiniSectors]);
            }
            $seen[$sector] = true;
            $offset = $sector * $this->miniSectorSize;
            if ($offset + $this->miniSectorSize > strlen($this->rootMiniStream)) {
                throw InvalidCompoundFileException::because('Mini sector points beyond the root mini stream.', ['stream' => $entry->name, 'sector' => $sector]);
            }
            $result .= substr($this->rootMiniStream, $offset, $this->miniSectorSize);
            $sector = $this->miniFat[$sector];
        }
        return substr($result, 0, $entry->streamSize);
    }

    /** @param list<int> $allocationTable */
    private function readSectorChain(int $startSector, array $allocationTable, int $unitSize, string $label, ?int $declaredCount = null): string
    {
        if ($startSector === self::ENDOFCHAIN && ($declaredCount === null || $declaredCount === 0)) {
            return '';
        }
        $result = '';
        $sector = $startSector;
        $seen = [];
        $maxSectors = (int) ($this->options['max_chain_sectors'] ?? 2_000_000);
        while ($sector !== self::ENDOFCHAIN) {
            if ($sector === self::FREESECT || $sector === self::FATSECT || $sector === self::DIFSECT || $sector >= count($allocationTable)) {
                throw InvalidCompoundFileException::because('Invalid sector chain reference.', ['chain' => $label, 'sector' => $sector]);
            }
            $this->assertSectorId($sector, $label);
            if (isset($seen[$sector])) {
                throw InvalidCompoundFileException::because('Cyclic sector chain detected.', ['chain' => $label, 'sector' => $sector]);
            }
            if (count($seen) >= $maxSectors) {
                throw InvalidCompoundFileException::because('Sector chain exceeds the configured limit.', ['chain' => $label, 'limit' => $maxSectors]);
            }
            $seen[$sector] = true;
            $result .= $this->readSector($sector);
            $sector = $allocationTable[$sector];
            if ($declaredCount !== null && count($seen) >= $declaredCount) {
                break;
            }
        }
        if ($declaredCount !== null && count($seen) !== $declaredCount) {
            throw InvalidCompoundFileException::because('Sector chain length does not match the declared count.', [
                'chain' => $label,
                'declared' => $declaredCount,
                'actual' => count($seen),
            ]);
        }
        return $result;
    }

    private function findStream(string $name): ?DirectoryEntry
    {
        foreach ($this->directoryEntries as $entry) {
            if ($entry->type === DirectoryEntry::TYPE_STREAM && strcasecmp($entry->name, $name) === 0) {
                return $entry;
            }
        }
        return null;
    }

    private function assertSectorId(int $sectorId, string $label): void
    {
        $sectorCount = intdiv($this->fileSize - $this->sectorBaseOffset, $this->sectorSize);
        if ($sectorId < 0 || $sectorId >= $sectorCount) {
            throw InvalidCompoundFileException::because('Sector ID points beyond the file.', ['chain' => $label, 'sector' => $sectorId, 'sector_count' => $sectorCount]);
        }
    }

    private function readSector(int $sectorId): string
    {
        return $this->readAt($this->sectorBaseOffset + ($sectorId * $this->sectorSize), $this->sectorSize);
    }

    private function readAt(int $offset, int $length): string
    {
        if ($offset < 0 || $length < 0 || $offset + $length > $this->fileSize) {
            throw InvalidCompoundFileException::because('Binary read exceeds the compound file boundary.', ['offset' => $offset, 'length' => $length, 'file_size' => $this->fileSize]);
        }
        if (fseek($this->handle, $offset) !== 0) {
            throw InvalidCompoundFileException::because('Unable to seek within compound file.', ['offset' => $offset]);
        }
        $data = '';
        while (strlen($data) < $length && !feof($this->handle)) {
            $chunk = fread($this->handle, $length - strlen($data));
            if ($chunk === false) {
                throw InvalidCompoundFileException::because('Unable to read compound file.', ['offset' => $offset, 'length' => $length]);
            }
            if ($chunk === '') {
                break;
            }
            $data .= $chunk;
        }
        if (strlen($data) !== $length) {
            throw InvalidCompoundFileException::because('Compound file is truncated.', ['offset' => $offset, 'expected' => $length, 'actual' => strlen($data)]);
        }
        return $data;
    }
}
