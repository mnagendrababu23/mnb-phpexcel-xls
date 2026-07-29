<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Writer;

use Mnb\PHPExcel\Biff\BiffRecordWriter;
use Mnb\PHPExcel\Biff\RecordType;

final class SharedStringWriter
{
    /** @var array<string,int> */
    private array $indexes = [];
    /** @var list<string> */
    private array $strings = [];
    private int $totalReferences = 0;

    public function add(string $value): int
    {
        $this->totalReferences++;
        $key = strlen($value) . ':' . $value;
        if (isset($this->indexes[$key])) {
            return $this->indexes[$key];
        }
        $index = count($this->strings);
        $this->indexes[$key] = $index;
        $this->strings[] = $value;
        return $index;
    }

    public function index(string $value): int
    {
        $key = strlen($value) . ':' . $value;
        if (!isset($this->indexes[$key])) {
            throw new \LogicException('String was not registered in the shared string table.');
        }
        return $this->indexes[$key];
    }

    public function records(): string
    {
        $segments = [pack('V2', $this->totalReferences, count($this->strings))];
        foreach ($this->strings as $value) {
            $encoded = iconv('UTF-8', 'UTF-16LE', $value);
            if ($encoded === false) {
                throw new \InvalidArgumentException('Unable to encode shared string as UTF-16LE.');
            }
            $characterCount = intdiv(strlen($encoded), 2);
            if ($characterCount > 65535) {
                throw new \InvalidArgumentException('XLS shared string exceeds 65535 characters.');
            }
            $this->appendBytes($segments, pack('vC', $characterCount, 0x01));
            $remaining = $encoded;
            while ($remaining !== '') {
                $last = count($segments) - 1;
                $room = 8224 - strlen($segments[$last]);
                if ($room < 2) {
                    $segments[] = "\x01"; // continuation string option byte
                    $last++;
                    $room = 8223;
                }
                $take = min(strlen($remaining), $room - ($room % 2));
                if ($take <= 0) {
                    $segments[] = "\x01";
                    continue;
                }
                $segments[$last] .= substr($remaining, 0, $take);
                $remaining = substr($remaining, $take);
                if ($remaining !== '') {
                    $segments[] = "\x01";
                }
            }
        }

        $result = '';
        foreach ($segments as $index => $payload) {
            $result .= BiffRecordWriter::record($index === 0 ? RecordType::SST : RecordType::CONTINUE, $payload);
        }
        return $result;
    }

    /** @param list<string> $segments */
    private function appendBytes(array &$segments, string $bytes): void
    {
        while ($bytes !== '') {
            $last = count($segments) - 1;
            $room = 8224 - strlen($segments[$last]);
            if ($room === 0) {
                $segments[] = '';
                continue;
            }
            $take = min(strlen($bytes), $room);
            $segments[$last] .= substr($bytes, 0, $take);
            $bytes = substr($bytes, $take);
        }
    }
}
