<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Biff\Formula;

final class FunctionRegistry
{
    /** @var array<int,array{name:string,args:?int}> */
    private const BY_ID = [
        0 => ['name' => 'COUNT', 'args' => null],
        1 => ['name' => 'IF', 'args' => null],
        2 => ['name' => 'ISNA', 'args' => 1],
        3 => ['name' => 'ISERROR', 'args' => 1],
        4 => ['name' => 'SUM', 'args' => null],
        5 => ['name' => 'AVERAGE', 'args' => null],
        6 => ['name' => 'MIN', 'args' => null],
        7 => ['name' => 'MAX', 'args' => null],
        8 => ['name' => 'ROW', 'args' => null],
        9 => ['name' => 'COLUMN', 'args' => null],
        10 => ['name' => 'NA', 'args' => 0],
        11 => ['name' => 'NPV', 'args' => null],
        12 => ['name' => 'STDEV', 'args' => null],
        13 => ['name' => 'DOLLAR', 'args' => null],
        14 => ['name' => 'FIXED', 'args' => null],
        15 => ['name' => 'SIN', 'args' => 1],
        16 => ['name' => 'COS', 'args' => 1],
        17 => ['name' => 'TAN', 'args' => 1],
        18 => ['name' => 'ATAN', 'args' => 1],
        19 => ['name' => 'PI', 'args' => 0],
        20 => ['name' => 'SQRT', 'args' => 1],
        21 => ['name' => 'EXP', 'args' => 1],
        22 => ['name' => 'LN', 'args' => 1],
        23 => ['name' => 'LOG10', 'args' => 1],
        24 => ['name' => 'ABS', 'args' => 1],
        25 => ['name' => 'INT', 'args' => 1],
        26 => ['name' => 'SIGN', 'args' => 1],
        27 => ['name' => 'ROUND', 'args' => 2],
        28 => ['name' => 'LOOKUP', 'args' => null],
        29 => ['name' => 'INDEX', 'args' => null],
        30 => ['name' => 'REPT', 'args' => 2],
        31 => ['name' => 'MID', 'args' => 3],
        32 => ['name' => 'LEN', 'args' => 1],
        33 => ['name' => 'VALUE', 'args' => 1],
        34 => ['name' => 'TRUE', 'args' => 0],
        35 => ['name' => 'FALSE', 'args' => 0],
        36 => ['name' => 'AND', 'args' => null],
        37 => ['name' => 'OR', 'args' => null],
        38 => ['name' => 'NOT', 'args' => 1],
        39 => ['name' => 'MOD', 'args' => 2],
        56 => ['name' => 'PV', 'args' => null],
        57 => ['name' => 'FV', 'args' => null],
        58 => ['name' => 'NPER', 'args' => null],
        59 => ['name' => 'PMT', 'args' => null],
        60 => ['name' => 'RATE', 'args' => null],
        61 => ['name' => 'MIRR', 'args' => 3],
        62 => ['name' => 'IRR', 'args' => null],
        63 => ['name' => 'RAND', 'args' => 0],
        65 => ['name' => 'DATE', 'args' => 3],
        66 => ['name' => 'TIME', 'args' => 3],
        67 => ['name' => 'DAY', 'args' => 1],
        68 => ['name' => 'MONTH', 'args' => 1],
        69 => ['name' => 'YEAR', 'args' => 1],
        70 => ['name' => 'WEEKDAY', 'args' => null],
        71 => ['name' => 'HOUR', 'args' => 1],
        72 => ['name' => 'MINUTE', 'args' => 1],
        73 => ['name' => 'SECOND', 'args' => 1],
        74 => ['name' => 'NOW', 'args' => 0],
        97 => ['name' => 'ATAN2', 'args' => 2],
        98 => ['name' => 'ASIN', 'args' => 1],
        99 => ['name' => 'ACOS', 'args' => 1],
        100 => ['name' => 'CHOOSE', 'args' => null],
        101 => ['name' => 'HLOOKUP', 'args' => null],
        102 => ['name' => 'VLOOKUP', 'args' => null],
        105 => ['name' => 'ISREF', 'args' => 1],
        109 => ['name' => 'LOG', 'args' => null],
        111 => ['name' => 'CHAR', 'args' => 1],
        112 => ['name' => 'LOWER', 'args' => 1],
        113 => ['name' => 'UPPER', 'args' => 1],
        114 => ['name' => 'PROPER', 'args' => 1],
        115 => ['name' => 'LEFT', 'args' => null],
        116 => ['name' => 'RIGHT', 'args' => null],
        117 => ['name' => 'EXACT', 'args' => 2],
        118 => ['name' => 'TRIM', 'args' => 1],
        119 => ['name' => 'REPLACE', 'args' => 4],
        120 => ['name' => 'SUBSTITUTE', 'args' => null],
        121 => ['name' => 'CODE', 'args' => 1],
        124 => ['name' => 'FIND', 'args' => null],
        125 => ['name' => 'CELL', 'args' => null],
        126 => ['name' => 'ISERR', 'args' => 1],
        127 => ['name' => 'ISTEXT', 'args' => 1],
        128 => ['name' => 'ISNUMBER', 'args' => 1],
        129 => ['name' => 'ISBLANK', 'args' => 1],
        130 => ['name' => 'T', 'args' => 1],
        131 => ['name' => 'N', 'args' => 1],
        162 => ['name' => 'CLEAN', 'args' => 1],
        163 => ['name' => 'MDETERM', 'args' => 1],
        164 => ['name' => 'MINVERSE', 'args' => 1],
        165 => ['name' => 'MMULT', 'args' => 2],
        169 => ['name' => 'COUNTA', 'args' => null],
        183 => ['name' => 'PRODUCT', 'args' => null],
        184 => ['name' => 'FACT', 'args' => 1],
        189 => ['name' => 'DPRODUCT', 'args' => 3],
        190 => ['name' => 'ISNONTEXT', 'args' => 1],
        197 => ['name' => 'TRUNC', 'args' => null],
        198 => ['name' => 'ISLOGICAL', 'args' => 1],
        212 => ['name' => 'ROUNDUP', 'args' => 2],
        213 => ['name' => 'ROUNDDOWN', 'args' => 2],
        221 => ['name' => 'TODAY', 'args' => 0],
        229 => ['name' => 'SUMIF', 'args' => null],
        255 => ['name' => 'EXTERNAL', 'args' => null],
        336 => ['name' => 'CONCATENATE', 'args' => null],
        337 => ['name' => 'POWER', 'args' => 2],
        342 => ['name' => 'RADIANS', 'args' => 1],
        343 => ['name' => 'DEGREES', 'args' => 1],
        344 => ['name' => 'SUBTOTAL', 'args' => null],
        345 => ['name' => 'SUMPRODUCT', 'args' => null],
        346 => ['name' => 'SERIESSUM', 'args' => 4],
        347 => ['name' => 'FACTDOUBLE', 'args' => 1],
        354 => ['name' => 'ROMAN', 'args' => null],
    ];

    /** @return array{name:string,args:?int}|null */
    public static function byId(int $id): ?array
    {
        return self::BY_ID[$id] ?? null;
    }

    /** @return array{id:int,args:?int}|null */
    public static function byName(string $name): ?array
    {
        $name = strtoupper($name);
        foreach (self::BY_ID as $id => $definition) {
            if ($definition['name'] === $name) {
                return ['id' => $id, 'args' => $definition['args']];
            }
        }
        return null;
    }
}
