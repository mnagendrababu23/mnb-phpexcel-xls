<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Format;

use Mnb\PHPExcel\Compatibility\XlsReader;
use Mnb\PHPExcel\Compatibility\XlsWriter;
use Mnb\PHPExcel\Core\WorkbookData;
use Mnb\PHPExcel\Core\WorkbookFactory;
use Mnb\PHPExcel\Reader\Options\ReaderOptions;
use Mnb\PHPExcel\Reader\ReadSession;

final class Xls
{
    /** @param array<string,mixed>|ReaderOptions $options */
    public static function read(string $path,array|ReaderOptions $options=[]):ReadSession{return new ReadSession($path,new XlsReader(),$options);}
    /** @param iterable<array<int|string,mixed>|mixed>|WorkbookData $rows @param array<string,mixed> $options */
    public static function write(iterable|WorkbookData $rows,string $path,array $options=[]):void{$workbook=$rows instanceof WorkbookData?$rows:WorkbookFactory::workbook($rows,(string)($options['sheet_name']??'Sheet1'),(bool)($options['with_header']??false));(new XlsWriter())->write($workbook,$path,$options);}
}
