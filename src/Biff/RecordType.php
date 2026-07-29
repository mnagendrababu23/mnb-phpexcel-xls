<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Biff;

final class RecordType
{
    public const BOF = 0x0809;
    public const EOF = 0x000A;
    public const BOUNDSHEET = 0x0085;
    public const CODEPAGE = 0x0042;
    public const DATEMODE = 0x0022;
    public const SST = 0x00FC;
    public const CONTINUE = 0x003C;
    public const EXTSST = 0x00FF;
    public const FONT = 0x0031;
    public const FORMAT = 0x041E;
    public const XF = 0x00E0;
    public const STYLE = 0x0293;
    public const WINDOW1 = 0x003D;
    public const WINDOW2 = 0x023E;
    public const DIMENSIONS = 0x0200;
    public const ROW = 0x0208;
    public const COLINFO = 0x007D;
    public const NUMBER = 0x0203;
    public const RK = 0x027E;
    public const MULRK = 0x00BD;
    public const LABEL = 0x0204;
    public const LABELSST = 0x00FD;
    public const BOOLERR = 0x0205;
    public const BLANK = 0x0201;
    public const MULBLANK = 0x00BE;
    public const FORMULA = 0x0006;
    public const STRING = 0x0207;
    public const MERGEDCELLS = 0x00E5;
    public const PANE = 0x0041;
    public const SELECTION = 0x001D;
    public const AUTOFILTERINFO = 0x009D;
    public const FILTERMODE = 0x009B;
    public const CALCMODE = 0x000D;
    public const CALCCOUNT = 0x000C;
    public const REFMODE = 0x000F;
    public const ITERATION = 0x0011;
    public const DELTA = 0x0010;
    public const SAVERECALC = 0x005F;

    private function __construct()
    {
    }
}
