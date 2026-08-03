<?php

declare(strict_types=1);
require __DIR__ . '/../vendor/autoload.php';
use Mnb\PHPExcel\Core\CellValue;
use Mnb\PHPExcel\Core\WorkbookData;
use Mnb\PHPExcel\Core\WorksheetData;
use Mnb\PHPExcel\Format\Xls;

$dir = sys_get_temp_dir() . '/mnb-xls-visual-' . bin2hex(random_bytes(4));
mkdir($dir, 0775, true);
$source = $dir . '/styled.xls';
$copy = $dir . '/copy.xls';
try {
    $sheet = new WorksheetData(
        name: 'Sales',
        rows: [
            ['Order ID', 'Order Date', 'Total', 'Status'],
            ['ORD-1', CellValue::date('2026-07-01', ['format' => 'yyyy-mm-dd']), 9996, 'Completed'],
        ],
        hasHeader: false,
        mergeCells: ['A3:D3'],
        columnWidths: ['A'=>15,'B'=>14,'C'=>16,'D'=>14],
        rowHeights: [1=>30],
        cellStyles: [
            'A1'=>['font'=>['bold'=>true,'color'=>'#FFFFFF'],'fill'=>['color'=>'#1F4E78'],'alignment'=>['horizontal'=>'center']],
            'B1'=>['font'=>['bold'=>true,'color'=>'#FFFFFF'],'fill'=>['color'=>'#1F4E78'],'alignment'=>['horizontal'=>'center']],
            'C1'=>['font'=>['bold'=>true,'color'=>'#FFFFFF'],'fill'=>['color'=>'#1F4E78'],'alignment'=>['horizontal'=>'center']],
            'D1'=>['font'=>['bold'=>true,'color'=>'#FFFFFF'],'fill'=>['color'=>'#1F4E78'],'alignment'=>['horizontal'=>'center']],
        ],
        rangeStyles: ['A2:D2'=>['borders'=>['all'=>['style'=>'thin','color'=>'#808080']]]],
        freezeRows: 1,
        freezeTopLeftCell: 'A2',
    );
    Xls::write(new WorkbookData([$sheet], ['_mnb_active_sheet'=>1]), $source);
    $snapshot = Xls::visualSnapshot($source);
    assert($snapshot['sheets'][0]['dimension'] === 'A1:D3');
    assert((float) ($snapshot['sheets'][0]['layout']['column_widths']['D'] ?? 0) === 14.0);
    assert((float) ($snapshot['sheets'][0]['layout']['row_heights'][1] ?? 0) === 30.0);
    assert(($snapshot['sheets'][0]['layout']['freeze_panes']['rows'] ?? null) === 1);
    $styleId = $snapshot['sheets'][0]['cells']['A1']['style_id'];
    $style = $snapshot['styles'][$styleId];
    assert(($style['font']['bold'] ?? false) === true);
    assert(($style['fill']['pattern'] ?? null) === 'solid');
    assert(($style['fill']['foreground']['rgb'] ?? null) === 'FF1F4E78');
    $borderId = $snapshot['sheets'][0]['cells']['A2']['style_id'];
    assert(($snapshot['styles'][$borderId]['border']['left']['style'] ?? null) === 'thin');
    assert(($snapshot['sheets'][0]['cells']['B2']['type'] ?? null) === 'date');
    Xls::createFromVisualSnapshot($snapshot, $copy);
    $round = Xls::visualSnapshot($copy);
    assert(($round['styles'][$round['sheets'][0]['cells']['A1']['style_id']]['font']['bold'] ?? false) === true);
    assert(($round['styles'][$round['sheets'][0]['cells']['A2']['style_id']]['border']['bottom']['style'] ?? null) === 'thin');
    assert(($round['sheets'][0]['cells']['B2']['type'] ?? null) === 'date');
    echo "xls_visual_snapshot_smoke passed\n";
} finally {
    @unlink($source); @unlink($copy); @rmdir($dir);
}
