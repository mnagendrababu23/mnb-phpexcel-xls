<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Compatibility;

use Mnb\PHPExcel\Core\CellValue;
use Mnb\PHPExcel\Core\RichText;
use Mnb\PHPExcel\Core\WorkbookData;
use Mnb\PHPExcel\Core\WorksheetData;
use Mnb\PHPExcel\Support\AtomicFileWriter;
use Mnb\PHPExcel\Support\Coordinate;
use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;

/** Legacy BIFF8/XLS writer powered by the optional XLS compatibility module. */
final class XlsWriter
{
    /** @param array<string,mixed> $options */
    public function write(WorkbookData $workbook,string $path,array $options=[]):void
    {
        $this->ensureDependency();
        AtomicFileWriter::write($path,function(string $temporary)use($workbook,$options):void{
            $spreadsheet=new \PhpOffice\PhpSpreadsheet\Spreadsheet();
            while($spreadsheet->getSheetCount()>0){$spreadsheet->removeSheetByIndex(0);}
            foreach($workbook->sheets as $index=>$source){$sheet=new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet,$source->name);$spreadsheet->addSheet($sheet);$this->writeSheet($sheet,$source,$options);}
            if($spreadsheet->getSheetCount()===0){$spreadsheet->createSheet()->setTitle('Sheet1');}
            $spreadsheet->setActiveSheetIndex(0);
            $writer=\PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet,'Xls');
            if(method_exists($writer,'setPreCalculateFormulas')){$writer->setPreCalculateFormulas((bool)($options['precalculate_formulas']??false));}
            if(method_exists($writer,'setIncludeCharts')){$writer->setIncludeCharts((bool)($options['include_charts']??true));}
            try{$writer->save($temporary);}finally{$spreadsheet->disconnectWorksheets();}
        });
    }

    /** @param array<string,mixed> $options */
    private function writeSheet(object $sheet,WorksheetData $source,array $options):void
    {
        foreach($source->rows as $rowIndex=>$row){$excelRow=$rowIndex+1;foreach(array_values($row) as $columnIndex=>$value){$cell=Coordinate::columnIndexToName($columnIndex+1).$excelRow;$this->setValue($sheet,$cell,$value);}}
        foreach($source->mergeCells as $range){$sheet->mergeCells($range);}
        foreach($source->columnWidths as $column=>$width){$letter=is_int($column)||ctype_digit((string)$column)?Coordinate::columnIndexToName((int)$column):strtoupper((string)$column);$sheet->getColumnDimension($letter)->setWidth((float)$width);}
        foreach($source->rowHeights as $row=>$height){$sheet->getRowDimension((int)$row)->setRowHeight((float)$height);}
        if($source->freezeTopLeftCell!==null){$sheet->freezePane($source->freezeTopLeftCell);}elseif($source->freezeRows>0||$source->freezeColumns>0){$sheet->freezePane(Coordinate::columnIndexToName($source->freezeColumns+1).($source->freezeRows+1));}elseif($source->freezeHeader){$sheet->freezePane('A2');}
        $filter=$source->autoFilterRange;if($filter===null&&$source->autoFilter&&$source->rows!==[]){$max=max(array_map('count',$source->rows));$filter='A'.($source->headerRowIndex??1).':'.Coordinate::columnIndexToName(max(1,$max)).($source->headerRowIndex??1);}if($filter!==null&&$filter!==''){$sheet->setAutoFilter($filter);}
        $this->applyStyles($sheet,$source);
        foreach($source->hyperlinks as $item){$cell=$sheet->getCell((string)$item['cell']);$cell->getHyperlink()->setUrl((string)$item['url']);if(isset($item['tooltip']))$cell->getHyperlink()->setTooltip((string)$item['tooltip']);if(isset($item['display']))$cell->setValue((string)$item['display']);}
        foreach($source->comments as $item){$comment=$sheet->getComment((string)$item['cell']);$comment->setAuthor((string)($item['author']??'MNB PHPExcel'));$comment->getText()->createTextRun((string)$item['text']);if(isset($item['width']))$comment->setWidth((string)$item['width'].'pt');if(isset($item['height']))$comment->setHeight((string)$item['height'].'pt');}
        foreach($source->images as $item){if(!is_file((string)$item['path']))continue;$drawing=new \PhpOffice\PhpSpreadsheet\Worksheet\Drawing();$drawing->setPath((string)$item['path']);$drawing->setCoordinates((string)$item['cell']);$drawing->setName((string)($item['name']??basename((string)$item['path'])));if(isset($item['width']))$drawing->setWidth((int)$item['width']);if(isset($item['height']))$drawing->setHeight((int)$item['height']);$drawing->setWorksheet($sheet);}
        foreach($source->dataValidations as $definition){$range=(string)($definition['range']??'');if($range==='')continue;foreach($sheet->getCellCollection()->getCoordinates() as $coordinate){if(!$this->cellInRange($coordinate,$range))continue;$validation=$sheet->getCell($coordinate)->getDataValidation();$type=strtolower((string)($definition['type']??'custom'));$map=['list'=>\PhpOffice\PhpSpreadsheet\Cell\DataValidation::TYPE_LIST,'whole'=>\PhpOffice\PhpSpreadsheet\Cell\DataValidation::TYPE_WHOLE,'decimal'=>\PhpOffice\PhpSpreadsheet\Cell\DataValidation::TYPE_DECIMAL,'date'=>\PhpOffice\PhpSpreadsheet\Cell\DataValidation::TYPE_DATE,'time'=>\PhpOffice\PhpSpreadsheet\Cell\DataValidation::TYPE_TIME,'text_length'=>\PhpOffice\PhpSpreadsheet\Cell\DataValidation::TYPE_TEXTLENGTH,'custom'=>\PhpOffice\PhpSpreadsheet\Cell\DataValidation::TYPE_CUSTOM];$validation->setType($map[$type]??\PhpOffice\PhpSpreadsheet\Cell\DataValidation::TYPE_CUSTOM);$values=(array)($definition['values']??[]);if($type==='list'&&$values!==[])$validation->setFormula1('"'.str_replace('"','""',implode(',',$values)).'"');elseif(isset($definition['formula1']))$validation->setFormula1((string)$definition['formula1']);if(isset($definition['formula2']))$validation->setFormula2((string)$definition['formula2']);$validation->setAllowBlank((bool)($definition['allow_blank']??true));$validation->setShowErrorMessage((bool)($definition['show_error']??true));}}
    }

    private function setValue(object $sheet,string $cell,mixed $value):void
    {
        if($value instanceof RichText){$rich=new \PhpOffice\PhpSpreadsheet\RichText\RichText();foreach($value->runs as $run){$part=$rich->createTextRun($run->text);$font=$part->getFont();if(isset($run->style['bold']))$font->setBold((bool)$run->style['bold']);if(isset($run->style['italic']))$font->setItalic((bool)$run->style['italic']);if(isset($run->style['size']))$font->setSize((float)$run->style['size']);if(isset($run->style['color']))$font->getColor()->setARGB($this->argb((string)$run->style['color']));}$sheet->setCellValue($cell,$rich);return;}
        if($value instanceof CellValue){switch($value->type()){case CellValue::TYPE_BLANK:$sheet->setCellValue($cell,null);return;case CellValue::TYPE_FORMULA:$sheet->setCellValue($cell,'='.$value->value());return;case CellValue::TYPE_TEXT:$sheet->setCellValueExplicit($cell,(string)$value->value(),\PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);return;case CellValue::TYPE_BOOLEAN:$sheet->setCellValueExplicit($cell,(bool)$value->value(),\PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_BOOL);return;case CellValue::TYPE_ERROR:$sheet->setCellValueExplicit($cell,(string)$value->value(),\PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_ERROR);return;default:$value=$value->displayValue();}}
        if(is_string($value)&&str_starts_with($value,'=')){$sheet->setCellValue($cell,$value);}elseif(is_string($value)){$sheet->setCellValueExplicit($cell,$value,\PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);}else{$sheet->setCellValue($cell,$value);}
    }

    private function applyStyles(object $sheet,WorksheetData $source):void
    {
        foreach($source->columnStyles as $column=>$style){$letter=is_int($column)||ctype_digit((string)$column)?Coordinate::columnIndexToName((int)$column):strtoupper((string)$column);$sheet->getStyle($letter.':'.$letter)->applyFromArray($this->style($style,$source));}
        foreach($source->rowStyles as $row=>$style){$sheet->getStyle((int)$row.':'.(int)$row)->applyFromArray($this->style($style,$source));}
        foreach($source->cellStyles as $cell=>$style){$sheet->getStyle((string)$cell)->applyFromArray($this->style($style,$source));}
        foreach($source->rangeStyles as $range=>$style){$sheet->getStyle((string)$range)->applyFromArray($this->style($style,$source));}
        if($source->hasHeader&&$source->headerStyle!==[]){$row=$source->headerRowIndex??1;$max=max(1,...array_map('count',$source->rows));$sheet->getStyle('A'.$row.':'.Coordinate::columnIndexToName($max).$row)->applyFromArray($this->style($source->headerStyle,$source));}
    }
    private function style(array|string $style,WorksheetData $source):array
    {
        if(is_string($style))$style=$source->namedStyles[$style]??[];$font=(array)($style['font']??[]);$fill=(array)($style['fill']??[]);$alignment=(array)($style['alignment']??[]);$result=[];
        if($font!==[]){$result['font']=array_filter(['name'=>$font['name']??null,'size'=>$font['size']??null,'bold'=>$font['bold']??null,'italic'=>$font['italic']??null,'underline'=>$font['underline']??null,'color'=>isset($font['color'])?['argb'=>$this->argb((string)$font['color'])]:null],fn($v)=>$v!==null);}
        if($fill!==[]||isset($style['background_color'])){$color=(string)($fill['color']??$fill['start_color']??$style['background_color']??'FFFFFF');$result['fill']=['fillType'=>\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,'startColor'=>['argb'=>$this->argb($color)]];}
        if(isset($style['borders']))$result['borders']=$style['borders'];
        if($alignment!==[]||isset($style['horizontal'])||isset($style['vertical'])||isset($style['wrap_text'])){$result['alignment']=array_filter(['horizontal'=>$alignment['horizontal']??$style['horizontal']??null,'vertical'=>$alignment['vertical']??$style['vertical']??null,'wrapText'=>$alignment['wrap_text']??$style['wrap_text']??null],fn($v)=>$v!==null);}
        if(isset($style['number_format']))$result['numberFormat']=['formatCode'=>(string)$style['number_format']];
        if(isset($style['protection']))$result['protection']=$style['protection'];
        return $result;
    }
    private function argb(string $color):string{$color=strtoupper(ltrim($color,'#'));return strlen($color)===6?'FF'.$color:$color;}
    private function cellInRange(string $cell,string $range):bool{[$start,$end]=array_pad(explode(':',strtoupper($range),2),2,strtoupper($range));[$sc,$sr]=Coordinate::splitCellRef($start);[$ec,$er]=Coordinate::splitCellRef($end);[$cc,$cr]=Coordinate::splitCellRef(strtoupper($cell));return$cc>=min($sc,$ec)&&$cc<=max($sc,$ec)&&$cr>=min($sr,$er)&&$cr<=max($sr,$er);}
    private function ensureDependency():void{if(!class_exists('PhpOffice\\PhpSpreadsheet\\Spreadsheet'))throw MnbExcelException::withCode('Legacy XLS writing requires mnb/mnb-phpexcel-xls.',ErrorCode::EXTENSION_MISSING,[],null,'Install mnb/mnb-phpexcel-xls to write .xls files.');}
}
