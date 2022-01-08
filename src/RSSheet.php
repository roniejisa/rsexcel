<?php
namespace ExcelLaravel;

use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\RichText\Run;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use Illuminate\Support\Collection;

class RSSheet
{
    static function setBold(Run $richText , bool $bool = true) :Run{
        $richText->getFont()->setBold($bool);
        return $richText;
    }

    static function setItalic(Run $richText, bool $bool = true) :Run{
        $richText->getFont()->setItalic($bool);     
        return $richText;
    }

    static function setColor(Run $richText,Color $color) :Run{
        $richText->getFont()->setColor($color);
        return $richText;
    }

    static function setUnderline(Run $richText, bool $bool = true) :Run{
        $richText->getFont()->setUnderline($bool);
        return $richText;
    }

    static function setSize(Run $richText, int $size) :Run{
        $richText->getFont()->setSize($size);
        return $richText;
    }

    static function setHeightColumn(Worksheet $sheet, int $height) :void{
        foreach ($sheet->getRowIterator() as $column) {
            $sheet->getRowDimension($column->getRowIndex())->setRowHeight($height);
        }
    }

    static function setAutoSize(Worksheet $sheet) :void{
        foreach ($sheet->getColumnIterator() as $column) {
            $sheet->getColumnDimension($column->getColumnIndex())->setAutoSize(true);
        }
    }

    static function getValue(Worksheet $sheet,string $column) :?string{
        return $sheet->getCell($column)->getValue();
    }

    static function toCollection(Worksheet $sheet) :Collection{ 
        $collect = new Collection;
        foreach ($sheet->getRowIterator() as $key => $row) {
            $collect[$key] = new Collection;
            foreach ($row->getCellIterator() as $cell) {
                $column = $cell->getParent()->getCurrentCoordinate();
                $value = $cell->getValue();
                $collect[$key][$column] = $value;
            }
        }

        $collectNull = $collect->filter(function ($q) {
            return $q == $q->filter(function ($value){
                return $value == null;
            });
        });

        $collection = $collect->filter(function ($q, $key) use ($collectNull) {
            return !$collectNull->keys()->contains($key);
        });
        return $collection;
    }

    static function getListCell(Worksheet $sheet) :Collection{
        $listCells = new Collection;
        foreach ($sheet->getRowIterator() as $key => $row) {
            $collect[$key] = new Collection;
            foreach ($row->getCellIterator() as $cell) {
                $listCells[] = $cell->getCoordinate();
            }
        }
        return $listCells;
    }

    static function download(Spreadsheet $spreadsheet,string $filename) :void{
        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save($filename);
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="'. urlencode($filename).'"');
        $content = file_get_contents($filename);
        ob_end_clean();
        $writer->save('php://output');
        unlink($filename);
        exit($content);
    }

    static function groupCollection(Spreadsheet $sheets,int $sheetIndex,array $arraySpecial = []) :Collection{
        $collection = RSSheet::toCollection($sheets->getSheet($sheetIndex));
        $field = $collection[1];
        $field = $field->values();
        $collection = $collection->filter(function($q,$key){
            return $key !== 1;
        });

        $newCollection = new Collection;
        $i = 0;
        foreach($collection as $key => $data){
            $newCollection[$i] = new Collection;
            $a = 0;
            foreach($data as $key => $value){
                if(in_array($field[$a],$arraySpecial)){
                    if(!isset($newCollection[$i][$field[$a]])){
                        $newCollection[$i][$field[$a]] = new Collection;
                    }
                    if($value !== null){
                        $newCollection[$i][$field[$a]][] = [
                            'column' => $key,
                            'value' => $value
                        ];
                    }
                }else{
                    $newCollection[$i][$field[$a]] = [
                        'column' => $key,
                        'value' => $value
                    ];
                }
                $a++;
            }
            $i++;
        }
        return $newCollection;
    }

    static function loadSpreadSheet(string $html) :Spreadsheet{
        ob_start();
        $html_file = \storage_path('/framework/cache/laravel-excel/laravel-excel-') . \Str::random(60).'.html';
        \file_put_contents($html_file,$html);
        $reader = IOFactory::createReader('Html');
        $spreadsheet =  $reader->load($html_file);
        return $spreadsheet;
    }
}