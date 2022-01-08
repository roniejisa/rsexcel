<?php
namespace ExcelLaravel;
use ExcelLaravel\RSSheet;
class RSExcel{
    public static function excelOfTableHTML(string $html, string $filename = 'excel.xlsx'){
        RSSheet::loadSpreadSheet($html);
        $sheet = $spreadsheet->getSheet(0);
        self::setHeightColumn(20);
        self::setAutoSize($sheet);
        RSSheet::download($spreadsheet,$filename);
    }
}


    