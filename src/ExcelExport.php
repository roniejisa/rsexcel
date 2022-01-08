<?php
    namespace ExcelLaravel;
    use PhpOffice\PhpSpreadsheet\RichText\RichText;
    class ExcelExport{

        private $html;
        public function __construct($html){
            $this->html = $html;
        } 

        public function htmlToExcel($filename){
            $spreadsheet = RSSheet::loadSpreadSheet($this->html);
            $sheet = $spreadsheet->getSheet(0);
            $listCells = RSSheet::getListCell($sheet);
            foreach($listCells as $cell){
                if(strpos(RSSheet::getValue($sheet,$cell),'ITALIC') !== false){
                    $listItems = \Str::of(RSSheet::getValue($sheet,$cell))->explode("\n");
                    $new = new RichText();
                    $i = 0;
                    foreach($listItems as $key => $item){
                        if(strpos($item,'ITALIC') === 0){
                            $new->createTextRun(\Str::of($item)->replace('ITALIC_','')."\n");
                            RSSheet::setItalic($new->getRichTextElements()[$i]);
                            $i++;
                        }elseif(strpos($item,'BOLD') === 0){
                            $listBold = \Str::of($item)->explode(' ');
                            foreach($listBold as $keyBold => $itemBold){
                                if(strpos($itemBold,'ANCHOR') !== false){
                                    $hasAnchor = true;
                                }else{
                                    $hasAnchor = false;
                                }
                                if($keyBold === $listBold->count() - 1){
                                    if($keyBold == 1){
                                        $itemBold = ' '.$itemBold;
                                    }
                                    $new->createTextRun(\Str::of($itemBold)->replace(['BOLD_','_ANCHOR'],['',''])."\n");
                                    RSSheet::setBold($new->getRichTextElements()[$i]);
                                }else{
                                    if($keyBold !== 0){
                                        $itemBold = ' '.$itemBold.' '; 
                                    }
                                    $new->createTextRun(\Str::of($itemBold)->replace(['BOLD_','_ANCHOR'],['','']));
                                    RSSheet::setBold($new->getRichTextElements()[$i]);
                                }
                                if($hasAnchor){
                                    RSSheet::setSize($new->getRichTextElements()[$i],14);
                                }
                                $i++;
                            }
                        }else{
                            $new->createTextRun($item."\n");
                            $i++;
                        }
                    }
                    $sheet->setCellValue($cell, $new);
                }
            }
            RSSheet::setHeightColumn($sheet,30);
            RSSheet::setAutoSize($sheet);
            RSSheet::download($spreadsheet, $filename);
        }
    }
?>