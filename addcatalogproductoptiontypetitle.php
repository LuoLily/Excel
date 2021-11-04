<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx; 

$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();

$spreadsheetinput = $reader->load("file/input/catalog_product_option_type_value new with added.xlsx");

//$d=$spreadsheet->getSheet(0)->toArray();

//echo count($d);

$sheetDatainput = $spreadsheetinput->getActiveSheet()->toArray();

$i=0;


$spreadsheetoutput = new Spreadsheet(); 
$sheetoutput = $spreadsheetoutput->getActiveSheet();
$sheetoutput->setCellValueByColumnAndRow(1,1,'option_type_id');
$sheetoutput->setCellValueByColumnAndRow(2,1,'store_id');
$sheetoutput->setCellValueByColumnAndRow(3,1,'title');
    //$sheetoutput->setCellValueByColumnAndRow(4,1,'disabled');
    $j=2;
foreach ($sheetDatainput as $t) {
 // process element here;
 //echo $t[2];
 if($t[8]==12762){
 $sheetoutput->setCellValueByColumnAndRow(1,$j,$t[0]);
 $sheetoutput->setCellValueByColumnAndRow(2,$j,0);
$sheetoutput->setCellValueByColumnAndRow(3,$j,'Fast Ship');
$j++;
$sheetoutput->setCellValueByColumnAndRow(1,$j,$t[0]);
$sheetoutput->setCellValueByColumnAndRow(2,$j,1);
$sheetoutput->setCellValueByColumnAndRow(3,$j,'Fast Ship');
$j++;}
else if($t[8]==12763){
$sheetoutput->setCellValueByColumnAndRow(1,$j,$t[0]);
$sheetoutput->setCellValueByColumnAndRow(2,$j,0);
$sheetoutput->setCellValueByColumnAndRow(3,$j,'Custom Order');
$j++;
$sheetoutput->setCellValueByColumnAndRow(1,$j,$t[0]);
$sheetoutput->setCellValueByColumnAndRow(2,$j,1);
$sheetoutput->setCellValueByColumnAndRow(3,$j,'Custom Order');
$j++;
}
}
// Write an .xlsx file  
$writer = new Xlsx($spreadsheetoutput); 
  
// Save .xlsx file to the files directory 
$writer->save('file/output/catalog_product_option_type_title_add.xlsx'); 

?>

