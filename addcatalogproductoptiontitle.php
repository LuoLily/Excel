<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx; 

$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();

$spreadsheetinput = $reader->load("file/input/catalog_product_option product with custom order and fast ship.xlsx");

//$d=$spreadsheet->getSheet(0)->toArray();

//echo count($d);

$sheetDatainput = $spreadsheetinput->getActiveSheet()->toArray();

$i=0;


$spreadsheetoutput = new Spreadsheet(); 
$sheetoutput = $spreadsheetoutput->getActiveSheet();
$sheetoutput->setCellValueByColumnAndRow(1,1,'option_title_id');
$sheetoutput->setCellValueByColumnAndRow(2,1,'option_id');
$sheetoutput->setCellValueByColumnAndRow(3,1,'store_id');
$sheetoutput->setCellValueByColumnAndRow(4,1,'title');
    //$sheetoutput->setCellValueByColumnAndRow(4,1,'disabled');
    $j=2;
foreach ($sheetDatainput as $t) {
 // process element here;
 //echo $t[2];
 $sheetoutput->setCellValueByColumnAndRow(1,$j,$t[0]);
 $sheetoutput->setCellValueByColumnAndRow(2,$j,0);
$sheetoutput->setCellValueByColumnAndRow(3,$j,'Material Options:');
$j++;
$sheetoutput->setCellValueByColumnAndRow(1,$j,$t[0]);
$sheetoutput->setCellValueByColumnAndRow(2,$j,1);
$sheetoutput->setCellValueByColumnAndRow(3,$j,'Material Options:');
$j++;
	
}
// Write an .xlsx file  
$writer = new Xlsx($spreadsheetoutput); 
  
// Save .xlsx file to the files directory 
$writer->save('file/output/catalog_product_option_title_add.xlsx'); 

?>

