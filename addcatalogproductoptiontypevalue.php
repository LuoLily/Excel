<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx; 

$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();

$spreadsheetinput = $reader->load("file/input/catalog_product_option product with custom order and fast ship without dupicate.xlsx");

//$d=$spreadsheet->getSheet(0)->toArray();

//echo count($d);

$sheetDatainput = $spreadsheetinput->getActiveSheet()->toArray();

$i=0;


$spreadsheetoutput = new Spreadsheet(); 
$sheetoutput = $spreadsheetoutput->getActiveSheet();
$sheetoutput->setCellValueByColumnAndRow(1,1,'option_id');
$sheetoutput->setCellValueByColumnAndRow(2,1,'sort_order');
$sheetoutput->setCellValueByColumnAndRow(3,1,'group_option_value_id');
$sheetoutput->setCellValueByColumnAndRow(4,1,'disabled');
    //$sheetoutput->setCellValueByColumnAndRow(4,1,'disabled');
    $j=2;
foreach ($sheetDatainput as $t) {
 // process element here;
 //echo $t[2];
 $sheetoutput->setCellValueByColumnAndRow(1,$j,$t[0]);
 $sheetoutput->setCellValueByColumnAndRow(2,$j,1);
$sheetoutput->setCellValueByColumnAndRow(3,$j,12762);
if($t[1]==1)
$sheetoutput->setCellValueByColumnAndRow(4,$j,0);
else
$sheetoutput->setCellValueByColumnAndRow(4,$j,1);
$j++;
$sheetoutput->setCellValueByColumnAndRow(1,$j,$t[0]);
$sheetoutput->setCellValueByColumnAndRow(2,$j,2);
$sheetoutput->setCellValueByColumnAndRow(3,$j,12763);
if($t[2]==1)
$sheetoutput->setCellValueByColumnAndRow(4,$j,0);
else
$sheetoutput->setCellValueByColumnAndRow(4,$j,1);
$j++;
 /* if ($t[3]==1||$t[4]==1)
  {$sheetoutput->setCellValueByColumnAndRow(1,$j,$t[0]);
      $sheetoutput->setCellValueByColumnAndRow(2,$j,$t[1]);
    $sheetoutput->setCellValueByColumnAndRow(3,$j,$t[3]);
    $sheetoutput->setCellValueByColumnAndRow(4,$j,$t[4]);
    $sheetoutput->setCellValueByColumnAndRow(5,$j,$t[8]);
    $sheetoutput->setCellValueByColumnAndRow(6,$j,$t[19]);
    $j++;
}*/
 //output
  // echo $i."---".$t[0].",".$t[3]." <br>";
	
}
// Write an .xlsx file  
$writer = new Xlsx($spreadsheetoutput); 
  
// Save .xlsx file to the files directory 
$writer->save('file/output/catalog_product_option_type_value_add.xlsx'); 

?>

