<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx; 

$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();

$spreadsheetinput = $reader->load("file/input/option_id and product id which has fast and custom order.xlsx");
$spreadsheetinput1 = $reader->load("file/output/catalog_product_option_type_value.xlsx");
//$d=$spreadsheet->getSheet(0)->toArray();

//echo count($d);

$sheetDatainput = $spreadsheetinput->getActiveSheet()->toArray();
$sheetDatainput1 = $spreadsheetinput1->getActiveSheet()->toArray();

$i=0;


$spreadsheetoutput = new Spreadsheet(); 
$sheetoutput = $spreadsheetoutput->getActiveSheet();
$sheetoutput->setCellValueByColumnAndRow(1,1,'option_id');
$sheetoutput->setCellValueByColumnAndRow(2,1,'custom_tab');
$sheetoutput->setCellValueByColumnAndRow(3,1,'stock_tab');
    //$sheetoutput->setCellValueByColumnAndRow(4,1,'disabled');
    $j=2;
foreach ($sheetDatainput as $t) {
 // process element here;
 //echo $t[2];
 foreach($sheetDatainput1 as $f)
 {
   if ($t[2]==$f[1]&&$f[3]==1)
   { $sheetoutput->setCellValueByColumnAndRow(1,$j,$t[2]);
    //$sheetoutput->setCellValueByColumnAndRow(2,$j,$t[1]);
  $sheetoutput->setCellValueByColumnAndRow(3,$j,1);

  $j++;
     break;
    }
    
 }

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
$writer->save('file/output/catalog_product_option_type_value_fast.xlsx'); 

?>

