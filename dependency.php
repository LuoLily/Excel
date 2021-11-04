<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx; 

$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();

$spreadsheetinput = $reader->load("file/input/mageworx_optiontemplates_group_option dependency.xlsx");


//$d=$spreadsheet->getSheet(0)->toArray();

//echo count($d);

$sheetDatainput = $spreadsheetinput->getActiveSheet()->toArray();

//fast ship
//$sheetDatainput1 = $spreadsheetinput->getSheet(1)->toArray();

//custom order
$sheetDatainput1 = $spreadsheetinput->getSheet(2)->toArray();

$i=0;


$spreadsheetoutput = new Spreadsheet(); 
$sheetoutput = $spreadsheetoutput->getActiveSheet();
$sheetoutput->setCellValueByColumnAndRow(1,1,'child_option_id');
$sheetoutput->setCellValueByColumnAndRow(2,1,'child_option_type_id');
    $sheetoutput->setCellValueByColumnAndRow(3,1,'parent_option_id');
    $sheetoutput->setCellValueByColumnAndRow(4,1,'parent_option_type_id');
    $sheetoutput->setCellValueByColumnAndRow(5,1,'product_id');
    $sheetoutput->setCellValueByColumnAndRow(6, 1,'group_id');
    $j=2;
foreach ($sheetDatainput as $t) {
 // process element here;

 //fast ship
 //if ($t[5]==1&&($t[4]!=12762||$t[4]!=12763))
 //custom order
  if ($t[6]==1&&($t[4]!=12762||$t[4]!=12763))
  {   foreach($sheetDatainput1 as $f)
    {  
      //fast ship
     // if($f[2]==$t[2]&&$f[4]==12762)
      //custom order
         if($f[2]==$t[2]&&$f[4]==12763)
      {$sheetoutput->setCellValueByColumnAndRow(1,$j,$t[0]);
      $sheetoutput->setCellValueByColumnAndRow(2,$j,$t[1]);
    $sheetoutput->setCellValueByColumnAndRow(3,$j,$f[0]);
    $sheetoutput->setCellValueByColumnAndRow(4,$j,$f[1]);
    $sheetoutput->setCellValueByColumnAndRow(5,$j,$t[2]);
    $sheetoutput->setCellValueByColumnAndRow(6,$j,$t[7]);
    $j++;}
    }
}
/*if ($t[6]==1)
  { foreach($sheetDatainput2 as $g)
    {  if($g[2]==$t[2]&&$g[4]==12763)
      $sheetoutput->setCellValueByColumnAndRow(1,$j,$t[0]);
      $sheetoutput->setCellValueByColumnAndRow(2,$j,$t[1]);
    $sheetoutput->setCellValueByColumnAndRow(3,$j,$g[0]);
    $sheetoutput->setCellValueByColumnAndRow(4,$j,$g[1]);
    $sheetoutput->setCellValueByColumnAndRow(5,$j,$t[2]);
    $sheetoutput->setCellValueByColumnAndRow(6,$j,$t[7]);
    $j++;
    }
}*/
 //output

  // echo $i."---".$t[0].",".$t[3]." <br>";
	$i++;
}
// Write an .xlsx file  
$writer = new Xlsx($spreadsheetoutput); 
  
// Save .xlsx file to the files directory 
$writer->save('file/output/mageworx_option_dependency custom order.xlsx'); 

?>

