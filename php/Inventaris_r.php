<?php
require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;

function leesXLSX($file)
{
  $reader = new Xlsx();
  if(file_exists($file))
// try-catch voor als $file geen xlsx is...
    $spreadsheet = $reader->load($file);
  else 
  {
    $spreadsheet = new Spreadsheet();
  }    
  $sheet_names = $spreadsheet->getSheetNames();
  $sheet_count = $spreadsheet->getSheetCount();
 
  foreach ($sheet_names as $sheet_name)
  {
    echo $sheet_name,"\n";
  }
  var_dump($sheet_names);
  echo "\n";
  var_export($sheet_names,false);
  echo "\n";
  var_export($sheet_names,true);
  echo "\n";
  print_r($sheet_names,false);
  echo "\n";
  print_r($sheet_names,true);
  echo "\n";
  
  for($i=0;i<$sheet_count;$i++)
  {
    $activeSheet = $spreadsheet->getSheet($i);
  }
/*
  foreach ($data as $rijIndex => $rijData) 
  {
    $rij = $rijIndex + 1;
    foreach ($rijData as $kolomIndex => $waarde)
    {
      $kolomLetter = chr(65 + $kolomIndex); // A, B, C...
      $cel = $kolomLetter . $rij;
      $sheet->setCellValue($cel, (is_array($waarde)?str_replace(",","\n",implode(",",$waarde)):$waarde));
      $sheet->getStyle($cel)->getAlignment()->setVertical(PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP);
      $sheet->getStyle($cel)->getAlignment()->setWrapText(true);
      $sheet->getColumnDimension($kolomLetter)->setAutoSize(true);
    }
  }
  $spreadsheet->getDefaultStyle()->getAlignment()->setWrapText(true);
  //$spreadsheet->getColumnDimension($col)->setAutoSize(true);
  return $spreadsheet;
*/
}

//function schrijfXLSX($data,$file,$tab)
//{
//  $spreadsheet = maakXLSX($data,$file,$tab);
//  $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
//  $writer->save($file);
//}

$file = "Inventaris.xlsx";
leesXLSX($file);
?>