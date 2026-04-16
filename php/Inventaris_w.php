<?php
require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
//use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
//use PhpOffice\PhpSpreadsheet\Writer\Ods;

function schrijfCSV($data,$file,$del=',',$idel=';')
{
  $fp = fopen($file, 'w');
  foreach ($data as $rij)
  {
    $kolom=0;
    $rrij=$rij;
    foreach ($rij as $cel)
    {
      if(is_array($cel))
      {
        $ccel=array($kolom=>implode($idel,$cel));
        $rrij=array_replace($rij,$ccel);
      }
      $kolom++;
    }
    fputcsv($fp, $rrij, $del);
  }
  fclose($fp);
}

function schrijfJSON($data,$file)
{
  $jsonData = [];
  foreach ($data as $index => $rij)
  {
    if ($index === 0)
    {
        // Header rij gebruiken als keys
        $headers = $rij;
    }
    else
    {
        // Data rij omzetten naar associatieve array
        $rowData = [];
        foreach ($headers as $colIndex => $header) {
            $rowData[$header] = $rij[$colIndex];
        }
        $jsonData[] = $rowData;
    }
  }
  file_put_contents($file, json_encode($jsonData, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE));
}

function schrijfXLSX($data,$file,$tab)
{
  $reader = new Xlsx();
//  $spreadsheet = new Spreadsheet();
  if(file_exists($file))
    $spreadsheet = $reader->load($file);
  else 
  {
    $spreadsheet = new Spreadsheet();
  }    
  $sheet_names = $spreadsheet->getSheetNames();
  if(in_array($tab, $sheet_names))
  {
    echo "Tabnaam bestaat al.\n";
    $sheet = $spreadsheet->getSheetByName($tab);
  }
  else
  {
    echo "Tabnaam bestaat nog niet, nieuw gemaakt.\n";
    $sheet = new Worksheet($spreadsheet, $tab);
    $spreadsheet->addSheet($sheet,$spreadsheet->getSheetCount()+1);
//    $sheet = $spreadsheet->createSheet();
//    $sheet->setTitle($tab);
  }
//  foreach ($sheet_names as $sheet_name)
//  {
//    echo $sheet_name,"\n";
//  }

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
//  $spreadsheet->getColumnDimension($col)->setAutoSize(true);
  $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
  $writer->save($file);
//  $reader->save($file);

}
  

// Je bron-data
$bronData = [
    ['Naam', 'Leeftijd', 'Stad'],
    ['Jan', 30, 'Amsterdam'],
    ['Eva', [25, 36], 'Rotterdam'],
    [['Pieter','Piet'], 35, 'Utrecht'],
    ['Piert', 46, ['Amsterdam', 'Rotterdam', 'Utrecht']]
];

schrijfCSV($bronData,"Inventaris_w.csv");
schrijfJSON($bronData,"Inventaris_w.json");
schrijfXLSX($bronData,"Inventaris_w.xlsx","Namen5");

echo "✓ Bestanden zijn aangemaakt\n";
?>
