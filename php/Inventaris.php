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

function maakXLSX($data,$file,$tab)
{
  $reader = new Xlsx();
//  $spreadsheet = new Spreadsheet();
  if(file_exists($file))
// try-catch voor als $file geen xlsx is...
    $spreadsheet = $reader->load($file);
  else 
  {
    $spreadsheet = new Spreadsheet();
  }    
  $sheet_names = $spreadsheet->getSheetNames();
  if(in_array($tab, $sheet_names))
  {
    //Tabnaam bestaat al
    $sheet = $spreadsheet->getSheetByName($tab);
  }
  else
  {
    //Tabnaam bestaat nog niet, nieuw aan te maken
    $sheet = new Worksheet($spreadsheet, $tab);
    $spreadsheet->addSheet($sheet,$spreadsheet->getSheetCount()+1);
    //$sheet = $spreadsheet->createSheet();
    //$sheet->setTitle($tab);
  }
  
//  foreach ($sheet_names as $sheet_name)
//  {
//    //echo $sheet_name,"\n";
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
  //$spreadsheet->getColumnDimension($col)->setAutoSize(true);
  return $spreadsheet;
}

function schrijfXLSX($data,$file,$tab)
{
  $spreadsheet = maakXLSX($data,$file,$tab);
  $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
  $writer->save($file);
}

function copyXLSX($data,$file,$tab)
{
  $spreadsheet = maakXLSX($data,$file,$tab);
  $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
  header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  header("Content-Disposition: attachment;filename=$file");
  header('Cache-Control: max-age=0');
// Sla de spreadsheet naar output
  $writer->save('php://output');
  //exit;
}

// Je bron-data
$bronData = [
    ['Naam', 'Leeftijd', 'Stad'],
    ['Jan', 30, 'Amsterdam'],
    ['Eva', [25, 36], 'Rotterdam'],
    [['Pieter','Piet'], 35, 'Utrecht'],
    ['Piert', 46, ['Amsterdam', 'Rotterdam', 'Utrecht']]
];

if (isset($_POST['bewaarkeuze']))
{
  switch($_POST['bewaarkeuze'])
  {
    case 'xlsx': //bewaar xlsx serverside
      schrijfXLSX($bronData,"Inventaris.xlsx","kieken");
      break;
    case 'csv': //bewaar csv serverside
      schrijfCSV($bronData,"Inventaris.csv");
      break;
    case 'json': //bewaar json serverside
      schrijfJSON($bronData,"Inventaris.json");
      break;
    case 'alles': //bewaar alle formaten serversie
      schrijfCSV($bronData,"Inventaris.csv");
      schrijfJSON($bronData,"Inventaris.json");
      schrijfXLSX($bronData,"Inventaris.xlsx","kieken");
      break;
    case 'lokaal': //bewaar lokale copy
      copyXLSX($bronData,"Inventaris_d.xlsx","Namen");
      break;
  }
}
//echo "✓ Bestanden zijn aangemaakt\n";
?>

