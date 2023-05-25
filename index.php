<?php

require_once "Classes/PHPExcel.php";
require('fpdf184/fpdf.php');

$path = "list.xlsx";
date_default_timezone_set("Asia/Kolkata");
$pdf = new FPDF('L', 'mm', 'A4');

function AddText($pdf, $text, $x, $y, $a, $f, $t, $s, $r, $g, $b)
{
    $pdf->SetFont($f, $t, $s);
    $pdf->SetXY($x, $y);
    $pdf->SetTextColor($r, $g, $b);
    $pdf->Cell(299, 19, $text, 0, 0, $a);
}

function CreatePage($pdf, $name,$course)
{
    $pdf->AddPage();
    $pdf->SetFont('Arial', 'B', 16);
    $pdf->SetCreator('');
    $pdf->Image('certificate.png', 0, 0, 297);
    AddText($pdf, ucwords($name), 0, 76, 'C', 'times', 'B', 35, 248, 207, 64);
    AddText($pdf, ucwords($course), 0, 114, 'C', 'times', 'B', 28, 28, 33, 67);
}

$reader = PHPExcel_IOFactory::createReaderForFile($path);
$excel_Obj = $reader->load($path);

$worksheet = $excel_Obj->getSheet('0');   

$lastRow = $worksheet->getHighestRow();
$totalData = $lastRow - 1;
$colomncount = $worksheet->getHighestDataColumn();
$colomncount_number = PHPExcel_Cell::columnIndexFromString($colomncount);

for ($row = 2; $row <= $lastRow; $row++) {

    $pdf = new FPDF('L','mm','A4');

    $person = $worksheet->getCell(PHPExcel_Cell::stringFromColumnIndex(0) . $row)->getValue();
    $course = $worksheet->getCell(PHPExcel_Cell::stringFromColumnIndex(1) . $row)->getValue();

    $person=trim($person);
    $course=trim($course);

     if($person==""){
        break;
     }
    
    CreatePage($pdf, $person, $course);

     $filename="certificates/".$person."-".$course.".pdf";
     $pdf->Output($filename,'F');

 }
 echo "Done";
 ?>
