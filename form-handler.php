<?php
require 'vendor/autoload.php'; // This loads PhpSpreadsheet

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$name = $_POST['name'];
$email = $_POST['email'];
$phone = $_POST['phone'];
$message = $_POST['message'];

$file = 'form-data.xlsx';

if (file_exists($file)) {
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($file);
    $sheet = $spreadsheet->getActiveSheet();
    $lastRow = $sheet->getHighestRow() + 1;
} else {
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setCellValue('A1', 'Name');
    $sheet->setCellValue('B1', 'Email');
    $sheet->setCellValue('C1', 'phone');
    $sheet->setCellValue('D1', 'Message');
    $lastRow = 2;
}

$sheet->setCellValue("A$lastRow", $name);
$sheet->setCellValue("B$lastRow", $email);
$sheet->setCellValue("C$lastRow", $phone);
$sheet->setCellValue("D$lastRow", $message);

$writer = new Xlsx($spreadsheet);
$writer->save($file);

echo "Form data saved successfully!";
?>
