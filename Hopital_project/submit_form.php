<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

function appendToExcel($name, $email, $message) {
    $filePath = 'contact_detail.xlsx';

    if (!file_exists($filePath)) {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setCellValue('A1', 'Name');
        $sheet->setCellValue('B1', 'Email');
        $sheet->setCellValue('C1', 'Message');
    } else {
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($filePath);
        $sheet = $spreadsheet->getActiveSheet();
    }

    $lastRow = $sheet->getHighestRow();
    $sheet->setCellValue('A' . ($lastRow + 1), $name);
    $sheet->setCellValue('B' . ($lastRow + 1), $email);
    $sheet->setCellValue('C' . ($lastRow + 1), $message);

    $writer = new Xlsx($spreadsheet);
    $writer->save($filePath);
}

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $name = $_POST['name'];
    $email = $_POST['email'];
    $message = $_POST['message'];

    appendToExcel($name, $email, $message);
    
    header('Location: contact.html');
}
?>
