<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

function appendToExcel($name, $email, $phone, $date, $time, $message) {
    $filePath = 'appointments.xlsx';

    if (!file_exists($filePath)) {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setCellValue('A1', 'Name');
        $sheet->setCellValue('B1', 'Email');
        $sheet->setCellValue('C1', 'Phone');
        $sheet->setCellValue('D1', 'Date');
        $sheet->setCellValue('E1', 'Time');
        $sheet->setCellValue('F1', 'Message');
    } else {
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($filePath);
        $sheet = $spreadsheet->getActiveSheet();
    }

    $lastRow = $sheet->getHighestRow();
    $sheet->setCellValue('A' . ($lastRow + 1), $name);
    $sheet->setCellValue('B' . ($lastRow + 1), $email);
    $sheet->setCellValue('C' . ($lastRow + 1), $phone);
    $sheet->setCellValue('D' . ($lastRow + 1), $date);
    $sheet->setCellValue('E' . ($lastRow + 1), $time);
    $sheet->setCellValue('F' . ($lastRow + 1), $message);

    $writer = new Xlsx($spreadsheet);
    $writer->save($filePath);
}

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $name = $_POST['name'];
    $email = $_POST['email'];
    $phone = $_POST['phone'];
    $date = $_POST['date'];
    $time = $_POST['time'];
    $message = $_POST['message'];

    appendToExcel($name, $email, $phone, $date, $time, $message);
    
    header('Location: appointment.html');
}
?>
