<?php

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

require_once "vendor/autoload.php";
[$email, $token, $url] = require("config.php");
Unirest\Request::auth($email, $token);

$headers = array(
    'Accept' => 'application/json',
);

$fetchUser = Unirest\Request::get(
     "{$url}/rest/api/3/user/search",
    $headers,
    ['query' => $email]
);
$user = $fetchUser->body[0];

$body = [
    "jql" => 'type != epik AND project in (PR, SERWIS, WD, ZDR) AND status in ("Do potwierdzenia", "DO WGRANIA", Done, "do testowania", "Do aktualizacji - krytyczne") AND (assignee in (currentUser()) OR "Osoba sprawdzajaca[People]" in (currentUser()) AND status was "code review" after -1d) AND status changed after -1d ORDER BY due ASC'
];


$response = Unirest\Request::get(
    "{$url}/rest/api/3/search",
    $headers,
    $body
);

$now = new DateTime();

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'Hello World !');

$row = 1;
$letter = "A";
$issues = $response->body->issues;
$sheet->getColumnDimension($letter)->setAutoSize(true);
$sheet->setCellValue($letter++ . $row, "Klucz");
$sheet->getColumnDimension($letter)->setWidth(260, "px");
$sheet->setCellValue($letter++ . $row, "Podsumowanie");
$sheet->getColumnDimension($letter)->setAutoSize(true);
$sheet->setCellValue($letter++ . $row, "Status");
$sheet->getColumnDimension($letter)->setAutoSize(true);
$sheet->setCellValue($letter++ . $row, "Typ");
$sheet->getColumnDimension($letter)->setAutoSize(true);
$sheet->setCellValue($letter++ . $row, "Osoba przypisana");
$sheet->getColumnDimension($letter)->setAutoSize(true);
$sheet->setCellValue($letter++ . $row, "Termin");
$sheet->setCellValue($letter++ . $row, "Etykiety");
$sheet->getColumnDimension($letter)->setAutoSize(true);
$sheet->getStyle("A{$row}:{$letter}{$row}")->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_MEDIUM);
$sheet->setCellValue($letter++ . $row, "Element nadrzędny");

$row++;
$firstRowOfIssues = $row;
foreach($issues as $issue) {
    $letter = "A";
    $sheet->setCellValue($letter++ . $row, $issue->key);
    $sheet->setCellValue($letter++ . $row, $issue->fields->summary);
    $sheet->setCellValue($letter++ . $row, $issue->fields->status->name);

    if ($user->accountId === $issue->fields->assignee->accountId) {
        if ($issue->fields->project->key === "SERWIS") {
            $type = "SERWIS";
        } else {
            $type = "INNE";
        }
    } else {
        $type = "SPRAWDZONE CODE REVIEW";
    }

    $sheet->setCellValue($letter++ . $row, $type);
    $sheet->setCellValue($letter++ . $row, $issue->fields->assignee->displayName);
    $sheet->setCellValue($letter++ . $row, $issue->fields->duedate);
    $sheet->setCellValue($letter++ . $row, implode(', ', $issue->fields->labels));
    $sheet->setCellValue($letter . $row, $issue->fields->parent->key ?? '');

    $row++;
}
$lastRow = $row - 1;
$sheet->getStyle("A{$firstRowOfIssues}:{$letter}{$lastRow}")->getBorders()->getOutline()->setBorderStyle(
    Border::BORDER_MEDIUM
);

$filesystem = new \Symfony\Component\Filesystem\Filesystem();

$writer = new Xlsx($spreadsheet);
$directoryYear = __DIR__."/tasks/" . $now->format("Y");

if (!$filesystem->exists($directoryYear)) {
    $filesystem->mkdir($directoryYear);
    if(!$filesystem->exists($directoryYear)) {
        throw new \RuntimeException(sprintf('Directory "%s" was not created', $directoryYear));
    }
}
$directoryMonth = $directoryYear . "/" . $now->format('m');
if (!$filesystem->exists($directoryMonth)) {
    $filesystem->mkdir($directoryMonth);
    if(!$filesystem->exists($directoryMonth)) {
        throw new \RuntimeException(sprintf('Directory "%s" was not created', $directoryMonth));
    }
}

$fileName = $directoryMonth . '/' . $now->format('Y_m_d') . '_done.xlsx';

$writer->save( $fileName);