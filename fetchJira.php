<?php

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

require_once "vendor/autoload.php";
[$email, $token, $url, $users, $hook, $siteUrl] = require("config.php");
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
$files = '';
foreach ($users as $u) {
    $body = [
        "jql" => 'type != epik AND project in (PR, SERWIS, WD, ZDR, ZD, ER) AND status in ("Do potwierdzenia", "DO WGRANIA", Done, "do testowania", "Do aktualizacji - krytyczne", "Gotowe do testowania", "code review") AND (assignee in ('.$u['id'].') OR "Osoba sprawdzajaca[People]" in ('.$u['id'].') AND status was "code review" after -1d) AND status changed after -1d ORDER BY due ASC'
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

        if ($u['id'] === $issue->fields->assignee->accountId) {
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
    $absoluteDirectoryYear = __DIR__."/tasks/".$now->format("Y");
    $directoryYear = "/tasks/".$now->format("Y");

    if (!$filesystem->exists($absoluteDirectoryYear)) {
        $filesystem->mkdir($absoluteDirectoryYear);
        if(!$filesystem->exists($absoluteDirectoryYear)) {
            throw new \RuntimeException(sprintf('Directory "%s" was not created', $absoluteDirectoryYear));
        }
    }
    $absoluteDirectoryMonth = $absoluteDirectoryYear . "/" . $now->format('m');
    $directoryMonth = $directoryYear . "/" . $now->format('m');
    if (!$filesystem->exists($absoluteDirectoryMonth)) {
        $filesystem->mkdir($absoluteDirectoryMonth);
        if(!$filesystem->exists($absoluteDirectoryMonth)) {
            throw new \RuntimeException(sprintf('Directory "%s" was not created', $absoluteDirectoryMonth));
        }
    }
    $absoluteDirectoryDay = $absoluteDirectoryMonth . "/" . $now->format('d');
    $directoryDay = $directoryMonth . "/" . $now->format('d');
    if (!$filesystem->exists($absoluteDirectoryDay)) {
        $filesystem->mkdir($absoluteDirectoryDay);
        if(!$filesystem->exists($absoluteDirectoryDay)) {
            throw new \RuntimeException(sprintf('Directory "%s" was not created', $absoluteDirectoryDay));
        }
    }

    $absoluteFileName = $absoluteDirectoryDay . '/' .$u['name'].'_'.$now->format('Y_m_d') . '_done.xlsx';
    $fileName = $directoryDay . '/' .$u['name'].'_'.$now->format('Y_m_d') . '_done.xlsx';

    $writer->save($absoluteFileName);

    $files .= $siteUrl.$fileName."\n";
}
$client = new Maknz\Slack\Client($hook);
$client->send($files);
