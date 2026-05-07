<?php
require_once __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle('Listing');

// eBay Germany flat file columns (row 1 = headers)
$columns = [
    'A'  => '*Action',
    'B'  => '*Category',
    'C'  => '*Title',
    'D'  => 'Description',
    'E'  => '*ConditionID',
    'F'  => 'PicURL',
    'G'  => '*Quantity',
    'H'  => '*Format',
    'I'  => '*StartPrice',
    'J'  => '*Duration',
    'K'  => '*Location',
    'L'  => 'ShippingProfileName',
    'M'  => 'ReturnProfileName',
    'N'  => 'PaymentProfileName',
    'O'  => 'C:Marke',
    'P'  => 'C:Produktart',
    'Q'  => 'Custom label (SKU)',
    'R'  => 'ConditionDescription',
    'S'  => 'C:Herstellernummer',
    'T'  => 'C:Produkt',
    'U'  => 'C:Modellkompatibilität',
    'V'  => 'C:Farbe',
    'W'  => 'C:Hersteller',
    'X'  => 'C:Material',
    'Y'  => 'C:Markenkompatibilität',
    'Z'  => 'C:Installationsart',
];

foreach ($columns as $col => $header) {
    $sheet->setCellValue($col . '1', $header);
}

// Style header row bold
$sheet->getStyle('A1:Y1')->getFont()->setBold(true);

// Auto-width for all columns
foreach (array_keys($columns) as $col) {
    $sheet->getColumnDimension($col)->setAutoSize(true);
}

$writer = new Xlsx($spreadsheet);
$writer->save(__DIR__ . '/flat_template.xlsx');

echo "flat_template.xlsx created successfully.\n";
echo "Columns:\n";
foreach ($columns as $col => $header) {
    echo "  {$col}: {$header}\n";
}
