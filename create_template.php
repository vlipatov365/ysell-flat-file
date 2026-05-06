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
    'B'  => 'Custom label (SKU)',
    'C'  => '*Category',
    'D'  => 'Description',
    'E'  => '*Title',
    'F'  => '*ConditionID',
    'G'  => 'ConditionDescription',
    'H'  => '*StartPrice',
    'I'  => '*Quantity',
    'J'  => '*Format',
    'K'  => '*Duration',
    'L'  => '*Location',
    'M'  => 'ShippingProfileName',
    'N'  => 'ReturnProfileName',
    'O'  => 'PaymentProfileName',
    'P'  => 'PicURL',
    'Q'  => 'C:Marke',
    'R'  => 'C:Hersteller',
    'S'  => 'C:Produktart',
    'T'  => 'C:Produkt',
    'U'  => 'C:Farbe',
    'V'  => 'C:Material',
    'W'  => 'C:Markenkompatibilität',
    'X'  => 'C:Modellkompatibilität',
    'Y'  => 'C:Herstellernummer',
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
