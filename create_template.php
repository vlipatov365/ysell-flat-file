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
    'D'  => '*Title',
    'E'  => '*ConditionID',
    'F'  => 'ConditionDescription',
    'G'  => '*StartPrice',
    'H'  => '*Quantity',
    'I'  => '*Format',
    'J'  => '*Duration',
    'K'  => '*Location',
    'L'  => 'ShippingProfileName',
    'M'  => 'ReturnProfileName',
    'N'  => 'PaymentProfileName',
    'O'  => 'PicURL',
    'P'  => 'C:Marke',
    'Q'  => 'C:Hersteller',
    'R'  => 'C:Produktart',
    'S'  => 'C:Produkt',
    'T'  => 'C:Farbe',
    'U'  => 'C:Material',
    'V'  => 'C:Markenkompatibilität',
    'W'  => 'C:Modellkompatibilität',
    'X'  => 'C:Herstellernummer',
    'Y'  => 'C:Installationsart',
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
