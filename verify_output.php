<?php
require_once __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

$file = __DIR__ . '/flat_filled.xlsx';
if (!file_exists($file)) { echo "flat_filled.xlsx not found.\n"; exit(1); }

$sheet   = IOFactory::load($file)->getActiveSheet();
$highRow = $sheet->getHighestDataRow();

// Column map per instruction order
$colMap = [
    'A' => '*Action',          'B' => '*Category',         'C' => '*Title',
    'D' => 'Description',      'E' => '*ConditionID',       'F' => 'PicURL',
    'G' => '*Quantity',        'H' => '*Format',            'I' => '*StartPrice',
    'J' => '*Duration',        'K' => '*Location',          'L' => 'ShippingProfileName',
    'M' => 'ReturnProfileName','N' => 'PaymentProfileName', 'O' => 'C:Marke',
    'P' => 'C:Produktart',     'Q' => 'Custom label (SKU)', 'R' => 'ConditionDescription',
    'S' => 'C:Herstellernummer','T' => 'C:Produkt',         'U' => 'C:Modellkompatibilität',
    'V' => 'C:Farbe',          'W' => 'C:Hersteller',       'X' => 'C:Material',
    'Y' => 'C:Markenkompatibilität', 'Z' => 'C:Installationsart',
];

echo "=== flat_filled.xlsx verification ===\n\n";

// 1. Row count
$dataRows = $highRow - 1;
echo "1. Data rows: {$dataRows}\n";

// 2. Required columns check
$required = [
    'A' => '*Action', 'B' => '*Category', 'C' => '*Title',
    'E' => '*ConditionID', 'G' => '*Quantity', 'H' => '*Format',
    'I' => '*StartPrice', 'J' => '*Duration', 'K' => '*Location',
];
echo "\n2. Required columns:\n";
foreach ($required as $col => $name) {
    $empty = [];
    for ($row = 2; $row <= $highRow; $row++) {
        if ((string)$sheet->getCell($col . $row)->getValue() === '') $empty[] = $row;
    }
    if (empty($empty)) {
        echo "  {$col} ({$name}): OK\n";
    } else {
        echo "  {$col} ({$name}): EMPTY rows " . implode(', ', array_slice($empty, 0, 5))
             . (count($empty) > 5 ? '…' : '') . " [" . count($empty) . " total]\n";
    }
}

// 3. Category coverage
echo "\n3. Category (*Category) coverage (col B):\n";
$noCat = 0;
for ($row = 2; $row <= $highRow; $row++) {
    if ((string)$sheet->getCell('B' . $row)->getValue() === '') $noCat++;
}
$filled = $dataRows - $noCat;
echo "  Filled: {$filled}/{$dataRows}" . ($noCat > 0 ? " ({$noCat} without category)" : " ✓") . "\n";

// 4. Description presence
echo "\n4. Description column (D):\n";
$noDesc = 0;
for ($row = 2; $row <= $highRow; $row++) {
    if (strlen((string)$sheet->getCell('D' . $row)->getValue()) < 100) $noDesc++;
}
echo "  " . ($noDesc === 0 ? "All rows have HTML description ✓" : "{$noDesc} rows missing/short description") . "\n";

// 5. Sample rows
echo "\n5. Sample rows:\n";
for ($row = 2; $row <= min(4, $highRow); $row++) {
    $idx = $row - 2;
    echo "  Row {$row} (product #{$idx}):\n";
    echo "    title    = " . mb_substr((string)$sheet->getCell('C' . $row)->getValue(), 0, 70) . "\n";
    echo "    category = " . $sheet->getCell('B' . $row)->getValue() . "\n";
    echo "    price    = " . $sheet->getCell('I' . $row)->getValue() . "\n";
    echo "    sku      = " . $sheet->getCell('Q' . $row)->getValue() . "\n";
    echo "    condID   = " . $sheet->getCell('E' . $row)->getValue() . "\n";
    echo "    picURL   = " . mb_substr((string)$sheet->getCell('F' . $row)->getValue(), 0, 60) . "\n";
    echo "    brand    = " . $sheet->getCell('O' . $row)->getValue() . "\n";
    echo "    material = " . $sheet->getCell('X' . $row)->getValue() . "\n";
    echo "    compat   = " . $sheet->getCell('Y' . $row)->getValue() . "\n";
    echo "    models   = " . $sheet->getCell('U' . $row)->getValue() . "\n";
}

// 6. Title length
echo "\n6. Title length (≤80 chars, col C):\n";
$tooLong = [];
for ($row = 2; $row <= $highRow; $row++) {
    $t = (string)$sheet->getCell('C' . $row)->getValue();
    if (mb_strlen($t) > 80) $tooLong[] = "Row {$row}: " . mb_strlen($t) . " chars";
}
echo "  " . (empty($tooLong) ? "All ≤ 80 ✓" : implode("\n  ", $tooLong)) . "\n";

echo "\nDone.\n";
