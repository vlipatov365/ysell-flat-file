<?php
require_once __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

$file = __DIR__ . '/flat_filled.xlsx';
if (!file_exists($file)) { echo "flat_filled.xlsx not found.\n"; exit(1); }

$sheet   = IOFactory::load($file)->getActiveSheet();
$highRow = $sheet->getHighestDataRow();

// Read only columns we care about
$colMap = [
    'A' => '*Action', 'B' => 'Custom label (SKU)', 'C' => '*Category',
    'D' => 'Description', 'E' => '*Title', 'F' => '*ConditionID',
    'G' => 'ConditionDescription', 'H' => '*StartPrice', 'I' => '*Quantity',
    'J' => '*Format', 'K' => '*Duration', 'L' => '*Location',
    'M' => 'ShippingProfileName', 'N' => 'ReturnProfileName', 'O' => 'PaymentProfileName',
    'P' => 'PicURL', 'Q' => 'C:Marke', 'R' => 'C:Hersteller',
    'S' => 'C:Produktart', 'T' => 'C:Produkt', 'U' => 'C:Farbe',
    'V' => 'C:Material', 'W' => 'C:Markenkompatibilität', 'X' => 'C:Modellkompatibilität',
    'Y' => 'C:Herstellernummer', 'Z' => 'C:Installationsart',
];

echo "=== flat_filled.xlsx verification ===\n\n";

// 1. Row count
$dataRows = $highRow - 1;
echo "1. Data rows: {$dataRows}\n";

// 2. Required columns check
$required = ['A' => '*Action', 'E' => '*Title', 'F' => '*ConditionID',
             'H' => '*StartPrice', 'I' => '*Quantity', 'J' => '*Format',
             'K' => '*Duration', 'L' => '*Location'];
echo "\n2. Required columns:\n";
$issues = [];
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
        $issues[] = $col;
    }
}

// 3. Category coverage
echo "\n3. Category (*Category) coverage:\n";
$noCat = 0;
for ($row = 2; $row <= $highRow; $row++) {
    if ((string)$sheet->getCell('C' . $row)->getValue() === '') $noCat++;
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
echo "\n5. Sample rows (vs products_raw.json):\n";
$raw = json_decode(file_get_contents(__DIR__ . '/products_raw.json'), true) ?? [];
for ($row = 2; $row <= min(4, $highRow); $row++) {
    $idx = $row - 2;
    echo "  Row {$row} (product #{$idx}):\n";
    echo "    title    = " . mb_substr((string)$sheet->getCell('E' . $row)->getValue(), 0, 70) . "\n";
    echo "    category = " . $sheet->getCell('C' . $row)->getValue() . "\n";
    echo "    price    = " . $sheet->getCell('H' . $row)->getValue() . "\n";
    echo "    sku      = " . $sheet->getCell('B' . $row)->getValue() . "\n";
    echo "    brand    = " . $sheet->getCell('Q' . $row)->getValue() . "\n";
    echo "    condID   = " . $sheet->getCell('F' . $row)->getValue() . "\n";
    echo "    compat   = " . $sheet->getCell('W' . $row)->getValue() . "\n";
    echo "    models   = " . $sheet->getCell('X' . $row)->getValue() . "\n";
}

// 6. Title length
echo "\n6. Title length (≤80 chars):\n";
$tooLong = [];
for ($row = 2; $row <= $highRow; $row++) {
    $t = (string)$sheet->getCell('E' . $row)->getValue();
    if (mb_strlen($t) > 80) $tooLong[] = "Row {$row}: " . mb_strlen($t) . " chars";
}
echo "  " . (empty($tooLong) ? "All ≤ 80 ✓" : implode("\n  ", $tooLong)) . "\n";

echo "\nDone.\n";
