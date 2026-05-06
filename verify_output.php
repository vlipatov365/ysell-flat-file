<?php
require_once __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

$file = __DIR__ . '/flat_filled.xlsx';
if (!file_exists($file)) {
    echo "flat_filled.xlsx not found.\n";
    exit(1);
}

$spreadsheet = IOFactory::load($file);
$sheet       = $spreadsheet->getActiveSheet();
$highRow     = $sheet->getHighestRow();
$highCol     = $sheet->getHighestColumn();

echo "=== flat_filled.xlsx verification ===\n\n";

// Step 5.1: Row count
$dataRows = $highRow - 1;
echo "1. Data rows (excluding header): {$dataRows}\n";

// Get header row
$headers = [];
for ($col = 'A'; $col <= $highCol; $col++) {
    $val = $sheet->getCell($col . '1')->getValue();
    if ($val !== '' && $val !== null) {
        $headers[$col] = $val;
    }
}

echo "\nColumns found:\n";
foreach ($headers as $col => $name) {
    echo "  {$col}: {$name}\n";
}

// Step 5.2: Check required columns for empty cells
$required = ['A', 'D', 'E', 'G', 'H', 'I', 'J', 'K'];
echo "\n2. Required column check (cols with *):\n";
$issues = [];
foreach ($required as $col) {
    $empty = [];
    for ($row = 2; $row <= $highRow; $row++) {
        $val = $sheet->getCell($col . $row)->getValue();
        if ($val === '' || $val === null) {
            $empty[] = $row;
        }
    }
    $colName = $headers[$col] ?? $col;
    if (empty($empty)) {
        echo "  {$col} ({$colName}): OK — all filled\n";
    } else {
        echo "  {$col} ({$colName}): EMPTY in rows " . implode(', ', $empty) . "\n";
        $issues[] = $col;
    }
}
echo empty($issues) ? "  → All required columns filled.\n" : "  → Some required columns have empty cells.\n";

// Step 5.3: Sample 3 rows vs products_raw.json
echo "\n3. Sample comparison (first 3 rows vs products_raw.json):\n";
$raw = json_decode(file_get_contents(__DIR__ . '/products_raw.json'), true) ?? [];
for ($row = 2; $row <= min(4, $highRow); $row++) {
    $idx = $row - 2;
    $title  = $sheet->getCell('D' . $row)->getValue();
    $price  = $sheet->getCell('G' . $row)->getValue();
    $sku    = $sheet->getCell('B' . $row)->getValue();
    $brand  = $sheet->getCell('P' . $row)->getValue();
    $cond   = $sheet->getCell('E' . $row)->getValue();
    echo "  Row {$row} (product #{$idx}): title='{$title}' | price={$price} | sku='{$sku}' | brand='{$brand}' | conditionID={$cond}\n";
}

// Step 5.4: Title length check
echo "\n4. Title length check (max 80 chars):\n";
$tooLong = [];
for ($row = 2; $row <= $highRow; $row++) {
    $title = (string)$sheet->getCell('D' . $row)->getValue();
    if (mb_strlen($title) > 80) {
        $tooLong[] = "Row {$row}: " . mb_strlen($title) . " chars — '{$title}'";
    }
}
if (empty($tooLong)) {
    echo "  All titles ≤ 80 chars. OK\n";
} else {
    foreach ($tooLong as $item) {
        echo "  OVER LIMIT: {$item}\n";
    }
}

echo "\nVerification complete.\n";
