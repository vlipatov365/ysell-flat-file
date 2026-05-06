<?php
declare(strict_types=1);

require_once __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// ─── Config ────────────────────────────────────────────────────────────────
$apiKey  = getenv('YSELL_API_KEY')  ?: '149185983e29fed80f5010edd60410b7';
$baseUrl = getenv('YSELL_BASE_URL') ?: 'https://4457.test1.ysell.pro';

// ─── HTTP helper ───────────────────────────────────────────────────────────
function apiGet(string $url, string $apiKey): array
{
    $ch = curl_init($url);
    curl_setopt_array($ch, [
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_HTTPHEADER     => [
            'Authorization: Bearer ' . $apiKey,
            'Accept: application/json',
            'Content-Type: application/json',
        ],
        CURLOPT_SSL_VERIFYPEER => false,
        CURLOPT_TIMEOUT        => 30,
        CURLOPT_FOLLOWLOCATION => true,
    ]);
    $body  = curl_exec($ch);
    $code  = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    $error = curl_error($ch);
    curl_close($ch);

    if ($error) {
        throw new RuntimeException("cURL error: {$error}");
    }
    if ($code >= 400) {
        throw new RuntimeException("API returned HTTP {$code}: {$body}");
    }
    $data = json_decode($body, true);
    if ($data === null) {
        throw new RuntimeException("Invalid JSON: {$body}");
    }
    return $data;
}

// ─── Fetch all products (paginated) ────────────────────────────────────────
function fetchAllProducts(string $baseUrl, string $apiKey): array
{
    $products = [];
    $page     = 1;

    do {
        echo "  Fetching page {$page}...\n";
        $response = apiGet("{$baseUrl}/api/v1/product?per_page=50&page={$page}", $apiKey);

        // API returns a plain array (not wrapped in data/meta)
        $items    = isset($response['data']) ? $response['data'] : $response;
        $lastPage = $response['meta']['last_page'] ?? $response['last_page'] ?? 1;

        if (!is_array($items) || empty($items)) {
            break;
        }

        $products = array_merge($products, $items);
        $page++;
    } while ($page <= $lastPage);

    return $products;
}

// ─── Fetch manufacturers: returns [id => name] ──────────────────────────────
function fetchManufacturers(string $baseUrl, string $apiKey): array
{
    try {
        $response = apiGet("{$baseUrl}/api/v1/manufacturer", $apiKey);
        $items    = isset($response['data']) ? $response['data'] : $response;
        $map      = [];
        foreach ($items as $m) {
            if (isset($m['id'])) {
                $map[(int)$m['id']] = (string)($m['name'] ?? $m['title'] ?? '');
            }
        }
        return $map;
    } catch (RuntimeException $e) {
        echo "  Warning: could not fetch manufacturers: " . $e->getMessage() . "\n";
        return [];
    }
}

// ─── Extract OEM part numbers from title ───────────────────────────────────
// Titles follow patterns like:
//   "... wie Bosch 00600436 für ..."
//   "... wie IGNIS 488000525918 EGO 20.40943.000 Bauknecht 488000525918 ..."
//   "... Liebherr 7433698 9193354"
function parseTitleForCompatibility(string $title): array
{
    $brands = [];
    $models = [];

    // Match everything after "wie " up to a preposition or end of string
    if (preg_match('/\bwie\s+(.+?)(?:\s+(?:für|in|an|von|mit|zu)\b|$)/u', $title, $wie)) {
        $segment = $wie[1];

        // Part numbers: must contain at least one digit (filters out pure brand words like IGNIS)
        preg_match_all('/\b[A-Z0-9]*\d[A-Z0-9.\-\/]*\b/u', $segment, $numMatches);
        $models = array_slice(array_unique($numMatches[0]), 0, 3);

        // Brands: letter-only tokens (no digits), any capitalization (Bosch, IGNIS, EGO)
        preg_match_all('/\b([A-ZÄÖÜ][A-Za-zäöüÄÖÜß]*)\b/u', $segment, $brandMatches);
        $brands = array_unique(array_filter($brandMatches[1], fn($w) => !preg_match('/\d/', $w)));
    }

    // Fallback: extract part numbers directly from title (e.g. "Bosch 17002715 / 17001433 ...")
    if (empty($models)) {
        preg_match_all('/\b[A-Z0-9]*\d[A-Z0-9.\-\/]{4,}\b/', $title, $numMatches);
        $models = array_slice($numMatches[0], 0, 3);
    }

    return [
        'brands' => array_values($brands),
        'models' => $models,
    ];
}

// ─── Condition ─────────────────────────────────────────────────────────────
function resolveCondition(array $product): array
{
    $raw   = strtolower((string)($product['condition'] ?? ''));
    $isNew = ($raw === '' || str_contains($raw, 'new') || str_contains($raw, 'neu'));

    return [
        'id'          => $isNew ? '1000' : '3000',
        'description' => $isNew ? 'Neu' : 'Gebraucht',
    ];
}

// ─── Price ─────────────────────────────────────────────────────────────────
function resolvePrice(array $product): string
{
    // Real API field is purchase_price (string|null)
    $price = $product['purchase_price'] ?? $product['netto'] ?? null;
    if ($price === null || $price === '' || (float)$price === 0.0) {
        return '';
    }
    return number_format((float)$price, 2, '.', '');
}

// ─── Image ─────────────────────────────────────────────────────────────────
function resolveImage(array $product): string
{
    // Real API: "image" is a plain string (or null)
    return (string)($product['image'] ?? '');
}

// ─── SKU / Herstellernummer ─────────────────────────────────────────────────
function resolveSkuAndHerstellernummer(array $product, string $manufacturerName): array
{
    $id      = (string)($product['id'] ?? '');
    $extId   = (string)($product['ext_id'] ?? '');
    $isViox  = stripos($manufacturerName, 'viox') !== false
            || stripos($manufacturerName, 'виокс') !== false;

    // Extract first OEM article number from title (after "wie BRAND NUMBER")
    $compat  = parseTitleForCompatibility((string)($product['title'] ?? ''));
    $oemNum  = $compat['models'][0] ?? '';

    if ($isViox) {
        return [
            'custom_label'     => "{$id}-MP",
            'herstellernummer' => $oemNum,   // SKU = OEM article extracted from title
        ];
    }

    return [
        'custom_label'     => $extId !== '' ? $extId : $id,
        'herstellernummer' => $oemNum,
    ];
}

// ─── Main mapping ──────────────────────────────────────────────────────────
function mapProductToRow(array $product, array $manufacturers): array
{
    $manufacturerId   = (int)($product['manufacturer_id'] ?? 0);
    $manufacturerName = $manufacturers[$manufacturerId] ?? '';

    $condition = resolveCondition($product);
    $sku       = resolveSkuAndHerstellernummer($product, $manufacturerName);
    $compat    = parseTitleForCompatibility((string)($product['title'] ?? ''));

    // Title: real field is "title", truncate to 80 chars
    $title     = mb_substr((string)($product['title'] ?? ''), 0, 80);

    // C:Produktart / C:Produkt — first word of title
    $firstWord = explode(' ', trim($title))[0] ?? '';

    // Compatible brands: parsed from title (e.g. "wie Bosch ..." → "Bosch")
    $compatBrands = implode(', ', $compat['brands']);

    // Compatible models: up to 3 part numbers from title
    $compatModels = implode(', ', $compat['models']);

    return [
        '*Action'                => 'Add (SiteID=Germany|Country=DE|Currency=EUR|Version=941)',
        'Custom label (SKU)'     => $sku['custom_label'],
        '*Category'              => '',                        // not in API — fill manually or leave blank
        '*Title'                 => $title,
        '*ConditionID'           => $condition['id'],
        'ConditionDescription'   => $condition['description'],
        '*StartPrice'            => resolvePrice($product),
        '*Quantity'              => '1',
        '*Format'                => 'FixedPrice',
        '*Duration'              => 'GTC',
        '*Location'              => 'Bremen',
        'ShippingProfileName'    => 'DHL-Standardvorlage - (ID: 76770166018)',
        'ReturnProfileName'      => 'Standard-Widerruf 30 Tage für eBay Plus - (ID: 247364069018)',
        'PaymentProfileName'     => 'UK Gutschein',
        'PicURL'                 => resolveImage($product),
        'C:Marke'                => $manufacturerName,
        'C:Hersteller'           => $manufacturerName,
        'C:Produktart'           => $firstWord,
        'C:Produkt'              => $firstWord,
        'C:Farbe'                => '',                        // not in list API — leave blank
        'C:Material'             => '',                        // not in list API — leave blank
        'C:Markenkompatibilität' => $compatBrands,
        'C:Modellkompatibilität' => $compatModels,
        'C:Herstellernummer'     => $sku['herstellernummer'],
        'C:Installationsart'     => 'Vollintegriert',
    ];
}

// ─── Column map ────────────────────────────────────────────────────────────
const COL_MAP = [
    '*Action'                => 'A',
    'Custom label (SKU)'     => 'B',
    '*Category'              => 'C',
    '*Title'                 => 'D',
    '*ConditionID'           => 'E',
    'ConditionDescription'   => 'F',
    '*StartPrice'            => 'G',
    '*Quantity'              => 'H',
    '*Format'                => 'I',
    '*Duration'              => 'J',
    '*Location'              => 'K',
    'ShippingProfileName'    => 'L',
    'ReturnProfileName'      => 'M',
    'PaymentProfileName'     => 'N',
    'PicURL'                 => 'O',
    'C:Marke'                => 'P',
    'C:Hersteller'           => 'Q',
    'C:Produktart'           => 'R',
    'C:Produkt'              => 'S',
    'C:Farbe'                => 'T',
    'C:Material'             => 'U',
    'C:Markenkompatibilität' => 'V',
    'C:Modellkompatibilität' => 'W',
    'C:Herstellernummer'     => 'X',
    'C:Installationsart'     => 'Y',
];

// ─── Entry point ───────────────────────────────────────────────────────────
$useSample = in_array('--sample', $argv ?? [], true);

if ($useSample) {
    $sampleFile = __DIR__ . '/sample_products.json';
    if (!file_exists($sampleFile)) {
        fwrite(STDERR, "ERROR: sample_products.json not found.\n");
        exit(1);
    }
    $products      = json_decode(file_get_contents($sampleFile), true);
    $manufacturers = json_decode(file_get_contents(__DIR__ . '/sample_manufacturers.json'), true) ?? [];
    echo "Using sample data (" . count($products) . " products)\n";
} else {
    echo "Fetching manufacturers...\n";
    $manufacturers = fetchManufacturers($baseUrl, $apiKey);
    echo "  Found " . count($manufacturers) . " manufacturers.\n";

    echo "Fetching all products...\n";
    try {
        $products = fetchAllProducts($baseUrl, $apiKey);
    } catch (RuntimeException $e) {
        fwrite(STDERR, "ERROR: " . $e->getMessage() . "\n");
        exit(1);
    }
}

echo "Total products: " . count($products) . "\n";

file_put_contents(__DIR__ . '/products_raw.json', json_encode($products, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE));
echo "Saved products_raw.json\n";

$spreadsheet = IOFactory::load(__DIR__ . '/flat_template.xlsx');
$sheet       = $spreadsheet->getActiveSheet();

$row = 2;
foreach ($products as $product) {
    $mapped = mapProductToRow($product, $manufacturers);
    foreach ($mapped as $header => $value) {
        $col = COL_MAP[$header] ?? null;
        if ($col !== null) {
            $sheet->setCellValue($col . $row, (string)$value);
        }
    }
    $row++;
}

$writer = new Xlsx($spreadsheet);
$writer->save(__DIR__ . '/flat_filled.xlsx');

$count = $row - 2;
echo "Готово! Записано товаров: {$count}\n";
