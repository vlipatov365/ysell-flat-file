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
    $body    = curl_exec($ch);
    $code    = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    $error   = curl_error($ch);
    curl_close($ch);

    if ($error) {
        throw new RuntimeException("cURL error: {$error}");
    }
    if ($code >= 400) {
        throw new RuntimeException("API returned HTTP {$code}: {$body}");
    }

    $data = json_decode($body, true);
    if ($data === null) {
        throw new RuntimeException("Invalid JSON response: {$body}");
    }
    return $data;
}

// ─── Pagination ────────────────────────────────────────────────────────────
function fetchAllProducts(string $baseUrl, string $apiKey): array
{
    $products = [];
    $page     = 1;

    do {
        echo "  Fetching page {$page}...\n";
        $response = apiGet("{$baseUrl}/api/v1/product?per_page=50&page={$page}", $apiKey);

        // Support both {data:[...], meta:{...}} and direct array responses
        $items    = $response['data']        ?? $response['products'] ?? $response;
        $lastPage = $response['meta']['last_page']
            ?? $response['last_page']
            ?? ($response['meta']['total_pages'] ?? 1);

        if (!is_array($items)) {
            break;
        }

        $products = array_merge($products, $items);
        $page++;
    } while ($page <= $lastPage);

    return $products;
}

// ─── Nested value helper ────────────────────────────────────────────────────
/**
 * Resolve a dot-notation path with multiple fallback keys.
 * e.g. get($p, 'brand.name', 'brand', 'manufacturer') → first non-empty value
 */
function get(array $data, string ...$paths): mixed
{
    foreach ($paths as $path) {
        $keys  = explode('.', $path);
        $value = $data;
        foreach ($keys as $key) {
            if (is_array($value) && array_key_exists($key, $value)) {
                $value = $value[$key];
            } else {
                $value = null;
                break;
            }
        }
        if ($value !== null && $value !== '') {
            return $value;
        }
    }
    return '';
}

// ─── Image URL ─────────────────────────────────────────────────────────────
function resolveImage(array $product): string
{
    // Try various common image field structures
    foreach (['images', 'photos', 'pictures', 'gallery'] as $key) {
        if (!empty($product[$key]) && is_array($product[$key])) {
            $first = $product[$key][0];
            return is_array($first)
                ? ($first['url'] ?? $first['path'] ?? $first['link'] ?? $first['src'] ?? '')
                : (string)$first;
        }
    }
    foreach (['image', 'photo', 'picture', 'image_url', 'photo_url', 'picture_url'] as $key) {
        if (!empty($product[$key])) {
            return (string)$product[$key];
        }
    }
    return '';
}

// ─── Price from suppliers ──────────────────────────────────────────────────
function resolvePrice(array $product): string
{
    // Try suppliers array first
    foreach (['suppliers', 'supplier', 'provider'] as $key) {
        if (!empty($product[$key]) && is_array($product[$key])) {
            $first = $product[$key][0];
            $price = is_array($first)
                ? ($first['price'] ?? $first['cost'] ?? $first['purchase_price'] ?? null)
                : null;
            if ($price !== null) {
                return number_format((float)$price, 2, '.', '');
            }
        }
    }
    // Fallback to direct price fields
    foreach (['price', 'sell_price', 'retail_price', 'cost', 'purchase_price'] as $key) {
        if (isset($product[$key]) && $product[$key] !== '') {
            return number_format((float)$product[$key], 2, '.', '');
        }
    }
    return '';
}

// ─── Condition ─────────────────────────────────────────────────────────────
function resolveCondition(array $product): array
{
    $raw = strtolower((string)get($product, 'condition', 'state', 'product_condition', 'status_condition'));
    $isNew = in_array($raw, ['new', 'neu', 'новый', '1', 'new_other'], true)
        || $raw === ''   // default to new if not specified
        || str_contains($raw, 'new')
        || str_contains($raw, 'neu');

    return [
        'id'          => $isNew ? '1000' : '3000',
        'description' => $isNew ? 'Neu' : 'Gebraucht',
    ];
}

// ─── Compatible models ("Passend für") ────────────────────────────────────
function resolveCompatibleModels(array $product): string
{
    // Try various field names for compatibility data
    foreach (['compatible_models', 'compatibility', 'passend_fur', 'suitable_for', 'fits', 'fit_for'] as $key) {
        if (!empty($product[$key])) {
            $val = $product[$key];
            if (is_array($val)) {
                // Take up to 3 model numbers
                $models = array_slice(array_filter(array_map(function($m) {
                    return is_array($m) ? ($m['model'] ?? $m['number'] ?? $m['name'] ?? '') : (string)$m;
                }, $val)), 0, 3);
                return implode(', ', $models);
            }
            if (is_string($val)) {
                // If comma-separated string, take first 3
                $parts = array_slice(array_map('trim', explode(',', $val)), 0, 3);
                return implode(', ', $parts);
            }
        }
    }
    // Also check nested attributes
    foreach (['attributes', 'characteristics', 'specs', 'properties', 'custom_fields'] as $attrKey) {
        if (!empty($product[$attrKey]) && is_array($product[$attrKey])) {
            foreach (['compatible_models', 'passend_fur', 'fits', 'compatibility', 'models'] as $key) {
                if (!empty($product[$attrKey][$key])) {
                    $val = $product[$attrKey][$key];
                    if (is_array($val)) {
                        return implode(', ', array_slice($val, 0, 3));
                    }
                    $parts = array_slice(array_map('trim', explode(',', (string)$val)), 0, 3);
                    return implode(', ', $parts);
                }
            }
        }
    }
    return '';
}

// ─── Attribute resolver ────────────────────────────────────────────────────
function resolveAttr(array $product, string ...$keys): string
{
    // Direct field check
    foreach ($keys as $key) {
        $val = get($product, $key);
        if ($val !== '') return (string)$val;
    }
    // Check nested attribute blocks
    foreach (['attributes', 'characteristics', 'specs', 'properties', 'custom_fields', 'extra'] as $attrBlock) {
        if (!empty($product[$attrBlock]) && is_array($product[$attrBlock])) {
            foreach ($keys as $key) {
                // Handle both associative {key: value} and indexed [{name: key, value: val}] formats
                if (isset($product[$attrBlock][$key]) && $product[$attrBlock][$key] !== '') {
                    return (string)$product[$attrBlock][$key];
                }
                // Indexed format
                foreach ($product[$attrBlock] as $attr) {
                    if (is_array($attr)) {
                        $attrName = strtolower((string)($attr['name'] ?? $attr['key'] ?? $attr['code'] ?? ''));
                        if (in_array($attrName, array_map('strtolower', [$key]), true)) {
                            return (string)($attr['value'] ?? $attr['val'] ?? '');
                        }
                    }
                }
            }
        }
    }
    return '';
}

// ─── Brand / Manufacturer ──────────────────────────────────────────────────
function resolveBrand(array $product): string
{
    return resolveAttr($product, 'brand', 'manufacturer', 'brand_name', 'make', 'hersteller', 'marke');
}

// ─── SKU logic ─────────────────────────────────────────────────────────────
function resolveSkuAndHerstellernummer(array $product): array
{
    $id    = (string)($product['id'] ?? '');
    $sku   = (string)get($product, 'sku', 'article_number', 'article', 'code', 'product_code');
    $brand = strtolower(resolveBrand($product));

    $isViox = str_contains($brand, 'viox') || str_contains($brand, 'виокс');

    if ($isViox) {
        return [
            'custom_label'    => "{$id}-MP",
            'herstellernummer' => $sku,
        ];
    }

    return [
        'custom_label'    => $sku !== '' ? "sku={$sku}" : "sku={$id}",
        'herstellernummer' => resolveAttr($product, 'model_number', 'model', 'oem_number', 'part_number', 'article_number') ?: $sku,
    ];
}

// ─── Main mapping ──────────────────────────────────────────────────────────
function mapProductToRow(array $product): array
{
    $condition = resolveCondition($product);
    $sku       = resolveSkuAndHerstellernummer($product);
    $brand     = resolveBrand($product);
    $title     = substr((string)get($product, 'name', 'title', 'product_name', 'label'), 0, 80);

    // C:Produktart — first word of product name
    $firstWord = explode(' ', trim($title))[0] ?? '';
    $category  = (string)get($product, 'category.id', 'category_id', 'ebay_category', 'category');

    return [
        '*Action'                  => 'Add (SiteID=Germany|Country=DE|Currency=EUR|Version=941)',
        'Custom label (SKU)'       => $sku['custom_label'],
        '*Category'                => $category,
        '*Title'                   => $title,
        '*ConditionID'             => $condition['id'],
        'ConditionDescription'     => $condition['description'],
        '*StartPrice'              => resolvePrice($product),
        '*Quantity'                => '1',
        '*Format'                  => 'FixedPrice',
        '*Duration'                => 'GTC',
        '*Location'                => 'Bremen',
        'ShippingProfileName'      => 'DHL-Standardvorlage - (ID: 76770166018)',
        'ReturnProfileName'        => 'Standard-Widerruf 30 Tage für eBay Plus - (ID: 247364069018)',
        'PaymentProfileName'       => 'UK Gutschein',
        'PicURL'                   => resolveImage($product),
        'C:Marke'                  => $brand,
        'C:Hersteller'             => resolveAttr($product, 'manufacturer', 'brand', 'hersteller'),
        'C:Produktart'             => $firstWord,
        'C:Produkt'                => $firstWord,
        'C:Farbe'                  => resolveAttr($product, 'color', 'colour', 'farbe'),
        'C:Material'               => resolveAttr($product, 'material', 'material_type'),
        'C:Markenkompatibilität'   => resolveAttr($product, 'compatible_brands', 'brand_compatibility', 'fits_brands', 'markenkompatibilitaet'),
        'C:Modellkompatibilität'   => resolveCompatibleModels($product),
        'C:Herstellernummer'       => $sku['herstellernummer'],
        'C:Installationsart'       => 'Vollintegriert',
    ];
}

// ─── Column letter map (must match flat_template.xlsx) ─────────────────────
const COL_MAP = [
    '*Action'                  => 'A',
    'Custom label (SKU)'       => 'B',
    '*Category'                => 'C',
    '*Title'                   => 'D',
    '*ConditionID'             => 'E',
    'ConditionDescription'     => 'F',
    '*StartPrice'              => 'G',
    '*Quantity'                => 'H',
    '*Format'                  => 'I',
    '*Duration'                => 'J',
    '*Location'                => 'K',
    'ShippingProfileName'      => 'L',
    'ReturnProfileName'        => 'M',
    'PaymentProfileName'       => 'N',
    'PicURL'                   => 'O',
    'C:Marke'                  => 'P',
    'C:Hersteller'             => 'Q',
    'C:Produktart'             => 'R',
    'C:Produkt'                => 'S',
    'C:Farbe'                  => 'T',
    'C:Material'               => 'U',
    'C:Markenkompatibilität'   => 'V',
    'C:Modellkompatibilität'   => 'W',
    'C:Herstellernummer'       => 'X',
    'C:Installationsart'       => 'Y',
];

// ─── Entry point ───────────────────────────────────────────────────────────
$useSample = in_array('--sample', $argv ?? [], true);

if ($useSample) {
    $sampleFile = __DIR__ . '/sample_products.json';
    if (!file_exists($sampleFile)) {
        fwrite(STDERR, "ERROR: sample_products.json not found.\n");
        exit(1);
    }
    $products = json_decode(file_get_contents($sampleFile), true);
    echo "Using sample data from sample_products.json (" . count($products) . " products)\n";
} else {
    echo "Fetching all products from YSELL API...\n";
    try {
        $products = fetchAllProducts($baseUrl, $apiKey);
    } catch (RuntimeException $e) {
        fwrite(STDERR, "ERROR: " . $e->getMessage() . "\n");
        exit(1);
    }
}

echo "Total products fetched: " . count($products) . "\n";

// Save raw JSON
file_put_contents(__DIR__ . '/products_raw.json', json_encode($products, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE));
echo "Raw response saved to products_raw.json\n";

// Load template
$spreadsheet = IOFactory::load(__DIR__ . '/flat_template.xlsx');
$sheet       = $spreadsheet->getActiveSheet();

// Fill from row 2
$row = 2;
foreach ($products as $product) {
    $mapped = mapProductToRow($product);
    foreach ($mapped as $header => $value) {
        $col = COL_MAP[$header] ?? null;
        if ($col !== null) {
            $sheet->setCellValue($col . $row, (string)$value);
        }
    }
    $row++;
}

// Save filled file
$writer = new Xlsx($spreadsheet);
$writer->save(__DIR__ . '/flat_filled.xlsx');

$count = $row - 2;
echo "Готово! Записано товаров: {$count}\n";
