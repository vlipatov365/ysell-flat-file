<?php
declare(strict_types=1);

require_once __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// ─── Config ────────────────────────────────────────────────────────────────
$apiKey  = getenv('YSELL_API_KEY')  ?: '149185983e29fed80f5010edd60410b7';
$baseUrl = getenv('YSELL_BASE_URL') ?: 'https://4457.test1.ysell.pro';

// ─── HTTP: YSELL API ───────────────────────────────────────────────────────
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

    if ($error) throw new RuntimeException("cURL error: {$error}");
    if ($code >= 400) throw new RuntimeException("HTTP {$code}: {$body}");
    $data = json_decode($body, true);
    if ($data === null) throw new RuntimeException("Invalid JSON: {$body}");
    return $data;
}

// ─── HTTP: plain HTML fetch (eBay scraping) ────────────────────────────────
function httpGet(string $url): string
{
    $ch = curl_init($url);
    curl_setopt_array($ch, [
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_SSL_VERIFYPEER => false,
        CURLOPT_TIMEOUT        => 15,
        CURLOPT_FOLLOWLOCATION => true,
        CURLOPT_HTTPHEADER     => [
            'Accept: text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language: de-DE,de;q=0.9,en;q=0.7',
            'User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
            'Accept-Encoding: gzip, deflate',
            'Cache-Control: no-cache',
        ],
        CURLOPT_ENCODING       => '',   // auto-decompress gzip
        CURLOPT_COOKIEJAR      => sys_get_temp_dir() . '/ebay_cookie.txt',
        CURLOPT_COOKIEFILE     => sys_get_temp_dir() . '/ebay_cookie.txt',
    ]);
    $html  = curl_exec($ch);
    $error = curl_error($ch);
    curl_close($ch);
    return (!$error && $html) ? (string)$html : '';
}

// ─── Fetch all products (paginated, with productSuppliers) ─────────────────
function fetchAllProducts(string $baseUrl, string $apiKey): array
{
    $products = [];
    $page     = 1;
    do {
        echo "  Fetching page {$page}...\n";
        // Include productSuppliers relation so we can resolve price via supplier endpoint
        $url      = "{$baseUrl}/api/v1/product?per_page=50&page={$page}&with[]=productSuppliers";
        $response = apiGet($url, $apiKey);
        $items    = $response['data'] ?? $response;
        $lastPage = $response['meta']['last_page'] ?? $response['last_page'] ?? 1;
        if (!is_array($items) || empty($items)) break;
        $products = array_merge($products, $items);
        $page++;
    } while ($page <= $lastPage);
    return $products;
}

// ─── Fetch manufacturers: [id => name] ─────────────────────────────────────
function fetchManufacturers(string $baseUrl, string $apiKey): array
{
    try {
        $response = apiGet("{$baseUrl}/api/v1/manufacturer", $apiKey);
        $items    = $response['data'] ?? $response;
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

// ─── Category: Variant A — keyword map ────────────────────────────────────
function getCategoryByKeywords(string $title): string
{
    $t = mb_strtolower($title, 'UTF-8');

    // More specific rules first
    $rules = [
        '116026' => ['geschirrspüler', 'geschirrspueler', 'spülmaschine', 'spuelmachine',
                     'sprüharm', 'sprüharm', 'seilzug', 'reedplatine'],
        '99697'  => ['waschmaschine', 'trockner', 'waschtrockner', 'schleuder',
                     'flusensieb', 'spannrolle', 'bottich', 'waschmitteleinsatz',
                     'flüssigwaschmittel', 'zweibandmotor'],
        '184666' => ['kühlschrank', 'kühlgerät', 'gefriergerät', 'gefrierschrank',
                     'gefrierraum', 'kühl-gefrier', 'kühlaggregat', 'türfach'],
        '43566'  => ['backofen', 'backrohr', 'herd', 'backblech', 'grillpfanne',
                     'fettpfanne', 'oberhitze', 'unterhitze', 'thermostat',
                     'türinnengitter', 'backofenthermostat'],
        '71253'  => ['dunstabzugshaube', 'dunstabzug', 'abzugshaube', 'dampfabzug',
                     'lampenabdeckung'],
        '99565'  => ['kaffeemaschine', 'kaffeevollautomat', 'espresso', 'kaffeeautomat',
                     'heißwasserdüse', 'staubsauger', 'filterbeutel', 'sauger',
                     'mehrzwecksauger', 'brotbackautomat', 'brotbackmaschine',
                     'fleischwolf', 'küchenmaschine', 'küchmaschine', 'knethaken', 'zahnrad'],
    ];

    foreach ($rules as $categoryId => $keywords) {
        foreach ($keywords as $kw) {
            if (str_contains($t, $kw)) {
                return (string)$categoryId;
            }
        }
    }

    return '';
}

// ─── Category + Material: eBay.de scraping ────────────────────────────────
// Returns ['category' => '...', 'material' => '...']
// Both fields are fetched in a single pass and cached together.
$_ebayDataCache = [];

function scrapeEbayData(string $title): array
{
    global $_ebayDataCache;

    $cacheKey = md5($title);
    if (isset($_ebayDataCache[$cacheKey])) {
        return $_ebayDataCache[$cacheKey];
    }

    $data = ['category' => '', 'material' => ''];

    // Step 1: search eBay.de
    $searchUrl  = 'https://www.ebay.de/sch/i.html?_nkw=' . urlencode($title) . '&_sacat=0';
    $searchHtml = httpGet($searchUrl);

    if ($searchHtml === '') {
        echo "    eBay scraping: no response from search.\n";
        $_ebayDataCache[$cacheKey] = $data;
        return $data;
    }

    // Step 2: extract first product page URL from search results
    preg_match('/href="(https:\/\/www\.ebay\.de\/itm\/[^"?]+)"/', $searchHtml, $itemMatch);
    $itemUrl = $itemMatch[1] ?? '';

    if ($itemUrl !== '') {
        // Step 3: fetch the product page
        $itemHtml = httpGet($itemUrl);

        if ($itemHtml !== '') {
            // Step 4: parse category from JSON-LD BreadcrumbList
            if (preg_match_all(
                '/<script[^>]+type=["\']application\/ld\+json["\'][^>]*>(.*?)<\/script>/si',
                $itemHtml,
                $scripts
            )) {
                foreach ($scripts[1] as $jsonRaw) {
                    $jsonData = json_decode(trim($jsonRaw), true);
                    if (
                        is_array($jsonData)
                        && ($jsonData['@type'] ?? '') === 'BreadcrumbList'
                        && !empty($jsonData['itemListElement'])
                    ) {
                        $lastItem = end($jsonData['itemListElement']);
                        $itemHref = $lastItem['item'] ?? '';
                        if (preg_match('/\/b\/[^\/]+\/(\d{4,7})\//', $itemHref, $idMatch)) {
                            $data['category'] = $idMatch[1];
                        }
                        break;
                    }
                }
            }

            // Fallback: scan /b/Name/ID/ hrefs in breadcrumb area, take last
            if ($data['category'] === '') {
                preg_match_all(
                    '/href="https:\/\/www\.ebay\.de\/b\/[^\/]+\/(\d{4,7})\/[^"]*"/',
                    $itemHtml,
                    $allCats
                );
                if (!empty($allCats[1])) {
                    $data['category'] = end($allCats[1]);
                }
            }

            // Step 5: extract Material from <dl class="...ux-labels-values--material...">
            // Structure: <dl class="...ux-labels-values--material...">
            //              <dt>...<span>Material</span>...</dt>
            //              <dd>...<span>VALUE</span>...</dd>
            //            </dl>
            if (preg_match(
                '/<dl[^>]+ux-labels-values--material[^>]*>.*?<dd[^>]*>.*?<span[^>]*>(.*?)<\/span>/si',
                $itemHtml,
                $matMatch
            )) {
                $data['material'] = trim(strip_tags($matMatch[1]));
            }
        }
    }

    // Last resort for category: _sacat= in search results
    if ($data['category'] === '') {
        preg_match_all('/_sacat=(\d{4,6})\b/', $searchHtml, $sacatMatches);
        foreach ($sacatMatches[1] as $cat) {
            if ($cat !== '0') {
                $data['category'] = $cat;
                break;
            }
        }
    }

    $_ebayDataCache[$cacheKey] = $data;
    return $data;
}

function getCategoryByEbayScraping(string $title): string
{
    return scrapeEbayData($title)['category'];
}

function getMaterialByEbayScraping(string $title): string
{
    return scrapeEbayData($title)['material'];
}

// ─── Category: combined resolver (A then B) ────────────────────────────────
function resolveCategory(string $title): string
{
    $cat = getCategoryByKeywords($title);
    if ($cat !== '') {
        return $cat;
    }

    echo "  → Keyword miss, scraping eBay for: " . mb_substr($title, 0, 60, 'UTF-8') . "\n";
    $cat = getCategoryByEbayScraping($title);

    if ($cat === '') {
        echo "    → eBay scraping returned no category, leaving blank.\n";
    }

    return $cat;
}

// ─── Parse title for brands and OEM numbers ───────────────────────────────
function parseTitleForCompatibility(string $title): array
{
    $brands = [];
    $models = [];

    if (preg_match('/\bwie\s+(.+?)(?:\s+(?:für|in|an|von|mit|zu)\b|$)/u', $title, $wie)) {
        $segment = $wie[1];
        preg_match_all('/\b[A-Z0-9]*\d[A-Z0-9.\-\/]*\b/u', $segment, $numMatches);
        $models = array_slice(array_unique($numMatches[0]), 0, 3);
        preg_match_all('/\b([A-ZÄÖÜ][A-Za-zäöüÄÖÜß]*)\b/u', $segment, $brandMatches);
        $brands = array_unique(array_filter($brandMatches[1], fn($w) => !preg_match('/\d/', $w)));
    }

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

// Cache supplier product responses: [supplierNum => price_string]
$_supplierPriceCache = [];

function fetchSupplierPrice(int|string $supplierNum, string $baseUrl, string $apiKey): string
{
    global $_supplierPriceCache;

    $key = (string)$supplierNum;
    if (array_key_exists($key, $_supplierPriceCache)) {
        return $_supplierPriceCache[$key];
    }

    try {
        $response = apiGet("{$baseUrl}/api/v1/supplier/{$supplierNum}/product", $apiKey);
        // Response may be wrapped or direct; price field may vary
        $data  = $response['data'] ?? $response;
        $price = $data['price'] ?? $data['purchase_price'] ?? $data['cost']
              ?? $data['netto']  ?? null;

        if ($price !== null && (float)$price > 0) {
            $result = number_format((float)$price, 2, '.', '');
            $_supplierPriceCache[$key] = $result;
            return $result;
        }
    } catch (RuntimeException) {
        // supplier endpoint failed — fall through to ''
    }

    $_supplierPriceCache[$key] = '';
    return '';
}

function resolvePrice(array $product, string $baseUrl, string $apiKey): string
{
    // Level 1: purchase_price directly on the product
    $direct = $product['purchase_price'] ?? null;
    if ($direct !== null && $direct !== '' && (float)$direct > 0) {
        return number_format((float)$direct, 2, '.', '');
    }

    // Level 2: productSuppliers relation → GET /api/v1/supplier/{num}/product
    $suppliers = $product['productSuppliers'] ?? [];
    foreach ($suppliers as $ps) {
        // supplierNum is the correct field per API spec; fallbacks for safety
        $num = $ps['supplierNum'] ?? $ps['supplier_id'] ?? $ps['supplier_num'] ?? $ps['id'] ?? null;
        if ($num === null) continue;

        $price = fetchSupplierPrice($num, $baseUrl, $apiKey);
        if ($price !== '') return $price;
    }

    return '';
}

// ─── SKU / Herstellernummer ─────────────────────────────────────────────────
function resolveSkuAndHerstellernummer(array $product, string $manufacturerName, array $compat): array
{
    $id    = (string)($product['id'] ?? '');
    $extId = (string)($product['ext_id'] ?? '');
    $isViox = stripos($manufacturerName, 'viox') !== false
           || stripos($manufacturerName, 'виокс') !== false;
    $oemNum = $compat['models'][0] ?? '';

    if ($isViox) {
        return ['custom_label' => "{$id}-MP", 'herstellernummer' => $oemNum];
    }
    return ['custom_label' => $extId ?: $id, 'herstellernummer' => $oemNum];
}

// ─── HTML Description ─────────────────────────────────────────────────────
function buildDescription(array $product, array $compat, string $manufacturerName): string
{
    $title    = htmlspecialchars((string)($product['title'] ?? ''), ENT_QUOTES, 'UTF-8');
    $imageUrl = htmlspecialchars((string)($product['image'] ?? ''), ENT_QUOTES, 'UTF-8');
    $brands   = $compat['brands'];
    $models   = $compat['models'];

    // Bullet 2: compatibility text
    $compatText = !empty($brands)
        ? 'Kompatibel mit Produkten der Marke(n): ' . htmlspecialchars(implode(', ', $brands), ENT_QUOTES, 'UTF-8')
        : 'Geeignet als Ersatzteil für viele gängige Haushaltsgeräte';

    // Bullet 1: OEM reference
    $oemRef = !empty($models)
        ? 'Passend als Ersatz für Artikel-Nr.: ' . htmlspecialchars(implode(', ', $models), ENT_QUOTES, 'UTF-8')
        : htmlspecialchars($title, ENT_QUOTES, 'UTF-8');

    // Description paragraph
    $descText = !empty($brands)
        ? 'Dieser Artikel ist ein hochwertiges Ersatzteil, das kompatibel mit Geräten der Marke(n) '
          . htmlspecialchars(implode(', ', $brands), ENT_QUOTES, 'UTF-8')
          . ' ist. Er entspricht den Original-Spezifikationen und eignet sich als direkter Austausch.'
        : 'Dieser Artikel ist ein hochwertiges Ersatzteil für Haushaltsgeräte. Er entspricht den Original-Spezifikationen und eignet sich als direkter Austausch.';

    // OEM numbers for "Kompatibel mit" block
    $oemLines = implode('<br/>', array_map(
        fn($m) => htmlspecialchars($m, ENT_QUOTES, 'UTF-8'),
        $models
    ));
    if ($oemLines === '') {
        $oemLines = '—';
    }

    $noImageHtml = $imageUrl !== ''
        ? "<img id=\"default\" src=\"{$imageUrl}\"/>"
        : "<p style=\"color:#999\">Kein Bild vorhanden</p>";

    return <<<HTML
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>eBay</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="https://ersatzteil-check.de/ebay_template/css/style.css">
    <link rel="stylesheet" href="https://ersatzteil-check.de/ebay_template/css/new_style.css">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet" crossorigin="anonymous">
</head>
<body>
<div class="ebay-content">
    <header class="container header">
        <img src="https://ersatzteil-check.de/ebay_template/img/logo.svg">
        <div class="menu">
            <a class="menu-item" href="https://www.ebay.de/str/bremenhaushalt">Über uns</a>
            <a class="menu-item" href="https://www.ebay.de/cnt/intermediatedFAQ?requested=ersatzteil-check-de">Kontakt</a>
            <a class="menu-item" href="https://www.ebay.de/str/bremenhaushalt?_tab=2">Bewertungen</a>
            <a class="menu-item-btn" href="https://www.ebay.de/str/bremenhaushalt">eBay Shop</a>
        </div>
    </header>
    <div class="container">
        <img class="banner" src="https://ersatzteil-check.de/ebay_template/img/banner.png">
    </div>
    <div class="container gears">
        <div class="column-3">
            <img src="https://ersatzteil-check.de/ebay_template/img/Group1.png">
            <p>Schnelle Lieferung</p>
        </div>
        <div class="column-3">
            <img src="https://ersatzteil-check.de/ebay_template/img/Group2.png">
            <p>Riesige Auswahl</p>
        </div>
        <div class="column-3">
            <img src="https://ersatzteil-check.de/ebay_template/img/Group3.png">
            <p>30 Tage Widerruffrist</p>
        </div>
    </div>
    <footer class="footer">
        <div class="container">
            <div class="footer-menu">
                <a class="footer-item" href="https://www.ebay.de/sch/i.html?_ssn=ersatzteil-check-de&store_name=bremenhaushalt&_oac=1&_trksid=p4429486.m3561.l2562">Alle Artikel</a>
                <a class="footer-item" href="https://www.ebay.de/str/bremenhaushalt/Backofen-Herdteile/_i.html?_sacat=43566">Backofen- &amp; Herdeteile</a>
                <a class="footer-item" href="https://www.ebay.de/str/bremenhaushalt/Waschmaschinen-Trocknerteile/_i.html?_sacat=99697">Waschmaschinen- &amp; Trocknerteile</a>
                <a class="footer-item" href="https://www.ebay.de/str/bremenhaushalt/Gefriergerate-Kuhlschrankteile/_i.html?_sacat=184666">Gefriergeräte- &amp; Kühlschrankteile</a>
                <a class="footer-item" href="https://www.ebay.de/str/bremenhaushalt/Geschirrspulerteile/_i.html?_sacat=116026">Geschirrspülerteile</a>
                <a class="footer-item" href="https://www.ebay.de/str/bremenhaushalt/Ersatzteile/_i.html?_sacat=99565">Kaffeemaschinen Ersatzteile</a>
                <a class="footer-item" href="https://www.ebay.de/str/bremenhaushalt/Dunstabzugshauben/_i.html?_sacat=71253">Dunstabzugshauben Ersatzteile</a>
            </div>
        </div>
    </footer>
</div>
<div class="container-fluid content-bg" align="center">
    <br/>
    <div class="container">
        <div class="row">
            <div class="col-md-12 maincontent">
                <div class="row top_head">
                    <h1><b style="color:#ff0000">{$title}</b></h1>
                </div>
                <div class="textbox main">
                    <div class="row artdesc">
                        <div class="col-md-6 artdesc-2">
                            <div class="image-gallery">
                                <div class="big-image">
                                    {$noImageHtml}
                                </div>
                            </div>
                        </div>
                        <div class="col-md-6 artdesc-3">
                            <div class="right_info">
                                <div class="row">
                                    <div class="col">
                                        <a href="#" class="button-1" id="ersatz-tocart">
                                            <img src="https://ersatzteil-check.de/ebay_template/img/cart.svg" alt="">
                                            Kaufen
                                        </a>
                                    </div>
                                    <div class="col">
                                        <a href="#" class="button-2" id="ersatz-towatch">
                                            <img src="https://ersatzteil-check.de/ebay_template/img/globalization.svg" alt="">
                                            Beobachten
                                        </a>
                                    </div>
                                    <div class="col">
                                        <a href="#" class="button-3" id="ersatz-toask">
                                            <img src="https://ersatzteil-check.de/ebay_template/img/support.svg" alt="">
                                            Frage stellen
                                        </a>
                                    </div>
                                </div>
                                <ul style="list-style-type:none; padding-left:0;">
                                    <li style="color:#ff0000;font-weight:700">&#9989; <b>Originalgetreue Passform:</b> {$oemRef}</li>
                                    <li style="color:#ff0000;font-weight:700">&#9989; <b>Vielseitig einsetzbar:</b> {$compatText}</li>
                                    <li style="color:#ff0000;font-weight:700">&#9989; <b>Geprüfte Qualität:</b> Hochwertige Verarbeitung in vergleichbarer Hersteller-Qualität</li>
                                    <li style="color:#ff0000;font-weight:700">&#9989; <b>Schnelle Lösung:</b> Direkter Austausch ohne Modifikation möglich</li>
                                    <li style="color:#ff0000;font-weight:700">&#9989; <b>Sind Sie unsicher?</b> Wir helfen Ihnen gerne, das exakt benötigte Ersatzteil zu finden.</li>
                                </ul>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="row bottom_text">
                    <div class="col-md-12">
                        <div class="tbc-1">
                            <h2>Beschreibung</h2>
                            <b style="color:#ff0000">{$title}</b><br/><br/>
                            <span style="color:#ff0000">{$descText}</span><br/><br/>
                            Wir sind ein professioneller Hersteller und Vertreiber von Original-Ersatzteilen und Alternativprodukten für Haushaltsgeräte. Egal ob Waschmaschine, Herd oder Kühlschrank: Wir haben diverse preisgünstige Original- und Alternativteile in vergleichbarer Hersteller-Qualität im Angebot. Überzeugen Sie sich selbst von der Wertigkeit unserer Ersatzteile und verschaffen Sie Ihrem Haushaltsgerät ein zweites Leben.<br/><br/>
                            <b>Kompatibel mit:</b><br/>
                            <span style="color:#ff0000">{$oemLines}</span>
                        </div>
                    </div>
                </div>
                <div class="row bottom_text">
                    <div class="col-md-12">
                        <div class="tbc-2">
                            <h2>Zahlungsoptionen</h2>
                            <img src="https://ersatzteil-check.de/ebay_template/img/zahlung-paypal_n.png">
                            <img src="https://ersatzteil-check.de/ebay_template/img/zahlung-ueberweisung_n.png">
                        </div>
                        <div class="tbc-3">
                            <h2>Versandoptionen</h2>
                            <img src="https://ersatzteil-check.de/ebay_template/img/versand-dhl_n.png">
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
</div>
</body>
</html>
HTML;
}

// ─── Main mapping ──────────────────────────────────────────────────────────
function mapProductToRow(array $product, array $manufacturers, string $baseUrl, string $apiKey): array
{
    $manufacturerId   = (int)($product['manufacturer_id'] ?? 0);
    $manufacturerName = $manufacturers[$manufacturerId] ?? '';

    $condition = resolveCondition($product);
    $compat    = parseTitleForCompatibility((string)($product['title'] ?? ''));
    $sku       = resolveSkuAndHerstellernummer($product, $manufacturerName, $compat);

    $rawTitle  = (string)($product['title'] ?? '');
    $title     = mb_substr($rawTitle, 0, 80, 'UTF-8');
    $firstWord = explode(' ', trim($title))[0] ?? '';

    $category    = resolveCategory($rawTitle);
    $description = buildDescription($product, $compat, $manufacturerName);

    return [
        '*Action'                => 'Add (SiteID=Germany|Country=DE|Currency=EUR|Version=941)',
        'Custom label (SKU)'     => $sku['custom_label'],
        '*Category'              => $category,
        'Description'            => $description,
        '*Title'                 => $title,
        '*ConditionID'           => $condition['id'],
        'ConditionDescription'   => $condition['description'],
        '*StartPrice'            => resolvePrice($product, $baseUrl, $apiKey),
        '*Quantity'              => '1',
        '*Format'                => 'FixedPrice',
        '*Duration'              => 'GTC',
        '*Location'              => 'Bremen',
        'ShippingProfileName'    => 'DHL-Standardvorlage - (ID: 76770166018)',
        'ReturnProfileName'      => 'Standard-Widerruf 30 Tage für eBay Plus - (ID: 247364069018)',
        'PaymentProfileName'     => 'UK Gutschein',
        'PicURL'                 => (string)($product['image'] ?? ''),
        'C:Marke'                => $manufacturerName,
        'C:Hersteller'           => $manufacturerName,
        'C:Produktart'           => $firstWord,
        'C:Produkt'              => $firstWord,
        'C:Farbe'                => '',
        'C:Material'             => getMaterialByEbayScraping($rawTitle),
        'C:Markenkompatibilität' => implode(', ', $compat['brands']),
        'C:Modellkompatibilität' => implode(', ', $compat['models']),
        'C:Herstellernummer'     => $sku['herstellernummer'],
        'C:Installationsart'     => 'Vollintegriert',
    ];
}

// ─── Column map ────────────────────────────────────────────────────────────
const COL_MAP = [
    '*Action'                => 'A',
    'Custom label (SKU)'     => 'B',
    '*Category'              => 'C',
    'Description'            => 'D',
    '*Title'                 => 'E',
    '*ConditionID'           => 'F',
    'ConditionDescription'   => 'G',
    '*StartPrice'            => 'H',
    '*Quantity'              => 'I',
    '*Format'                => 'J',
    '*Duration'              => 'K',
    '*Location'              => 'L',
    'ShippingProfileName'    => 'M',
    'ReturnProfileName'      => 'N',
    'PaymentProfileName'     => 'O',
    'PicURL'                 => 'P',
    'C:Marke'                => 'Q',
    'C:Hersteller'           => 'R',
    'C:Produktart'           => 'S',
    'C:Produkt'              => 'T',
    'C:Farbe'                => 'U',
    'C:Material'             => 'V',
    'C:Markenkompatibilität' => 'W',
    'C:Modellkompatibilität' => 'X',
    'C:Herstellernummer'     => 'Y',
    'C:Installationsart'     => 'Z',
];

// ─── Entry point ───────────────────────────────────────────────────────────
$useSample = in_array('--sample', $argv ?? [], true);

if ($useSample) {
    $products      = json_decode(file_get_contents(__DIR__ . '/sample_products.json'), true);
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

file_put_contents(
    __DIR__ . '/products_raw.json',
    json_encode($products, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE)
);
echo "Saved products_raw.json\n";

$spreadsheet = IOFactory::load(__DIR__ . '/flat_template.xlsx');
$sheet       = $spreadsheet->getActiveSheet();

$row      = 2;
$nocat    = 0;
foreach ($products as $product) {
    $mapped = mapProductToRow($product, $manufacturers, $baseUrl, $apiKey);
    foreach ($mapped as $header => $value) {
        $col = COL_MAP[$header] ?? null;
        if ($col !== null) {
            $sheet->setCellValue($col . $row, (string)$value);
        }
    }
    if ($mapped['*Category'] === '') $nocat++;
    $row++;
}

$writer = new Xlsx($spreadsheet);
$writer->save(__DIR__ . '/flat_filled.xlsx');

$count = $row - 2;
echo "Готово! Записано товаров: {$count}\n";
if ($nocat > 0) {
    echo "Внимание: {$nocat} товаров без категории (заполни вручную колонку C).\n";
}
