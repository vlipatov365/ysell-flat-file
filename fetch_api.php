<?php
$apiKey = '149185983e29fed80f5010edd60410b7';
$baseUrl = 'https://4457.test1.ysell.pro';

function apiGet(string $url, string $apiKey): array {
    $ch = curl_init($url);
    curl_setopt_array($ch, [
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_HTTPHEADER => [
            'Authorization: Bearer ' . $apiKey,
            'Content-Type: application/json',
            'Accept: application/json',
        ],
        CURLOPT_SSL_VERIFYPEER => false,
        CURLOPT_TIMEOUT => 30,
    ]);
    $response = curl_exec($ch);
    $httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    $error = curl_error($ch);
    curl_close($ch);

    if ($error) {
        return ['error' => $error, 'http_code' => $httpCode];
    }
    return ['body' => $response, 'http_code' => $httpCode];
}

// Step 1: Get first product (list)
echo "=== GET /api/v1/product?per_page=1 ===\n";
$result = apiGet($baseUrl . '/api/v1/product?per_page=1', $apiKey);
echo "HTTP Code: " . $result['http_code'] . "\n";
echo $result['body'] ?? ('Error: ' . $result['error']);
echo "\n\n";

$data = json_decode($result['body'], true);
if (!empty($data['data'][0])) {
    $product = $data['data'][0];
    echo "=== Keys of first product ===\n";
    echo implode(', ', array_keys($product)) . "\n\n";

    // Step 2: Get single product by ID
    $id = $product['id'];
    echo "=== GET /api/v1/product/{$id} ===\n";
    $single = apiGet($baseUrl . '/api/v1/product/' . $id, $apiKey);
    echo "HTTP Code: " . $single['http_code'] . "\n";
    echo $single['body'] ?? ('Error: ' . $single['error']);
    echo "\n";
}
