<?php
/**
 * Reenhanced Azure Service Tags Proxy
 *
 * Drop this file alongside the widget on your WordPress site (or any PHP host).
 * It proxies the Azure Service Tags JSON from download.microsoft.com to avoid
 * browser CORS restrictions. All processing still happens in the browser.
 *
 * Usage: proxy.php?url=download/7/1/d/71d86715-5596-4529-9b13-da13a5de5b63/ServiceTags_Public_20260406.json
 */

// Only allow GET
if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') {
    header('Access-Control-Allow-Origin: *');
    header('Access-Control-Allow-Methods: GET, OPTIONS');
    header('Access-Control-Max-Age: 86400');
    http_response_code(204);
    exit;
}

if ($_SERVER['REQUEST_METHOD'] !== 'GET') {
    http_response_code(405);
    echo 'Method not allowed';
    exit;
}

$path = isset($_GET['url']) ? $_GET['url'] : '';

// Validate: only allow paths under download.microsoft.com that match the ServiceTags pattern
if (!preg_match('#^download/[0-9a-f]/[0-9a-f]/[0-9a-f]/[0-9a-f\-]+/ServiceTags_Public_\d{8}\.json$#', $path)) {
    http_response_code(400);
    echo 'Invalid path. Only ServiceTags downloads are allowed.';
    exit;
}

// Hardcode the target domain — never let user input control the host
$url = 'https://download.microsoft.com/' . $path;

$ch = curl_init($url);
curl_setopt_array($ch, [
    CURLOPT_RETURNTRANSFER => true,
    CURLOPT_FOLLOWLOCATION => false,
    CURLOPT_TIMEOUT        => 30,
    CURLOPT_USERAGENT      => 'ReenhancedAzureCFWidget/1.0',
]);

$body = curl_exec($ch);
$code = curl_getinfo($ch, CURLINFO_HTTP_CODE);
curl_close($ch);

if ($body === false || $code !== 200) {
    http_response_code(502);
    echo 'Failed to fetch from Microsoft (HTTP ' . $code . ')';
    exit;
}

// Validate the response is actually JSON before passing it through
json_decode($body);
if (json_last_error() !== JSON_ERROR_NONE) {
    http_response_code(502);
    echo 'Upstream response was not valid JSON.';
    exit;
}

header('Content-Type: application/json');
header('Access-Control-Allow-Origin: *');
header('Cache-Control: public, max-age=3600');
echo $body;
