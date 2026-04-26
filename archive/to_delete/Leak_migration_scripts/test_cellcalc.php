<?php
// Quick test: check if build_dashboard.php has the cellCalc fix
error_reporting(E_ALL);
ini_set('display_errors', 1);
header('Content-Type: application/json');

$buildFile = __DIR__ . DIRECTORY_SEPARATOR . 'build_dashboard.php';
$buildContent = file_get_contents($buildFile);

// Check for the fix markers
$hasOldCalc = strpos($buildContent, 'getOldCalculatedValue') !== false;
$hasCatchThrowable = strpos($buildContent, 'Throwable') !== false;

// Extract the cellCalc function
$funcStart = strpos($buildContent, 'function cellCalc(');
$funcBody = '';
if ($funcStart !== false) {
    // Get 800 chars from function start
    $funcBody = substr($buildContent, $funcStart, 800);
}

// Also check the api.php cellCalc
$apiFile = __DIR__ . DIRECTORY_SEPARATOR . 'api.php';
$apiContent = file_get_contents($apiFile);
$apiHasOldCalc = strpos($apiContent, 'getOldCalculatedValue') !== false;
$apiFuncStart = strpos($apiContent, 'function cellCalc(');
$apiFuncBody = '';
if ($apiFuncStart !== false) {
    $apiFuncBody = substr($apiContent, $apiFuncStart, 800);
}

echo json_encode([
    'build' => [
        'hasOldCalcValue' => $hasOldCalc,
        'hasThrowable' => $hasCatchThrowable,
        'cellCalcSnippet' => substr($funcBody, 0, 500),
    ],
    'api' => [
        'hasOldCalcValue' => $apiHasOldCalc,
        'cellCalcSnippet' => substr($apiFuncBody, 0, 500),
    ],
], JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT);
