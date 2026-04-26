<?php
error_reporting(E_ALL);
ini_set('display_errors', 1);
header('Content-Type: application/json');
set_time_limit(120);

// Try multiple autoload paths
$paths = [
    __DIR__ . '/vendor/autoload.php',
    __DIR__ . '/../vendor/autoload.php',
    'C:/xampp/htdocs/vendor/autoload.php',
];
$loaded = false;
foreach ($paths as $p) {
    if (file_exists($p)) { require_once $p; $loaded = true; break; }
}

// Also try the build_dashboard.php loader
if (!$loaded) {
    // Include the load function from build_dashboard
    $buildContent = file_get_contents(__DIR__ . '/build_dashboard.php');
    if (preg_match('/function load_phsspreadsheet\(\).*?^}/ms', $buildContent, $m)) {
        eval($m[0]);
        $loaded = load_phsspreadsheet();
    }
}

if (!$loaded) {
    // Try to find composer autoload
    $candidates = glob(__DIR__ . '/*/autoload.php');
    foreach ($candidates as $c) {
        require_once $c;
        $loaded = true;
        break;
    }
}

if (!$loaded || !class_exists('\\PhpOffice\\PhpSpreadsheet\\IOFactory')) {
    echo json_encode(['error' => 'PhpSpreadsheet not loaded', 'tried' => $paths]);
    exit;
}

$rlFolder = __DIR__ . DIRECTORY_SEPARATOR . 'ข้อมูลดิบ' . DIRECTORY_SEPARATOR . 'Real Leak';
$files = glob($rlFolder . DIRECTORY_SEPARATOR . 'RL_2568.*');
if (empty($files)) {
    echo json_encode(['error' => 'No RL_2568 file found', 'folder' => $rlFolder, 'contents' => scandir($rlFolder)]);
    exit;
}

$rlFile = $files[0];
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($rlFile);
$sheet = $spreadsheet->getSheet(0);
$sheetName = $sheet->getTitle();

// Volume col = 7, Rate col = 8 (based on build log col_map)
$results = [];
for ($r = 4; $r <= 6; $r++) {
    $branch = trim((string)($sheet->getCell([2, $r])->getValue() ?? ''));

    // Check volume cell (col 7)
    $vCell = $sheet->getCell([7, $r]);
    $vRaw = $vCell->getValue();
    $vIsFormula = is_string($vRaw) && strlen($vRaw) > 0 && $vRaw[0] === '=';

    $vOldCalc = null;
    $vCalcVal = null;
    try { $vOldCalc = $vCell->getOldCalculatedValue(); } catch (\Throwable $e) { $vOldCalc = 'ERR'; }
    try { $vCalcVal = $vCell->getCalculatedValue(); } catch (\Throwable $e) { $vCalcVal = 'ERR'; }

    // Also check production cell (col 3) and supplied cell (col 4) and sold cell (col 5)
    $pCell = $sheet->getCell([3, $r]);
    $pRaw = $pCell->getValue();
    $pIsFormula = is_string($pRaw) && strlen($pRaw) > 0 && $pRaw[0] === '=';
    $pOldCalc = null;
    try { $pOldCalc = $pCell->getOldCalculatedValue(); } catch (\Throwable $e) { $pOldCalc = 'ERR'; }
    try { $pCalcVal = $pCell->getCalculatedValue(); } catch (\Throwable $e) { $pCalcVal = 'ERR'; }

    $results[] = [
        'row' => $r,
        'branch' => mb_substr($branch, 0, 15),
        'v_raw' => is_string($vRaw) ? mb_substr($vRaw, 0, 60) : $vRaw,
        'v_isFormula' => $vIsFormula,
        'v_oldCalc' => $vOldCalc,
        'v_calcVal' => $vCalcVal,
        'p_raw' => is_string($pRaw) ? mb_substr($pRaw, 0, 60) : $pRaw,
        'p_isFormula' => $pIsFormula,
        'p_oldCalc' => $pOldCalc,
        'p_calcVal' => $pCalcVal,
    ];
}

echo json_encode([
    'file' => basename($rlFile),
    'sheet' => $sheetName,
    'cells' => $results,
], JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT);
