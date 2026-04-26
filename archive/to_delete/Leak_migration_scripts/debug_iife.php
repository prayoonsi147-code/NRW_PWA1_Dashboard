<?php
// Extract the Dual Mode IIFE from the actual served index.html
$html = file_get_contents(__DIR__ . '/index.html');
$lines = explode("\n", $html);

echo "Total lines: " . count($lines) . "\n\n";

// Find embedded data lines
echo "=== EMBEDDED DATA LINES ===\n";
for ($i = 0; $i < count($lines); $i++) {
    if (preg_match('/^var (D|RL|EU|MNF|P3|KPI2)\s*=/', $lines[$i])) {
        echo "Line " . ($i+1) . ": " . substr($lines[$i], 0, 80) . "... (" . strlen($lines[$i]) . " chars)\n";
    }
}

echo "\n=== DUAL MODE IIFE ===\n";
// Find the IIFE
$iifeStart = null;
$iifeEnd = null;
for ($i = 0; $i < count($lines); $i++) {
    if (strpos($lines[$i], 'Dual Mode: Try API first') !== false) {
        $iifeStart = $i;
    }
    if ($iifeStart !== null && $i > $iifeStart + 2 && preg_match('/^\}\)\(\);?\s*$/', trim($lines[$i]))) {
        $iifeEnd = $i;
        break;
    }
}

if ($iifeStart !== null) {
    echo "IIFE starts at line " . ($iifeStart+1) . ", ends at line " . ($iifeEnd+1) . "\n";
    echo "IIFE content:\n";
    for ($i = $iifeStart; $i <= $iifeEnd; $i++) {
        echo "L" . ($i+1) . ": " . $lines[$i] . "\n";
    }
} else {
    echo "IIFE not found!\n";
}

// Find what comes after the IIFE
echo "\n=== AFTER IIFE (next 10 lines) ===\n";
if ($iifeEnd !== null) {
    for ($i = $iifeEnd + 1; $i <= min($iifeEnd + 10, count($lines) - 1); $i++) {
        echo "L" . ($i+1) . ": " . substr($lines[$i], 0, 120) . "\n";
    }
}

// Check what allData is and how it's used
echo "\n=== allData references ===\n";
$count = 0;
for ($i = 0; $i < count($lines); $i++) {
    if (strpos($lines[$i], 'allData') !== false && $count < 15) {
        echo "Line " . ($i+1) . ": " . substr($lines[$i], 0, 150) . "\n";
        $count++;
    }
}

// Find the init/DOMContentLoaded or other chart init
echo "\n=== Chart init patterns ===\n";
$patterns = ['DOMContentLoaded', 'initChart', 'initDashboard', 'renderChart', 'buildChart', 'switchMainTab'];
foreach ($patterns as $p) {
    for ($i = 0; $i < count($lines); $i++) {
        if (strpos($lines[$i], $p) !== false) {
            echo "Line " . ($i+1) . " [$p]: " . substr($lines[$i], 0, 120) . "\n";
        }
    }
}
