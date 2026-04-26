<?php
$file = __DIR__ . '/index.html';
$lines = explode("\n", file_get_contents($file));
$changes = 0;

echo "Total lines: " . count($lines) . "\n\n";

// Show current state around the problem area
echo "=== Current state L2508-2525 ===\n";
for ($i = 2507; $i < min(2525, count($lines)); $i++) {
    echo "L" . ($i+1) . ": " . $lines[$i] . "\n";
}

echo "\n=== Current state L2548-2558 ===\n";
for ($i = 2547; $i < min(2558, count($lines)); $i++) {
    echo "L" . ($i+1) . ": " . $lines[$i] . "\n";
}

// Fix 1: Remove any remaining leftover }); after rebuildAllData line
for ($i = 0; $i < count($lines); $i++) {
    if (strpos($lines[$i], 'rebuildAllData(); // initial build') !== false) {
        // Remove any orphan lines immediately after
        $j = $i + 1;
        while ($j < count($lines) && (trim($lines[$j]) === '' || trim($lines[$j]) === '});' || trim($lines[$j]) === 'allData[y]=sheets;')) {
            if (trim($lines[$j]) !== '') {
                echo "\nFix 1: Removing leftover line L" . ($j+1) . ": " . $lines[$j] . "\n";
                array_splice($lines, $j, 1);
                $changes++;
            } else {
                $j++;
            }
        }
        break;
    }
}

// Fix 2: Move activeYears and activeCalYears declarations BEFORE rebuildAllData function
// Find where they are currently declared (as let)
$ayLine = null;
$acyLine = null;
for ($i = 0; $i < count($lines); $i++) {
    if (preg_match('/^let\s+activeYears\s*=/', $lines[$i])) $ayLine = $i;
    if (preg_match('/^let\s+activeCalYears\s*=/', $lines[$i])) $acyLine = $i;
}

echo "\nactiveYears at line " . ($ayLine+1) . "\n";
echo "activeCalYears at line " . ($acyLine+1) . "\n";

// Find rebuildAllData function declaration
$fnLine = null;
for ($i = 0; $i < count($lines); $i++) {
    if (strpos($lines[$i], 'function rebuildAllData()') !== false) {
        $fnLine = $i;
        break;
    }
}
echo "rebuildAllData() at line " . ($fnLine+1) . "\n";

if ($ayLine !== null && $acyLine !== null && $fnLine !== null && $ayLine > $fnLine) {
    // Remove old declarations
    // Remove in reverse order to keep indices correct
    $toRemove = [$acyLine, $ayLine];
    sort($toRemove);
    $toRemove = array_reverse($toRemove);
    foreach ($toRemove as $idx) {
        echo "Fix 2: Removing old declaration at L" . ($idx+1) . ": " . $lines[$idx] . "\n";
        array_splice($lines, $idx, 1);
        $changes++;
    }

    // Re-find rebuildAllData function (indices may have shifted)
    for ($i = 0; $i < count($lines); $i++) {
        if (strpos($lines[$i], 'function rebuildAllData()') !== false) {
            $fnLine = $i;
            break;
        }
    }

    // Insert declarations BEFORE the function
    $insertAt = $fnLine;
    $newDecls = [
        'let activeYears=new Set();',
        'let activeCalYears=new Set();'
    ];
    array_splice($lines, $insertAt, 0, $newDecls);
    $changes += 2;
    echo "Fix 2: Inserted activeYears+activeCalYears at line " . ($insertAt+1) . "\n";
}

// Fix 3: Also check the "// (moved into rebuildAllData function)" line that might be between them
for ($i = 0; $i < count($lines); $i++) {
    if (trim($lines[$i]) === '// (moved into rebuildAllData function)') {
        echo "Fix 3: Removing comment placeholder at L" . ($i+1) . "\n";
        array_splice($lines, $i, 1);
        $changes++;
        break;
    }
}

if ($changes > 0) {
    $result = implode("\n", $lines);
    file_put_contents($file, $result);
    echo "\n✅ Applied $changes fixes. File saved (" . strlen($result) . " bytes)\n";
} else {
    echo "\n⚠️ No changes needed\n";
}

// Final verification
echo "\n=== VERIFICATION ===\n";
$lines2 = explode("\n", file_get_contents($file));
// Show around rebuildAllData
for ($i = 0; $i < count($lines2); $i++) {
    if (strpos($lines2[$i], 'function rebuildAllData()') !== false) {
        echo "--- rebuildAllData area ---\n";
        for ($j = max(0, $i-3); $j <= min(count($lines2)-1, $i+20); $j++) {
            echo "L" . ($j+1) . ": " . $lines2[$j] . "\n";
        }
        break;
    }
}
