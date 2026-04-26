<?php
$file = __DIR__ . '/index.html';
$lines = explode("\n", file_get_contents($file));

// Find and remove the leftover lines after rebuildAllData():
// L2523: allData[y]=sheets;
// L2524: });
$removed = 0;
for ($i = 0; $i < count($lines); $i++) {
    // Find the rebuildAllData() call
    if (strpos($lines[$i], 'rebuildAllData(); // initial build') !== false) {
        // Check if next lines are leftover code
        $next1 = isset($lines[$i+1]) ? trim($lines[$i+1]) : '';
        $next2 = isset($lines[$i+2]) ? trim($lines[$i+2]) : '';

        echo "Found rebuildAllData() at line " . ($i+1) . "\n";
        echo "  Next: '$next1'\n";
        echo "  Next+1: '$next2'\n";

        // Remove leftover lines
        $toRemove = [];
        $j = $i + 1;
        while ($j < count($lines)) {
            $trimmed = trim($lines[$j]);
            // These are leftover fragments from the old allData init
            if ($trimmed === 'allData[y]=sheets;' || $trimmed === '});' || $trimmed === '') {
                $toRemove[] = $j;
                $j++;
                if ($trimmed === '});') break; // stop after closing
            } else {
                break;
            }
        }

        if (!empty($toRemove)) {
            echo "Removing " . count($toRemove) . " leftover lines:\n";
            foreach (array_reverse($toRemove) as $idx) {
                echo "  L" . ($idx+1) . ": " . $lines[$idx] . "\n";
                array_splice($lines, $idx, 1);
                $removed++;
            }
        }
        break;
    }
}

if ($removed > 0) {
    $result = implode("\n", $lines);
    file_put_contents($file, $result);
    echo "\n✅ Removed $removed leftover lines. File saved (" . strlen($result) . " bytes)\n";
} else {
    echo "\n⚠️ No leftover lines found\n";
}

// Verify
$lines2 = explode("\n", file_get_contents($file));
echo "\nVerification (lines around rebuildAllData):\n";
for ($i = 0; $i < count($lines2); $i++) {
    if (strpos($lines2[$i], 'rebuildAllData()') !== false) {
        for ($j = max(0, $i-2); $j <= min(count($lines2)-1, $i+5); $j++) {
            echo "L" . ($j+1) . ": " . $lines2[$j] . "\n";
        }
        break;
    }
}
