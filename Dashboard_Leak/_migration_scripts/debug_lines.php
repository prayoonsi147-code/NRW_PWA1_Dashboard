<?php
$lines = explode("\n", file_get_contents(__DIR__ . '/index.html'));
echo "Lines 2505-2570:\n";
for ($i = 2504; $i < min(2570, count($lines)); $i++) {
    echo "L" . ($i+1) . ": " . $lines[$i] . "\n";
}
