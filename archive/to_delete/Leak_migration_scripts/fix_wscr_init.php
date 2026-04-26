<?php
/**
 * Fix: Add wscRInit() call to API fetch callbacks for RL and OIS
 * The WSC-R chart depends on both OIS (D) and RL data,
 * so it needs to be re-initialized when either data source loads.
 */
$file = __DIR__ . '/index.html';
$content = file_get_contents($file);

// Add wscR reinit after RL fetch success
$rlCallback = "rlcInitialized=false;\n            if(typeof rlcInit==='function') rlcInit();";
$rlReplacement = "rlcInitialized=false;\n            if(typeof rlcInit==='function') rlcInit();\n            wscRInitialized=false;\n            if(typeof wscRInit==='function') wscRInit();";

if (strpos($content, $rlCallback) !== false) {
    $content = str_replace($rlCallback, $rlReplacement, $content);
    echo "✅ Added wscRInit to RL callback\n";
} else {
    echo "⚠️ RL callback pattern not found\n";
}

// Add wscR reinit after OIS fetch success
$oisCallback = "if(typeof onSheetChange==='function') onSheetChange();";
$oisReplacement = "if(typeof onSheetChange==='function') onSheetChange();\n            wscRInitialized=false;\n            if(typeof wscRInit==='function') wscRInit();";

if (strpos($content, $oisCallback) !== false) {
    $content = str_replace($oisCallback, $oisReplacement, $content);
    echo "✅ Added wscRInit to OIS callback\n";
} else {
    echo "⚠️ OIS callback pattern not found\n";
}

// Also add EU reinit with null check - the euInitControls error is because
// DOM elements haven't been created yet. Wrap in try/catch.
$euCallback = "if(typeof euInitControls==='function'){euInitControls();euInitialized=true;}";
$euReplacement = "try{ if(typeof euInitControls==='function'){euInitControls();euInitialized=true;} }catch(e){ console.warn('[API] EU init deferred:',e.message); }";

if (strpos($content, $euCallback) !== false) {
    $content = str_replace($euCallback, $euReplacement, $content);
    echo "✅ Wrapped EU init in try/catch\n";
} else {
    echo "⚠️ EU callback pattern not found\n";
}

file_put_contents($file, $content);
echo "\n✅ File saved (" . strlen($content) . " bytes)\n";

// Verify
echo "\n=== Verification: RL callback ===\n";
$pos = strpos($content, 'wscRInitialized=false;');
if ($pos !== false) {
    echo substr($content, $pos - 100, 300) . "\n";
}
