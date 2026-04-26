<?php
/**
 * Migration: Remove Dual Mode, switch to API-fetch-only
 *
 * This script modifies index.html to:
 * 1. Replace embedded data (var D={...}, var RL={...}, etc.) with empty objects
 * 2. Replace the Dual Mode IIFE with a mandatory API-fetch initialization
 * 3. Keep the allData rebuild logic but wrap it in a function
 *
 * Run once: http://localhost/.../migrate_api_only.php
 * After running, build_dashboard.php no longer needs to embed data.
 */

$indexFile = __DIR__ . '/index.html';
$backupFile = __DIR__ . '/index.html.bak_' . date('Ymd_His');

// Read the actual served index.html
$html = file_get_contents($indexFile);
if (!$html) {
    die("ERROR: Cannot read index.html\n");
}

$lines = explode("\n", $html);
$totalLines = count($lines);
echo "Read index.html: $totalLines lines, " . strlen($html) . " bytes\n";

// Backup
file_put_contents($backupFile, $html);
echo "Backup saved to: $backupFile\n";

// ──────────────────────────────────────────────────────
// Step 1: Replace embedded data lines with empty objects
// ──────────────────────────────────────────────────────
$dataVarMap = [
    'D'   => '{}',
    'RL'  => '{}',
    'EU'  => '{}',
    'MNF' => '{}',
    'KPI' => '{}',
    'P3'  => '{}'
];

$replacedVars = [];
for ($i = 0; $i < $totalLines; $i++) {
    foreach ($dataVarMap as $varName => $emptyVal) {
        if (preg_match('/^var\s+' . $varName . '\s*=/', $lines[$i])) {
            $oldLen = strlen($lines[$i]);
            $lines[$i] = "var $varName=$emptyVal; // [API-only mode] data loaded from api.php/$varName";
            $replacedVars[] = "$varName (line " . ($i+1) . ", was $oldLen chars)";
        }
    }
}

echo "\nStep 1 — Replaced embedded data:\n";
foreach ($replacedVars as $v) {
    echo "  ✅ $v\n";
}

// ──────────────────────────────────────────────────────
// Step 2: Replace the allData inline initialization
//   (lines 2508-2515 area) with a reusable function
// ──────────────────────────────────────────────────────
// Find the allData initialization block:
//   let allData={};
//   Object.keys(D).forEach(ys=>{...});
$allDataStart = null;
$allDataEnd = null;
for ($i = 0; $i < $totalLines; $i++) {
    if (preg_match('/^let\s+allData\s*=\s*\{\}/', $lines[$i])) {
        $allDataStart = $i;
    }
    if ($allDataStart !== null && $i > $allDataStart && preg_match('/^\}\);?\s*$/', trim($lines[$i]))) {
        // Check if this closes the Object.keys(D).forEach
        $allDataEnd = $i;
        break;
    }
}

if ($allDataStart !== null && $allDataEnd !== null) {
    // Replace with function-based version
    $newAllDataLines = [
        'let allData={};',
        'function rebuildAllData(){',
        '    allData={};',
        '    Object.keys(D).forEach(ys=>{',
        '        const y=parseInt(ys),sheets={};',
        '        Object.entries(D[ys]).forEach(([s,rows])=>{',
        '            sheets[s]={rows:rows.map(r=>({label:r.l,unit:r.u,monthly:r.m,total:r.t,ty:r.ty!=null?r.ty:null,tm:r.tm!=null?r.tm:null,hasData:r.m.some(v=>v!=null&&v!==0)}))};',
        '        });',
        '        allData[y]=sheets;',
        '    });',
        '    activeYears=new Set(Object.keys(allData).map(Number));',
        '    activeCalYears=new Set();',
        '    Object.keys(allData).map(Number).forEach(fy=>{activeCalYears.add(fy);activeCalYears.add(fy-1);});',
        '}',
        'rebuildAllData(); // initial build (empty until API data arrives)'
    ];
    array_splice($lines, $allDataStart, $allDataEnd - $allDataStart + 1, $newAllDataLines);
    echo "\nStep 2 — Replaced allData init (lines " . ($allDataStart+1) . "-" . ($allDataEnd+1) . ") with rebuildAllData() function\n";

    // Recalculate total lines since we changed count
    $totalLines = count($lines);
}

// ──────────────────────────────────────────────────────
// Step 3: Replace Dual Mode IIFE with API-fetch-only init
// ──────────────────────────────────────────────────────
$iifeStart = null;
$iifeEnd = null;
for ($i = 0; $i < $totalLines; $i++) {
    if (strpos($lines[$i], 'Dual Mode: Try API first') !== false) {
        // Go back one line to include the comment opener
        $iifeStart = $i - 1;
    }
    if ($iifeStart !== null && $i > $iifeStart + 5 && preg_match('/^\}\)\(\);?\s*$/', trim($lines[$i]))) {
        $iifeEnd = $i;
        break;
    }
}

if ($iifeStart !== null && $iifeEnd !== null) {
    $newIife = <<<'JSEOF'
/* ═══════════════════════════════════════════════════════════════════════════
   API-Only Mode: โหลดข้อมูลทั้งหมดจาก API (ไม่ใช้ embedded data)

   ข้อมูลทุกหมวด (OIS, RL, EU, MNF, KPI, P3) จะถูกโหลดจาก api.php
   เมื่อโหลดเสร็จแต่ละหมวดจะ re-init กราฟที่เกี่ยวข้องทันที

   ไม่ต้อง rebuild index.html เมื่อ upload ข้อมูลใหม่ —
   แค่ refresh หน้าก็จะดึงข้อมูลล่าสุดจาก API + cache
   ═══════════════════════════════════════════════════════════════════════════ */
(function(){
    if(location.protocol==='file:'){
        console.warn('[API Mode] Cannot load from API in file:// mode');
        return;
    }

    var _loaded=0, _total=6, _errors=[];
    var _apiBase = 'api.php/';

    function _done(name, ok){
        if(ok){
            _loaded++;
            console.log('[API] ✅ '+name+' ('+_loaded+'/'+_total+')');
        } else {
            _errors.push(name);
            console.warn('[API] ❌ '+name+' failed');
        }
        // เมื่อโหลดครบทุกหมวด (สำเร็จหรือไม่ก็ตาม)
        if((_loaded + _errors.length) >= _total){
            console.log('[API] All done: '+_loaded+' loaded, '+_errors.length+' errors');
            // ลบ loading overlay ถ้ามี
            var overlay = document.getElementById('api-loading-overlay');
            if(overlay) overlay.style.display='none';
        }
    }

    // ── OIS (D) ──
    fetch(_apiBase+'ois-data')
        .then(function(r){return r.json();})
        .then(function(res){
            if(!res.ok||!res.has_data){ _done('OIS',true); return; }
            D=res.data;
            rebuildAllData();
            brInitialized=false; wlInitialized=false;
            if(typeof brInit==='function') brInit();
            if(typeof wlInit==='function') wlInit();
            if(typeof buildYearRangeSelects==='function') buildYearRangeSelects();
            if(typeof onSheetChange==='function') onSheetChange();
            _done('OIS',true);
        }).catch(function(e){ console.error('[API] OIS error:',e); _done('OIS',false); });

    // ── Real Leak (RL) ──
    fetch(_apiBase+'rl-data')
        .then(function(r){return r.json();})
        .then(function(res){
            if(!res.ok||!res.has_data){ _done('RL',true); return; }
            RL=res.data;
            rlcInitialized=false;
            if(typeof rlcInit==='function') rlcInit();
            _done('RL',true);
        }).catch(function(e){ console.error('[API] RL error:',e); _done('RL',false); });

    // ── EU (ค่าไฟ) ──
    fetch(_apiBase+'eu-data')
        .then(function(r){return r.json();})
        .then(function(res){
            if(!res.ok||!res.has_data){ _done('EU',true); return; }
            EU=res.data;
            euInitialized=false;
            if(typeof euInitControls==='function'){euInitControls();euInitialized=true;}
            _done('EU',true);
        }).catch(function(e){ console.error('[API] EU error:',e); _done('EU',false); });

    // ── MNF ──
    fetch(_apiBase+'mnf-data')
        .then(function(r){return r.json();})
        .then(function(res){
            if(!res.ok||!res.has_data){ _done('MNF',true); return; }
            MNF=res.data;
            mnfInitialized=false;
            if(typeof mnfInit==='function') mnfInit();
            _done('MNF',true);
        }).catch(function(e){ console.error('[API] MNF error:',e); _done('MNF',false); });

    // ── KPI ──
    fetch(_apiBase+'kpi-data')
        .then(function(r){return r.json();})
        .then(function(res){
            if(!res.ok||!res.has_data){ _done('KPI',true); return; }
            KPI=res.data;
            _done('KPI',true);
        }).catch(function(e){ console.error('[API] KPI error:',e); _done('KPI',false); });

    // ── P3 ──
    fetch(_apiBase+'p3-data')
        .then(function(r){return r.json();})
        .then(function(res){
            if(!res.ok||!res.has_data){ _done('P3',true); return; }
            P3=res.data;
            p3Initialized=false;
            if(typeof p3Init==='function') p3Init();
            _done('P3',true);
        }).catch(function(e){ console.error('[API] P3 error:',e); _done('P3',false); });
})();
JSEOF;

    $newLines = explode("\n", $newIife);
    array_splice($lines, $iifeStart, $iifeEnd - $iifeStart + 1, $newLines);
    echo "\nStep 3 — Replaced Dual Mode IIFE (lines " . ($iifeStart+1) . "-" . ($iifeEnd+1) . ") with API-only init\n";
    $totalLines = count($lines);
}

// ──────────────────────────────────────────────────────
// Step 4: Also handle the activeYears/activeCalYears
//   that are initialized right after allData
// ──────────────────────────────────────────────────────
// These should already be handled by rebuildAllData() but let's check
// that the standalone declarations still exist and don't conflict
for ($i = 0; $i < $totalLines; $i++) {
    // If there's a standalone activeYears=new Set() that's not inside rebuildAllData
    if (preg_match('/^let\s+activeYears\s*=\s*new\s+Set/', $lines[$i])) {
        $lines[$i] = 'let activeYears=new Set(); // rebuilt by rebuildAllData()';
        echo "\nStep 4a — Simplified activeYears declaration (line " . ($i+1) . ")\n";
    }
    if (preg_match('/^let\s+activeCalYears\s*=\s*new\s+Set/', $lines[$i])) {
        $lines[$i] = 'let activeCalYears=new Set(); // rebuilt by rebuildAllData()';
        echo "Step 4b — Simplified activeCalYears declaration (line " . ($i+1) . ")\n";
    }
    // Remove standalone Object.keys(allData)... lines after activeYears/activeCalYears
    if (preg_match('/^Object\.keys\(allData\)\.map\(Number\)\.forEach/', $lines[$i])) {
        $lines[$i] = '// (moved into rebuildAllData function)';
        echo "Step 4c — Commented out standalone forEach (line " . ($i+1) . ")\n";
    }
}

// ──────────────────────────────────────────────────────
// Write the modified file
// ──────────────────────────────────────────────────────
$newHtml = implode("\n", $lines);
$written = file_put_contents($indexFile, $newHtml);

echo "\n═══════════════════════════════════════════\n";
echo "✅ Migration complete!\n";
echo "New file: $totalLines lines, $written bytes\n";
echo "Old file: " . strlen($html) . " bytes\n";
echo "Saved: " . (strlen($html) - $written) . " bytes removed\n";
echo "Backup: $backupFile\n";
echo "\nNext steps:\n";
echo "  1. Refresh the dashboard page\n";
echo "  2. Data will load from API endpoints\n";
echo "  3. build_dashboard.php no longer needs to embed data\n";
