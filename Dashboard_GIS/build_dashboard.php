<?php
/**
 * build_dashboard.php — Auto-update fallback data in index.html
 * =============================================================
 * อ่านข้อมูลจาก SQLite database แล้ว update ตัวแปร fallback ทั้งหมดใน index.html
 * เพื่อให้ GitHub Pages แสดงข้อมูลตรงกับ local เสมอ
 *
 * ใช้โดย push_to_github.bat (Step 2: BUILD)
 * หรือรันตรง: php build_dashboard.php
 *
 * Logic ทุกส่วนเหมือนกับ api.php ทุกประการ:
 *   - PD1/PD2: pending-chart (สะสมรายเดือน/เทียบเดือนก่อน)
 *   - PD3: pending-table (นับตาม status + finish_ce)
 *   - PD4: pending-detail (รายละเอียดทั้งหมด)
 *   - PD5: pending-nojob (ท่อแตกรั่วยังไม่เปิดงาน)
 */

// ─── Configuration ────────────────────────────────────────────────────────
$BASE_DIR = __DIR__;
$RAW_DATA_DIR = $BASE_DIR . DIRECTORY_SEPARATOR . 'ข้อมูลดิบ';
$INDEX_FILE = $BASE_DIR . DIRECTORY_SEPARATOR . 'index.html';
$PENDING_DIR = $RAW_DATA_DIR . DIRECTORY_SEPARATOR . 'ซ่อมท่อค้างระบบ';

const BRANCH_LIST = [
    "ชลบุรี","พัทยา","บ้านบึง","พนัสนิคม","ศรีราชา","แหลมฉบัง",
    "ฉะเชิงเทรา","บางปะกง","บางคล้า","พนมสารคาม","ระยอง","บ้านฉาง",
    "ปากน้ำประแสร์","จันทบุรี","ขลุง","ตราด","คลองใหญ่",
    "สระแก้ว","วัฒนานคร","อรัญประเทศ","ปราจีนบุรี","กบินทร์บุรี"
];

echo "=== build_dashboard.php (Dashboard_GIS) ===\n";

// ─── Verify files exist ───────────────────────────────────────────────────
if (!file_exists($INDEX_FILE)) {
    echo "  [ERROR] index.html not found\n";
    exit(1);
}
if (!is_dir($PENDING_DIR)) {
    echo "  [ERROR] Pending data directory not found: $PENDING_DIR\n";
    exit(1);
}

// ─── Find SQLite database ─────────────────────────────────────────────────
function findLatestSqlite($dir) {
    $dbs = [];
    foreach (scandir($dir) as $f) {
        if (preg_match('/ค้างซ่อม_fy_(\d+)\.sqlite$/u', $f, $m)) {
            $dbs[(int)$m[1]] = $dir . DIRECTORY_SEPARATOR . $f;
        }
    }
    if (empty($dbs)) return null;
    krsort($dbs);
    $fy = array_key_first($dbs);
    return ['fy' => $fy, 'path' => $dbs[$fy]];
}

$dbInfo = findLatestSqlite($PENDING_DIR);
if (!$dbInfo) {
    echo "  [ERROR] No SQLite database found in: $PENDING_DIR\n";
    exit(1);
}

$fy = $dbInfo['fy'];
$dbPath = $dbInfo['path'];
echo "  Found DB: " . basename($dbPath) . " (FY $fy)\n";

// ─── Open database ────────────────────────────────────────────────────────
try {
    $db = new PDO('sqlite:' . $dbPath, null, null, [
        PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION
    ]);
} catch (Exception $e) {
    echo "  [ERROR] Cannot open database: " . $e->getMessage() . "\n";
    exit(1);
}

// ─── Get update_date ──────────────────────────────────────────────────────
$update_date = '';
$stmt = $db->query("SELECT value FROM meta WHERE key='update_date'");
$row = $stmt->fetch(PDO::FETCH_NUM);
if ($row) $update_date = $row[0];
echo "  Update date: $update_date\n";

// ─── Fiscal year config ───────────────────────────────────────────────────
$fy_be = $fy;
$fy_ce = $fy_be - 543;
$fy_start_ce = $fy_ce - 1;
$count_start = "$fy_start_ce-10-01";

$fy_yy_start = $fy_be - 2500 - 1;
$fy_yy_end = $fy_yy_start + 1;
$fy_months = [];
for ($mm = 10; $mm <= 12; $mm++) {
    $fy_months[] = sprintf('%02d-%02d', $fy_yy_start, $mm);
}
for ($mm = 1; $mm <= 9; $mm++) {
    $fy_months[] = sprintf('%02d-%02d', $fy_yy_end, $mm);
}

// ═══════════════════════════════════════════════════════════════════════════
//  1) PD1 + PD2 (pending-chart logic)
// ═══════════════════════════════════════════════════════════════════════════
echo "  Computing PD1/PD2 (pending-chart)...\n";

$records = [];
$stmt = $db->query("SELECT date_ce, finish_ce, status, branch FROM pending_rows WHERE date_ce >= '$count_start' ORDER BY date_ce");
while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
    $records[] = $row;
}

$early_records = [];
$stmt2 = $db->query("SELECT date_ce, finish_ce, status, branch FROM pending_rows WHERE date_ce < '$count_start'");
while ($row = $stmt2->fetch(PDO::FETCH_ASSOC)) {
    $early_records[] = $row;
}

// Find months with data
$month_set = [];
foreach ($records as $rec) {
    $dt = new DateTime($rec['date_ce']);
    $y = $dt->format('Y');
    $m = $dt->format('m');
    $month_set["$y-$m"] = [$y, $m];
}
ksort($month_set);

// Build pd2_data
$pd2_months = [];
$pd2_data = [];
$all_records = array_merge($early_records, $records);

foreach ($month_set as [$y, $m]) {
    $yy = ($y + 543) % 100;
    $mk = sprintf('%02d-%02d', $yy, $m);
    $pd2_months[] = $mk;

    $end_of_month_str = "$y-$m-" . date('t', mktime(0, 0, 0, $m, 1, $y));
    $branch_counts = [];

    foreach ($all_records as $rec) {
        if ($rec['date_ce'] < $count_start || $rec['date_ce'] > $end_of_month_str) continue;

        $is_pending = false;
        if ($rec['finish_ce'] && $rec['finish_ce'] > $end_of_month_str) {
            $is_pending = true;
        } elseif (strpos($rec['status'], 'ซ่อมไม่เสร็จ') !== false) {
            $is_pending = true;
        }

        if ($is_pending) {
            $b = $rec['branch'];
            $branch_counts[$b] = ($branch_counts[$b] ?? 0) + 1;
        }
    }

    $pd2_data[$mk] = $branch_counts;
}

// Derive pd1_data
$pd1_data = [];
foreach ($pd2_months as $i => $mk) {
    $prev_mk = $i > 0 ? $pd2_months[$i - 1] : null;
    $prev_snap = $prev_mk ? ($pd2_data[$prev_mk] ?? []) : [];
    $curr_snap = $pd2_data[$mk] ?? [];
    $branch_pairs = [];

    foreach (BRANCH_LIST as $b) {
        $pv = $prev_snap[$b] ?? 0;
        $cv = $curr_snap[$b] ?? 0;
        $branch_pairs[$b] = [$pv, $cv];
    }
    $pd1_data[$mk] = $branch_pairs;
}

echo "  PD1/PD2: " . count($pd2_months) . " months\n";

// ═══════════════════════════════════════════════════════════════════════════
//  2) PD3 (pending-table logic)
// ═══════════════════════════════════════════════════════════════════════════
echo "  Computing PD3 (pending-table)...\n";

$fy_months_sql = "'" . implode("','", $fy_months) . "'";

// Status-based count
$sql = "SELECT branch, month_key, COUNT(*) as cnt
        FROM pending_rows
        WHERE status LIKE '%ซ่อมไม่เสร็จ%'
          AND month_key IN ($fy_months_sql)
        GROUP BY branch, month_key";
$result = $db->query($sql);

$pd3_data = [];
foreach (BRANCH_LIST as $b) {
    $pd3_data[$b] = [];
    foreach ($fy_months as $mk) {
        $pd3_data[$b][$mk] = 0;
    }
}

while ($row = $result->fetch(PDO::FETCH_ASSOC)) {
    $b = $row['branch'];
    $mk = $row['month_key'];
    if (isset($pd3_data[$b]) && isset($pd3_data[$b][$mk])) {
        $pd3_data[$b][$mk] = (int)$row['cnt'];
    }
}

// Extra: finished after end of month
$today_eom = date('Y-m-t');
$extra_sql = "SELECT branch, month_key, COUNT(*) as cnt
              FROM pending_rows
              WHERE status NOT LIKE '%ซ่อมไม่เสร็จ%'
                AND finish_ce IS NOT NULL AND finish_ce != '' AND finish_ce > '$today_eom'
                AND month_key IN ($fy_months_sql)
              GROUP BY branch, month_key";
try {
    $extra_result = $db->query($extra_sql);
    while ($row = $extra_result->fetch(PDO::FETCH_ASSOC)) {
        $b = $row['branch'];
        $mk = $row['month_key'];
        if (isset($pd3_data[$b]) && isset($pd3_data[$b][$mk])) {
            $pd3_data[$b][$mk] += (int)$row['cnt'];
        }
    }
} catch (Exception $e) {
    // finish_ce column might not exist — ignore
}

// Build sparse output (only branches with data > 0)
$pd3_out = [];
foreach (BRANCH_LIST as $branch) {
    $bd = $pd3_data[$branch] ?? [];
    $sparse = [];
    foreach ($bd as $mk => $v) {
        if ($v > 0) $sparse[$mk] = $v;
    }
    if (!empty($sparse)) {
        $pd3_out[$branch] = $sparse;
    }
}

echo "  PD3: " . count($pd3_out) . " branches with data\n";

// ═══════════════════════════════════════════════════════════════════════════
//  3) PD4 (pending-detail logic)
// ═══════════════════════════════════════════════════════════════════════════
echo "  Computing PD4 (pending-detail)...\n";

// Get all pending records (status contains ซ่อมไม่เสร็จ)
$sql = "SELECT branch, notify_no, date_val, job_no, type_val, side, team, tech, pipe, status, month_key
        FROM pending_rows WHERE status LIKE '%ซ่อมไม่เสร็จ%' ORDER BY date_ce";
$result = $db->query($sql);
$pd4_records = [];

while ($row = $result->fetch(PDO::FETCH_ASSOC)) {
    $pd4_records[] = [
        'branch' => $row['branch'],
        'notify_no' => $row['notify_no'],
        'date' => $row['date_val'],
        'job_no' => $row['job_no'],
        'type' => $row['type_val'],
        'aspect' => $row['side'],
        'team' => $row['team'],
        'tech' => $row['tech'],
        'pipe' => $row['pipe'],
        'status' => $row['status'],
        'month' => $row['month_key']
    ];
}

// Extra: finished after end of month
$extra_sql = "SELECT branch, notify_no, date_val, job_no, type_val, side, team, tech, pipe, status, month_key
              FROM pending_rows
              WHERE status NOT LIKE '%ซ่อมไม่เสร็จ%'
                AND finish_ce IS NOT NULL AND finish_ce != '' AND finish_ce > '$today_eom'
              ORDER BY date_ce";
try {
    $extra_result = $db->query($extra_sql);
    while ($row = $extra_result->fetch(PDO::FETCH_ASSOC)) {
        $pd4_records[] = [
            'branch' => $row['branch'],
            'notify_no' => $row['notify_no'],
            'date' => $row['date_val'],
            'job_no' => $row['job_no'],
            'type' => $row['type_val'],
            'aspect' => $row['side'],
            'team' => $row['team'],
            'tech' => $row['tech'],
            'pipe' => $row['pipe'],
            'status' => $row['status'],
            'month' => $row['month_key']
        ];
    }
} catch (Exception $e) {
    // ignore
}

echo "  PD4: " . count($pd4_records) . " records\n";

// ═══════════════════════════════════════════════════════════════════════════
//  4) PD5 (pending-nojob logic)
// ═══════════════════════════════════════════════════════════════════════════
echo "  Computing PD5 (pending-nojob)...\n";

$latest_mk_stmt = $db->query("SELECT month_key FROM pending_rows ORDER BY date_ce DESC LIMIT 1");
$latest_mk_row = $latest_mk_stmt->fetch(PDO::FETCH_NUM);
$latest_mk = $latest_mk_row ? $latest_mk_row[0] : '';

$pd5_by_branch = [];
$pd5_records = [];
$pd5_month = $latest_mk;

if ($latest_mk) {
    $safe_mk = str_replace("'", "''", $latest_mk);
    $sql = "SELECT branch, notify_no, date_val, side, topic, detail, status
            FROM pending_rows
            WHERE month_key = '$safe_mk'
              AND (side = 'ด้านท่อแตกรั่ว'
                   OR topic LIKE '%ท่อแตก%' OR topic LIKE '%ท่อรั่ว%' OR topic LIKE '%แตกรั่ว%'
                   OR detail LIKE '%ท่อแตก%' OR detail LIKE '%ท่อรั่ว%' OR detail LIKE '%แตกรั่ว%')
              AND (job_no IS NULL OR job_no = '')
              AND status NOT LIKE '%ดำเนินการแล้วเสร็จ%'
            ORDER BY branch, date_ce";

    $result = $db->query($sql);
    while ($row = $result->fetch(PDO::FETCH_ASSOC)) {
        $branch = $row['branch'];
        $pd5_by_branch[$branch] = ($pd5_by_branch[$branch] ?? 0) + 1;
        $pd5_records[] = [
            'branch' => $branch,
            'notify_no' => $row['notify_no'],
            'date' => $row['date_val'],
            'side' => $row['side'] ?? '',
            'status' => $row['status']
        ];
    }
}

echo "  PD5: " . count($pd5_records) . " records (month: $pd5_month)\n";

$db = null; // close

// ═══════════════════════════════════════════════════════════════════════════
//  UPDATE index.html
// ═══════════════════════════════════════════════════════════════════════════
echo "  Updating index.html...\n";

$html = file_get_contents($INDEX_FILE);
if (!$html) {
    echo "  [ERROR] Cannot read index.html\n";
    exit(1);
}

$json_opts = JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES;
$changes = 0;

// --- Update PENDING_UPDATE_DATE ---
$html = preg_replace(
    "/const PENDING_UPDATE_DATE='[^']*'/",
    "const PENDING_UPDATE_DATE='" . addslashes($update_date) . "'",
    $html, 1, $cnt
);
$changes += $cnt;

// --- Update PD1_DATA_FALLBACK ---
$pd1_json = json_encode($pd1_data, $json_opts);
$html = preg_replace(
    '/var PD1_DATA_FALLBACK=\{[^\n]*\};/',
    'var PD1_DATA_FALLBACK=' . $pd1_json . ';',
    $html, 1, $cnt
);
$changes += $cnt;

// --- Update PD1_MONTHS_FALLBACK ---
$pd1_months_json = json_encode($pd2_months, $json_opts);
$html = preg_replace(
    '/var PD1_MONTHS_FALLBACK=\[[^\]]*\];/',
    'var PD1_MONTHS_FALLBACK=' . $pd1_months_json . ';',
    $html, 1, $cnt
);
$changes += $cnt;

// --- Update PD2_DATA_FALLBACK ---
$pd2_json = json_encode($pd2_data, $json_opts);
$html = preg_replace(
    '/var PD2_DATA_FALLBACK=\{[^\n]*\};/',
    'var PD2_DATA_FALLBACK=' . $pd2_json . ';',
    $html, 1, $cnt
);
$changes += $cnt;

// --- Update PD2_MONTHS_FALLBACK ---
$html = preg_replace(
    '/var PD2_MONTHS_FALLBACK=\[[^\]]*\];/',
    'var PD2_MONTHS_FALLBACK=' . $pd1_months_json . ';',
    $html, 1, $cnt
);
$changes += $cnt;

// --- Update PD3_FALLBACK ---
// PD3 uses a multi-line block with JS variable reference PENDING_UPDATE_DATE
// Build the replacement manually
$pd3_lines = "var PD3_FALLBACK={\n";
$pd3_lines .= '    "' . $fy . '":{' . "\n";
$pd3_lines .= '        months:' . json_encode($fy_months, $json_opts) . ",\n";
$pd3_lines .= "        update_date:PENDING_UPDATE_DATE,\n";
$pd3_lines .= "        data:{\n";
$branch_entries = [];
foreach ($pd3_out as $branch => $mkdata) {
    $pairs = [];
    foreach ($mkdata as $mk => $v) {
        $pairs[] = '"' . $mk . '":' . $v;
    }
    $branch_entries[] = '            "' . $branch . '":{' . implode(',', $pairs) . '}';
}
$pd3_lines .= implode(",\n", $branch_entries) . "\n";
$pd3_lines .= "        }\n";
$pd3_lines .= "    }\n";
$pd3_lines .= "};";

$html = preg_replace(
    '/var PD3_FALLBACK=\{.*?\};/s',
    $pd3_lines,
    $html, 1, $cnt
);
$changes += $cnt;

// --- Update PD4_FALLBACK (between markers) ---
$pd4_json = json_encode($pd4_records, $json_opts);
$pd4_replacement = '/*FALLBACK_PD4_START*/var PD4_FALLBACK=' . $pd4_json . ';/*FALLBACK_PD4_END*/';
$html = preg_replace(
    '/\/\*FALLBACK_PD4_START\*\/.*?\/\*FALLBACK_PD4_END\*\//',
    $pd4_replacement,
    $html, 1, $cnt
);
$changes += $cnt;

// --- Update PD5_FALLBACK (between markers) ---
$pd5_by_branch_json = json_encode($pd5_by_branch, $json_opts);
$pd5_records_json = json_encode($pd5_records, $json_opts);
$pd5_month_json = json_encode($pd5_month, $json_opts);
$pd5_replacement = '/*FALLBACK_PD5_START*/var PD5_BY_BRANCH_FALLBACK=' . $pd5_by_branch_json
    . ';var PD5_RECORDS_FALLBACK=' . $pd5_records_json
    . ';var PD5_MONTH_FALLBACK=' . $pd5_month_json
    . ';/*FALLBACK_PD5_END*/';
$html = preg_replace(
    '/\/\*FALLBACK_PD5_START\*\/.*?\/\*FALLBACK_PD5_END\*\//',
    $pd5_replacement,
    $html, 1, $cnt
);
$changes += $cnt;

// ═══════════════════════════════════════════════════════════════════════════
//  TAB 1: KPI data (from repair_data.json cache)
// ═══════════════════════════════════════════════════════════════════════════
$cache_dir = $BASE_DIR . DIRECTORY_SEPARATOR . '.cache';
$repair_cache = $cache_dir . DIRECTORY_SEPARATOR . 'repair_data.json';
if (file_exists($repair_cache)) {
    $repair_raw = json_decode(file_get_contents($repair_cache), true);
    if ($repair_raw && !empty($repair_raw['data'])) {
        // Build DATA object matching JS format: {months, branches, data, month_names}
        $repair_data = [
            'months' => $repair_raw['months'] ?? [],
            'branches' => $repair_raw['branches'] ?? [],
            'data' => $repair_raw['data'] ?? [],
            'month_names' => $repair_raw['month_names'] ?? [
                "01"=>"ม.ค.","02"=>"ก.พ.","03"=>"มี.ค.","04"=>"เม.ย.",
                "05"=>"พ.ค.","06"=>"มิ.ย.","07"=>"ก.ค.","08"=>"ส.ค.",
                "09"=>"ก.ย.","10"=>"ต.ค.","11"=>"พ.ย.","12"=>"ธ.ค."
            ]
        ];
        $repair_json = json_encode($repair_data, $json_opts);
        // Replace "const DATA = {...};" using strpos (safe for large JSON)
        $replaced = false;
        foreach (['const DATA ', 'const DATA=', 'var DATA ', 'var DATA='] as $needle) {
            $pos = strpos($html, $needle);
            if ($pos === false) continue;
            $eq = strpos($html, '=', $pos + 4);
            if ($eq === false || $eq - $pos > 20) continue;
            $brace = strpos($html, '{', $eq);
            if ($brace === false || $brace - $eq > 5) continue;
            $depth = 0; $i = $brace; $len = strlen($html);
            while ($i < $len) {
                $ch = $html[$i];
                if ($ch === '{') $depth++;
                elseif ($ch === '}') { $depth--; if ($depth === 0) break; }
                elseif ($ch === '"' || $ch === "'") {
                    $q = $ch; $i++;
                    while ($i < $len && $html[$i] !== $q) { if ($html[$i] === '\\') $i++; $i++; }
                }
                $i++;
            }
            if ($depth !== 0) continue;
            $end = $i + 1;
            if ($end < $len && $html[$end] === ';') $end++;
            $new_val = 'const DATA = ' . $repair_json . ';';
            echo "  [OK] TAB1 KPI DATA embedded (" . number_format(strlen($repair_json)) . " bytes)\n";
            $html = substr($html, 0, $pos) . $new_val . substr($html, $end);
            $changes++;
            $replaced = true;
            break;
        }
        if (!$replaced) {
            echo "  [WARNING] Cannot find DATA variable in index.html for TAB1\n";
        }
    } else {
        echo "  [SKIP] repair_data.json has no data\n";
    }
} else {
    echo "  [SKIP] repair_data.json not found (TAB1 KPI)\n";
}

// ═══════════════════════════════════════════════════════════════════════════
//  TAB 2: Pressure/OIS data (from pressure_data.json cache)
// ═══════════════════════════════════════════════════════════════════════════
$pressure_cache = $cache_dir . DIRECTORY_SEPARATOR . 'pressure_data.json';
if (file_exists($pressure_cache)) {
    $pressure_raw = json_decode(file_get_contents($pressure_cache), true);
    if ($pressure_raw && !empty($pressure_raw['data'])) {
        $pressure_data = $pressure_raw['data'];
        $pressure_months = $pressure_raw['months'] ?? array_keys($pressure_data);

        // Replace PRESSURE_DATA
        $p_json = json_encode($pressure_data, $json_opts);
        $html = preg_replace(
            '/var PRESSURE_DATA=\{[^\n]*\};/',
            'var PRESSURE_DATA=' . $p_json . ';',
            $html, 1, $cnt
        );
        if ($cnt > 0) {
            $changes += $cnt;
            echo "  [OK] TAB2 PRESSURE_DATA embedded (" . number_format(strlen($p_json)) . " bytes)\n";
        } else {
            echo "  [WARNING] Cannot find PRESSURE_DATA in index.html\n";
        }

        // Replace PRESSURE_MONTHS
        $pm_json = json_encode($pressure_months, $json_opts);
        $html = preg_replace(
            '/var PRESSURE_MONTHS=\[[^\]]*\];/',
            'var PRESSURE_MONTHS=' . $pm_json . ';',
            $html, 1, $cnt
        );
        if ($cnt > 0) {
            $changes += $cnt;
            echo "  [OK] TAB2 PRESSURE_MONTHS embedded (" . count($pressure_months) . " months)\n";
        } else {
            echo "  [WARNING] Cannot find PRESSURE_MONTHS in index.html\n";
        }
    } else {
        echo "  [SKIP] pressure_data.json has no data\n";
    }
} else {
    echo "  [SKIP] pressure_data.json not found (TAB2 OIS)\n";
}

// ─── Write updated HTML ───────────────────────────────────────────────────
if ($changes > 0) {
    file_put_contents($INDEX_FILE, $html);
    echo "  [OK] Updated $changes sections in index.html\n";
} else {
    echo "  [WARNING] No sections were updated!\n";
    exit(1);
}

echo "=== Done ===\n";
exit(0);
