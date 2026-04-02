<?php
/**
 * build_dashboard.php — Dashboard Meter
 * อ่าน METER_*.xlsx + OIS → ฝัง DEAD_METER, TOTAL_METERS, DEAD_METER_DATE ลง index.html
 *
 * ใช้ PhpSpreadsheet (require composer autoload จาก parent)
 * Safety: backup → build → validate → write (หรือ restore ถ้าพัง)
 */

ini_set('display_errors', '0');
error_reporting(E_ALL);

// ── PhpSpreadsheet ──
$autoload = dirname(__DIR__) . DIRECTORY_SEPARATOR . 'vendor' . DIRECTORY_SEPARATOR . 'autoload.php';
if (!file_exists($autoload)) {
    echo "[ERROR] vendor/autoload.php not found. Run: composer install\n";
    exit(1);
}
require $autoload;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

// ── Constants (same as api.php) ──
$BRANCH_CODE_MAP = [
    "1102" => "ชลบุรี(พ)",      "1103" => "พัทยา(พ)",       "1104" => "บ้านบึง",       "1105" => "พนัสนิคม",
    "1106" => "ศรีราชา",        "1107" => "แหลมฉบัง",       "1108" => "ฉะเชิงเทรา",     "1109" => "บางปะกง",
    "1110" => "บางคล้า",        "1111" => "พนมสารคาม",     "1112" => "ระยอง",        "1113" => "บ้านฉาง",
    "1114" => "ปากน้ำประแสร์",   "1115" => "จันทบุรี",       "1116" => "ขลุง",         "1117" => "ตราด",
    "1118" => "คลองใหญ่",        "1119" => "สระแก้ว",        "1120" => "วัฒนานคร",      "1121" => "อรัญประเทศ",
    "1122" => "ปราจีนบุรี",      "1123" => "กบินทร์บุรี"
];

$METER_SIZES = ["1/2", "3/4", "1", "1 1/2", "2", "2 1/2", "3", "4", "6", "8"];

$OIS_SHEET_MAP = [
    'ป.ชลบุรี น.3' => 'ชลบุรี(พ)', 'ป.บ้านบึง น.4' => 'บ้านบึง',
    'ป.พนัสนิคม น.5' => 'พนัสนิคม', 'ป.ศรีราชา น.6' => 'ศรีราชา',
    'ป.แหลมฉบัง น.7' => 'แหลมฉบัง', 'ป.พัทยา น.8' => 'พัทยา(พ)',
    'ป.ฉะเชิงเทรา น.9' => 'ฉะเชิงเทรา', 'ป.บางปะกง น.10' => 'บางปะกง',
    'ป.บางคล้า น.11' => 'บางคล้า', 'ป.พนมสารคาม น.12' => 'พนมสารคาม',
    'ป.ระยอง น.13' => 'ระยอง', 'ป.บ้านฉาง น.14' => 'บ้านฉาง',
    'ป.ปากน้ำประแสร์ น.15' => 'ปากน้ำประแสร์', 'ป.จันทบุรี น.16' => 'จันทบุรี',
    'ป.ขลุง น.17' => 'ขลุง', 'ป.ตราด น.18' => 'ตราด',
    'ป.คลองใหญ่ น.19' => 'คลองใหญ่', 'ป.สระแก้ว น.20' => 'สระแก้ว',
    'ป.วัฒนา น.21' => 'วัฒนานคร', 'ป.อรัญประเทศ น.22' => 'อรัญประเทศ',
    'ป.ปราจีน น.23' => 'ปราจีนบุรี', 'ป.กบินทร์ น.24' => 'กบินทร์บุรี',
];

$TH_MONTHS = [
    '', 'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
    'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
];

// ── Paths ──
$base_dir = __DIR__;
$raw_dir = $base_dir . DIRECTORY_SEPARATOR . 'ข้อมูลดิบ';
$meter_dir = $raw_dir . DIRECTORY_SEPARATOR . 'มาตรวัดน้ำผิดปกติ';
$ois_dir = dirname($base_dir) . DIRECTORY_SEPARATOR . 'Dashboard_Leak'
         . DIRECTORY_SEPARATOR . 'ข้อมูลดิบ' . DIRECTORY_SEPARATOR . 'OIS';
$html_file = $base_dir . DIRECTORY_SEPARATOR . 'index.html';

echo "=== Build Dashboard Meter ===\n";

// ── Helper: normalize meter size ──
function normalize_size($val) {
    global $METER_SIZES;
    $s = trim(str_replace(['"', "'", ' นิ้ว', 'นิ้ว'], '', (string)$val));
    if (in_array($s, $METER_SIZES)) return $s;
    $s = str_replace(',', '.', $s);
    if (is_numeric($s)) {
        $n = floatval($s);
        foreach ($METER_SIZES as $ms) {
            if (abs(eval("return $ms;") - $n) < 0.01) return $ms;
        }
    }
    return null;
}

// ── Helper: detect_meter_columns (same as api.php) ──
function detect_meter_columns($worksheet) {
    $keywords = [
        'cid'       => ['CA', 'รหัสผู้ใช้น้ำ', 'เลขที่ผู้ใช้น้ำ', 'customer'],
        'size'      => ['ขนาดมาตร', 'ขนาด', 'meter size'],
        'condition' => ['สภาพมาตร', 'สภาพ', 'condition'],
        'change'    => ['เปลี่ยนมาตร', 'การเปลี่ยน', 'change'],
    ];
    $maxScan = min(10, $worksheet->getHighestRow());
    $maxCol = Coordinate::columnIndexFromString($worksheet->getHighestColumn());
    $maxCol = min($maxCol, 30);

    for ($r = 1; $r <= $maxScan; $r++) {
        $found = [];
        for ($c = 1; $c <= $maxCol; $c++) {
            $val = trim((string)($worksheet->getCell([$c, $r])->getValue() ?? ''));
            if ($val === '') continue;
            $lower = mb_strtolower($val);
            foreach ($keywords as $key => $kws) {
                if (isset($found[$key])) continue;
                foreach ($kws as $kw) {
                    if (mb_strpos($lower, mb_strtolower($kw)) !== false) {
                        $found[$key] = $c;
                        break 2;
                    }
                }
            }
        }
        if (isset($found['cid']) && count($found) >= 2) {
            return ['header_row' => $r, 'cols' => $found, 'fallback' => false];
        }
    }
    return ['header_row' => 1, 'cols' => ['cid' => 2, 'size' => 9, 'condition' => 12, 'change' => 16], 'fallback' => true];
}

// ══════════════════════════════════════════════════════════════
//  1. Parse DEAD_METER from METER_*.xlsx
// ══════════════════════════════════════════════════════════════
$dead_meter = [];
$latest_file_date = 0;

if (is_dir($meter_dir)) {
    $files = glob($meter_dir . DIRECTORY_SEPARATOR . 'METER_*.xlsx') ?: [];
    sort($files);
    echo "  Found " . count($files) . " METER files\n";

    foreach ($files as $file) {
        $basename = pathinfo($file, PATHINFO_FILENAME);
        $code_part = preg_replace('/^METER_/', '', $basename);
        // Extract branch code (first 4 digits)
        if (preg_match('/(\d{4})/', $code_part, $cm)) {
            $code = $cm[1];
        } else {
            continue;
        }
        $branch = isset($BRANCH_CODE_MAP[$code]) ? $BRANCH_CODE_MAP[$code] : null;
        if (!$branch) { echo "  [SKIP] $basename: unknown code $code\n"; continue; }

        // Track latest file modification for date
        $mt = filemtime($file);
        if ($mt > $latest_file_date) $latest_file_date = $mt;

        try {
            $spreadsheet = IOFactory::load($file);
            $ws = $spreadsheet->getActiveSheet();
            $maxRow = $ws->getHighestRow();

            $det = detect_meter_columns($ws);
            $hdr = $det['header_row'];
            $cCid       = $det['cols']['cid']       ?? 2;
            $cSize      = $det['cols']['size']       ?? 9;
            $cCondition = $det['cols']['condition']   ?? 12;
            $cChange    = $det['cols']['change']     ?? 16;

            $seen = [];
            $sizes = [];
            foreach ($METER_SIZES as $sz) { $sizes[$sz] = 0; }
            $total = 0;

            for ($r = $hdr + 1; $r <= $maxRow; $r++) {
                $cid = $ws->getCell([$cCid, $r])->getValue();
                if ($cid === null) continue;
                $cid = trim((string)$cid);
                if (isset($seen[$cid])) continue;

                $condition = $ws->getCell([$cCondition, $r])->getValue();
                if ($condition === null || trim((string)$condition) !== 'มาตรไม่เดิน') continue;

                $change = $ws->getCell([$cChange, $r])->getValue();
                if ($change !== null && trim((string)$change) === 'เปลี่ยนแล้ว') continue;

                $seen[$cid] = true;
                $total++;

                $sv = $ws->getCell([$cSize, $r])->getValue();
                if ($sv !== null) {
                    $ns = normalize_size($sv);
                    if ($ns !== null && isset($sizes[$ns])) { $sizes[$ns]++; }
                }
            }

            $dead_meter[$branch] = ['total' => $total, 'sizes' => $sizes];
            $spreadsheet->disconnectWorksheets();
            unset($spreadsheet);
            echo "  [OK] $branch: $total dead meters\n";
        } catch (\Throwable $e) {
            echo "  [ERROR] $basename: " . $e->getMessage() . "\n";
            $sizes = [];
            foreach ($METER_SIZES as $sz) { $sizes[$sz] = 0; }
            $dead_meter[$branch] = ['total' => 0, 'sizes' => $sizes];
        }
    }
} else {
    echo "  [WARNING] Meter data dir not found: $meter_dir\n";
}

echo "  Dead meter branches: " . count($dead_meter) . "\n\n";

// ══════════════════════════════════════════════════════════════
//  2. Parse TOTAL_METERS from OIS (latest file)
// ══════════════════════════════════════════════════════════════
$total_meters = [];

if (is_dir($ois_dir)) {
    $ois_files = array_merge(
        glob($ois_dir . DIRECTORY_SEPARATOR . 'OIS_*.xls') ?: [],
        glob($ois_dir . DIRECTORY_SEPARATOR . 'OIS_*.xlsx') ?: [],
        glob($ois_dir . DIRECTORY_SEPARATOR . '*.xls') ?: []
    );
    $ois_files = array_unique($ois_files);
    sort($ois_files);
    $latest_ois = !empty($ois_files) ? end($ois_files) : null;

    if ($latest_ois) {
        echo "  Reading OIS: " . basename($latest_ois) . "\n";
        try {
            $spreadsheet = IOFactory::load($latest_ois);

            $month_col = 6; // default ต.ค. (col F = 6 in 1-based)
            $first_sheet_name = array_keys($OIS_SHEET_MAP)[0];
            try {
                $first_sheet = $spreadsheet->getSheetByName($first_sheet_name);
                if ($first_sheet) {
                    for ($c = 6; $c <= 17; $c++) {
                        $v = $first_sheet->getCell([$c, 6])->getValue();
                        if ($v !== null && $v !== '' && $v != 0) {
                            $month_col = $c;
                        }
                    }
                }
            } catch (\Throwable $e) {}

            foreach ($OIS_SHEET_MAP as $sheetName => $branchName) {
                try {
                    $ws = $spreadsheet->getSheetByName($sheetName);
                    $val = $ws->getCell([$month_col, 6])->getValue();
                    $total_meters[$branchName] = ($val !== null && $val !== '') ? intval($val) : 0;
                } catch (\Throwable $e) {
                    $total_meters[$branchName] = 0;
                }
            }

            $spreadsheet->disconnectWorksheets();
            unset($spreadsheet);
            echo "  [OK] Total meters for " . count($total_meters) . " branches\n\n";
        } catch (\Throwable $e) {
            echo "  [ERROR] OIS: " . $e->getMessage() . "\n\n";
        }
    } else {
        echo "  [WARNING] No OIS file found in $ois_dir\n\n";
    }
} else {
    echo "  [WARNING] OIS dir not found — TOTAL_METERS will not be updated\n\n";
}

// ══════════════════════════════════════════════════════════════
//  3. Generate date string
// ══════════════════════════════════════════════════════════════
if ($latest_file_date > 0) {
    $d = $latest_file_date;
    $day = date('j', $d);
    $month = intval(date('n', $d));
    $year = intval(date('Y', $d)) + 543;
    $date_str = "ณ วันที่ $day " . $TH_MONTHS[$month] . " $year";
} else {
    $day = date('j');
    $month = intval(date('n'));
    $year = intval(date('Y')) + 543;
    $date_str = "ณ วันที่ $day " . $TH_MONTHS[$month] . " $year";
}
echo "  Date: $date_str\n\n";

// ══════════════════════════════════════════════════════════════
//  4. Embed data into index.html (with safety checks)
// ══════════════════════════════════════════════════════════════
if (!file_exists($html_file)) {
    echo "[ERROR] index.html not found!\n";
    exit(1);
}

// Read original
$html = file_get_contents($html_file);
$original_len = strlen($html);

if ($original_len < 1000) {
    echo "[ERROR] index.html too small ($original_len bytes) — possibly corrupted\n";
    exit(1);
}

// Safety: check DOCTYPE
if (stripos($html, '<!DOCTYPE html>') === false) {
    echo "[ERROR] index.html missing DOCTYPE — possibly corrupted\n";
    exit(1);
}

echo "  Embedding data into index.html ($original_len bytes)...\n";

// Backup
$backup_file = $html_file . '.bak';
copy($html_file, $backup_file);

// ── Replace DEAD_METER ──
$dead_json = json_encode($dead_meter, JSON_UNESCAPED_UNICODE);
if (!empty($dead_meter)) {
    // Match: var DEAD_METER={...};  (single line)
    $pattern = '/^var DEAD_METER=\{.*\};$/m';
    $replacement = 'var DEAD_METER=' . $dead_json . ';';
    $new_html = preg_replace($pattern, $replacement, $html, 1);
    if ($new_html !== null && $new_html !== $html) {
        $html = $new_html;
        echo "  [OK] DEAD_METER replaced\n";
    } else {
        // Try multiline: var DEAD_METER={...newlines...};
        $pos = strpos($html, 'var DEAD_METER={');
        if ($pos !== false) {
            $start = $pos;
            // Find matching closing };
            $depth = 0;
            $end = $start;
            $in_str = false;
            $str_char = '';
            for ($i = strpos($html, '{', $start); $i < strlen($html); $i++) {
                $ch = $html[$i];
                if ($in_str) {
                    if ($ch === '\\') { $i++; continue; }
                    if ($ch === $str_char) $in_str = false;
                    continue;
                }
                if ($ch === '"' || $ch === "'") { $in_str = true; $str_char = $ch; continue; }
                if ($ch === '{') $depth++;
                if ($ch === '}') { $depth--; if ($depth === 0) { $end = $i + 1; break; } }
            }
            // Include trailing ;
            if ($end < strlen($html) && $html[$end] === ';') $end++;
            $html = substr($html, 0, $start) . 'var DEAD_METER=' . $dead_json . ';' . substr($html, $end);
            echo "  [OK] DEAD_METER replaced (multiline)\n";
        } else {
            echo "  [SKIP] DEAD_METER marker not found\n";
        }
    }
}

// ── Replace TOTAL_METERS ──
$tm_json = json_encode($total_meters, JSON_UNESCAPED_UNICODE);
if (!empty($total_meters)) {
    $pattern = '/^var TOTAL_METERS=\{.*\};$/m';
    $replacement = 'var TOTAL_METERS=' . $tm_json . ';';
    $new_html = preg_replace($pattern, $replacement, $html, 1);
    if ($new_html !== null && $new_html !== $html) {
        $html = $new_html;
        echo "  [OK] TOTAL_METERS replaced\n";
    } else {
        // Multiline fallback
        $pos = strpos($html, 'var TOTAL_METERS={');
        if ($pos !== false) {
            $start = $pos;
            $depth = 0;
            $end = $start;
            $in_str = false;
            $str_char = '';
            for ($i = strpos($html, '{', $start); $i < strlen($html); $i++) {
                $ch = $html[$i];
                if ($in_str) {
                    if ($ch === '\\') { $i++; continue; }
                    if ($ch === $str_char) $in_str = false;
                    continue;
                }
                if ($ch === '"' || $ch === "'") { $in_str = true; $str_char = $ch; continue; }
                if ($ch === '{') $depth++;
                if ($ch === '}') { $depth--; if ($depth === 0) { $end = $i + 1; break; } }
            }
            if ($end < strlen($html) && $html[$end] === ';') $end++;
            $html = substr($html, 0, $start) . 'var TOTAL_METERS=' . $tm_json . ';' . substr($html, $end);
            echo "  [OK] TOTAL_METERS replaced (multiline)\n";
        } else {
            echo "  [SKIP] TOTAL_METERS marker not found\n";
        }
    }
}

// ── Replace DEAD_METER_DATE ──
$pattern = '/^var DEAD_METER_DATE="[^"]*";$/m';
$replacement = 'var DEAD_METER_DATE="' . $date_str . '";';
$new_html = preg_replace($pattern, $replacement, $html, 1);
if ($new_html !== null && $new_html !== $html) {
    $html = $new_html;
    echo "  [OK] DEAD_METER_DATE replaced\n";
} else {
    echo "  [SKIP] DEAD_METER_DATE marker not found\n";
}

// ── Validate result ──
$new_len = strlen($html);

if ($new_len < 1000) {
    echo "[FAIL] Result too small ($new_len bytes) — restoring backup\n";
    copy($backup_file, $html_file);
    @unlink($backup_file);
    exit(1);
}

if (stripos($html, '<!DOCTYPE html>') === false) {
    echo "[FAIL] Result missing DOCTYPE — restoring backup\n";
    copy($backup_file, $html_file);
    @unlink($backup_file);
    exit(1);
}

if ($new_len < $original_len * 0.3) {
    echo "[FAIL] Result shrunk to " . round($new_len / $original_len * 100) . "% — restoring backup\n";
    copy($backup_file, $html_file);
    @unlink($backup_file);
    exit(1);
}

// ── Write ──
$written = file_put_contents($html_file, $html);
if ($written === false) {
    echo "[FAIL] Cannot write index.html — restoring backup\n";
    copy($backup_file, $html_file);
    @unlink($backup_file);
    exit(1);
}

@unlink($backup_file);
echo "\n  [DONE] index.html updated: " . round($new_len / 1024, 1) . "KB\n";
echo "=== Build Meter Complete ===\n";
