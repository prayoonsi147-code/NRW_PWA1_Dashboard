<?php
/**
 * Dashboard Meter — XAMPP (Apache + PHP) Backend API
 * ========================================================================
 * PHP เทียบเท่า Flask server.py
 * รับ upload ไฟล์ Excel → parse ข้อมูลมาตรตายทันที → บันทึก data.json
 *
 * Architecture:
 *   - Single file handling ALL API routes via PATH_INFO
 *   - .htaccess rewrites /api/* to this file
 *   - Static files (index.html, manage.html) served directly by Apache
 *
 * Setup:
 *   1. Install via Composer: composer require phpoffice/phpspreadsheet
 *   2. Place composer vendor/ at project root (../vendor/autoload.php)
 *   3. Create .htaccess with rewrite rules
 */

// ─── Error Handling ────────────────────────────────────────────────────────
ini_set('display_errors', '0');
error_reporting(E_ALL);
ini_set('log_errors', '1');

// ─── Configuration ─────────────────────────────────────────────────────────

define('BASE_DIR', __DIR__);
define('RAW_DATA_DIR', BASE_DIR . DIRECTORY_SEPARATOR . 'ข้อมูลดิบ');
define('DATA_FILE', BASE_DIR . DIRECTORY_SEPARATOR . 'data.json');

// Meter size standards
const METER_SIZES = ["1/2", "3/4", "1", "1 1/2", "2", "2 1/2", "3", "4", "6", "8"];

// Branch code mapping
const BRANCH_CODE_MAP = [
    "1102" => "ชลบุรี(พ)",      "1103" => "พัทยา(พ)",       "1104" => "บ้านบึง",       "1105" => "พนัสนิคม",
    "1106" => "ศรีราชา",        "1107" => "แหลมฉบัง",       "1108" => "ฉะเชิงเทรา",     "1109" => "บางปะกง",
    "1110" => "บางคล้า",        "1111" => "พนมสารคาม",     "1112" => "ระยอง",        "1113" => "บ้านฉาง",
    "1114" => "ปากน้ำประแสร์",   "1115" => "จันทบุรี",       "1116" => "ขลุง",         "1117" => "ตราด",
    "1118" => "คลองใหญ่",        "1119" => "สระแก้ว",        "1120" => "วัฒนานคร",      "1121" => "อรัญประเทศ",
    "1122" => "ปราจีนบุรี",      "1123" => "กบินทร์บุรี"
];

// Branch display order
const BRANCH_ORDER = [
    "ชลบุรี(พ)", "พัทยา(พ)", "บ้านบึง", "พนัสนิคม", "ศรีราชา", "แหลมฉบัง",
    "ฉะเชิงเทรา", "บางปะกง", "บางคล้า", "พนมสารคาม",
    "ระยอง", "บ้านฉาง", "ปากน้ำประแสร์",
    "จันทบุรี", "ขลุง", "ตราด", "คลองใหญ่",
    "สระแก้ว", "วัฒนานคร", "อรัญประเทศ",
    "ปราจีนบุรี", "กบินทร์บุรี"
];

// Thai month names
const TH_MONTHS = [
    '', 'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
    'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
];

// Category mapping: URL slug → Thai folder name
const CATEGORY_MAP = [
    'abnormal' => 'มาตรวัดน้ำผิดปกติ'
];

// ─── Cache Setup ──────────────────────────────────────────────────────────
define('CACHE_DIR', BASE_DIR . DIRECTORY_SEPARATOR . '.cache');
define('CACHE_TTL', 86400); // 1 day — mtime check handles invalidation when Excel files change
if (!is_dir(CACHE_DIR)) { mkdir(CACHE_DIR, 0755, true); }

function get_folder_mtime($folder_path) {
    $latest = 0;
    if (!is_dir($folder_path)) return 0;
    foreach (scandir($folder_path) as $f) {
        if ($f[0] === '.') continue;
        $fp = $folder_path . DIRECTORY_SEPARATOR . $f;
        if (is_file($fp)) { $mt = filemtime($fp); if ($mt > $latest) $latest = $mt; }
    }
    return $latest;
}

function load_cache($cache_key, $folder_path) {
    $cache_file = CACHE_DIR . DIRECTORY_SEPARATOR . $cache_key . '.json';
    if (!file_exists($cache_file)) return null;
    $cache_mtime = filemtime($cache_file);
    $folder_mtime = get_folder_mtime($folder_path);
    if ($folder_mtime <= $cache_mtime && (time() - $cache_mtime) < CACHE_TTL) {
        $data = json_decode(file_get_contents($cache_file), true);
        if ($data !== null) return $data;
    }
    return null;
}

function save_cache($cache_key, $data) {
    $cache_file = CACHE_DIR . DIRECTORY_SEPARATOR . $cache_key . '.json';
    file_put_contents($cache_file, json_encode($data, JSON_UNESCAPED_UNICODE));
}

// ─── Auto-enable zip extension (required for .xlsx) ───────────────────────
if (!extension_loaded('zip')) {
    $ini = php_ini_loaded_file();
    if ($ini && is_writable($ini)) {
        $content = file_get_contents($ini);
        if (preg_match('/^;extension=zip/m', $content)) {
            $content = preg_replace('/^;extension=zip/m', 'extension=zip', $content);
            file_put_contents($ini, $content);
            error_log("Dashboard: auto-enabled extension=zip in php.ini — restart Apache to activate");
        }
    }
}

// ─── PhpSpreadsheet Loader ─────────────────────────────────────────────────

$composerAutoload = dirname(BASE_DIR) . DIRECTORY_SEPARATOR . 'vendor' . DIRECTORY_SEPARATOR . 'autoload.php';
$phpSpreadsheet = null;

if (file_exists($composerAutoload)) {
    require_once $composerAutoload;
    try {
        $phpSpreadsheet = true;
    } catch (\Throwable $e) {
        error_log("Warning: PhpSpreadsheet not available: " . $e->getMessage());
    }
} else {
    error_log("Warning: Composer vendor/ not found at " . dirname(BASE_DIR));
}

// ─── Setup ─────────────────────────────────────────────────────────────────

// Create directories
if (!is_dir(RAW_DATA_DIR)) {
    mkdir(RAW_DATA_DIR, 0755, true);
}
foreach (CATEGORY_MAP as $thai_name) {
    $folder = RAW_DATA_DIR . DIRECTORY_SEPARATOR . $thai_name;
    if (!is_dir($folder)) {
        mkdir($folder, 0755, true);
    }
}

// ─── Helper Functions ──────────────────────────────────────────────────────

/**
 * Send JSON response with CORS headers
 */
function json_response($data, $status_code = 200) {
    http_response_code($status_code);
    header('Content-Type: application/json; charset=utf-8');
    header('Access-Control-Allow-Origin: *');
    header('Access-Control-Allow-Methods: GET, POST, DELETE, OPTIONS');
    header('Access-Control-Allow-Headers: Content-Type');
    echo json_encode($data, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT);
    exit;
}

/**
 * Load data from data.json
 */
function load_data() {
    if (!file_exists(DATA_FILE)) {
        return [
            'dead_meter' => [
                'snapshots' => [],
                'latest' => ''
            ]
        ];
    }

    try {
        $content = file_get_contents(DATA_FILE);
        $data = json_decode($content, true);

        // Migrate old format to snapshots format
        if (!isset($data['dead_meter']['snapshots'])) {
            $data['dead_meter'] = [
                'snapshots' => [],
                'latest' => ''
            ];
        }

        return $data;
    } catch (\Throwable $e) {
        error_log("Error loading data.json: " . $e->getMessage());
        return [
            'dead_meter' => [
                'snapshots' => [],
                'latest' => ''
            ]
        ];
    }
}

/**
 * Save data to data.json
 */
function save_data($data) {
    $json = json_encode($data, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT);
    file_put_contents(DATA_FILE, $json);
    chmod(DATA_FILE, 0644);
}

/**
 * Parse date string (YYYY-MM-DD or DD/MM/YYYY format in Buddhist era)
 * Returns array: [date_key, date_label] or [null, null] if invalid
 */
function parse_date_key($date_str) {
    $date_str = trim($date_str);

    // Try YYYY-MM-DD format
    if (preg_match('/^(\d{4})-(\d{1,2})-(\d{1,2})$/', $date_str, $m)) {
        $year = (int)$m[1];
        $month = (int)$m[2];
        $day = (int)$m[3];

        if ($month >= 1 && $month <= 12 && $day >= 1 && $day <= 31) {
            $date_key = sprintf("%04d-%02d-%02d", $year, $month, $day);
            $date_label = sprintf("ณ วันที่ %d %s %d", $day, TH_MONTHS[$month], $year);
            return [$date_key, $date_label];
        }
    }

    // Try DD/MM/YYYY format
    if (preg_match('/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/', $date_str, $m)) {
        $day = (int)$m[1];
        $month = (int)$m[2];
        $year = (int)$m[3];

        if ($month >= 1 && $month <= 12 && $day >= 1 && $day <= 31) {
            $date_key = sprintf("%04d-%02d-%02d", $year, $month, $day);
            $date_label = sprintf("ณ วันที่ %d %s %d", $day, TH_MONTHS[$month], $year);
            return [$date_key, $date_label];
        }
    }

    return [null, null];
}

/**
 * Normalize meter size to match METER_SIZES
 */
function normalize_size($s) {
    $s = trim((string)$s);

    if (in_array($s, METER_SIZES)) {
        return $s;
    }

    // Try removing spaces
    $clean = str_replace(' ', '', $s);
    foreach (METER_SIZES as $ms) {
        if ($clean === str_replace(' ', '', $ms)) {
            return $ms;
        }
    }

    // Check for 8 inch meter
    if (strpos($s, '8') !== false && (strpos($s, 'ตั้งแต่') !== false || strpos($s, 'นิ้ว') !== false)) {
        return '8';
    }

    return null;
}

/**
 * Parse Excel file for dead meter data
 * Conditions:
 *   1. สภาพมาตร (col 12) = "มาตรไม่เดิน"
 *   2. การเปลี่ยนมาตร (col 16) ≠ "เปลี่ยนแล้ว"
 *   3. Unique customer IDs (col 2)
 */
// ── Smart Header Detection สำหรับไฟล์มาตรตาย ──
// ค้นหาคอลัมน์จาก keyword แทนตำแหน่งตายตัว
function detect_meter_columns($worksheet) {
    $keywords = [
        'cid'       => ['CA', 'รหัสผู้ใช้น้ำ', 'เลขที่ผู้ใช้น้ำ', 'CA_NO', 'เลขที่'],
        'size'      => ['ขนาดมาตร', 'ขนาด'],
        'condition'  => ['สภาพมาตร', 'สภาพ'],
        'change'    => ['การเปลี่ยน', 'เปลี่ยนมาตร'],
        'billing'   => ['รอบบิล', 'งวด', 'BILLING', 'รหัสงวด'],
    ];
    $maxScan = min(5, $worksheet->getHighestRow());
    $maxCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($worksheet->getHighestColumn());
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
        if (isset($found['cid']) && (isset($found['condition']) || isset($found['size']))) {
            return ['header_row' => $r, 'cols' => $found, 'fallback' => false];
        }
    }
    // fallback: ตำแหน่งเดิม
    return ['header_row' => 1, 'cols' => ['billing' => 1, 'cid' => 2, 'size' => 9, 'condition' => 12, 'change' => 16], 'fallback' => true];
}

// ── Validate file format before upload ──
function validate_meter_file($tmp_path) {
    if (!class_exists('PhpOffice\\PhpSpreadsheet\\IOFactory')) {
        return ['valid' => false, 'message' => 'PhpSpreadsheet ไม่พร้อมใช้งาน — ไม่สามารถตรวจสอบไฟล์ได้'];
    }
    if (!class_exists('ZipArchive')) {
        return ['valid' => false, 'message' => 'PHP zip extension ไม่ได้เปิด — ไม่สามารถอ่านไฟล์ .xlsx ได้ (กรุณาเปิด extension=zip ใน php.ini แล้ว restart Apache)'];
    }
    try {
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($tmp_path);
        // สแกนทุกชีท (ไฟล์อาจมีหลายชีท ข้อมูลอาจไม่อยู่ชีทแรก)
        $meter_skip = ['สารบัญ', 'สรุป', 'ปก', 'cover', 'summary', 'chart', 'graph', 'index'];
        $det = null;
        for ($si = 0; $si < $spreadsheet->getSheetCount(); $si++) {
            $ws = $spreadsheet->getSheet($si);
            $wsName = mb_strtolower($ws->getTitle());
            $skip = false;
            foreach ($meter_skip as $sk) {
                if (mb_strpos($wsName, $sk) !== false) { $skip = true; break; }
            }
            if ($skip) continue;
            $det = detect_meter_columns($ws);
            if (!$det['fallback']) break;
        }
        $spreadsheet->disconnectWorksheets();

        if ($det === null || $det['fallback']) {
            return [
                'valid' => false,
                'message' => 'ไม่พบหัวคอลัมน์ที่คาดหวัง (CA/รหัสผู้ใช้น้ำ, สภาพมาตร, ขนาดมาตร) — รูปแบบไฟล์ไม่ตรงกับที่ระบบรองรับ'
            ];
        }
        return ['valid' => true, 'message' => ''];
    } catch (\Throwable $e) {
        return [
            'valid' => false,
            'message' => 'ไม่สามารถอ่านไฟล์ Excel ได้: ' . $e->getMessage()
        ];
    }
}

function parse_dead_meter_file($file_path) {
    if (!extension_loaded('zip')) {
        throw new Exception("ZIP extension not loaded");
    }

    try {
        if (function_exists('PhpOffice\\PhpSpreadsheet\\IOFactory')) {
            $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($file_path);
            // สแกนทุกชีท หาชีทที่มี header ถูกต้อง
            $meter_skip = ['สารบัญ', 'สรุป', 'ปก', 'cover', 'summary', 'chart', 'graph', 'index'];
            $worksheet = null;
            $det = null;
            for ($si = 0; $si < $spreadsheet->getSheetCount(); $si++) {
                $ws = $spreadsheet->getSheet($si);
                $wsName = mb_strtolower($ws->getTitle());
                $skip = false;
                foreach ($meter_skip as $sk) {
                    if (mb_strpos($wsName, $sk) !== false) { $skip = true; break; }
                }
                if ($skip) continue;
                $det = detect_meter_columns($ws);
                if (!$det['fallback']) { $worksheet = $ws; break; }
            }
            if ($worksheet === null) {
                $worksheet = $spreadsheet->getActiveSheet();
                $det = detect_meter_columns($worksheet);
            }
            $hdr = $det['header_row'];
            $cCid       = $det['cols']['cid']       ?? 2;
            $cSize      = $det['cols']['size']       ?? 9;
            $cCondition = $det['cols']['condition']   ?? 12;
            $cChange    = $det['cols']['change']     ?? 16;
            $cBilling   = $det['cols']['billing']    ?? 1;

            $seen = [];
            $sizes = [];
            $total = 0;

            foreach (METER_SIZES as $sz) {
                $sizes[$sz] = 0;
            }

            $max_row = $worksheet->getHighestRow();

            for ($r = $hdr + 1; $r <= $max_row; $r++) {
                $cid = $worksheet->getCell([$cCid, $r])->getValue();
                if ($cid === null) continue;

                $cid = trim((string)$cid);
                if (isset($seen[$cid])) continue;

                // Condition 1: สภาพมาตร = "มาตรไม่เดิน"
                $condition = $worksheet->getCell([$cCondition, $r])->getValue();
                if ($condition === null || trim((string)$condition) !== "มาตรไม่เดิน") continue;

                // Condition 2: การเปลี่ยนมาตร ≠ "เปลี่ยนแล้ว"
                $change = $worksheet->getCell([$cChange, $r])->getValue();
                if ($change !== null && trim((string)$change) === "เปลี่ยนแล้ว") continue;

                $seen[$cid] = true;
                $total++;

                $sv = $worksheet->getCell([$cSize, $r])->getValue();
                if ($sv !== null) {
                    $ns = normalize_size($sv);
                    if ($ns !== null && isset($sizes[$ns])) {
                        $sizes[$ns]++;
                    }
                }
            }

            // Extract billing month
            $billing_month = null;
            for ($r = $hdr + 1; $r < min($max_row + 1, 20); $r++) {
                $v = $worksheet->getCell([$cBilling, $r])->getValue();
                if ($v) {
                    $vs = trim((string)$v);
                    if (strlen($vs) === 6 && ctype_digit($vs)) {
                        $billing_month = substr($vs, 0, 4) . "-" . substr($vs, 4);
                        break;
                    }
                }
            }

            $spreadsheet->disconnectWorksheets();
            unset($spreadsheet);

            return [
                'total' => $total,
                'sizes' => $sizes,
                'billing_month' => $billing_month
            ];
        } else {
            throw new Exception("PhpSpreadsheet not available");
        }
    } catch (\Throwable $e) {
        throw new Exception("parse ล้มเหลว: " . $e->getMessage());
    }
}

// ─── Route Handling ────────────────────────────────────────────────────────

// Handle CORS preflight
if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') {
    http_response_code(204);
    header('Access-Control-Allow-Origin: *');
    header('Access-Control-Allow-Methods: GET, POST, DELETE, OPTIONS');
    header('Access-Control-Allow-Headers: Content-Type');
    exit;
}

// Set CORS headers for all responses
header('Access-Control-Allow-Origin: *');
header('Access-Control-Allow-Methods: GET, POST, DELETE, OPTIONS');
header('Access-Control-Allow-Headers: Content-Type');

// Get request method and path
$method = $_SERVER['REQUEST_METHOD'];

// Try PATH_INFO first, then parse from REQUEST_URI
if (!empty($_SERVER['PATH_INFO'])) {
    $path_info = $_SERVER['PATH_INFO'];
} else {
    // Extract path after api.php from REQUEST_URI
    $req_uri = urldecode($_SERVER['REQUEST_URI']);
    $pos = strpos($req_uri, 'api.php');
    $path_info = ($pos !== false) ? substr($req_uri, $pos + 7) : '/';
    if ($path_info === '' || $path_info === false) $path_info = '/';
}
$path_parts = array_values(array_filter(explode('/', $path_info), fn($p) => $p !== ''));
if (count($path_parts) > 0 && $path_parts[0] === 'api') { array_shift($path_parts); $path_parts = array_values($path_parts); }

// Route: GET /api/ping
if ($method === 'GET' && count($path_parts) === 1 && $path_parts[0] === 'ping') {
    json_response([
        'ok' => true,
        'version' => '2.0',
        'timestamp' => (new DateTime('now', new DateTimeZone('UTC')))->format('c')
    ]);
}

// Route: GET /api/data
if ($method === 'GET' && count($path_parts) === 1 && $path_parts[0] === 'data') {
    $data = load_data();

    // Build inventory
    $inventory = [];
    foreach (CATEGORY_MAP as $slug => $thai_name) {
        $folder_path = RAW_DATA_DIR . DIRECTORY_SEPARATOR . $thai_name;
        $files = [];
        $last_modified = null;

        if (is_dir($folder_path)) {
            try {
                $dir_contents = scandir($folder_path);
                foreach ($dir_contents as $f) {
                    if ($f[0] === '.') continue;

                    $fp = $folder_path . DIRECTORY_SEPARATOR . $f;
                    if (is_file($fp)) {
                        $ext = strtolower(pathinfo($f, PATHINFO_EXTENSION));
                        if (!in_array($ext, ['xlsx', 'xls', 'csv'])) continue;
                        $files[] = $f;
                        $mtime = filemtime($fp);
                        if ($last_modified === null || $mtime > $last_modified) {
                            $last_modified = $mtime;
                        }
                    }
                }
                sort($files);
            } catch (\Throwable $e) {
                error_log("Error reading folder: " . $e->getMessage());
            }
        }

        $last_modified_str = null;
        if ($last_modified !== null) {
            $dt = new DateTime('@' . $last_modified);
            $dt->setTimezone(new DateTimeZone('Asia/Bangkok'));
            $last_modified_str = $dt->format('d/m/Y H:i');
        }

        $inventory[$slug] = [
            'thai_name' => $thai_name,
            'files' => $files,
            'count' => count($files),
            'last_modified' => $last_modified_str
        ];
    }

    // Load notes
    $notes = [];
    $notes_file = RAW_DATA_DIR . DIRECTORY_SEPARATOR . 'notes.json';
    if (file_exists($notes_file)) {
        try {
            $notes = json_decode(file_get_contents($notes_file), true) ?: [];
        } catch (\Throwable $e) {
            error_log("Error loading notes.json: " . $e->getMessage());
        }
    }

    json_response([
        'ok' => true,
        'inventory' => $inventory,
        'dead_meter' => $data['dead_meter'],
        'notes' => $notes
    ]);
}

// Route: POST /api/pre-check/<category>
// 2-step upload flow: pre-check files and prepare temp batch directory
if ($method === 'POST' && count($path_parts) === 2 && $path_parts[0] === 'pre-check') {
    $category = $path_parts[1];

    if (!isset(CATEGORY_MAP[$category])) {
        json_response([
            'ok' => false,
            'error' => 'ไม่รู้จัก category: ' . $category
        ], 400);
    }

    // Get date from form data
    $data_date = isset($_POST['data_date']) ? trim($_POST['data_date']) : '';
    if (!$data_date) {
        json_response([
            'ok' => false,
            'error' => 'กรุณาระบุวันที่ดึงข้อมูล (data_date)'
        ], 400);
    }

    [$date_key, $date_label] = parse_date_key($data_date);
    if (!$date_key) {
        json_response([
            'ok' => false,
            'error' => 'รูปแบบวันที่ไม่ถูกต้อง: ' . $data_date . ' (ใช้ YYYY-MM-DD เช่น 2569-01-16)'
        ], 400);
    }

    // Get uploaded files
    if (!isset($_FILES['files'])) {
        json_response([
            'ok' => false,
            'error' => 'ไม่ได้เลือกไฟล์'
        ], 400);
    }

    $files = $_FILES['files'];

    // Normalize to array of files
    if (!is_array($files['name'])) {
        $files = [
            'name' => [$files['name']],
            'type' => [$files['type']],
            'tmp_name' => [$files['tmp_name']],
            'error' => [$files['error']],
            'size' => [$files['size']]
        ];
    }

    $folder_path = RAW_DATA_DIR . DIRECTORY_SEPARATOR . CATEGORY_MAP[$category];
    if (!is_dir($folder_path)) {
        mkdir($folder_path, 0755, true);
    }

    $PREFIX_MAP = ['abnormal' => 'METER'];

    // Create batch ID and temp directory
    $batch_id = bin2hex(random_bytes(8)); // e.g., "a1b2c3d4e5f6g7h8"
    $tmp_batch_dir = BASE_DIR . DIRECTORY_SEPARATOR . '__tmp_upload' . DIRECTORY_SEPARATOR . $batch_id;
    if (!is_dir($tmp_batch_dir)) {
        mkdir($tmp_batch_dir, 0755, true);
    }

    $preview = [];
    $errors = [];

    for ($i = 0; $i < count($files['name']); $i++) {
        if ($files['error'][$i] !== UPLOAD_ERR_OK) {
            $errors[] = [
                'filename' => $files['name'][$i],
                'error' => 'Upload failed (error code: ' . $files['error'][$i] . ')'
            ];
            continue;
        }

        $filename = trim($files['name'][$i]);
        if (!$filename) continue;

        try {
            // ── ตรวจสอบรูปแบบไฟล์ก่อน upload ──
            if ($category === 'abnormal' && preg_match('/\.xlsx?$/i', $filename)) {
                $validation = validate_meter_file($files['tmp_name'][$i]);
                if (!$validation['valid']) {
                    $errors[] = [
                        'filename' => $filename,
                        'error' => '⚠️ ' . $validation['message']
                    ];
                    continue;
                }
            }

            $prefix = isset(PREFIX_MAP[$category]) ? PREFIX_MAP[$category] : strtoupper($category);
            $pathinfo = pathinfo($filename);
            $name_only = $pathinfo['filename'];
            $ext = isset($pathinfo['extension']) ? '.' . $pathinfo['extension'] : '.xlsx';

            // Extract branch code (4 digits like 1102)
            $branch = null;
            if (preg_match('/(\d{4})/', $name_only, $m)) {
                $code = $m[1];
                $branch = isset(BRANCH_CODE_MAP[$code]) ? BRANCH_CODE_MAP[$code] : null;
            }

            // ── For abnormal files: try to extract billing_month from Excel first ──
            $effective_date_key = $date_key;
            $parsed = null;
            if ($category === 'abnormal' && $phpSpreadsheet) {
                try {
                    $parsed = parse_dead_meter_file($files['tmp_name'][$i]);
                    if ($parsed && isset($parsed['billing_month']) && $parsed['billing_month']) {
                        // billing_month is in format "YYYY-MM" e.g., "2569-03"
                        $effective_date_key = str_replace('-', '', $parsed['billing_month']); // e.g., "256903"
                    }
                } catch (\Throwable $pe) {
                    // If parse fails, just use the form date_key
                }
            }

            // Create date suffix for filename
            $date_suffix = $effective_date_key;

            // Build the new filename
            $new_name = null;
            if ($branch) {
                $code = null;
                if (preg_match('/(\d{4})/', $name_only, $m)) {
                    $code = $m[1];
                }
                if ($code) {
                    $new_name = $prefix . '_' . $code . '_' . $date_suffix . $ext;
                }
            }
            if (!$new_name) {
                $clean = preg_replace('/[^\w\-.]/', '_', $name_only);
                $clean = trim($clean, '_');
                if (strlen($clean) > 30) {
                    $clean = substr($clean, 0, 30);
                }
                $new_name = $prefix . '_' . $clean . '_' . $date_suffix . $ext;
            }

            $final_path = $folder_path . DIRECTORY_SEPARATOR . $new_name;

            // Check for overwrite
            $will_overwrite = file_exists($final_path);
            $overwrite_file = null;
            if ($will_overwrite) {
                $overwrite_file = $new_name;
            }

            // Copy uploaded file to temp batch directory
            $tmp_file_path = $tmp_batch_dir . DIRECTORY_SEPARATOR . $new_name;
            if (!copy($files['tmp_name'][$i], $tmp_file_path)) {
                throw new Exception('Failed to copy file to temp batch directory');
            }
            chmod($tmp_file_path, 0644);

            // Add to preview
            $preview[] = [
                'original' => $filename,
                'new_name' => $new_name,
                'will_overwrite' => $will_overwrite,
                'overwrite_file' => $overwrite_file,
                'branch' => $branch
            ];
        } catch (\Throwable $e) {
            $errors[] = [
                'filename' => $filename,
                'error' => $e->getMessage()
            ];
        }
    }

    json_response([
        'ok' => true,
        'batch_id' => $batch_id,
        'category' => $category,
        'date_key' => $date_key,
        'date_label' => $date_label,
        'preview' => $preview,
        'errors' => $errors
    ]);
}

// Route: POST /api/upload-confirm/<batch_id>
// 2-step upload flow: confirm and move files from temp to final directory
if ($method === 'POST' && count($path_parts) === 2 && $path_parts[0] === 'upload-confirm') {
    $batch_id = $path_parts[1];
    $category = isset($_POST['category']) ? trim($_POST['category']) : '';
    $date_key = isset($_POST['date_key']) ? trim($_POST['date_key']) : '';

    if (!$category || !isset(CATEGORY_MAP[$category])) {
        json_response([
            'ok' => false,
            'error' => 'ไม่รู้จัก category: ' . $category
        ], 400);
    }

    $tmp_batch_dir = BASE_DIR . DIRECTORY_SEPARATOR . '__tmp_upload' . DIRECTORY_SEPARATOR . $batch_id;
    if (!is_dir($tmp_batch_dir)) {
        json_response([
            'ok' => false,
            'error' => 'Batch directory not found'
        ], 400);
    }

    $folder_path = RAW_DATA_DIR . DIRECTORY_SEPARATOR . CATEGORY_MAP[$category];
    if (!is_dir($folder_path)) {
        mkdir($folder_path, 0755, true);
    }

    $data = load_data();

    // Ensure snapshot exists
    if (!isset($data['dead_meter']['snapshots'][$date_key])) {
        $data['dead_meter']['snapshots'][$date_key] = [
            'date_label' => '',
            'data' => [],
            'total_meters' => [],
            'files' => []
        ];
    }
    $snapshot = &$data['dead_meter']['snapshots'][$date_key];

    $results = [];
    $errors = [];

    // Get all files from temp batch directory
    $batch_files = array_diff(scandir($tmp_batch_dir), ['.', '..']);

    foreach ($batch_files as $new_name) {
        try {
            $tmp_file_path = $tmp_batch_dir . DIRECTORY_SEPARATOR . $new_name;
            $final_path = $folder_path . DIRECTORY_SEPARATOR . $new_name;

            // Check if file will overwrite
            $will_overwrite = file_exists($final_path);

            // If overwriting, delete old files with same stem but different extension
            if ($will_overwrite) {
                $stem = preg_replace('/\.[^.]+$/', '', $new_name);
                foreach (scandir($folder_path) as $existing) {
                    if ($existing[0] === '.') continue;
                    $existing_stem = preg_replace('/\.[^.]+$/', '', $existing);
                    if ($existing_stem === $stem) {
                        $existing_path = $folder_path . DIRECTORY_SEPARATOR . $existing;
                        if (is_file($existing_path)) {
                            unlink($existing_path);
                        }
                    }
                }
            }

            // Move file from temp to final directory
            if (!rename($tmp_file_path, $final_path)) {
                throw new Exception('Failed to move file to final directory');
            }
            chmod($final_path, 0644);

            // Parse Excel and store data (for abnormal category)
            if ($category === 'abnormal' && preg_match('/\.xlsx?$/i', $new_name) && $phpSpreadsheet) {
                try {
                    $parsed = parse_dead_meter_file($final_path);
                    if ($parsed) {
                        // Extract branch code from filename
                        $m = null;
                        if (preg_match('/METER_(\d{4})_/', $new_name, $m)) {
                            $code = $m[1];
                            $branch = isset(BRANCH_CODE_MAP[$code]) ? BRANCH_CODE_MAP[$code] : null;
                            if ($branch) {
                                $snapshot['data'][$branch] = $parsed;
                                $snapshot['files'][$branch] = $new_name;
                            }
                        }
                    }
                } catch (\Throwable $pe) {
                    // Log parse error but don't fail the upload
                    error_log('Parse error for ' . $new_name . ': ' . $pe->getMessage());
                }
            }

            $results[] = [
                'filename' => $new_name,
                'status' => $will_overwrite ? 'overwrite' : 'success',
                'message' => $new_name
            ];
        } catch (\Throwable $e) {
            $errors[] = [
                'filename' => $new_name,
                'error' => $e->getMessage()
            ];
        }
    }

    // Update data
    $data['dead_meter']['latest'] = $date_key;
    save_data($data);

    // Write upload log
    if (!empty($results)) {
        $log_file = __DIR__ . DIRECTORY_SEPARATOR . 'upload_log.json';
        $log = file_exists($log_file) ? (json_decode(file_get_contents($log_file), true) ?: []) : [];
        $log[] = [
            'time' => date('Y-m-d H:i:s'),
            'category' => $category,
            'files' => array_map(function($r) { return $r['filename']; }, $results),
            'count' => count($results)
        ];
        if (count($log) > 200) $log = array_slice($log, -200);
        file_put_contents($log_file, json_encode($log, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT));
    }

    // Cleanup temp batch directory
    if (is_dir($tmp_batch_dir)) {
        foreach (scandir($tmp_batch_dir) as $f) {
            if ($f[0] !== '.') {
                $fpath = $tmp_batch_dir . DIRECTORY_SEPARATOR . $f;
                if (is_file($fpath)) unlink($fpath);
            }
        }
        rmdir($tmp_batch_dir);
    }

    json_response([
        'ok' => true,
        'category' => $category,
        'results' => $results,
        'errors' => $errors,
        'dead_meter' => $data['dead_meter']
    ]);
}

// Route: POST /api/upload/<category>
if ($method === 'POST' && count($path_parts) === 2 && $path_parts[0] === 'upload') {
    $category = $path_parts[1];

    if (!isset(CATEGORY_MAP[$category])) {
        json_response([
            'ok' => false,
            'error' => 'ไม่รู้จัก category: ' . $category
        ], 400);
    }

    // Get date from form data
    $data_date = isset($_POST['data_date']) ? trim($_POST['data_date']) : '';
    if (!$data_date) {
        json_response([
            'ok' => false,
            'error' => 'กรุณาระบุวันที่ดึงข้อมูล (data_date)'
        ], 400);
    }

    [$date_key, $date_label] = parse_date_key($data_date);
    if (!$date_key) {
        json_response([
            'ok' => false,
            'error' => 'รูปแบบวันที่ไม่ถูกต้อง: ' . $data_date . ' (ใช้ YYYY-MM-DD เช่น 2569-01-16)'
        ], 400);
    }

    // Get uploaded files
    if (!isset($_FILES['files'])) {
        json_response([
            'ok' => false,
            'error' => 'ไม่ได้เลือกไฟล์'
        ], 400);
    }

    $files = $_FILES['files'];

    // Normalize to array of files
    if (!is_array($files['name'])) {
        $files = [
            'name' => [$files['name']],
            'type' => [$files['type']],
            'tmp_name' => [$files['tmp_name']],
            'error' => [$files['error']],
            'size' => [$files['size']]
        ];
    }

    $folder_path = RAW_DATA_DIR . DIRECTORY_SEPARATOR . CATEGORY_MAP[$category];
    if (!is_dir($folder_path)) {
        mkdir($folder_path, 0755, true);
    }

    $PREFIX_MAP = ['abnormal' => 'METER'];
    $data = load_data();

    // Ensure snapshot exists for this date
    $snapshots = &$data['dead_meter']['snapshots'];
    if (!isset($snapshots[$date_key])) {
        $snapshots[$date_key] = [
            'date_label' => $date_label,
            'data' => [],
            'total_meters' => [],
            'files' => []
        ];
    }
    $snapshot = &$snapshots[$date_key];

    $results = [];
    $errors = [];

    for ($i = 0; $i < count($files['name']); $i++) {
        if ($files['error'][$i] !== UPLOAD_ERR_OK) {
            $errors[] = [
                'filename' => $files['name'][$i],
                'error' => 'Upload failed (error code: ' . $files['error'][$i] . ')'
            ];
            continue;
        }

        $filename = trim($files['name'][$i]);
        if (!$filename) continue;

        try {
            // ── ตรวจสอบรูปแบบไฟล์ก่อน upload ──
            if ($category === 'abnormal' && preg_match('/\.xlsx?$/i', $filename)) {
                $validation = validate_meter_file($files['tmp_name'][$i]);
                if (!$validation['valid']) {
                    $errors[] = [
                        'filename' => $filename,
                        'error' => '⚠️ ' . $validation['message']
                    ];
                    continue;
                }
            }

            $prefix = isset(PREFIX_MAP[$category]) ? PREFIX_MAP[$category] : strtoupper($category);
            $pathinfo = pathinfo($filename);
            $name_only = $pathinfo['filename'];
            $ext = isset($pathinfo['extension']) ? '.' . $pathinfo['extension'] : '.xlsx';

            // Extract branch code (4 digits like 1102)
            $branch = null;
            if (preg_match('/(\d{4})/', $name_only, $m)) {
                $code = $m[1];
                $branch = isset(BRANCH_CODE_MAP[$code]) ? BRANCH_CODE_MAP[$code] : null;
            }

            // ── For abnormal files: try to extract billing_month from Excel first ──
            // This gives us accurate date info if the form date was wrong
            $effective_date_key = $date_key;
            $parsed = null;
            if ($category === 'abnormal' && $phpSpreadsheet) {
                try {
                    $parsed = parse_dead_meter_file($files['tmp_name'][$i]);
                    if ($parsed && isset($parsed['billing_month']) && $parsed['billing_month']) {
                        // billing_month is in format "YYYY-MM" e.g., "2569-03"
                        $effective_date_key = str_replace('-', '', $parsed['billing_month']); // e.g., "256903"
                    }
                } catch (\Throwable $pe) {
                    // If parse fails, just use the form date_key
                    // File will still be moved with form date
                }
            }

            // Create date suffix for filename (e.g., "25690317" or "256903" from billing)
            $date_suffix = $effective_date_key;

            // Build the new filename
            $new_name = null;
            if ($branch) {
                $code = null;
                if (preg_match('/(\d{4})/', $name_only, $m)) {
                    $code = $m[1];
                }
                if ($code) {
                    $new_name = $prefix . '_' . $code . '_' . $date_suffix . $ext;
                }
            }
            if (!$new_name) {
                $clean = preg_replace('/[^\w\-.]/', '_', $name_only);
                $clean = trim($clean, '_');
                if (strlen($clean) > 30) {
                    $clean = substr($clean, 0, 30);
                }
                $new_name = $prefix . '_' . $clean . '_' . $date_suffix . $ext;
            }

            $dest_path = $folder_path . DIRECTORY_SEPARATOR . $new_name;

            // Check if overwriting
            $overwrite = file_exists($dest_path);

            // Move uploaded file
            if (!move_uploaded_file($files['tmp_name'][$i], $dest_path)) {
                throw new Exception('Failed to move uploaded file');
            }
            chmod($dest_path, 0644);

            // Parse Excel file again (if not already done above) and store parsed data
            if ($category === 'abnormal' && $branch && $phpSpreadsheet && !$parsed) {
                try {
                    $parsed = parse_dead_meter_file($dest_path);
                    $snapshot['data'][$branch] = $parsed;
                    $snapshot['files'][$branch] = $new_name;
                } catch (\Throwable $pe) {
                    $errors[] = [
                        'filename' => $new_name,
                        'error' => 'parse ล้มเหลว: ' . $pe->getMessage()
                    ];
                }
            } else if ($category === 'abnormal' && $branch && $parsed) {
                // Use already-parsed data
                $snapshot['data'][$branch] = $parsed;
                $snapshot['files'][$branch] = $new_name;
            }

            $results[] = [
                'filename' => $new_name,
                'original' => $filename,
                'status' => $overwrite ? 'overwrite' : 'success',
                'message' => $filename . ' → ' . $new_name,
                'branch' => $branch,
                'dead_count' => $parsed ? $parsed['total'] : null
            ];
        } catch (\Throwable $e) {
            $errors[] = [
                'filename' => $filename,
                'error' => $e->getMessage()
            ];
        }
    }

    // Update latest snapshot
    $data['dead_meter']['latest'] = $date_key;
    save_data($data);

    // Write upload log
    if (!empty($results)) {
        $log_file = __DIR__ . DIRECTORY_SEPARATOR . 'upload_log.json';
        $log = file_exists($log_file) ? (json_decode(file_get_contents($log_file), true) ?: []) : [];
        $log[] = [
            'time' => date('Y-m-d H:i:s'),
            'category' => $category,
            'files' => array_map(function($r) { return $r['filename']; }, $results),
            'count' => count($results)
        ];
        if (count($log) > 200) $log = array_slice($log, -200);
        file_put_contents($log_file, json_encode($log, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT));
    }

    json_response([
        'ok' => true,
        'category' => $category,
        'thai_name' => CATEGORY_MAP[$category],
        'date_key' => $date_key,
        'date_label' => $date_label,
        'results' => $results,
        'errors' => $errors,
        'dead_meter' => $data['dead_meter']
    ]);
}

// Route: DELETE /api/data/<category>/<snapshot_date>/<filename>
if ($method === 'DELETE' && count($path_parts) === 4 && $path_parts[0] === 'data') {
    $category = $path_parts[1];
    $snapshot_date = $path_parts[2];
    $filename = $path_parts[3];

    if (!isset(CATEGORY_MAP[$category])) {
        json_response([
            'ok' => false,
            'error' => 'ไม่รู้จัก category: ' . $category
        ], 400);
    }

    $folder_path = RAW_DATA_DIR . DIRECTORY_SEPARATOR . CATEGORY_MAP[$category];
    $file_path = $folder_path . DIRECTORY_SEPARATOR . $filename;

    // Safety check: ensure file is within folder
    $abs_file = realpath($file_path);
    $abs_folder = realpath($folder_path);

    if (!$abs_file || strpos($abs_file, $abs_folder) !== 0) {
        json_response([
            'ok' => false,
            'error' => 'ไม่อนุญาต'
        ], 403);
    }

    try {
        if (file_exists($file_path)) {
            unlink($file_path);

            // Update data if abnormal category
            if ($category === 'abnormal') {
                if (preg_match('/(\d{4})/', $filename, $m)) {
                    $code = $m[1];
                    $branch = isset(BRANCH_CODE_MAP[$code]) ? BRANCH_CODE_MAP[$code] : null;

                    if ($branch) {
                        $data = load_data();
                        $snapshots = &$data['dead_meter']['snapshots'];

                        if (isset($snapshots[$snapshot_date])) {
                            $snap = &$snapshots[$snapshot_date];
                            unset($snap['data'][$branch]);
                            unset($snap['files'][$branch]);
                            unset($snap['total_meters'][$branch]);

                            // Remove snapshot if empty
                            if (empty($snap['data'])) {
                                unset($snapshots[$snapshot_date]);

                                // Update latest
                                if ($data['dead_meter']['latest'] === $snapshot_date) {
                                    $data['dead_meter']['latest'] = !empty($snapshots) ? max(array_keys($snapshots)) : '';
                                }
                            }

                            save_data($data);
                        }
                    }
                }
            }

            json_response([
                'ok' => true,
                'filename' => $filename,
                'deleted' => true
            ]);
        } else {
            json_response([
                'ok' => false,
                'error' => 'ไม่พบไฟล์'
            ], 404);
        }
    } catch (\Throwable $e) {
        json_response([
            'ok' => false,
            'error' => $e->getMessage()
        ], 500);
    }
}

// Route: DELETE /api/data/<category>/<filename> (backward compat)
if ($method === 'DELETE' && count($path_parts) === 3 && $path_parts[0] === 'data') {
    $category = $path_parts[1];
    $filename = $path_parts[2];

    if (!isset(CATEGORY_MAP[$category])) {
        json_response([
            'ok' => false,
            'error' => 'ไม่รู้จัก category: ' . $category
        ], 400);
    }

    $folder_path = RAW_DATA_DIR . DIRECTORY_SEPARATOR . CATEGORY_MAP[$category];
    $file_path = $folder_path . DIRECTORY_SEPARATOR . $filename;

    $abs_file = realpath($file_path);
    $abs_folder = realpath($folder_path);

    if (!$abs_file || strpos($abs_file, $abs_folder) !== 0) {
        json_response([
            'ok' => false,
            'error' => 'ไม่อนุญาต'
        ], 403);
    }

    try {
        if (file_exists($file_path)) {
            unlink($file_path);

            if ($category === 'abnormal') {
                if (preg_match('/(\d{4})/', $filename, $m)) {
                    $code = $m[1];
                    $branch = isset(BRANCH_CODE_MAP[$code]) ? BRANCH_CODE_MAP[$code] : null;

                    if ($branch) {
                        $data = load_data();
                        $snapshots = &$data['dead_meter']['snapshots'];

                        // Remove from all snapshots with this file
                        foreach ($snapshots as $sk => &$snap) {
                            if (isset($snap['files'][$branch]) && $snap['files'][$branch] === $filename) {
                                unset($snap['data'][$branch]);
                                unset($snap['files'][$branch]);
                                unset($snap['total_meters'][$branch]);
                            }
                        }

                        // Remove empty snapshots
                        $empty_keys = [];
                        foreach ($snapshots as $k => $v) {
                            if (empty($v['data'])) {
                                $empty_keys[] = $k;
                            }
                        }

                        foreach ($empty_keys as $k) {
                            unset($snapshots[$k]);
                        }

                        if (in_array($data['dead_meter']['latest'], $empty_keys)) {
                            $data['dead_meter']['latest'] = !empty($snapshots) ? max(array_keys($snapshots)) : '';
                        }

                        save_data($data);
                    }
                }
            }

            json_response([
                'ok' => true,
                'filename' => $filename,
                'deleted' => true
            ]);
        } else {
            json_response([
                'ok' => false,
                'error' => 'ไม่พบไฟล์'
            ], 404);
        }
    } catch (\Throwable $e) {
        json_response([
            'ok' => false,
            'error' => $e->getMessage()
        ], 500);
    }
}

// Route: POST /api/notes/<slug>
// Accepts category slugs (e.g. 'abnormal') and derived keys (e.g. 'abnormal_source_url')
if ($method === 'POST' && count($path_parts) === 2 && $path_parts[0] === 'notes') {
    $slug = $path_parts[1];

    // Validate: must be a known category OR a derived key like {category}_source_url
    $base_slug = preg_replace('/_source_url$/', '', $slug);
    if (!isset(CATEGORY_MAP[$base_slug])) {
        json_response([
            'ok' => false,
            'error' => 'invalid slug'
        ], 400);
    }

    $body = json_decode(file_get_contents('php://input'), true) ?: [];
    $text = isset($body['text']) ? $body['text'] : '';

    $notes_file = RAW_DATA_DIR . DIRECTORY_SEPARATOR . 'notes.json';
    $notes = [];

    if (file_exists($notes_file)) {
        try {
            $notes = json_decode(file_get_contents($notes_file), true) ?: [];
        } catch (\Throwable $e) {
            error_log("Error loading notes: " . $e->getMessage());
        }
    }

    $notes[$slug] = $text;

    $json = json_encode($notes, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT);
    file_put_contents($notes_file, $json);
    chmod($notes_file, 0644);

    json_response(['ok' => true]);
}

// Route: POST /api/open-folder
if ($method === 'POST' && count($path_parts) === 1 && $path_parts[0] === 'open-folder') {
    $folder = realpath(RAW_DATA_DIR);
    if ($folder && is_dir($folder)) {
        $folder_win = str_replace('/', '\\', $folder);
        pclose(popen('start explorer "' . $folder_win . '"', 'r'));
        json_response(['ok' => true, 'path' => $folder]);
    } else {
        json_response(['ok' => false, 'error' => 'Folder not found: ' . RAW_DATA_DIR]);
    }
}

// Route: POST /api/open-main
if ($method === 'POST' && count($path_parts) === 1 && $path_parts[0] === 'open-main') {
    // PHP on XAMPP can't safely open browser
    // Return parent directory path instead
    $parent_dir = dirname(BASE_DIR);
    json_response([
        'ok' => true,
        'path' => $parent_dir,
        'note' => 'Parent directory path returned; OS-specific opening not available in PHP'
    ]);
}

// ─── Route: GET /api/meter-data (Dual Mode) ──────────────────────────────
// Parses METER_XXXX.xlsx from ข้อมูลดิบ/มาตรวัดน้ำผิดปกติ/ directly
// Also reads OIS for TOTAL_METERS
// Returns: {dead_meter: {branch: {total, sizes}}, total_meters: {branch: N}}

if ($method === 'GET' && count($path_parts) === 1 && $path_parts[0] === 'meter-data') {
    if (!$phpSpreadsheet) {
        json_response(['ok' => false, 'error' => 'PhpSpreadsheet not available'], 500);
    }

    $meter_dir = RAW_DATA_DIR . DIRECTORY_SEPARATOR . CATEGORY_MAP['abnormal'];
    if (!is_dir($meter_dir)) {
        json_response(['ok' => true, 'has_data' => false, 'data' => new stdClass()]);
    }

    // Check cache (combine meter + OIS folder mtimes)
    $cached = load_cache('meter_data', $meter_dir);
    if ($cached !== null) {
        json_response($cached);
    }

    // OIS sheet name → branch name mapping
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

    // --- 1. Parse DEAD_METER from METER_XXXX.xlsx ---
    $dead_meter = [];
    $files = glob($meter_dir . DIRECTORY_SEPARATOR . 'METER_*.xlsx') ?: [];
    sort($files);

    foreach ($files as $file) {
        $basename = pathinfo($file, PATHINFO_FILENAME);
        $code = str_replace('METER_', '', $basename);
        $branch = isset(BRANCH_CODE_MAP[$code]) ? BRANCH_CODE_MAP[$code] : null;
        if (!$branch) continue;

        try {
            $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($file);
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
            foreach (METER_SIZES as $sz) { $sizes[$sz] = 0; }
            $total = 0;

            for ($r = $hdr + 1; $r <= $maxRow; $r++) {
                $cid = $ws->getCell([$cCid, $r])->getValue();
                if ($cid === null) continue;
                $cid = trim((string)$cid);
                if (isset($seen[$cid])) continue;

                // Condition 1: สภาพมาตร = "มาตรไม่เดิน"
                $condition = $ws->getCell([$cCondition, $r])->getValue();
                if ($condition === null || trim((string)$condition) !== 'มาตรไม่เดิน') continue;

                // Condition 2: การเปลี่ยนมาตร ≠ "เปลี่ยนแล้ว"
                $change = $ws->getCell([$cChange, $r])->getValue();
                if ($change !== null && trim((string)$change) === 'เปลี่ยนแล้ว') continue;

                $seen[$cid] = true;
                $total++;

                // ขนาดมาตร
                $sv = $ws->getCell([$cSize, $r])->getValue();
                if ($sv !== null) {
                    $ns = normalize_size($sv);
                    if ($ns !== null && isset($sizes[$ns])) { $sizes[$ns]++; }
                }
            }

            $dead_meter[$branch] = ['total' => $total, 'sizes' => $sizes];
            $spreadsheet->disconnectWorksheets();
            unset($spreadsheet);
        } catch (\Throwable $e) {
            error_log("Meter: Cannot parse $file: " . $e->getMessage());
            $sizes = [];
            foreach (METER_SIZES as $sz) { $sizes[$sz] = 0; }
            $dead_meter[$branch] = ['total' => 0, 'sizes' => $sizes];
        }
    }

    // --- 2. Parse TOTAL_METERS from OIS ---
    $total_meters = [];
    $ois_dir = dirname(BASE_DIR) . DIRECTORY_SEPARATOR . 'Dashboard_Leak'
             . DIRECTORY_SEPARATOR . 'ข้อมูลดิบ' . DIRECTORY_SEPARATOR . 'OIS';

    if (is_dir($ois_dir)) {
        $ois_files = array_merge(
            glob($ois_dir . DIRECTORY_SEPARATOR . 'OIS_*.xls') ?: [],
            glob($ois_dir . DIRECTORY_SEPARATOR . 'OIS_*.xlsx') ?: []
        );
        $ois_files = array_unique($ois_files);
        sort($ois_files);
        $latest_ois = !empty($ois_files) ? end($ois_files) : null;

        if ($latest_ois) {
            try {
                $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($latest_ois);

                // ╔══════════════════════════════════════════════════════════════╗
                // ║ ⚠️  [OIS METER DATA] MAPPED SHEET ACCESS                 ║
                // ║                                                              ║
                // ║ อ่านจำนวนมิเตอร์จากไฟล์ OIS สาขาต่างๆ                     ║
                // ║ ใช้ $OIS_SHEET_MAP เพื่อเข้าถึงชีทเฉพาะตามชื่อ            ║
                // ║ ค้นหาคอลัมน์เดือนล่าสุด (ต.ค.-ก.ย.) จากแถว 6             ║
                // ║                                                              ║
                // ║ Sheets to PROCESS: Sheets in OIS_SHEET_MAP (branch sheets) ║
                // ║ Sheets to AVOID: "กราฟ", "สรุป", "รวม", "Chart", "Summary" ║
                // ╚══════════════════════════════════════════════════════════════╝
                // Auto-detect latest month column from first branch sheet
                // Python xlrd: 0-based → row 5=6th row, col 5=6th col (F)
                // PhpSpreadsheet: 1-based → row 6, col 6 (F) to col 17 (Q)
                $month_col = 6; // default ต.ค. (col F = 6 in 1-based)
                $first_sheet_name = array_keys($OIS_SHEET_MAP)[0];
                $first_sheet = null;
                try { $first_sheet = $spreadsheet->getSheetByName($first_sheet_name); } catch (\Throwable $e) {}

                if ($first_sheet) {
                    // Row 6 = "ผู้ใช้น้ำต้นงวด" (1-based), check cols 6-17 for data
                    for ($c = 6; $c <= 17; $c++) {
                        $v = $first_sheet->getCell([$c, 6])->getValue();
                        if ($v !== null && $v !== '' && $v != 0) {
                            $month_col = $c;
                        }
                    }
                }

                // ╔══════════════════════════════════════════════════════════════╗
                // ║ ⚠️  [BRANCH ITERATION] MAPPED SHEET LOOP                   ║
                // ║                                                              ║
                // ║ อ่านจำนวนมิเตอร์จากแต่ละสาขา โดยใช้ $OIS_SHEET_MAP      ║
                // ║ กำหนดการแมพชีทกับชื่อสาขา ไม่อ่านชีตสรุป/กราฟ           ║
                // ║ ข้อมูลจำนวนมิเตอร์อยู่ที่แถว 6 คอลัมน์ $month_col        ║
                // ║                                                              ║
                // ║ Sheets to PROCESS: Branch sheets in OIS_SHEET_MAP          ║
                // ║ Sheets to AVOID: "กราฟ", "สรุป", "รวม", "Chart", "Summary" ║
                // ╚══════════════════════════════════════════════════════════════╝
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
            } catch (\Throwable $e) {
                error_log("Meter: Cannot read OIS: " . $e->getMessage());
            }
        }
    }

    $response = [
        'ok' => true,
        'has_data' => !empty($dead_meter),
        'dead_meter' => $dead_meter,
        'total_meters' => $total_meters
    ];
    save_cache('meter_data', $response);
    json_response($response);
}

// ───────────────────────────────────────────────────────────────────────────
// Route: POST /api/clear-cache — clear API cache for fresh data
// ───────────────────────────────────────────────────────────────────────────
if ($method === 'POST' && count($path_parts) === 1 && $path_parts[0] === 'clear-cache') {
    $cache_dir = defined('CACHE_DIR') ? CACHE_DIR : (__DIR__ . DIRECTORY_SEPARATOR . '.cache');
    $cleared = 0;
    if (is_dir($cache_dir)) {
        foreach (glob($cache_dir . '/meter_*.json') as $cf) {
            if (@unlink($cf)) $cleared++;
        }
    }
    json_response(['ok' => true, 'cleared' => $cleared]);
}

// 404 - Route not found
json_response([
    'ok' => false,
    'error' => 'Route not found: ' . $method . ' ' . $path_info
], 404);
