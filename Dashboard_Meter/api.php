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
define('CACHE_TTL', 60);
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

// ─── PhpSpreadsheet Loader ─────────────────────────────────────────────────

$composerAutoload = dirname(BASE_DIR) . DIRECTORY_SEPARATOR . 'vendor' . DIRECTORY_SEPARATOR . 'autoload.php';
$phpSpreadsheet = null;

if (file_exists($composerAutoload)) {
    require_once $composerAutoload;
    try {
        $phpSpreadsheet = true;
    } catch (Exception $e) {
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
    } catch (Exception $e) {
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
            $val = trim((string)($worksheet->getCellByColumnAndRow($c, $r)->getValue() ?? ''));
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
        return ['valid' => true, 'message' => ''];
    }
    try {
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($tmp_path);
        $worksheet = $spreadsheet->getActiveSheet();
        $det = detect_meter_columns($worksheet);
        $spreadsheet->disconnectWorksheets();

        if ($det['fallback']) {
            return [
                'valid' => false,
                'message' => 'ไม่พบหัวคอลัมน์ที่คาดหวัง (CA/รหัสผู้ใช้น้ำ, สภาพมาตร, ขนาดมาตร) — รูปแบบไฟล์ไม่ตรงกับที่ระบบรองรับ'
            ];
        }
        return ['valid' => true, 'message' => ''];
    } catch (Exception $e) {
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
            $worksheet = $spreadsheet->getActiveSheet();

            $det = detect_meter_columns($worksheet);
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
                $cid = $worksheet->getCellByColumnAndRow($cCid, $r)->getValue();
                if ($cid === null) continue;

                $cid = trim((string)$cid);
                if (isset($seen[$cid])) continue;

                // Condition 1: สภาพมาตร = "มาตรไม่เดิน"
                $condition = $worksheet->getCellByColumnAndRow($cCondition, $r)->getValue();
                if ($condition === null || trim((string)$condition) !== "มาตรไม่เดิน") continue;

                // Condition 2: การเปลี่ยนมาตร ≠ "เปลี่ยนแล้ว"
                $change = $worksheet->getCellByColumnAndRow($cChange, $r)->getValue();
                if ($change !== null && trim((string)$change) === "เปลี่ยนแล้ว") continue;

                $seen[$cid] = true;
                $total++;

                $sv = $worksheet->getCellByColumnAndRow($cSize, $r)->getValue();
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
                $v = $worksheet->getCellByColumnAndRow($cBilling, $r)->getValue();
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
    } catch (Exception $e) {
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
            } catch (Exception $e) {
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
        } catch (Exception $e) {
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

            // Create date suffix for filename (e.g., "25690317")
            $date_suffix = str_replace('-', '', $date_key);

            // Extract branch code (4 digits like 1102)
            $branch = null;
            $new_name = null;

            if (preg_match('/(\d{4})/', $name_only, $m)) {
                $code = $m[1];
                $new_name = $prefix . '_' . $code . '_' . $date_suffix . $ext;
                $branch = isset(BRANCH_CODE_MAP[$code]) ? BRANCH_CODE_MAP[$code] : null;
            } else {
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

            // Parse Excel file if abnormal category
            $parsed = null;
            if ($category === 'abnormal' && $branch && $phpSpreadsheet) {
                try {
                    $parsed = parse_dead_meter_file($dest_path);
                    $snapshot['data'][$branch] = $parsed;
                    $snapshot['files'][$branch] = $new_name;
                } catch (Exception $pe) {
                    $errors[] = [
                        'filename' => $new_name,
                        'error' => 'parse ล้มเหลว: ' . $pe->getMessage()
                    ];
                }
            }

            $results[] = [
                'filename' => $new_name,
                'original' => $filename,
                'status' => $overwrite ? 'overwrite' : 'success',
                'message' => $filename . ' → ' . $new_name,
                'branch' => $branch,
                'dead_count' => $parsed ? $parsed['total'] : null
            ];
        } catch (Exception $e) {
            $errors[] = [
                'filename' => $filename,
                'error' => $e->getMessage()
            ];
        }
    }

    // Update latest snapshot
    $data['dead_meter']['latest'] = $date_key;
    save_data($data);

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
    } catch (Exception $e) {
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
    } catch (Exception $e) {
        json_response([
            'ok' => false,
            'error' => $e->getMessage()
        ], 500);
    }
}

// Route: POST /api/notes/<slug>
if ($method === 'POST' && count($path_parts) === 2 && $path_parts[0] === 'notes') {
    $slug = $path_parts[1];

    if (!isset(CATEGORY_MAP[$slug])) {
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
        } catch (Exception $e) {
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
        } catch (Exception $e) {
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

                // Auto-detect latest month column from first branch sheet
                // Python xlrd: 0-based → row 5=6th row, col 5=6th col (F)
                // PhpSpreadsheet: 1-based → row 6, col 6 (F) to col 17 (Q)
                $month_col = 6; // default ต.ค. (col F = 6 in 1-based)
                $first_sheet_name = array_keys($OIS_SHEET_MAP)[0];
                $first_sheet = null;
                try { $first_sheet = $spreadsheet->getSheetByName($first_sheet_name); } catch (Exception $e) {}

                if ($first_sheet) {
                    // Row 6 = "ผู้ใช้น้ำต้นงวด" (1-based), check cols 6-17 for data
                    for ($c = 6; $c <= 17; $c++) {
                        $v = $first_sheet->getCell([$c, 6])->getValue();
                        if ($v !== null && $v !== '' && $v != 0) {
                            $month_col = $c;
                        }
                    }
                }

                foreach ($OIS_SHEET_MAP as $sheetName => $branchName) {
                    try {
                        $ws = $spreadsheet->getSheetByName($sheetName);
                        $val = $ws->getCell([$month_col, 6])->getValue();
                        $total_meters[$branchName] = ($val !== null && $val !== '') ? intval($val) : 0;
                    } catch (Exception $e) {
                        $total_meters[$branchName] = 0;
                    }
                }

                $spreadsheet->disconnectWorksheets();
                unset($spreadsheet);
            } catch (Exception $e) {
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

// 404 - Route not found
json_response([
    'ok' => false,
    'error' => 'Route not found: ' . $method . ' ' . $path_info
], 404);
