<?php
/**
 * Dashboard Leak — XAMPP (Apache + PHP) Backend API
 * ========================================================================
 * PHP เทียบเท่า Flask server.py
 * รับ upload ไฟล์ → auto-rename ตามหมวดหมู่และวันที่ → ไม่ parse (ให้ build_dashboard.py จัดการ)
 *
 * Architecture:
 *   - Single file handling ALL API routes via PATH_INFO
 *   - .htaccess rewrites /api/* to this file
 *   - Static files (index.html, manage.html) served directly by Apache
 *
 * Categories (auto-created):
 *   - ois:   OIS
 *   - rl:    Real Leak
 *   - mnf:   MNF
 *   - p3:    P3 (special handling)
 *   - activities: Activities
 *   - eu:    หน่วยไฟ
 *   - kpi2:  เกณฑ์วัดน้ำสูญเสีย
 */

// ─── Configuration ─────────────────────────────────────────────────────────

define('BASE_DIR', __DIR__);
define('RAW_DATA_DIR', BASE_DIR . DIRECTORY_SEPARATOR . 'ข้อมูลดิบ');
define('CACHE_DIR', BASE_DIR . DIRECTORY_SEPARATOR . '.cache');
define('CACHE_TTL', 60); // seconds

// Category mapping: URL slug → Thai folder name
const CATEGORY_MAP = [
    'ois' => 'OIS',
    'rl' => 'Real Leak',
    'mnf' => 'MNF',
    'p3' => 'P3',
    'activities' => 'Activities',
    'eu' => 'หน่วยไฟ',
    'kpi2' => 'เกณฑ์วัดน้ำสูญเสีย',
];

// Auto-rename prefix map
const PREFIX_MAP = [
    'ois' => 'OIS',
    'rl' => 'RL',
    'mnf' => 'MNF',
    'p3' => 'P3',
    'activities' => 'ACT',
    'eu' => 'EU',
    'kpi2' => 'KPI2'
];

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

// ─── PhpSpreadsheet Loader ─────────────────────────────────────────────────

$composerAutoload = dirname(BASE_DIR) . DIRECTORY_SEPARATOR . 'vendor' . DIRECTORY_SEPARATOR . 'autoload.php';
$phpSpreadsheet = false;

if (file_exists($composerAutoload)) {
    require_once $composerAutoload;
    try {
        $phpSpreadsheet = class_exists('PhpOffice\\PhpSpreadsheet\\IOFactory');
    } catch (Exception $e) {
        error_log("Warning: PhpSpreadsheet not available: " . $e->getMessage());
    }
} else {
    error_log("Warning: Composer vendor/ not found at " . dirname(BASE_DIR));
}

// ─── Cache Setup ──────────────────────────────────────────────────────────

if (!is_dir(CACHE_DIR)) {
    mkdir(CACHE_DIR, 0755, true);
}

// ─── Branch Normalization ─────────────────────────────────────────────────

const STANDARD_BRANCHES = [
    'ชลบุรี(พ)', 'พัทยา(พ)', 'พนัสนิคม', 'บ้านบึง', 'ศรีราชา', 'แหลมฉบัง',
    'ฉะเชิงเทรา', 'บางปะกง', 'บางคล้า', 'พนมสารคาม', 'ระยอง', 'บ้านฉาง',
    'ปากน้ำประแสร์', 'จันทบุรี', 'ขลุง', 'ตราด', 'คลองใหญ่', 'สระแก้ว',
    'วัฒนานคร', 'อรัญประเทศ', 'ปราจีนบุรี', 'กบินทร์บุรี'
];

const BRANCH_ALIASES = [
    'พนัมสารคาม' => 'พนมสารคาม',
];

function normalize_branch_name($raw_name) {
    if (!$raw_name || !is_string($raw_name)) return null;
    $name = trim($raw_name);
    // Remove leading numbers and dots like "1.", "01."
    $name = preg_replace('/^\d+\.?\s*/', '', $name);
    // Check alias table
    foreach (BRANCH_ALIASES as $alias => $std) {
        if (mb_strpos($name, $alias) !== false) return $std;
    }
    // Try exact match
    foreach (STANDARD_BRANCHES as $std) {
        if ($name === $std) return $std;
    }
    // Try matching core name
    foreach (STANDARD_BRANCHES as $std) {
        $core = preg_replace('/\(.*?\)/', '', $std);
        $core = trim($core);
        $raw_core = preg_replace('/\(.*?\)/', '', $name);
        $raw_core = preg_replace('/\s*(พ|น\.\d+)$/', '', trim($raw_core));
        $raw_core = trim($raw_core);
        if ($raw_core === $core) return $std;
    }
    return null;
}

// ─── Cell Access Helper (PhpSpreadsheet 2.x compatible) ──────────────────

function cellVal($sheet, $col, $row) {
    // PhpSpreadsheet 2.x: use getCell([col, row]) instead of getCellByColumnAndRow
    return $sheet->getCell([$col, $row])->getValue();
}

function cellCalc($sheet, $col, $row) {
    return $sheet->getCell([$col, $row])->getCalculatedValue();
}

// ─── Cache Helpers ────────────────────────────────────────────────────────

function get_folder_mtime($folder_path) {
    $latest = 0;
    if (!is_dir($folder_path)) return 0;
    foreach (scandir($folder_path) as $f) {
        if ($f[0] === '.') continue;
        $fp = $folder_path . DIRECTORY_SEPARATOR . $f;
        if (is_file($fp)) {
            $mt = filemtime($fp);
            if ($mt > $latest) $latest = $mt;
        }
    }
    return $latest;
}

function load_cache($cache_key, $folder_path) {
    $cache_file = CACHE_DIR . DIRECTORY_SEPARATOR . $cache_key . '.json';
    if (!file_exists($cache_file)) return null;
    $cache_mtime = filemtime($cache_file);
    $folder_mtime = get_folder_mtime($folder_path);
    // Cache valid if folder hasn't changed and within TTL
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
header('Content-Type: application/json; charset=utf-8');

// Get request method and path
$method = $_SERVER['REQUEST_METHOD'];

// Try PATH_INFO first, then parse from REQUEST_URI
if (!empty($_SERVER['PATH_INFO'])) {
    $path_info = $_SERVER['PATH_INFO'];
} else {
    $req_uri = urldecode($_SERVER['REQUEST_URI']);
    $pos = strpos($req_uri, 'api.php');
    $path_info = ($pos !== false) ? substr($req_uri, $pos + 7) : '/';
    if ($path_info === '' || $path_info === false) $path_info = '/';
}
$path_parts = array_values(array_filter(explode('/', $path_info), fn($p) => $p !== ''));
if (count($path_parts) > 0 && $path_parts[0] === 'api') { array_shift($path_parts); $path_parts = array_values($path_parts); }

// ───────────────────────────────────────────────────────────────────────────
// Route: GET /api/ping
// ───────────────────────────────────────────────────────────────────────────
if ($method === 'GET' && count($path_parts) === 1 && $path_parts[0] === 'ping') {
    json_response([
        'ok' => true,
        'version' => '1.0',
        'timestamp' => (new DateTime('now', new DateTimeZone('UTC')))->format('c')
    ]);
}

// ───────────────────────────────────────────────────────────────────────────
// Route: GET /api/data
// ───────────────────────────────────────────────────────────────────────────
if ($method === 'GET' && count($path_parts) === 1 && $path_parts[0] === 'data') {
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

    // Load saved notes
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
        'notes' => $notes
    ]);
}

// ───────────────────────────────────────────────────────────────────────────
// ── Validate file format before upload ──
function validate_leak_file($tmp_path, $category) {
    if (!class_exists('PhpOffice\\PhpSpreadsheet\\IOFactory')) {
        return ['valid' => true, 'message' => '']; // ไม่มี library ให้ข้ามการตรวจสอบ
    }
    if (!preg_match('/\.xlsx?$/i', basename($tmp_path))) {
        return ['valid' => true, 'message' => '']; // ไม่ใช่ Excel
    }

    // ตรวจเฉพาะหมวดที่มีการ parse
    $parseable = ['eu', 'kpi2', 'ois', 'rl', 'mnf'];
    if (!in_array($category, $parseable)) {
        return ['valid' => true, 'message' => ''];
    }

    try {
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($tmp_path);
        $sheet = $spreadsheet->getSheet(0);
        $highRow = $sheet->getHighestDataRow();
        $highCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($sheet->getHighestDataColumn());

        if ($highRow < 2) {
            $spreadsheet->disconnectWorksheets();
            return ['valid' => false, 'message' => 'ไฟล์ว่างเปล่าหรือมีเฉพาะ header'];
        }

        if ($category === 'eu') {
            // ค้นหา "สาขา"/"หน่วยงาน" + เดือน (ต.ค.)
            $found_branch = false;
            $found_month = false;
            for ($r = 1; $r <= min($highRow, 10); $r++) {
                for ($c = 1; $c <= min($highCol, 20); $c++) {
                    $v = mb_strtolower(trim((string)($sheet->getCellByColumnAndRow($c, $r)->getValue() ?? '')));
                    if ($v === '') continue;
                    if (mb_strpos($v, 'สาขา') !== false || mb_strpos($v, 'หน่วยงาน') !== false) $found_branch = true;
                    if (mb_strpos($v, 'ต.ค') !== false || mb_strpos($v, 'ตุลาคม') !== false || mb_strpos($v, 'oct') !== false) $found_month = true;
                }
            }
            $spreadsheet->disconnectWorksheets();
            if (!$found_branch && !$found_month) {
                return ['valid' => false, 'message' => 'ไม่พบหัวคอลัมน์ "สาขา" หรือ "เดือน (ต.ค.)" — รูปแบบไฟล์ EU ไม่ตรงกับที่ระบบรองรับ'];
            }
        } elseif ($category === 'kpi2') {
            // ค้นหา "สาขา" + "เป้าหมาย"/"ระดับ"
            $found_branch = false;
            $found_kpi = false;
            for ($r = 1; $r <= min($highRow, 15); $r++) {
                for ($c = 1; $c <= min($highCol, 15); $c++) {
                    $v = mb_strtolower(trim((string)($sheet->getCellByColumnAndRow($c, $r)->getValue() ?? '')));
                    if ($v === '') continue;
                    if (mb_strpos($v, 'สาขา') !== false) $found_branch = true;
                    if (mb_strpos($v, 'เป้าหมาย') !== false || mb_strpos($v, 'ระดับ') !== false || mb_strpos($v, 'ผลดำเนินการ') !== false) $found_kpi = true;
                }
            }
            $spreadsheet->disconnectWorksheets();
            if (!$found_branch) {
                return ['valid' => false, 'message' => 'ไม่พบหัวคอลัมน์ "สาขา" — รูปแบบไฟล์ KPI ไม่ตรงกับที่ระบบรองรับ'];
            }
            if (!$found_kpi) {
                return ['valid' => false, 'message' => 'ไม่พบหัวคอลัมน์ "เป้าหมาย/ระดับ/ผลดำเนินการ" — รูปแบบไฟล์ KPI ไม่ตรงกับที่ระบบรองรับ'];
            }
        } elseif (in_array($category, ['ois', 'rl', 'mnf'])) {
            // OIS/RL/MNF: ตรวจว่ามีข้อมูลตัวเลขและชื่อสาขา
            $found_branch = false;
            for ($r = 1; $r <= min($highRow, 10); $r++) {
                for ($c = 1; $c <= min($highCol, 10); $c++) {
                    $v = mb_strtolower(trim((string)($sheet->getCellByColumnAndRow($c, $r)->getValue() ?? '')));
                    if (mb_strpos($v, 'สาขา') !== false || mb_strpos($v, 'หน่วยงาน') !== false) {
                        $found_branch = true;
                        break 2;
                    }
                }
            }
            $spreadsheet->disconnectWorksheets();
            if (!$found_branch) {
                return ['valid' => false, 'message' => 'ไม่พบหัวคอลัมน์ "สาขา/หน่วยงาน" — รูปแบบไฟล์ไม่ตรงกับที่ระบบรองรับ'];
            }
        } else {
            $spreadsheet->disconnectWorksheets();
        }

        return ['valid' => true, 'message' => ''];
    } catch (Exception $e) {
        return ['valid' => false, 'message' => 'ไม่สามารถอ่านไฟล์ Excel ได้: ' . $e->getMessage()];
    }
}

// Route: POST /api/upload/<category>
// ───────────────────────────────────────────────────────────────────────────
if ($method === 'POST' && count($path_parts) === 2 && $path_parts[0] === 'upload') {
    $category = $path_parts[1];

    if (!isset(CATEGORY_MAP[$category])) {
        json_response([
            'ok' => false,
            'error' => 'ไม่รู้จัก category: ' . $category
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
            if (preg_match('/\.xlsx?$/i', $filename)) {
                $validation = validate_leak_file($files['tmp_name'][$i], $category);
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

            $new_name = null;

            if ($category === 'p3') {
                // P3: extract branch name + YY-MM → P3_สาขา_YY-MM.xlsx
                $branch_name = null;
                $date_code = null;

                // Extract YY-MM pattern
                if (preg_match('/(\d{2}-\d{2})/', $name_only, $m)) {
                    $date_code = $m[1];
                }

                // Extract branch: try known branches or Thai text between underscores
                $parts = preg_split('/[_\-]/', $name_only);
                foreach ($parts as $p) {
                    $p = trim($p);
                    if ($p && !preg_match('/^(P3|p3|\d+)$/', $p) && !preg_match('/^\d{2}-\d{2}$/', $p)) {
                        $branch_name = $p;
                        break;
                    }
                }

                if ($branch_name && $date_code) {
                    $new_name = $prefix . '_' . $branch_name . '_' . $date_code . $ext;
                } elseif ($branch_name) {
                    $new_name = $prefix . '_' . $branch_name . $ext;
                } elseif ($date_code) {
                    $new_name = $prefix . '_' . $date_code . $ext;
                } else {
                    $clean = preg_replace('/[^\w\-.]/', '_', $name_only);
                    $clean = trim($clean, '_');
                    if (strlen($clean) > 30) {
                        $clean = substr($clean, 0, 30);
                    }
                    $new_name = $prefix . '_' . $clean . $ext;
                }
            } else {
                // Other categories: extract numbers/dates
                // Try YYMMDD format
                if (preg_match('/(\d{6})/', $name_only, $m)) {
                    $new_name = $prefix . '_' . $m[1] . $ext;
                } elseif (preg_match('/(\d{4})/', $name_only, $m)) {
                    // Try 4-digit year like 2569
                    $new_name = $prefix . '_' . $m[1] . $ext;
                } elseif (preg_match('/(\d{3,4})/', $name_only, $m)) {
                    // Try 3-4 digits
                    $new_name = $prefix . '_' . $m[1] . $ext;
                } else {
                    // Fallback: clean filename
                    $clean = preg_replace('/[^\w\-.]/', '_', $name_only);
                    $clean = trim($clean, '_');
                    if (strlen($clean) > 30) {
                        $clean = substr($clean, 0, 30);
                    }
                    $new_name = $prefix . '_' . $clean . $ext;
                }
            }

            $dest_path = $folder_path . DIRECTORY_SEPARATOR . $new_name;

            // Check if overwriting
            $overwrite = file_exists($dest_path);

            // Move uploaded file
            if (!move_uploaded_file($files['tmp_name'][$i], $dest_path)) {
                throw new Exception('Failed to move uploaded file');
            }
            chmod($dest_path, 0644);

            $msg = $filename . ' → ' . $new_name;
            if ($overwrite) {
                $msg .= ' (เขียนทับ)';
            }

            $results[] = [
                'filename' => $new_name,
                'original' => $filename,
                'status' => 'success',
                'message' => $msg
            ];
        } catch (Exception $e) {
            $errors[] = [
                'filename' => $filename,
                'error' => $e->getMessage()
            ];
        }
    }

    json_response([
        'ok' => true,
        'category' => $category,
        'thai_name' => CATEGORY_MAP[$category],
        'results' => $results,
        'errors' => $errors
    ]);
}

// ───────────────────────────────────────────────────────────────────────────
// Route: DELETE /api/data/<category>/<filename>
// ───────────────────────────────────────────────────────────────────────────
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

// ───────────────────────────────────────────────────────────────────────────
// Route: POST /api/notes/<slug>
// ───────────────────────────────────────────────────────────────────────────
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

// ───────────────────────────────────────────────────────────────────────────
// Route: POST /api/open-folder (not applicable in XAMPP, return path)
// ───────────────────────────────────────────────────────────────────────────
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

// ───────────────────────────────────────────────────────────────────────────
// Route: POST /api/open-main (not applicable in XAMPP, return path)
// ───────────────────────────────────────────────────────────────────────────
if ($method === 'POST' && count($path_parts) === 1 && $path_parts[0] === 'open-main') {
    $parent_dir = dirname(BASE_DIR);
    json_response([
        'ok' => true,
        'path' => $parent_dir,
        'note' => 'Parent directory path returned; OS-specific opening not available in PHP'
    ]);
}

// ───────────────────────────────────────────────────────────────────────────
// Route: POST /api/rebuild (not applicable in XAMPP)
// ───────────────────────────────────────────────────────────────────────────
if ($method === 'POST' && count($path_parts) === 1 && $path_parts[0] === 'rebuild') {
    json_response([
        'ok' => false,
        'error' => 'rebuild not available in XAMPP; run build_dashboard.py manually'
    ], 400);
}

// ═══════════════════════════════════════════════════════════════════════════
// DATA PARSING ENDPOINTS (Dual Mode: XAMPP = live API, GitHub Pages = fallback)
// ═══════════════════════════════════════════════════════════════════════════

// ───────────────────────────────────────────────────────────────────────────
// Route: GET /api/eu-data
// Parse EU (หน่วยไฟฟ้า/น้ำจำหน่าย) from Excel files
// Returns: {ok, has_data, data: {year: {branch: [12 monthly values]}}}
// ───────────────────────────────────────────────────────────────────────────
if ($method === 'GET' && count($path_parts) === 1 && $path_parts[0] === 'eu-data') {
    if (!$phpSpreadsheet) {
        json_response(['ok' => false, 'error' => 'PhpSpreadsheet not available'], 500);
    }

    $eu_folder = RAW_DATA_DIR . DIRECTORY_SEPARATOR . 'หน่วยไฟ';

    // Check cache
    $cached = load_cache('eu_data', $eu_folder);
    if ($cached !== null) {
        json_response($cached);
    }

    $result = [];

    if (is_dir($eu_folder)) {
        foreach (scandir($eu_folder) as $fname) {
            if ($fname[0] === '.') continue;
            // Match EU_YYYY.xlsx or EU-YYYY.xlsx
            if (!preg_match('/EU[-_](\d{4})\.xlsx?$/i', $fname, $m)) continue;
            $year_str = $m[1];

            $fpath = $eu_folder . DIRECTORY_SEPARATOR . $fname;
            try {
                $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($fpath);
                $sheet = $spreadsheet->getSheet(0); // First sheet (กราฟ)
                $highRow = $sheet->getHighestDataRow();

                $year_data = [];

                // --- Smart Header Detection for EU ---
                $highCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($sheet->getHighestDataColumn());
                $eu_branch_col = 2;  // default Col B
                $eu_month_start = 3; // default Col C
                $eu_data_start = 3;  // default row 3
                $eu_kw_branch = ['สาขา', 'หน่วยงาน', 'ภาพรวม', 'ชื่อสาขา'];
                $eu_kw_month  = ['ต.ค.', 'ต.ค', 'พ.ย.', 'ตุลาคม', 'oct'];

                for ($sr = 1; $sr <= min($highRow, 10); $sr++) {
                    $found_branch = false;
                    for ($sc = 1; $sc <= min($highCol, 20); $sc++) {
                        $hv = mb_strtolower(trim((string)(cellVal($sheet, $sc, $sr) ?? '')));
                        if ($hv === '') continue;
                        foreach ($eu_kw_branch as $kw) {
                            if (mb_strpos($hv, mb_strtolower($kw)) !== false) {
                                $eu_branch_col = $sc;
                                $found_branch = true;
                                break;
                            }
                        }
                        foreach ($eu_kw_month as $kw) {
                            if (mb_strpos($hv, mb_strtolower($kw)) !== false) {
                                $eu_month_start = $sc;
                                break;
                            }
                        }
                    }
                    if ($found_branch) {
                        $eu_data_start = $sr + 1;
                        break;
                    }
                }

                for ($row = $eu_data_start; $row <= $highRow; $row++) {
                    $raw_name = cellVal($sheet, $eu_branch_col, $row);
                    if (!$raw_name || !is_string($raw_name) || !trim($raw_name)) continue;

                    $name = trim($raw_name);
                    $is_regional = (mb_strpos($name, 'ภาพรวม') !== false);

                    if ($is_regional) {
                        $branch_key = '__regional__';
                    } else {
                        $branch_key = normalize_branch_name($name);
                        if (!$branch_key) continue;
                    }

                    $monthly = array_fill(0, 12, null);
                    for ($mi = 0; $mi < 12; $mi++) {
                        $col = $eu_month_start + $mi;
                        $val = cellCalc($sheet, $col, $row);
                        if (is_numeric($val) && !is_bool($val)) {
                            $monthly[$mi] = round((float)$val, 4);
                        }
                    }
                    $year_data[$branch_key] = $monthly;
                }

                if (!empty($year_data)) {
                    $result[$year_str] = $year_data;
                }
                $spreadsheet->disconnectWorksheets();
                unset($spreadsheet);
            } catch (Exception $e) {
                error_log("EU parse error ($fname): " . $e->getMessage());
            }
        }
    }

    $response = [
        'ok' => true,
        'has_data' => !empty($result),
        'data' => $result
    ];
    save_cache('eu_data', $response);
    json_response($response);
}

// ───────────────────────────────────────────────────────────────────────────
// Route: GET /api/rl-data
// Parse RL (Real Leak) from Excel files
// Returns: {ok, has_data, data: {year: {branch: {r:[12],v:[12],p:[12],s:[12],d:[12],b:[12]}}}}
// ───────────────────────────────────────────────────────────────────────────
if ($method === 'GET' && count($path_parts) === 1 && $path_parts[0] === 'rl-data') {
    if (!$phpSpreadsheet) {
        json_response(['ok' => false, 'error' => 'PhpSpreadsheet not available'], 500);
    }

    $rl_folder = RAW_DATA_DIR . DIRECTORY_SEPARATOR . 'Real Leak';

    // Check cache
    $cached = load_cache('rl_data', $rl_folder);
    if ($cached !== null) {
        json_response($cached);
    }

    // Month abbreviations → fiscal month index (0=ต.ค., 11=ก.ย.)
    $MONTH_ABBR = [
        'ต.ค.' => 0, 'พ.ย.' => 1, 'ธ.ค.' => 2, 'ม.ค.' => 3, 'ก.พ.' => 4, 'มี.ค.' => 5,
        'เม.ย.' => 6, 'พ.ค.' => 7, 'มิ.ย.' => 8, 'ก.ค.' => 9, 'ส.ค.' => 10, 'ก.ย.' => 11
    ];

    // Metric column keywords
    $METRIC_KEYWORDS = [
        'production' => 'น้ำผลิตรวม',
        'supplied'   => 'น้ำผลิตจ่ายสุทธิ',
        'sold'       => 'น้ำจำหน่าย',
        'blowoff'    => 'blow',
        'volume'     => 'ปริมาณ',
        'rate'       => 'อัตรา'
    ];

    $result = [];

    if (is_dir($rl_folder)) {
        foreach (scandir($rl_folder) as $fname) {
            if ($fname[0] === '.') continue;
            if (!preg_match('/RL[-_](\d{4})\.xlsx?$/i', $fname, $m)) continue;
            $file_year = $m[1];

            $fpath = $rl_folder . DIRECTORY_SEPARATOR . $fname;
            try {
                $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($fpath);

                foreach ($spreadsheet->getSheetNames() as $sheetName) {
                    // Parse month and fiscal year from sheet name
                    $mi = null;
                    foreach ($MONTH_ABBR as $abbr => $idx) {
                        if (mb_strpos($sheetName, $abbr) !== false) {
                            $mi = $idx;
                            break;
                        }
                    }
                    if ($mi === null) continue;

                    // Extract 2-digit calendar year from sheet name
                    $fy_str = $file_year;
                    if (preg_match('/(\d{2})\s*$/', trim($sheetName), $ym)) {
                        $cal_year = 2500 + intval($ym[1]);
                        $fy_str = ($mi <= 2) ? strval($cal_year + 1) : strval($cal_year);
                    }

                    $sheet = $spreadsheet->getSheetByName($sheetName);
                    $highRow = $sheet->getHighestDataRow();
                    $highCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($sheet->getHighestDataColumn());

                    // Find header row (contains "สาขา")
                    $headerRow = null;
                    $branchCol = null;
                    for ($r = 1; $r <= min($highRow, 15); $r++) {
                        for ($c = 1; $c <= min($highCol, 5); $c++) {
                            $v = cellVal($sheet, $c, $r);
                            if (is_string($v) && mb_strpos($v, 'สาขา') !== false) {
                                $headerRow = $r;
                                $branchCol = $c;
                                break 2;
                            }
                        }
                    }
                    if ($headerRow === null) continue;

                    // Map metric columns by scanning header rows
                    $metricCols = [];
                    $lossHeaderCol = null;
                    for ($r = $headerRow; $r <= min($headerRow + 2, $highRow); $r++) {
                        for ($c = $branchCol + 1; $c <= $highCol; $c++) {
                            $v = cellVal($sheet, $c, $r);
                            if (!is_string($v)) continue;
                            $v = trim($v);
                            $vl = mb_strtolower($v);

                            if (mb_strpos($v, 'น้ำผลิตรวม') !== false && !isset($metricCols['production']))
                                $metricCols['production'] = $c;
                            if (mb_strpos($v, 'น้ำผลิตจ่ายสุทธิ') !== false && !isset($metricCols['supplied']))
                                $metricCols['supplied'] = $c;
                            if (mb_strpos($v, 'น้ำจำหน่าย') !== false && !isset($metricCols['sold']))
                                $metricCols['sold'] = $c;
                            if (mb_strpos($vl, 'blow') !== false && !isset($metricCols['blowoff']))
                                $metricCols['blowoff'] = $c;
                            if (mb_strpos($v, 'น้ำสูญเสีย') !== false)
                                $lossHeaderCol = $c;
                            if ($lossHeaderCol !== null) {
                                if (mb_strpos($v, 'ปริมาณ') !== false && !isset($metricCols['volume']))
                                    $metricCols['volume'] = $c;
                                if (mb_strpos($v, 'อัตรา') !== false && !isset($metricCols['rate']))
                                    $metricCols['rate'] = $c;
                            }
                        }
                    }

                    // Read branch data
                    $dataStartRow = $headerRow + 2;
                    for ($r = $dataStartRow; $r <= $highRow; $r++) {
                        $raw_name = cellVal($sheet, $branchCol, $r);
                        if (!is_string($raw_name)) continue;
                        $branch = normalize_branch_name($raw_name);
                        if (!$branch) continue;

                        if (!isset($result[$fy_str])) $result[$fy_str] = [];
                        if (!isset($result[$fy_str][$branch])) {
                            $result[$fy_str][$branch] = [
                                'r' => array_fill(0, 12, null),
                                'v' => array_fill(0, 12, null),
                                'p' => array_fill(0, 12, null),
                                's' => array_fill(0, 12, null),
                                'd' => array_fill(0, 12, null),
                                'b' => array_fill(0, 12, null)
                            ];
                        }

                        $keyMap = [
                            'rate' => 'r', 'volume' => 'v', 'production' => 'p',
                            'supplied' => 's', 'sold' => 'd', 'blowoff' => 'b'
                        ];
                        foreach ($metricCols as $metric => $col) {
                            $val = cellCalc($sheet, $col, $r);
                            if (is_numeric($val) && !is_bool($val)) {
                                $k = $keyMap[$metric] ?? null;
                                if ($k) $result[$fy_str][$branch][$k][$mi] = round((float)$val, 4);
                            }
                        }
                    }
                }

                $spreadsheet->disconnectWorksheets();
                unset($spreadsheet);
            } catch (Exception $e) {
                error_log("RL parse error ($fname): " . $e->getMessage());
            }
        }
    }

    $response = [
        'ok' => true,
        'has_data' => !empty($result),
        'data' => $result
    ];
    save_cache('rl_data', $response);
    json_response($response);
}

// ───────────────────────────────────────────────────────────────────────────
// Route: GET /api/mnf-data
// Parse MNF (Minimum Night Flow) from Excel files
// Returns: {ok, has_data, data: {year: {branch: {a:[12],c:[12],t:[12],p:[12]}}}}
// ───────────────────────────────────────────────────────────────────────────
if ($method === 'GET' && count($path_parts) === 1 && $path_parts[0] === 'mnf-data') {
    if (!$phpSpreadsheet) {
        json_response(['ok' => false, 'error' => 'PhpSpreadsheet not available'], 500);
    }

    $mnf_folder = RAW_DATA_DIR . DIRECTORY_SEPARATOR . 'MNF';
    $cached = load_cache('mnf_data', $mnf_folder);
    if ($cached !== null) { json_response($cached); }

    // MNF sheet name → standard branch name
    $MNF_SHEET_MAP = [
        'ภาพรวมเขต' => '__regional__',
        '1.ชลบุรี' => 'ชลบุรี(พ)', '2.พัทยา' => 'พัทยา(พ)', '3.บ้านบึง' => 'บ้านบึง',
        '4.พนัสนิคม' => 'พนัสนิคม', '5.ศรีราชา' => 'ศรีราชา', '6.แหลมฉบัง' => 'แหลมฉบัง',
        '7.บางปะกง' => 'บางปะกง', '8.ฉะเชิงเทรา' => 'ฉะเชิงเทรา', '9.บางคล้า' => 'บางคล้า',
        '10.พนมสารคาม' => 'พนมสารคาม', '11.ระยอง' => 'ระยอง', '12.บ้านฉาง' => 'บ้านฉาง',
        '13.ปากน้ำประแสร์' => 'ปากน้ำประแสร์', '14.จันทบุรี' => 'จันทบุรี', '15.ขลุง' => 'ขลุง',
        '16.ตราด' => 'ตราด', '17.คลองใหญ่' => 'คลองใหญ่', '18.สระแก้ว' => 'สระแก้ว',
        '19.วัฒนานคร' => 'วัฒนานคร', '20.อรัญประเทศ' => 'อรัญประเทศ',
        '21.ปราจีนบุรี' => 'ปราจีนบุรี', '22.กบินทร์บุรี' => 'กบินทร์บุรี',
    ];

    // MNF row label → metric key
    $MNF_ROW_KEYWORDS = [
        'MNF เกิดจริง' => 'a',
        'MNF ที่ยอมรับได้' => 'c',
        'เป้าหมาย MNF' => 't',
        'น้ำผลิตจ่าย' => 'p',
    ];

    $result = [];

    if (is_dir($mnf_folder)) {
        foreach (scandir($mnf_folder) as $fname) {
            if ($fname[0] === '.') continue;
            if (!preg_match('/MNF[-_](\d{4})\.xlsx?$/i', $fname, $m)) continue;
            $year_str = $m[1];

            $fpath = $mnf_folder . DIRECTORY_SEPARATOR . $fname;
            try {
                $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($fpath);

                foreach ($spreadsheet->getSheetNames() as $sheetName) {
                    $branch_key = null;
                    $data_start = 4; // branch sheets: data starts row 4
                    foreach ($MNF_SHEET_MAP as $sn => $bk) {
                        if ($sheetName === $sn) { $branch_key = $bk; break; }
                    }
                    if ($branch_key === null) continue;
                    if ($branch_key === '__regional__') $data_start = 3; // regional: row 3

                    $sheet = $spreadsheet->getSheetByName($sheetName);
                    $highRow = $sheet->getHighestDataRow();

                    $metrics = ['a' => array_fill(0,12,null), 'c' => array_fill(0,12,null),
                                't' => array_fill(0,12,null), 'p' => array_fill(0,12,null)];

                    for ($r = $data_start; $r <= $highRow; $r++) {
                        $label = cellVal($sheet, 1, $r); // Col A
                        if (!is_string($label)) $label = ($label !== null) ? strval($label) : '';
                        $label = trim($label);

                        $metric_key = null;
                        foreach ($MNF_ROW_KEYWORDS as $kw => $mk) {
                            if (mb_strpos($label, $kw) !== false) { $metric_key = $mk; break; }
                        }
                        if (!$metric_key) continue;

                        for ($mi = 0; $mi < 12; $mi++) {
                            $col = 2 + $mi; // Col B(2)=ต.ค. ... Col M(13)=ก.ย.
                            $val = cellCalc($sheet, $col, $r);
                            if (is_numeric($val) && !is_bool($val)) {
                                $fv = round((float)$val, 4);
                                // MNF actual=0 means unfilled → null
                                if ($metric_key === 'a' && $fv == 0) $fv = null;
                                $metrics[$metric_key][$mi] = $fv;
                            }
                        }
                    }

                    if (!isset($result[$year_str])) $result[$year_str] = [];
                    $result[$year_str][$branch_key] = $metrics;
                }

                $spreadsheet->disconnectWorksheets();
                unset($spreadsheet);
            } catch (Exception $e) {
                error_log("MNF parse error ($fname): " . $e->getMessage());
            }
        }
    }

    $response = ['ok' => true, 'has_data' => !empty($result), 'data' => $result];
    save_cache('mnf_data', $response);
    json_response($response);
}

// ───────────────────────────────────────────────────────────────────────────
// Route: GET /api/kpi-data
// Parse KPI (เกณฑ์วัดน้ำสูญเสีย) from Excel files
// Returns: {ok, has_data, data: {year: {branch: {t:float, l:[5], a:float}}}}
// ───────────────────────────────────────────────────────────────────────────
if ($method === 'GET' && count($path_parts) === 1 && $path_parts[0] === 'kpi-data') {
    if (!$phpSpreadsheet) {
        json_response(['ok' => false, 'error' => 'PhpSpreadsheet not available'], 500);
    }

    $kpi_folder = RAW_DATA_DIR . DIRECTORY_SEPARATOR . 'เกณฑ์วัดน้ำสูญเสีย';
    $cached = load_cache('kpi_data', $kpi_folder);
    if ($cached !== null) { json_response($cached); }

    // KPI branch name normalization
    $KPI_BRANCH_MAP = [
        'ชลบุรี' => 'ชลบุรี(พ)', 'พัทยา' => 'พัทยา(พ)', 'บ้านบึง' => 'บ้านบึง',
        'พนัสนิคม' => 'พนัสนิคม', 'ศรีราชา' => 'ศรีราชา', 'แหลมฉบัง' => 'แหลมฉบัง',
        'ฉะเชิงเทรา' => 'ฉะเชิงเทรา', 'บางปะกง' => 'บางปะกง', 'บางคล้า' => 'บางคล้า',
        'พนมสารคาม' => 'พนมสารคาม', 'ระยอง' => 'ระยอง', 'บ้านฉาง' => 'บ้านฉาง',
        'ปากน้ำประแสร์' => 'ปากน้ำประแสร์', 'จันทบุรี' => 'จันทบุรี', 'ขลุง' => 'ขลุง',
        'ตราด' => 'ตราด', 'คลองใหญ่' => 'คลองใหญ่', 'สระแก้ว' => 'สระแก้ว',
        'วัฒนานคร' => 'วัฒนานคร', 'อรัญประเทศ' => 'อรัญประเทศ',
        'ปราจีนบุรี' => 'ปราจีนบุรี', 'กบินทร์บุรี' => 'กบินทร์บุรี',
    ];

    function normalize_kpi_branch($name, $map) {
        $name = trim($name);
        if (isset($map[$name])) return $map[$name];
        if (mb_strpos($name, 'รวม') !== false) return '__regional__';
        foreach ($map as $kn => $sn) {
            if (mb_strpos($name, $kn) !== false || mb_strpos($kn, $name) !== false) return $sn;
        }
        return $name;
    }

    function to_float($val) {
        if ($val === null) return null;
        if (is_numeric($val) && !is_bool($val)) return (float)$val;
        $s = str_replace(',', '', trim(strval($val)));
        return is_numeric($s) ? (float)$s : null;
    }

    $result = [];

    if (is_dir($kpi_folder)) {
        foreach (scandir($kpi_folder) as $fname) {
            if ($fname[0] === '.') continue;
            if (!preg_match('/(\d{4})/', $fname, $m)) continue;
            if (!preg_match('/\.xlsx?$/i', $fname)) continue;
            $year_str = $m[1];

            $fpath = $kpi_folder . DIRECTORY_SEPARATOR . $fname;
            try {
                $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($fpath);
                $sheet = $spreadsheet->getSheet(0);
                $highRow = $sheet->getHighestDataRow();
                $highCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($sheet->getHighestDataColumn());

                // --- Smart Header Detection for KPI ---
                $headerRow = null;
                $kpi_branch_col = 2;  // default Col B
                $kpi_target_col = 3;  // default Col C
                $kpi_l1_col = 4;      // default Col D
                $kpi_actual_col = 9;  // default Col I

                $kw_branch  = ['สาขา', 'หน่วยงาน', 'ชื่อสาขา'];
                $kw_target  = ['เป้าหมาย', 'target', 'เป้า'];
                $kw_level   = ['ระดับ', 'level', 'ระดับ 1', 'ระดับ1'];
                $kw_actual  = ['ผลดำเนินการ', 'ผลการดำเนินงาน', 'actual', 'ผล'];

                for ($r = 1; $r <= min($highRow, 15); $r++) {
                    for ($c = 1; $c <= min($highCol, 20); $c++) {
                        $v = trim((string)(cellVal($sheet, $c, $r) ?? ''));
                        if ($v === '') continue;
                        $lv = mb_strtolower($v);

                        foreach ($kw_branch as $kw) {
                            if (mb_strpos($lv, mb_strtolower($kw)) !== false && mb_strpos($lv, 'รวม') === false) {
                                $kpi_branch_col = $c;
                                $headerRow = $r;
                            }
                        }
                        foreach ($kw_target as $kw) {
                            if (mb_strpos($lv, mb_strtolower($kw)) !== false) {
                                $kpi_target_col = $c;
                            }
                        }
                        foreach ($kw_level as $kw) {
                            if (mb_strpos($lv, mb_strtolower($kw)) !== false) {
                                $kpi_l1_col = $c; // first level column
                            }
                        }
                        foreach ($kw_actual as $kw) {
                            if (mb_strpos($lv, mb_strtolower($kw)) !== false) {
                                $kpi_actual_col = $c;
                            }
                        }
                    }
                    if ($headerRow) break;
                }
                if (!$headerRow) { $spreadsheet->disconnectWorksheets(); continue; }

                $year_data = [];
                $dataStart = $headerRow + 2;
                for ($r = $dataStart; $r <= $highRow; $r++) {
                    $branch_raw = cellVal($sheet, $kpi_branch_col, $r);
                    if (!$branch_raw) {
                        // Check column before branch for "รวม"
                        $prev_col = max(1, $kpi_branch_col - 1);
                        $c0 = cellVal($sheet, $prev_col, $r);
                        if (is_string($c0) && mb_strpos($c0, 'รวม') !== false) $branch_raw = $c0;
                        else continue;
                    }
                    $branch = normalize_kpi_branch(strval($branch_raw), $KPI_BRANCH_MAP);

                    $target = to_float(cellCalc($sheet, $kpi_target_col, $r));
                    $l1 = to_float(cellCalc($sheet, $kpi_l1_col, $r));
                    $l2 = to_float(cellCalc($sheet, $kpi_l1_col + 1, $r));
                    $l3 = to_float(cellCalc($sheet, $kpi_l1_col + 2, $r));
                    $l4 = to_float(cellCalc($sheet, $kpi_l1_col + 3, $r));
                    $l5 = to_float(cellCalc($sheet, $kpi_l1_col + 4, $r));
                    $actual = to_float(cellCalc($sheet, $kpi_actual_col, $r));

                    if ($target === null && $l1 === null) continue;
                    $year_data[$branch] = ['t' => $target, 'l' => [$l1,$l2,$l3,$l4,$l5], 'a' => $actual];
                }

                if (!empty($year_data)) $result[$year_str] = $year_data;
                $spreadsheet->disconnectWorksheets();
                unset($spreadsheet);
            } catch (Exception $e) {
                error_log("KPI parse error ($fname): " . $e->getMessage());
            }
        }
    }

    $response = ['ok' => true, 'has_data' => !empty($result), 'data' => $result];
    save_cache('kpi_data', $response);
    json_response($response);
}

// ───────────────────────────────────────────────────────────────────────────
// Route: GET /api/p3-data
// Parse P3 (Pressure) from Excel files
// Returns: {ok, has_data, data: {year: {month_key: {branch: [{n,p,a,h}]}}}}
// ───────────────────────────────────────────────────────────────────────────
if ($method === 'GET' && count($path_parts) === 1 && $path_parts[0] === 'p3-data') {
    if (!$phpSpreadsheet) {
        json_response(['ok' => false, 'error' => 'PhpSpreadsheet not available'], 500);
    }

    $p3_folder = RAW_DATA_DIR . DIRECTORY_SEPARATOR . 'P3';
    $cached = load_cache('p3_data', $p3_folder);
    if ($cached !== null) { json_response($cached); }

    function clean_p3_name($name) {
        if (!is_string($name)) return $name;
        return trim(str_replace(['├','└','│','─'], '', $name));
    }

    function p3_val($v) {
        if ($v === null || $v === '' || $v === '-') return null;
        return is_numeric($v) ? round((float)$v, 4) : null;
    }

    function parse_p3_xlsx($fpath) {
        $points = [];
        try {
            $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($fpath);
            $sheet = $spreadsheet->getSheet(0);
            $highRow = $sheet->getHighestDataRow();

            // Find header row with "พื้นที่"
            $headerRow = null;
            for ($r = 1; $r <= min($highRow, 10); $r++) {
                $v = cellVal($sheet, 1, $r); // Col A
                if (is_string($v) && mb_strpos($v, 'พื้นที่') !== false) {
                    $headerRow = $r; break;
                }
            }
            if (!$headerRow) { $spreadsheet->disconnectWorksheets(); return $points; }

            for ($r = $headerRow + 1; $r <= $highRow; $r++) {
                $name = cellVal($sheet, 1, $r); // Col A
                if (!is_string($name) || mb_strpos($name, 'P3') === false) continue;
                $name = clean_p3_name($name);

                $avg_prev = p3_val(cellCalc($sheet, 2, $r)); // Col B
                $avg_day = p3_val(cellCalc($sheet, 3, $r));  // Col C

                $hourly = [];
                for ($col = 4; $col <= 27; $col++) { // Cols D-AA = 24 hours
                    $hourly[] = p3_val(cellCalc($sheet, $col, $r));
                }

                $points[] = ['n' => $name, 'p' => $avg_prev, 'a' => $avg_day, 'h' => $hourly];
            }
            $spreadsheet->disconnectWorksheets();
        } catch (Exception $e) {
            error_log("P3 parse error: " . $e->getMessage());
        }
        return $points;
    }

    $result = [];

    if (is_dir($p3_folder)) {
        // Scan files directly in P3/ folder (flat: P3_branch_YY-MM.xlsx)
        foreach (scandir($p3_folder) as $fname) {
            if ($fname[0] === '.' || $fname[0] === '~') continue;
            if (!preg_match('/^P3_(.+?)_((\d{2})-(\d{2}))\.xlsx$/i', $fname, $m)) continue;

            $branch = $m[1];
            $month_key = $m[2]; // "69-03"
            $yy = intval($m[3]);
            $year_str = strval(2500 + $yy);

            $fpath = $p3_folder . DIRECTORY_SEPARATOR . $fname;
            $points = parse_p3_xlsx($fpath);
            if (!empty($points)) {
                if (!isset($result[$year_str])) $result[$year_str] = [];
                if (!isset($result[$year_str][$month_key])) $result[$year_str][$month_key] = [];
                $result[$year_str][$month_key][$branch] = $points;
            }
        }

        // Also scan year subfolders
        foreach (scandir($p3_folder) as $subdir) {
            $subdirPath = $p3_folder . DIRECTORY_SEPARATOR . $subdir;
            if ($subdir[0] === '.' || !is_dir($subdirPath)) continue;
            foreach (scandir($subdirPath) as $fname) {
                if ($fname[0] === '.' || $fname[0] === '~') continue;
                if (!preg_match('/^(.+?)_((\d{2})-(\d{2}))\.xlsx$/i', $fname, $m)) continue;
                $branch = $m[1]; $month_key = $m[2]; $yy = intval($m[3]);
                $year_str = $subdir; // Use subfolder name as year key
                $fpath = $subdirPath . DIRECTORY_SEPARATOR . $fname;
                $points = parse_p3_xlsx($fpath);
                if (!empty($points)) {
                    if (!isset($result[$year_str])) $result[$year_str] = [];
                    if (!isset($result[$year_str][$month_key])) $result[$year_str][$month_key] = [];
                    $result[$year_str][$month_key][$branch] = $points;
                }
            }
        }
    }

    $response = ['ok' => true, 'has_data' => !empty($result), 'data' => $result];
    save_cache('p3_data', $response);
    json_response($response);
}

// ─── Route: GET /api/ois-data ─────────────────────────────────────────────
// D (OIS) data - the main dashboard dataset
// Parses OIS_YYYY.xls/.xlsx files from ข้อมูลดิบ/OIS/
// Output: {year: {sheet_name: [{l, u, m:[12], t, ty, tm}]}}

if ($method === 'GET' && count($path_parts) === 1 && $path_parts[0] === 'ois-data') {
    if (!$phpSpreadsheet) {
        json_response(['ok' => false, 'error' => 'PhpSpreadsheet not available'], 500);
    }

    $ois_dir = RAW_DATA_DIR . DIRECTORY_SEPARATOR . CATEGORY_MAP['ois'];
    if (!is_dir($ois_dir)) {
        json_response(['ok' => true, 'has_data' => false, 'data' => new stdClass()]);
    }

    // Check cache
    $cached = load_cache('ois_data', $ois_dir);
    if ($cached !== null) {
        json_response($cached);
    }

    // Month keywords for header detection
    $MONTH_KEYWORDS = ['ต.ค.', 'พ.ย.', 'ธ.ค.', 'ม.ค.', 'ก.พ.', 'มี.ค.',
        'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.',
        'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม', 'มกราคม', 'กุมภาพันธ์',
        'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม',
        'สิงหาคม', 'กันยายน'];

    // Short and long month names for column matching (fiscal order: ต.ค. to ก.ย.)
    $MONTH_SHORT = ['ต.ค.', 'พ.ย.', 'ธ.ค.', 'ม.ค.', 'ก.พ.', 'มี.ค.',
        'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.'];
    $MONTH_LONG = ['ตุลาคม', 'พฤศจิกายน', 'ธันวาคม', 'มกราคม', 'กุมภาพันธ์',
        'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม',
        'สิงหาคม', 'กันยายน'];

    // Label normalization map (some years use different names)
    $LABEL_NORMALIZE_MAP = [
        '2.5 อัตราการสูญเสีย (ต่อน้ำผลิตจ่าย)' => '2.5 อัตราน้ำสูญเสีย (ต่อน้ำผลิตจ่าย)',
        '2.2  ปริมาณน้ำจ่ายฟรี + Blowoff' => '2.2  ปริมาณน้ำจ่ายฟรี',
        '4.2 เงินเดือนและค่าจ้างประจำ' => '4.1 เงินเดือนและค่าจ้างประจำ',
        '4.3 ค่าจ้างชั่วคราว' => '4.2 ค่าจ้างชั่วคราว',
        '4.5 วัสดุการผลิต' => '4.4 วัสดุการผลิต',
    ];

    // Skip these sheet names
    $SKIP_SHEETS = ['เป้าหมาย'];

    // Helper: find month header row (row with 6+ month keywords)
    function ois_find_header_row($sheet, $highestRow, $highestCol) {
        global $MONTH_KEYWORDS;
        $maxCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestCol);
        for ($row = 1; $row <= min($highestRow, 20); $row++) {
            $rowText = '';
            for ($col = 1; $col <= $maxCol; $col++) {
                $val = cellVal($sheet, $col, $row);
                if (is_string($val)) $rowText .= ' ' . $val;
            }
            $count = 0;
            foreach ($MONTH_KEYWORDS as $kw) {
                if (mb_strpos($rowText, $kw) !== false) $count++;
            }
            if ($count >= 6) return $row;
        }
        return null;
    }

    // Helper: find month columns (map 12 fiscal months to column indices)
    function ois_find_month_cols($sheet, $headerRow, $highestCol) {
        global $MONTH_SHORT, $MONTH_LONG;
        $maxCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestCol);
        $monthCols = array_fill(0, 12, null);
        for ($col = 1; $col <= $maxCol; $col++) {
            $val = cellVal($sheet, $col, $headerRow);
            if (!is_string($val)) continue;
            for ($mi = 0; $mi < 12; $mi++) {
                if (mb_strpos($val, $MONTH_SHORT[$mi]) !== false ||
                    mb_strpos($val, $MONTH_LONG[$mi]) !== false) {
                    $monthCols[$mi] = $col;
                    break;
                }
            }
        }
        return $monthCols;
    }

    // Helper: find total column ("รวม")
    function ois_find_total_col($sheet, $headerRow, $highestCol) {
        $maxCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestCol);
        // Check row above header first
        if ($headerRow > 1) {
            for ($col = 1; $col <= $maxCol; $col++) {
                $val = cellVal($sheet, $col, $headerRow - 1);
                if (is_string($val) && mb_strpos($val, 'รวม') !== false) return $col;
            }
        }
        // Check header row itself
        for ($col = 1; $col <= $maxCol; $col++) {
            $val = cellVal($sheet, $col, $headerRow);
            if (is_string($val) && mb_strpos($val, 'รวม') !== false) return $col;
        }
        return null;
    }

    // Helper: extract data rows from sheet
    function ois_extract_rows($sheet, $headerRow, $monthCols, $totalCol, $highestRow) {
        $rows = [];
        $dataStart = $headerRow + 1;
        for ($r = $dataStart; $r <= $highestRow; $r++) {
            // Label = col 1 (A)
            $label = cellVal($sheet, 1, $r);
            if (is_numeric($label)) $label = strval($label);
            if (!is_string($label)) $label = '';
            $label = trim($label);
            if ($label === '') continue;
            if (mb_strpos($label, 'หมายเหตุ') !== false) continue;

            // Unit = col 2 (B)
            $unit = cellVal($sheet, 2, $r);
            if (is_numeric($unit)) $unit = strval($unit);
            if (!is_string($unit)) $unit = '';
            $unit = trim($unit);

            // 12 monthly values
            $monthly = [];
            for ($mi = 0; $mi < 12; $mi++) {
                $mc = $monthCols[$mi];
                if ($mc !== null) {
                    $v = cellCalc($sheet, $mc, $r);
                    if (is_numeric($v)) {
                        $monthly[] = floatval($v);
                    } else {
                        $monthly[] = null;
                    }
                } else {
                    $monthly[] = null;
                }
            }

            // Total value
            $total = null;
            if ($totalCol !== null) {
                $tv = cellCalc($sheet, $totalCol, $r);
                if (is_numeric($tv)) $total = floatval($tv);
            }

            // Target year = col 3 (C), target month = col 5 (E)
            $targetYear = null;
            $targetMonth = null;
            $tyv = cellVal($sheet, 3, $r);
            $tmv = cellVal($sheet, 5, $r);
            if (is_numeric($tyv)) $targetYear = floatval($tyv);
            if (is_numeric($tmv)) $targetMonth = floatval($tmv);

            $rows[] = [
                'l' => $label,
                'u' => $unit,
                'm' => $monthly,
                't' => $total,
                'ty' => $targetYear,
                'tm' => $targetMonth,
            ];
        }
        return $rows;
    }

    // Helper: fix trailing zeros for incomplete fiscal years
    function ois_fix_trailing_zeros(&$allData) {
        foreach ($allData as $yearStr => &$sheets) {
            foreach ($sheets as $sheetName => &$rows) {
                if (empty($rows)) continue;
                $numRows = count($rows);
                $lastRealMonth = -1;
                for ($mi = 0; $mi < 12; $mi++) {
                    $nonZeroCount = 0;
                    foreach ($rows as &$r) {
                        if ($r['m'][$mi] !== null && $r['m'][$mi] != 0) $nonZeroCount++;
                    }
                    unset($r);
                    if ($numRows > 0 && ($nonZeroCount / max($numRows, 1)) >= 0.30) {
                        $lastRealMonth = $mi;
                    }
                }
                if ($lastRealMonth < 11) {
                    foreach ($rows as &$r) {
                        for ($mi = $lastRealMonth + 1; $mi < 12; $mi++) {
                            if ($r['m'][$mi] === 0 || $r['m'][$mi] === 0.0) {
                                $r['m'][$mi] = null;
                            }
                        }
                    }
                    unset($r);
                }
            }
            unset($rows);
        }
        unset($sheets);
    }

    // Scan for OIS files
    $files = array_merge(
        glob($ois_dir . DIRECTORY_SEPARATOR . '*.xls') ?: [],
        glob($ois_dir . DIRECTORY_SEPARATOR . '*.xlsx') ?: []
    );
    // Remove duplicates (*.xls also matches *.xlsx on some systems)
    $files = array_unique($files);
    sort($files);

    $result = [];

    foreach ($files as $file) {
        $basename = pathinfo($file, PATHINFO_FILENAME);
        // Extract year from filename e.g. OIS_2569 -> 2569
        if (!preg_match('/(\d{4})/', $basename, $m)) continue;
        $yearStr = $m[1];

        try {
            $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($file);
        } catch (Exception $e) {
            error_log("OIS: Cannot load $file: " . $e->getMessage());
            continue;
        }

        $yearData = [];
        $sheetCount = $spreadsheet->getSheetCount();

        for ($si = 0; $si < $sheetCount; $si++) {
            $sheet = $spreadsheet->getSheet($si);
            $sheetName = $sheet->getTitle();

            // Skip target sheets
            $skip = false;
            foreach ($SKIP_SHEETS as $ss) {
                if (mb_strpos($sheetName, $ss) !== false) { $skip = true; break; }
            }
            if ($skip) continue;

            $highestRow = $sheet->getHighestRow();
            $highestCol = $sheet->getHighestColumn();
            if ($highestRow < 3) continue;

            // Find month header row
            $headerRow = ois_find_header_row($sheet, $highestRow, $highestCol);
            if ($headerRow === null) continue;

            // Find month columns
            $monthCols = ois_find_month_cols($sheet, $headerRow, $highestCol);
            if (count(array_filter($monthCols, function($v) { return $v !== null; })) === 0) continue;

            // Find total column
            $totalCol = ois_find_total_col($sheet, $headerRow, $highestCol);

            // Extract data rows
            $rows = ois_extract_rows($sheet, $headerRow, $monthCols, $totalCol, $highestRow);
            if (!empty($rows)) {
                // Normalize labels
                foreach ($rows as &$row) {
                    if (isset($LABEL_NORMALIZE_MAP[$row['l']])) {
                        $row['l'] = $LABEL_NORMALIZE_MAP[$row['l']];
                    }
                }
                unset($row);
                $yearData[$sheetName] = $rows;
            }
        }

        $spreadsheet->disconnectWorksheets();
        unset($spreadsheet);

        if (!empty($yearData)) {
            $result[$yearStr] = $yearData;
        }
    }

    // Fix trailing zeros for incomplete fiscal years
    ois_fix_trailing_zeros($result);

    $response = ['ok' => true, 'has_data' => !empty($result), 'data' => $result];
    save_cache('ois_data', $response);
    json_response($response);
}

// 404 - Route not found
json_response([
    'ok' => false,
    'error' => 'Route not found: ' . $method . ' ' . $path_info
], 404);
