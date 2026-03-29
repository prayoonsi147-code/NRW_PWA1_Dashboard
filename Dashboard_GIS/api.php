<?php
/**
 * Dashboard GIS — XAMPP (Apache + PHP) Backend API
 * ========================================================================
 * PHP เทียบเท่า Flask server.py
 * รับ upload ไฟล์ Excel → จัดเก็บลงโฟลเดอร์ ข้อมูลดิบ/
 *
 * Architecture:
 *   - Single file handling ALL API routes via PATH_INFO
 *   - .htaccess rewrites /api/* to this file
 *   - Static files (index.html, manage.html) served directly by Apache
 *   - File-based Excel caching with TTL for pending operations
 *
 * Setup:
 *   1. Install via Composer: composer require phpoffice/phpspreadsheet
 *   2. Place composer vendor/ at project root (../vendor/autoload.php)
 *   3. Create .htaccess with rewrite rules
 */

// ─── Configuration ─────────────────────────────────────────────────────────

define('BASE_DIR', __DIR__);
define('RAW_DATA_DIR', BASE_DIR . DIRECTORY_SEPARATOR . 'ข้อมูลดิบ');
define('CACHE_DIR', BASE_DIR . DIRECTORY_SEPARATOR . '.cache');
define('CACHE_TTL', 60); // seconds

// Category mapping: URL slug → Thai folder name
const CATEGORY_MAP = [
    'repair'   => 'ลงข้อมูลซ่อมท่อ',
    'pressure' => 'แรงดันน้ำ',
    'pending'  => 'ซ่อมท่อค้างระบบ',
];

// Thai month names
const TH_MONTHS = [
    '', 'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
    'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
];

// Branch list
const BRANCH_LIST = [
    "ชลบุรี","พัทยา","บ้านบึง","พนัสนิคม","ศรีราชา","แหลมฉบัง",
    "ฉะเชิงเทรา","บางปะกง","บางคล้า","พนมสารคาม","ระยอง","บ้านฉาง",
    "ปากน้ำประแสร์","จันทบุรี","ขลุง","ตราด","คลองใหญ่",
    "สระแก้ว","วัฒนานคร","อรัญประเทศ","ปราจีนบุรี","กบินทร์บุรี"
];

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
if (!is_dir(CACHE_DIR)) {
    mkdir(CACHE_DIR, 0755, true);
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
 * Parse Thai date (DD/MM/YYYY พ.ศ. or datetime object)
 * Returns: [datetime_obj, buddhist_year] or [null, null]
 */
function parse_thai_date($val) {
    if ($val instanceof DateTime) {
        $by = $val->format('Y') < 2500 ? $val->format('Y') + 543 : $val->format('Y');
        return [$val, $by];
    }

    if (is_string($val) && strpos($val, '/') !== false) {
        try {
            $parts = explode('/', trim($val));
            if (count($parts) === 3) {
                $dd = (int)$parts[0];
                $mm = (int)$parts[1];
                $yyyy = (int)$parts[2];

                $by = $yyyy > 2500 ? $yyyy : $yyyy + 543;
                $ce_year = $by - 543;

                $dt = new DateTime("$ce_year-$mm-$dd");
                return [$dt, $by];
            }
        } catch (Exception $e) {
            return [null, null];
        }
    }

    return [null, null];
}

/**
 * Read Excel file into memory (cached)
 * Returns: array of rows, or null on error
 */
function read_excel_cached($fpath) {
    global $phpSpreadsheet;

    $mtime = filemtime($fpath);

    // Check Python-style cache first (same folder, .cache.json suffix)
    $py_cache = $fpath . '.cache.json';
    if (file_exists($py_cache)) {
        try {
            $cached = json_decode(file_get_contents($py_cache), true);
            if (isset($cached['rows'])) {
                return $cached['rows'];
            }
        } catch (Exception $e) {}
    }

    // Check PHP-style cache (.cache/ folder with md5 name)
    $cache_file = CACHE_DIR . DIRECTORY_SEPARATOR . md5($fpath) . '.json';
    if (file_exists($cache_file)) {
        try {
            $cached = json_decode(file_get_contents($cache_file), true);
            if (isset($cached['mtime']) && $cached['mtime'] == $mtime &&
                isset($cached['time']) && (time() - $cached['time']) < CACHE_TTL) {
                return $cached['rows'];
            }
        } catch (Exception $e) {}
    }

    // Fallback: read Excel with PhpSpreadsheet
    if (!$phpSpreadsheet || !class_exists('PhpOffice\\PhpSpreadsheet\\IOFactory')) {
        return null;
    }

    try {
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($fpath);
        $worksheet = $spreadsheet->getActiveSheet();
        $rows = [];

        foreach ($worksheet->getRowIterator() as $row) {
            $row_data = [];
            foreach ($row->getCellIterator() as $cell) {
                $row_data[] = $cell->getValue();
            }
            $rows[] = $row_data;
        }

        $spreadsheet->disconnectWorksheets();
        unset($spreadsheet);

        // Save to Python-style cache for next time
        $cache_data = ['mtime' => $mtime, 'rows' => $rows];
        file_put_contents($py_cache, json_encode($cache_data, JSON_UNESCAPED_UNICODE));

        return $rows;
    } catch (Exception $e) {
        error_log("Error reading Excel: " . $e->getMessage());
        return null;
    }
}

/**
 * Get value from row at column index (0-based)
 */
function row_get($row, $col) {
    return isset($row[$col]) ? $row[$col] : null;
}

// ─── SQLite Helpers for Pending Data ──────────────────────────────────────

/**
 * Build SQLite database from pending data rows (called after upload merge)
 * @param array $all_data_rows - data rows from Excel (0-based columns)
 * @param string $folder_path - folder to save .sqlite file
 * @param string $excel_filename - e.g. ค้างซ่อม_10-68_to_03-69.xlsx
 * @return string|false - path to .sqlite file or false on error
 */
function build_pending_sqlite($all_data_rows, $folder_path, $excel_filename) {
    $db_name = preg_replace('/\.xlsx?$/i', '.sqlite', $excel_filename);
    $db_path = $folder_path . DIRECTORY_SEPARATOR . $db_name;

    // Delete old DB if exists
    if (file_exists($db_path)) {
        unlink($db_path);
    }

    try {
        $db = new SQLite3($db_path);
        $db->exec('PRAGMA journal_mode=WAL');
        $db->exec('PRAGMA synchronous=NORMAL');

        // Create table
        $db->exec('CREATE TABLE pending_rows (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            row_num INTEGER,
            notify_no TEXT,
            date_val TEXT,
            date_ce TEXT,
            date_by INTEGER,
            month_key TEXT,
            finish_val TEXT,
            finish_ce TEXT,
            job_no TEXT,
            type_val TEXT,
            side TEXT,
            topic TEXT,
            detail TEXT,
            branch TEXT,
            team TEXT,
            tech TEXT,
            pipe TEXT,
            status TEXT
        )');

        // Create indexes for common queries
        $db->exec('CREATE INDEX idx_branch ON pending_rows(branch)');
        $db->exec('CREATE INDEX idx_month_key ON pending_rows(month_key)');
        $db->exec('CREATE INDEX idx_status ON pending_rows(status)');
        $db->exec('CREATE INDEX idx_side ON pending_rows(side)');
        $db->exec('CREATE INDEX idx_job_no ON pending_rows(job_no)');
        $db->exec('CREATE INDEX idx_date_ce ON pending_rows(date_ce)');

        // Meta table
        $db->exec('CREATE TABLE meta (key TEXT PRIMARY KEY, value TEXT)');

        // Column indices (0-based)
        $C_NOTIFY = 2; $C_DATE = 3; $C_FINISH = 5; $C_JOB = 6;
        $C_TYPE = 7; $C_SIDE = 8; $C_TOPIC = 9; $C_DETAIL = 10;
        $C_BRANCH = 19; $C_TEAM = 20; $C_TECH = 21; $C_PIPE = 25; $C_STATUS = 26;

        $stmt = $db->prepare('INSERT INTO pending_rows
            (row_num, notify_no, date_val, date_ce, date_by, month_key,
             finish_val, finish_ce, job_no, type_val, side, topic, detail,
             branch, team, tech, pipe, status)
            VALUES (:row_num, :notify_no, :date_val, :date_ce, :date_by, :month_key,
                    :finish_val, :finish_ce, :job_no, :type_val, :side, :topic, :detail,
                    :branch, :team, :tech, :pipe, :status)');

        $db->exec('BEGIN TRANSACTION');

        $last_report_dt = null;
        $row_count = 0;

        foreach ($all_data_rows as $idx => $row) {
            $date_val = isset($row[$C_DATE]) ? $row[$C_DATE] : null;
            if (!$date_val) continue;

            [$dt, $by] = parse_thai_date($date_val);
            if (!$dt || !$by) continue;

            $date_ce = $dt->format('Y-m-d');
            $yy = $by % 100;
            $mm = (int)$dt->format('m');
            $month_key = sprintf('%02d-%02d', $yy, $mm);

            if ($last_report_dt === null || $dt > $last_report_dt) {
                $last_report_dt = $dt;
            }

            // Parse finish date
            $finish_val = isset($row[$C_FINISH]) ? (string)($row[$C_FINISH] ?? '') : '';
            $finish_ce = '';
            if ($finish_val) {
                [$fdt, $_] = parse_thai_date($finish_val);
                if ($fdt) {
                    $finish_ce = $fdt->format('Y-m-d');
                } elseif (is_string($finish_val) && strlen($finish_val) >= 10) {
                    [$fdt, $_] = parse_thai_date(substr($finish_val, 0, 10));
                    if ($fdt) $finish_ce = $fdt->format('Y-m-d');
                }
            }

            $branch = trim((string)(isset($row[$C_BRANCH]) ? $row[$C_BRANCH] : ''));
            if (!$branch) continue;

            $stmt->bindValue(':row_num', $idx, SQLITE3_INTEGER);
            $stmt->bindValue(':notify_no', (string)(isset($row[$C_NOTIFY]) ? $row[$C_NOTIFY] : ''), SQLITE3_TEXT);
            $stmt->bindValue(':date_val', is_string($date_val) ? $date_val : $dt->format('d/m/Y'), SQLITE3_TEXT);
            $stmt->bindValue(':date_ce', $date_ce, SQLITE3_TEXT);
            $stmt->bindValue(':date_by', $by, SQLITE3_INTEGER);
            $stmt->bindValue(':month_key', $month_key, SQLITE3_TEXT);
            $stmt->bindValue(':finish_val', $finish_val, SQLITE3_TEXT);
            $stmt->bindValue(':finish_ce', $finish_ce, SQLITE3_TEXT);
            $stmt->bindValue(':job_no', trim((string)(isset($row[$C_JOB]) ? $row[$C_JOB] : '')), SQLITE3_TEXT);
            $stmt->bindValue(':type_val', (string)(isset($row[$C_TYPE]) ? $row[$C_TYPE] : ''), SQLITE3_TEXT);
            $stmt->bindValue(':side', (string)(isset($row[$C_SIDE]) ? $row[$C_SIDE] : ''), SQLITE3_TEXT);
            $stmt->bindValue(':topic', (string)(isset($row[$C_TOPIC]) ? $row[$C_TOPIC] : ''), SQLITE3_TEXT);
            $stmt->bindValue(':detail', (string)(isset($row[$C_DETAIL]) ? $row[$C_DETAIL] : ''), SQLITE3_TEXT);
            $stmt->bindValue(':branch', $branch, SQLITE3_TEXT);
            $stmt->bindValue(':team', (string)(isset($row[$C_TEAM]) ? $row[$C_TEAM] : ''), SQLITE3_TEXT);
            $stmt->bindValue(':tech', (string)(isset($row[$C_TECH]) ? $row[$C_TECH] : ''), SQLITE3_TEXT);
            $stmt->bindValue(':pipe', (string)(isset($row[$C_PIPE]) ? $row[$C_PIPE] : ''), SQLITE3_TEXT);
            $stmt->bindValue(':status', (string)(isset($row[$C_STATUS]) ? $row[$C_STATUS] : ''), SQLITE3_TEXT);
            $stmt->execute();
            $stmt->reset();
            $row_count++;
        }

        $db->exec('COMMIT');

        // Save meta
        $meta_stmt = $db->prepare('INSERT INTO meta (key, value) VALUES (:k, :v)');
        $meta_stmt->bindValue(':k', 'row_count', SQLITE3_TEXT);
        $meta_stmt->bindValue(':v', (string)$row_count, SQLITE3_TEXT);
        $meta_stmt->execute();

        if ($last_report_dt) {
            $by_lrd = $last_report_dt->format('Y') + 543;
            $meta_stmt->reset();
            $meta_stmt->bindValue(':k', 'update_date', SQLITE3_TEXT);
            $meta_stmt->bindValue(':v', sprintf('%02d-%02d-%02d',
                $last_report_dt->format('d'), $last_report_dt->format('m'), $by_lrd % 100), SQLITE3_TEXT);
            $meta_stmt->execute();
        }

        $meta_stmt->reset();
        $meta_stmt->bindValue(':k', 'excel_file', SQLITE3_TEXT);
        $meta_stmt->bindValue(':v', $excel_filename, SQLITE3_TEXT);
        $meta_stmt->execute();

        $db->close();
        return $db_path;

    } catch (Exception $e) {
        error_log("SQLite build error: " . $e->getMessage());
        return false;
    }
}

/**
 * Find pending SQLite/Excel files by fiscal year
 * Returns: ['fy_list' => [...], 'fy_files' => [fy => ['sqlite' => path, 'excel' => path]], 'latest_fy' => int]
 */
function find_pending_files($pending_dir) {
    $fy_files = [];

    foreach (scandir($pending_dir) as $fname) {
        // Look for Excel files to determine FY
        if (!preg_match('/\.xlsx?$/i', $fname)) continue;
        if (!preg_match('/(\d{1,2})-(\d{2})_to_(\d{1,2})-(\d{2})/', $fname, $m)) continue;

        $fpath = $pending_dir . DIRECTORY_SEPARATOR . $fname;
        $start_mm = (int)$m[1];
        $start_yy = (int)$m[2];
        $fy_be = ($start_mm >= 10) ? (2500 + $start_yy + 1) : (2500 + $start_yy);

        // Check if SQLite version exists
        $sqlite_name = preg_replace('/\.xlsx?$/i', '.sqlite', $fname);
        $sqlite_path = $pending_dir . DIRECTORY_SEPARATOR . $sqlite_name;

        $fy_files[$fy_be] = [
            'excel' => $fpath,
            'sqlite' => file_exists($sqlite_path) ? $sqlite_path : null
        ];
    }

    $fy_list = array_keys($fy_files);
    sort($fy_list);

    return [
        'fy_list' => $fy_list,
        'fy_files' => $fy_files,
        'latest_fy' => !empty($fy_list) ? end($fy_list) : null
    ];
}

/**
 * Open SQLite DB for a given FY, building from cache/Excel if needed
 * Returns: SQLite3 object or null
 */
function open_pending_db($pending_dir, $fy_files_info, $fy) {
    $info = $fy_files_info['fy_files'][$fy] ?? null;
    if (!$info) return null;

    // If SQLite exists, use it
    if ($info['sqlite'] && file_exists($info['sqlite'])) {
        try {
            $db = new SQLite3($info['sqlite'], SQLITE3_OPEN_READONLY);
            return $db;
        } catch (Exception $e) {
            error_log("SQLite open error: " . $e->getMessage());
        }
    }

    // No SQLite — build from Excel cache or Excel file
    $excel_path = $info['excel'];
    if (!$excel_path || !file_exists($excel_path)) return null;

    // Try reading via cache
    $rows = read_excel_cached($excel_path);
    if ($rows === null) return null;

    // Build SQLite from rows (skip header rows 0-7, data starts at row 8)
    $DATA_START = 8;
    $data_rows = array_slice($rows, $DATA_START);
    unset($rows); // free memory

    $excel_filename = basename($excel_path);
    $db_path = build_pending_sqlite($data_rows, $pending_dir, $excel_filename);

    if ($db_path) {
        try {
            return new SQLite3($db_path, SQLITE3_OPEN_READONLY);
        } catch (Exception $e) {
            error_log("SQLite open after build error: " . $e->getMessage());
        }
    }

    return null;
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
        'version' => '1.0',
        'timestamp' => (new DateTime('now', new DateTimeZone('Asia/Bangkok')))->format('c')
    ]);
}

// Route: GET /api/data
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
            $dt = new DateTime('@' . $last_modified, new DateTimeZone('Asia/Bangkok'));
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
        'notes' => $notes
    ]);
}

// ── Validate file format before upload ──
function validate_gis_file($tmp_path, $category) {
    if (!class_exists('PhpOffice\\PhpSpreadsheet\\IOFactory')) {
        return ['valid' => true, 'message' => '']; // ไม่มี library ให้ข้ามการตรวจสอบ
    }
    if (!preg_match('/\.xlsx?$/i', $tmp_path) && !preg_match('/\.xlsx?$/i', basename($tmp_path))) {
        // ไม่ใช่ Excel — ข้ามการตรวจสอบ
        return ['valid' => true, 'message' => ''];
    }

    try {
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($tmp_path);
        $worksheet = $spreadsheet->getSheet(0);

        if ($category === 'repair') {
            $det = detect_repair_columns($worksheet);
            $spreadsheet->disconnectWorksheets();
            if ($det['fallback']) {
                return [
                    'valid' => false,
                    'message' => 'ไม่พบหัวคอลัมน์ที่คาดหวัง (สาขา, ปิดงาน, สำเร็จ, คะแนน) — รูปแบบไฟล์ไม่ตรงกับที่ระบบรองรับ'
                ];
            }
        } elseif ($category === 'pressure') {
            // ตรวจสอบว่ามี "ปีงบประมาณ" หรือชื่อสาขาในไฟล์
            $found_fy = false;
            $found_branch = false;
            $maxR = min(10, $worksheet->getHighestRow());
            $maxC = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($worksheet->getHighestColumn());
            $maxC = min($maxC, 10);
            for ($r = 1; $r <= $maxR; $r++) {
                for ($c = 1; $c <= $maxC; $c++) {
                    $v = trim((string)($worksheet->getCellByColumnAndRow($c, $r)->getValue() ?? ''));
                    if (preg_match('/ปีงบประมาณ/', $v)) $found_fy = true;
                    if (preg_match('/สาขา|แรงดัน|pressure/i', $v)) $found_branch = true;
                }
            }
            $spreadsheet->disconnectWorksheets();
            if (!$found_fy && !$found_branch) {
                return [
                    'valid' => false,
                    'message' => 'ไม่พบข้อมูล "ปีงบประมาณ" หรือ "สาขา/แรงดัน" ในไฟล์ — รูปแบบไฟล์ไม่ตรงกับที่ระบบรองรับ'
                ];
            }
        } elseif ($category === 'pending') {
            // Pending: ตรวจสอบว่ามีข้อมูลอย่างน้อย 1 แถว
            $highRow = $worksheet->getHighestRow();
            $spreadsheet->disconnectWorksheets();
            if ($highRow < 2) {
                return [
                    'valid' => false,
                    'message' => 'ไฟล์ไม่มีข้อมูล (มีเฉพาะ header หรือว่างเปล่า)'
                ];
            }
        } else {
            $spreadsheet->disconnectWorksheets();
        }

        return ['valid' => true, 'message' => ''];
    } catch (Exception $e) {
        return [
            'valid' => false,
            'message' => 'ไม่สามารถอ่านไฟล์ Excel ได้: ' . $e->getMessage()
        ];
    }
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

    // Parse date
    [$date_key, $date_label] = parse_date_key($data_date);
    if (!$date_key) {
        json_response([
            'ok' => false,
            'error' => 'รูปแบบวันที่ไม่ถูกต้อง: ' . $data_date . ' (ใช้ YYYY-MM-DD เช่น 2569-01-16 หรือ DD/MM/YYYY เช่น 16/01/2569)'
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

    $PREFIX_MAP = ['repair' => 'GIS', 'pressure' => 'PRESSURE', 'pending' => 'PENDING'];

    $results = [];
    $errors = [];
    $pending_batch = [];  // For pending category batch merge

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
                $validation = validate_gis_file($files['tmp_name'][$i], $category);
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

            // Category-specific handling
            if ($category === 'repair') {
                // repair: GIS_YYMMDD.xlsx
                if (preg_match('/(\d{6})/', $name_only, $m)) {
                    $new_name = $prefix . '_' . $m[1] . $ext;
                } else {
                    $today = date('ymd');
                    $new_name = $prefix . '_' . $today . $ext;
                }
            } elseif ($category === 'pressure') {
                // pressure: PRESSURE_สาขา_ปีงบYY.xlsx
                // Extract branch name from filename
                preg_match_all('/[\u0e00-\u0e7f]+/u', $name_only, $matches);
                $branch_name = !empty($matches[0]) ? $matches[0][count($matches[0]) - 1] : 'unknown';

                // Read fiscal year from file
                $fiscal_year = '';
                try {
                    $raw_bytes = file_get_contents($files['tmp_name'][$i]);
                    if (class_exists('PhpOffice\\PhpSpreadsheet\\IOFactory')) {
                        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($files['tmp_name'][$i]);
                        $worksheet = $spreadsheet->getActiveSheet();

                        for ($r = 1; $r <= min(6, $worksheet->getHighestRow()); $r++) {
                            for ($c = 1; $c <= min(6, $worksheet->getHighestColumn()); $c++) {
                                $cell_val = (string)($worksheet->getCellByColumnAndRow($c, $r)->getValue() ?: '');
                                if (preg_match('/ปีงบประมาณ\s*(\d{4})/', $cell_val, $m)) {
                                    $fiscal_year = $m[1];
                                    break 2;
                                }
                            }
                        }
                        $spreadsheet->disconnectWorksheets();
                    }
                } catch (Exception $e) {
                    // Continue without fiscal year
                }

                $fy_suffix = $fiscal_year ? '_ปีงบ' . substr($fiscal_year, -2) : '';
                $new_name = $prefix . '_' . $branch_name . $fy_suffix . $ext;
            } elseif ($category === 'pending') {
                // Pending: batch merge, store raw bytes for later
                $pending_batch[] = [
                    'filename' => $filename,
                    'ext' => $ext,
                    'tmp_path' => $files['tmp_name'][$i]
                ];
                continue;
            } else {
                // Fallback
                if (preg_match('/(\d{6})/', $name_only, $m)) {
                    $new_name = $prefix . '_' . $m[1] . $ext;
                } else {
                    $clean = preg_replace('/[^\w\-.]/', '_', $name_only);
                    $clean = trim($clean, '_');
                    if (strlen($clean) > 30) $clean = substr($clean, 0, 30);
                    $new_name = $prefix . '_' . $clean . $ext;
                }
            }

            $dest_path = $folder_path . DIRECTORY_SEPARATOR . $new_name;
            $overwritten = file_exists($dest_path);

            // Move uploaded file
            if (!move_uploaded_file($files['tmp_name'][$i], $dest_path)) {
                throw new Exception('Failed to move uploaded file');
            }
            chmod($dest_path, 0644);

            $results[] = [
                'filename' => $new_name,
                'original' => $filename,
                'status' => $overwritten ? 'overwrite' : 'success',
                'message' => $filename . ' → ' . $new_name
            ];
        } catch (Exception $e) {
            $errors[] = [
                'filename' => $filename,
                'error' => $e->getMessage()
            ];
        }
    }

    // Pending batch merge
    if ($category === 'pending' && !empty($pending_batch)) {
        try {
            $HEADER_ROWS = 8;  // Row 0-7 = header, Row 8+ = data
            $DATE_COL = 3;     // Col 3 = date (0-based)

            $all_data_rows = [];
            $header_rows = null;

            global $phpSpreadsheet;
            if ($phpSpreadsheet && class_exists('PhpOffice\\PhpSpreadsheet\\IOFactory')) {
                foreach ($pending_batch as $item) {
                    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($item['tmp_path']);
                    $worksheet = $spreadsheet->getActiveSheet();

                    // Capture header from first file
                    if ($header_rows === null) {
                        $header_rows = [];
                        for ($r = 1; $r <= min($HEADER_ROWS, $worksheet->getHighestRow()); $r++) {
                            $row_data = [];
                            for ($c = 1; $c <= $worksheet->getHighestColumn(); $c++) {
                                $row_data[] = $worksheet->getCellByColumnAndRow($c, $r)->getValue();
                            }
                            $header_rows[] = $row_data;
                        }
                    }

                    // Capture data rows
                    for ($r = $HEADER_ROWS + 1; $r <= $worksheet->getHighestRow(); $r++) {
                        $row_data = [];
                        for ($c = 1; $c <= $worksheet->getHighestColumn(); $c++) {
                            $row_data[] = $worksheet->getCellByColumnAndRow($c, $r)->getValue();
                        }
                        // Skip empty rows
                        if (!empty($row_data[2])) {
                            $all_data_rows[] = $row_data;
                        }
                    }

                    $spreadsheet->disconnectWorksheets();
                }

                if (empty($all_data_rows)) {
                    $errors[] = ['filename' => 'pending', 'error' => 'ไม่พบข้อมูลในไฟล์ที่อัปโหลด'];
                } else {
                    // Sort by date
                    usort($all_data_rows, function($a, $b) use ($DATE_COL) {
                        $dt_a = parse_thai_date($a[$DATE_COL] ?? '');
                        $dt_b = parse_thai_date($b[$DATE_COL] ?? '');
                        if ($dt_a[0] === null) return 1;
                        if ($dt_b[0] === null) return -1;
                        return $dt_a[0] <=> $dt_b[0];
                    });

                    // Determine month range for filename
                    $min_date = null;
                    $max_date = null;

                    foreach ($all_data_rows as $row) {
                        [$dt, $by] = parse_thai_date($row[$DATE_COL] ?? '');
                        if ($dt && $by) {
                            $yy = $by % 100;
                            $mm = $dt->format('m');

                            if ($min_date === null || [$yy, $mm] < $min_date) {
                                $min_date = [$yy, $mm];
                            }
                            if ($max_date === null || [$yy, $mm] > $max_date) {
                                $max_date = [$yy, $mm];
                            }
                        }
                    }

                    // Create filename
                    if ($min_date && $max_date) {
                        $new_name = sprintf('ค้างซ่อม_%02d-%02d_to_%02d-%02d.xlsx',
                            $min_date[1], $min_date[0], $max_date[1], $max_date[0]);
                    } else {
                        $today = new DateTime();
                        $new_name = sprintf('ค้างซ่อม_%02d-%02d.xlsx',
                            $today->format('m'), $today->format('y'));
                    }

                    // Write merged Excel file
                    $wb_out = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
                    $ws_out = $wb_out->getActiveSheet();
                    $ws_out->setTitle('ค้างซ่อม');

                    // Write header
                    if ($header_rows) {
                        foreach ($header_rows as $r_idx => $row) {
                            foreach ($row as $c_idx => $val) {
                                $ws_out->setCellValueByColumnAndRow($c_idx + 1, $r_idx + 1, $val);
                            }
                        }
                    }

                    // Write data
                    foreach ($all_data_rows as $r_idx => $row) {
                        foreach ($row as $c_idx => $val) {
                            $ws_out->setCellValueByColumnAndRow($c_idx + 1, $HEADER_ROWS + $r_idx + 1, $val);
                        }
                    }

                    $dest_path = $folder_path . DIRECTORY_SEPARATOR . $new_name;
                    $overwritten = file_exists($dest_path);

                    $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($wb_out);
                    $writer->save($dest_path);
                    $wb_out->disconnectWorksheets();

                    // Clear cache for new file
                    $cache_file = CACHE_DIR . DIRECTORY_SEPARATOR . md5($dest_path) . '.json';
                    if (file_exists($cache_file)) unlink($cache_file);
                    // Clear Python-style cache too
                    $py_cache = $dest_path . '.cache.json';
                    if (file_exists($py_cache)) unlink($py_cache);

                    // Build SQLite database for fast querying
                    $sqlite_result = build_pending_sqlite($all_data_rows, $folder_path, $new_name);
                    $sqlite_msg = $sqlite_result ? ' + SQLite OK' : ' (SQLite failed)';

                    $orig_names = implode(', ', array_column($pending_batch, 'filename'));
                    $results[] = [
                        'filename' => $new_name,
                        'original' => $orig_names,
                        'status' => $overwritten ? 'overwrite' : 'success',
                        'message' => sprintf('รวม %d ไฟล์ (%d แถว) → %s%s',
                            count($pending_batch), count($all_data_rows), $new_name, $sqlite_msg)
                    ];
                }
            }
        } catch (Exception $e) {
            $errors[] = ['filename' => 'pending-merge', 'error' => $e->getMessage()];
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

// Route: GET /api/pending-chart
if ($method === 'GET' && count($path_parts) === 1 && $path_parts[0] === 'pending-chart') {
    $pending_dir = RAW_DATA_DIR . DIRECTORY_SEPARATOR . CATEGORY_MAP['pending'];

    if (!is_dir($pending_dir)) {
        json_response(['ok' => false, 'error' => 'ไม่พบโฟลเดอร์ข้อมูลค้างซ่อม'], 404);
    }

    $pf = find_pending_files($pending_dir);
    if (empty($pf['fy_list'])) {
        json_response(['ok' => false, 'error' => 'ไม่พบไฟล์ข้อมูลค้างซ่อม'], 404);
    }

    $req_fy = $_GET['fy'] ?? '';
    $fy = (is_numeric($req_fy) && $req_fy > 0) ? (int)$req_fy : $pf['latest_fy'];

    $db = open_pending_db($pending_dir, $pf, $fy);
    if (!$db) {
        json_response(['ok' => false, 'error' => 'ไม่สามารถเปิดฐานข้อมูล'], 500);
    }

    // Get update_date from meta
    $update_date = '';
    $res = $db->querySingle("SELECT value FROM meta WHERE key='update_date'");
    if ($res) $update_date = $res;

    // Fiscal year CE range
    $fy_be = $fy > 0 ? $fy : 2569;
    $fy_ce = $fy_be - 543;
    $count_start = "$fy_ce-01-01";

    // Get all records with parsed dates for chart calculation
    $records = [];
    $result = $db->query("SELECT date_ce, finish_ce, status, branch FROM pending_rows WHERE date_ce >= '$count_start' ORDER BY date_ce");
    while ($row = $result->fetchArray(SQLITE3_ASSOC)) {
        $records[] = $row;
    }

    // Also get records before count_start that might still be pending
    $result2 = $db->query("SELECT date_ce, finish_ce, status, branch FROM pending_rows WHERE date_ce < '$count_start'");
    $early_records = [];
    while ($row = $result2->fetchArray(SQLITE3_ASSOC)) {
        $early_records[] = $row;
    }

    $db->close();

    // Find months with data
    $month_set = [];
    foreach ($records as $rec) {
        $dt = new DateTime($rec['date_ce']);
        $y = $dt->format('Y');
        $m = $dt->format('m');
        $month_set["$y-$m"] = [$y, $m];
    }
    ksort($month_set);

    // Build months list and data
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

    json_response([
        'ok' => true,
        'fy' => $fy,
        'fy_list' => $pf['fy_list'],
        'update_date' => $update_date,
        'pd2_months' => $pd2_months,
        'pd2_data' => $pd2_data,
        'pd1_data' => $pd1_data,
        'branches' => BRANCH_LIST
    ]);
}

// Route: GET /api/pending-table
if ($method === 'GET' && count($path_parts) === 1 && $path_parts[0] === 'pending-table') {
    $pending_dir = RAW_DATA_DIR . DIRECTORY_SEPARATOR . CATEGORY_MAP['pending'];

    if (!is_dir($pending_dir)) {
        json_response(['ok' => false, 'error' => 'ไม่พบโฟลเดอร์ข้อมูลค้างซ่อม'], 404);
    }

    $pf = find_pending_files($pending_dir);
    if (empty($pf['fy_list'])) {
        json_response(['ok' => false, 'error' => 'ไม่พบไฟล์ข้อมูลค้างซ่อม'], 404);
    }

    $req_fy = $_GET['fy'] ?? '';
    $fy = (is_numeric($req_fy) && $req_fy > 0) ? (int)$req_fy : $pf['latest_fy'];

    $db = open_pending_db($pending_dir, $pf, $fy);
    if (!$db) {
        json_response(['ok' => false, 'error' => 'ไม่สามารถเปิดฐานข้อมูล'], 500);
    }

    // Get update_date from meta
    $update_date = '';
    $res = $db->querySingle("SELECT value FROM meta WHERE key='update_date'");
    if ($res) $update_date = $res;

    // Fiscal year months
    $fy_yy_start = ($fy - 2500 - 1) > 0 ? ($fy - 2500 - 1) : 68;
    $fy_yy_end = $fy_yy_start + 1;
    $fy_months = [];

    for ($mm = 10; $mm <= 12; $mm++) {
        $fy_months[] = sprintf('%02d-%02d', $fy_yy_start, $mm);
    }
    for ($mm = 1; $mm <= 9; $mm++) {
        $fy_months[] = sprintf('%02d-%02d', $fy_yy_end, $mm);
    }

    // Query: count by branch × month where status contains 'ซ่อมไม่เสร็จ'
    $fy_months_sql = "'" . implode("','", $fy_months) . "'";
    $sql = "SELECT branch, month_key, COUNT(*) as cnt
            FROM pending_rows
            WHERE status LIKE '%ซ่อมไม่เสร็จ%'
              AND month_key IN ($fy_months_sql)
            GROUP BY branch, month_key";
    $result = $db->query($sql);

    // Initialize result
    $data = [];
    foreach (BRANCH_LIST as $b) {
        $data[$b] = [];
        foreach ($fy_months as $mk) {
            $data[$b][$mk] = 0;
        }
    }

    while ($row = $result->fetchArray(SQLITE3_ASSOC)) {
        $b = $row['branch'];
        $mk = $row['month_key'];
        if (isset($data[$b]) && isset($data[$b][$mk])) {
            $data[$b][$mk] = (int)$row['cnt'];
        }
    }

    $db->close();

    // Build totals
    $col_totals = [];
    $grand_total = 0;
    foreach ($fy_months as $mk) {
        $col_totals[$mk] = 0;
    }

    $data_out = [];
    foreach (BRANCH_LIST as $branch) {
        $bd = $data[$branch] ?? [];
        if (array_sum($bd) > 0) {
            $data_out[$branch] = $bd;
            foreach ($bd as $mk => $v) {
                $col_totals[$mk] += $v;
                $grand_total += $v;
            }
        }
    }

    json_response([
        'ok' => true,
        'fy' => $fy,
        'fy_be' => $fy,
        'fy_list' => $pf['fy_list'],
        'update_date' => $update_date,
        'branches' => BRANCH_LIST,
        'months' => $fy_months,
        'data' => $data_out,
        'col_totals' => $col_totals,
        'grand_total' => $grand_total
    ]);
}

// Route: GET /api/pending-detail
if ($method === 'GET' && count($path_parts) === 1 && $path_parts[0] === 'pending-detail') {
    $pending_dir = RAW_DATA_DIR . DIRECTORY_SEPARATOR . CATEGORY_MAP['pending'];

    if (!is_dir($pending_dir)) {
        json_response(['ok' => false, 'error' => 'ไม่พบโฟลเดอร์ข้อมูลค้างซ่อม'], 404);
    }

    $pf = find_pending_files($pending_dir);
    if (empty($pf['fy_list'])) {
        json_response(['ok' => false, 'error' => 'ไม่พบไฟล์ข้อมูลค้างซ่อม'], 404);
    }

    $req_fy = $_GET['fy'] ?? '';
    $fy = (is_numeric($req_fy) && $req_fy > 0) ? (int)$req_fy : $pf['latest_fy'];

    $req_month = $_GET['month'] ?? '';
    $req_branch = $_GET['branch'] ?? '';

    $db = open_pending_db($pending_dir, $pf, $fy);
    if (!$db) {
        json_response(['ok' => false, 'error' => 'ไม่สามารถเปิดฐานข้อมูล'], 500);
    }

    // Build query with filters
    $where = ["status LIKE '%ซ่อมไม่เสร็จ%'"];
    if ($req_month) $where[] = "month_key = '" . SQLite3::escapeString($req_month) . "'";
    if ($req_branch) $where[] = "branch = '" . SQLite3::escapeString($req_branch) . "'";

    $where_sql = implode(' AND ', $where);
    $sql = "SELECT branch, notify_no, date_val, job_no, type_val, side, team, tech, pipe, status, month_key
            FROM pending_rows WHERE $where_sql ORDER BY date_ce";

    $result = $db->query($sql);
    $records = [];

    while ($row = $result->fetchArray(SQLITE3_ASSOC)) {
        $records[] = [
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

    $db->close();

    json_response([
        'ok' => true,
        'records' => $records,
        'total' => count($records)
    ]);
}

// Route: GET /api/pending-nojob
if ($method === 'GET' && count($path_parts) === 1 && $path_parts[0] === 'pending-nojob') {
    $pending_dir = RAW_DATA_DIR . DIRECTORY_SEPARATOR . CATEGORY_MAP['pending'];

    if (!is_dir($pending_dir)) {
        json_response(['ok' => false, 'error' => 'ไม่พบโฟลเดอร์'], 404);
    }

    $pf = find_pending_files($pending_dir);
    if (empty($pf['fy_list'])) {
        json_response(['ok' => false, 'error' => 'ไม่พบไฟล์'], 404);
    }

    $req_fy = $_GET['fy'] ?? '';
    $fy = (is_numeric($req_fy) && $req_fy > 0) ? (int)$req_fy : $pf['latest_fy'];

    $db = open_pending_db($pending_dir, $pf, $fy);
    if (!$db) {
        json_response(['ok' => false, 'error' => 'ไม่สามารถเปิดฐานข้อมูล'], 500);
    }

    // Get update_date from meta
    $update_date = '';
    $res = $db->querySingle("SELECT value FROM meta WHERE key='update_date'");
    if ($res) $update_date = $res;

    // Find latest month_key
    $latest_mk = $db->querySingle("SELECT month_key FROM pending_rows ORDER BY date_ce DESC LIMIT 1");

    if (!$latest_mk) {
        $db->close();
        json_response(['ok' => true, 'by_branch' => [], 'records' => [], 'total' => 0,
                        'update_date' => $update_date, 'month_key' => '', 'fy' => $fy]);
    }

    // Query with all filters: latest month + pipe complaint + no job + not done
    $sql = "SELECT branch, notify_no, date_val, side, topic, detail, status
            FROM pending_rows
            WHERE month_key = :mk
              AND (side = 'ด้านท่อแตกรั่ว'
                   OR topic LIKE '%ท่อแตก%' OR topic LIKE '%ท่อรั่ว%' OR topic LIKE '%แตกรั่ว%'
                   OR detail LIKE '%ท่อแตก%' OR detail LIKE '%ท่อรั่ว%' OR detail LIKE '%แตกรั่ว%')
              AND (job_no IS NULL OR job_no = '')
              AND status NOT LIKE '%ดำเนินการแล้วเสร็จ%'
            ORDER BY branch, date_ce";

    $stmt = $db->prepare($sql);
    $stmt->bindValue(':mk', $latest_mk, SQLITE3_TEXT);
    $result = $stmt->execute();

    $by_branch = [];
    $records = [];

    while ($row = $result->fetchArray(SQLITE3_ASSOC)) {
        $branch = $row['branch'];
        $by_branch[$branch] = ($by_branch[$branch] ?? 0) + 1;
        $records[] = [
            'branch' => $branch,
            'notify_no' => $row['notify_no'],
            'date' => $row['date_val'],
            'side' => $row['side'],
            'topic' => $row['topic'],
            'detail' => mb_substr($row['detail'], 0, 100),
            'status' => $row['status']
        ];
    }

    $db->close();

    json_response([
        'ok' => true,
        'by_branch' => $by_branch,
        'records' => $records,
        'total' => count($records),
        'update_date' => $update_date,
        'month_key' => $latest_mk,
        'fy' => $fy
    ]);
}

// Route: DELETE /api/data/<category>/<filename>
if ($method === 'DELETE' && count($path_parts) === 3 && $path_parts[0] === 'data') {
    $category = $path_parts[1];
    $filename = $path_parts[2];

    if (!isset(CATEGORY_MAP[$category])) {
        json_response(['ok' => false, 'error' => 'ไม่รู้จัก category: ' . $category], 400);
    }

    $folder_path = RAW_DATA_DIR . DIRECTORY_SEPARATOR . CATEGORY_MAP[$category];
    $file_path = $folder_path . DIRECTORY_SEPARATOR . $filename;

    // Safety check
    $abs_file = realpath($file_path);
    $abs_folder = realpath($folder_path);

    if (!$abs_file || strpos($abs_file, $abs_folder) !== 0) {
        json_response(['ok' => false, 'error' => 'ไม่อนุญาต'], 403);
    }

    try {
        if (file_exists($file_path)) {
            unlink($file_path);

            // Clear cache
            $cache_file = CACHE_DIR . DIRECTORY_SEPARATOR . md5($file_path) . '.json';
            if (file_exists($cache_file)) unlink($cache_file);

            json_response(['ok' => true, 'filename' => $filename, 'deleted' => true]);
        } else {
            json_response(['ok' => false, 'error' => 'ไม่พบไฟล์'], 404);
        }
    } catch (Exception $e) {
        json_response(['ok' => false, 'error' => $e->getMessage()], 500);
    }
}

// Route: POST /api/notes/<slug>
if ($method === 'POST' && count($path_parts) === 2 && $path_parts[0] === 'notes') {
    $slug = $path_parts[1];

    if (!isset(CATEGORY_MAP[$slug])) {
        json_response(['ok' => false, 'error' => 'invalid slug'], 400);
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
    $parent_dir = dirname(BASE_DIR);
    json_response([
        'ok' => true,
        'path' => $parent_dir,
        'note' => 'Parent directory path returned; OS-specific opening not available in PHP'
    ]);
}

/**
 * Parse date string (YYYY-MM-DD or DD/MM/YYYY)
 * Returns: [$date_key, $date_label] or [null, null]
 */
function parse_date_key($date_str) {
    $date_str = trim($date_str);

    // Try YYYY-MM-DD
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

    // Try DD/MM/YYYY
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

// ───────────────────────────────────────────────────────────────────────────
// Route: GET /api/repair-data
// Returns KPI จุดซ่อมท่อ data (TAB 1) parsed live from GIS_YYMMDD.xlsx files
// ───────────────────────────────────────────────────────────────────────────
if ($method === 'GET' && count($path_parts) >= 1 && $path_parts[0] === 'repair-data') {
    $repair_dir = RAW_DATA_DIR . DIRECTORY_SEPARATOR . 'ลงข้อมูลซ่อมท่อ';

    if (!is_dir($repair_dir)) {
        json_response(['ok' => true, 'has_data' => false, 'message' => 'ไม่พบโฟลเดอร์ลงข้อมูลซ่อมท่อ']);
    }

    // Check if PhpSpreadsheet is available
    global $phpSpreadsheet;
    if (!$phpSpreadsheet) {
        json_response(['ok' => false, 'error' => 'PhpSpreadsheet ไม่พร้อมใช้งาน'], 500);
    }

    // Cache: use file modification time of repair folder
    $cache_file = CACHE_DIR . DIRECTORY_SEPARATOR . 'repair_data.json';
    $folder_mtime = 0;
    foreach (scandir($repair_dir) as $f) {
        if ($f[0] === '.' || !preg_match('/\.xlsx$/i', $f)) continue;
        $mt = filemtime($repair_dir . DIRECTORY_SEPARATOR . $f);
        if ($mt > $folder_mtime) $folder_mtime = $mt;
    }

    // Return cached data if still fresh
    if (file_exists($cache_file)) {
        $cache_mtime = filemtime($cache_file);
        if ($cache_mtime >= $folder_mtime) {
            $cached = json_decode(file_get_contents($cache_file), true);
            if ($cached) {
                json_response($cached);
            }
        }
    }

    // Parse all GIS repair files
    $all_data = [];    // month_key => branch => {closed, complete, score}
    $branches_set = [];
    $month_names = [
        '01' => 'ม.ค.', '02' => 'ก.พ.', '03' => 'มี.ค.', '04' => 'เม.ย.',
        '05' => 'พ.ค.', '06' => 'มิ.ย.', '07' => 'ก.ค.', '08' => 'ส.ค.',
        '09' => 'ก.ย.', '10' => 'ต.ค.', '11' => 'พ.ย.', '12' => 'ธ.ค.'
    ];

    // Find all xlsx files and pick latest per month
    $month_files = []; // month_key => [dd, filepath]
    foreach (scandir($repair_dir) as $fname) {
        if ($fname[0] === '.' || $fname[0] === '~') continue;
        if (!preg_match('/\.xlsx$/i', $fname)) continue;
        if (!preg_match('/(\d{6})/', $fname, $m)) continue;

        $digits = $m[1];
        $yy = (int)substr($digits, 0, 2);
        $mm = (int)substr($digits, 2, 2);
        $dd = (int)substr($digits, 4, 2);
        if ($mm < 1 || $mm > 12) continue;

        $month_key = sprintf("%02d-%02d", $yy, $mm);
        if (!isset($month_files[$month_key]) || $dd > $month_files[$month_key][0]) {
            $month_files[$month_key] = [$dd, $repair_dir . DIRECTORY_SEPARATOR . $fname];
        }
    }

    // ── Smart Header Detection for Repair ──
    // ค้นหาคอลัมน์จาก keyword แทนตำแหน่งตายตัว
    function detect_repair_columns($worksheet) {
        $keywords = [
            'branch'   => ['ชื่อสาขา', 'สาขา', 'หน่วยงาน', 'branch'],
            'closed'   => ['ปิดงาน', 'ปิด', 'closed'],
            'complete' => ['สำเร็จ', 'เสร็จสิ้น', 'เสร็จ', 'complete'],
            'score'    => ['คะแนน', 'score', 'ผลคะแนน'],
        ];
        $maxScan = min(10, $worksheet->getHighestRow());
        $maxCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($worksheet->getHighestColumn());
        $maxCol = min($maxCol, 20);

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
            if (isset($found['branch']) && count($found) >= 2) {
                return ['header_row' => $r, 'cols' => $found, 'fallback' => false];
            }
        }
        // fallback: ตำแหน่งเดิม
        return ['header_row' => 1, 'cols' => ['branch' => 1, 'closed' => 2, 'complete' => 3, 'score' => 4], 'fallback' => true];
    }

    // Parse each file
    foreach ($month_files as $month_key => $info) {
        $fpath = $info[1];
        try {
            $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($fpath);
            $worksheet = $spreadsheet->getSheet(0);
            $max_row = $worksheet->getHighestRow();
            $month_data = [];

            $det = detect_repair_columns($worksheet);
            $hdr = $det['header_row'];
            $cBranch   = $det['cols']['branch']   ?? 1;
            $cClosed   = $det['cols']['closed']    ?? 2;
            $cComplete = $det['cols']['complete']  ?? 3;
            $cScore    = $det['cols']['score']     ?? 4;

            for ($r = $hdr + 1; $r <= $max_row; $r++) {
                $branch = $worksheet->getCellByColumnAndRow($cBranch, $r)->getValue();
                if ($branch === null || !is_string($branch)) continue;
                $branch = trim($branch);
                if ($branch === '' || mb_strpos($branch, 'ชื่อสาขา') !== false) continue;

                $closed_v = $worksheet->getCellByColumnAndRow($cClosed, $r)->getValue();
                $complete_v = $worksheet->getCellByColumnAndRow($cComplete, $r)->getValue();
                $score_v = isset($det['cols']['score']) ? $worksheet->getCellByColumnAndRow($cScore, $r)->getValue() : 0;

                $closed = is_numeric($closed_v) ? (int)$closed_v : 0;
                $complete = is_numeric($complete_v) ? (int)$complete_v : 0;
                $score = is_numeric($score_v) ? round((float)$score_v, 2) : 0;

                $month_data[$branch] = [
                    'closed' => $closed,
                    'complete' => $complete,
                    'score' => $score
                ];

                if (!in_array($branch, $branches_set)) {
                    $branches_set[] = $branch;
                }
            }

            $all_data[$month_key] = $month_data;
            $spreadsheet->disconnectWorksheets();
            unset($spreadsheet);
        } catch (Exception $e) {
            error_log("Error parsing repair file $fpath: " . $e->getMessage());
        }
    }

    // Sort months
    $months = array_keys($all_data);
    sort($months);

    $result = [
        'ok' => true,
        'has_data' => !empty($all_data),
        'months' => $months,
        'branches' => $branches_set,
        'data' => $all_data,
        'month_names' => $month_names
    ];

    // Save cache
    if (!is_dir(CACHE_DIR)) mkdir(CACHE_DIR, 0755, true);
    file_put_contents($cache_file, json_encode($result, JSON_UNESCAPED_UNICODE));

    json_response($result);
}

// 404 - Route not found
json_response([
    'ok' => false,
    'error' => 'Route not found: ' . $method . ' ' . $path_info
], 404);
