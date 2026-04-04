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

// ─── Error Handling ────────────────────────────────────────────────────────
ini_set('display_errors', '0');
error_reporting(E_ALL);
ini_set('log_errors', '1');

// Catch fatal errors (e.g. memory exhaustion) and return JSON instead of empty response
ob_start();
register_shutdown_function(function() {
    $error = error_get_last();
    if ($error && in_array($error['type'], [E_ERROR, E_CORE_ERROR, E_COMPILE_ERROR])) {
        ob_end_clean();
        http_response_code(500);
        header('Content-Type: application/json; charset=utf-8');
        header('Access-Control-Allow-Origin: *');
        $msg = 'PHP Fatal Error: ' . $error['message'];
        if (stripos($error['message'], 'memory') !== false) {
            $msg = 'หน่วยความจำไม่พอ — ลองอัปโหลดทีละน้อยไฟล์ (เช่น 2-3 ไฟล์ต่อครั้ง)';
        }
        echo json_encode(['ok' => false, 'error' => $msg], JSON_UNESCAPED_UNICODE);
    }
});

// ─── Configuration ─────────────────────────────────────────────────────────

define('BASE_DIR', __DIR__);
define('RAW_DATA_DIR', BASE_DIR . DIRECTORY_SEPARATOR . 'ข้อมูลดิบ');
define('CACHE_DIR', BASE_DIR . DIRECTORY_SEPARATOR . '.cache');
define('CACHE_TTL', 86400); // 1 day — mtime check handles invalidation when Excel files change

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
        } catch (\Throwable $e) {
            return [null, null];
        }
    }

    return [null, null];
}

/**
 * Detect the dominant month/year from a pending (ซ่อมท่อค้างระบบ) Excel file.
 * Uses ReadFilter to load only first ~50 data rows (saves RAM for large .xls files).
 * Returns: ['mm' => int, 'yy' => int] (Buddhist Era 2-digit) or null
 */
function detect_pending_month($fpath) {
    if (!class_exists('PhpOffice\\PhpSpreadsheet\\IOFactory')) return null;
    try {
        ini_set('memory_limit', '1024M');

        $HEADER_ROWS = 8;
        $SAMPLE_ROWS = 50; // Read only 50 data rows to detect month
        $DATE_COL = 3; // 0-indexed → column D (1-indexed = 4)
        $MAX_ROW = $HEADER_ROWS + $SAMPLE_ROWS;

        // Use ReadFilter to limit rows loaded into memory
        $filter = new class($MAX_ROW) implements \PhpOffice\PhpSpreadsheet\Reader\IReadFilter {
            private $maxRow;
            public function __construct($maxRow) { $this->maxRow = $maxRow; }
            public function readCell(string $columnAddress, int $row, string $worksheetName = ''): bool {
                return $row <= $this->maxRow;
            }
        };

        $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReaderForFile($fpath);
        $reader->setReadFilter($filter);
        $reader->setReadDataOnly(true);
        $spreadsheet = $reader->load($fpath);
        $worksheet = $spreadsheet->getActiveSheet();

        $month_counts = []; // "MM-YY" => count

        for ($r = $HEADER_ROWS + 1; $r <= min($MAX_ROW, $worksheet->getHighestRow()); $r++) {
            $val = $worksheet->getCell([$DATE_COL + 1, $r])->getValue(); // +1 for 1-indexed
            [$dt, $by] = parse_thai_date($val);
            if ($dt && $by) {
                $key = sprintf('%02d-%02d', (int)$dt->format('m'), $by % 100);
                $month_counts[$key] = ($month_counts[$key] ?? 0) + 1;
            }
        }

        $spreadsheet->disconnectWorksheets();
        unset($spreadsheet);

        if (empty($month_counts)) return null;

        // Find most frequent month
        arsort($month_counts);
        $top = array_key_first($month_counts);
        [$mm, $yy] = explode('-', $top);
        return ['mm' => (int)$mm, 'yy' => (int)$yy];
    } catch (\Throwable $e) {
        return null;
    }
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
        } catch (\Throwable $e) {}
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
        } catch (\Throwable $e) {}
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
    } catch (\Throwable $e) {
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
function build_pending_sqlite($all_data_rows, $folder_path, $db_or_excel_name) {
    // Accept either .sqlite name directly or .xlsx name (for backward compat)
    if (preg_match('/\.sqlite$/i', $db_or_excel_name)) {
        $db_name = $db_or_excel_name;
    } else {
        $db_name = preg_replace('/\.xlsx?$/i', '.sqlite', $db_or_excel_name);
    }
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

    } catch (\Throwable $e) {
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

        $fpath = $pending_dir . DIRECTORY_SEPARATOR . $fname;

        // Format 1 (old merged): ค้างซ่อม_10-68_to_03-69.xlsx
        if (preg_match('/(\d{1,2})-(\d{2})_to_(\d{1,2})-(\d{2})/', $fname, $m)) {
            $start_mm = (int)$m[1];
            $start_yy = (int)$m[2];
            $fy_be = ($start_mm >= 10) ? (2500 + $start_yy + 1) : (2500 + $start_yy);

            $sqlite_name = preg_replace('/\.xlsx?$/i', '.sqlite', $fname);
            $sqlite_path = $pending_dir . DIRECTORY_SEPARATOR . $sqlite_name;

            if (!isset($fy_files[$fy_be])) {
                $fy_files[$fy_be] = ['excel_files' => [], 'sqlite' => null];
            }
            $fy_files[$fy_be]['excel_files'][] = $fpath;
            if (file_exists($sqlite_path)) {
                $fy_files[$fy_be]['sqlite'] = $sqlite_path;
            }
            continue;
        }

        // Format 2 (new per-month): ค้างซ่อม_MM-YY.xlsx
        if (preg_match('/ค้างซ่อม_(\d{2})-(\d{2})\.xlsx?$/i', $fname, $m)) {
            $mm = (int)$m[1];
            $yy = (int)$m[2];
            $fy_be = ($mm >= 10) ? (2500 + $yy + 1) : (2500 + $yy);

            if (!isset($fy_files[$fy_be])) {
                $fy_files[$fy_be] = ['excel_files' => [], 'sqlite' => null];
            }
            $fy_files[$fy_be]['excel_files'][] = $fpath;
            continue;
        }
    }

    // Check for combined SQLite per FY (ค้างซ่อม_fy_XXXX.sqlite)
    foreach ($fy_files as $fy_be => &$info) {
        $combined_sqlite = $pending_dir . DIRECTORY_SEPARATOR . "ค้างซ่อม_fy_{$fy_be}.sqlite";
        if (file_exists($combined_sqlite) && $info['sqlite'] === null) {
            // Verify it's still fresh (newest Excel must not be newer than SQLite)
            $sqlite_mtime = filemtime($combined_sqlite);
            $newest_excel = 0;
            foreach ($info['excel_files'] as $ef) {
                $newest_excel = max($newest_excel, filemtime($ef));
            }
            if ($sqlite_mtime >= $newest_excel) {
                $info['sqlite'] = $combined_sqlite;
            }
        }
    }
    unset($info);

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

    // If SQLite exists and is fresh, use it
    if ($info['sqlite'] && file_exists($info['sqlite'])) {
        try {
            $db = new SQLite3($info['sqlite'], SQLITE3_OPEN_READONLY);
            return $db;
        } catch (\Throwable $e) {
            error_log("SQLite open error: " . $e->getMessage());
        }
    }

    // No fresh SQLite — build from Excel files
    $excel_files = $info['excel_files'] ?? [];
    if (empty($excel_files)) return null;

    // Collect all data rows from all Excel files for this FY
    $all_data_rows = [];
    $DATA_START = 8;

    foreach ($excel_files as $excel_path) {
        if (!file_exists($excel_path)) continue;
        $rows = read_excel_cached($excel_path);
        if ($rows === null) continue;

        $data_rows = array_slice($rows, $DATA_START);
        foreach ($data_rows as $dr) {
            $all_data_rows[] = $dr;
        }
        unset($rows, $data_rows);
    }

    if (empty($all_data_rows)) return null;

    // Build combined SQLite for this FY
    $combined_name = "ค้างซ่อม_fy_{$fy}.sqlite";
    $db_path = build_pending_sqlite($all_data_rows, $pending_dir, $combined_name);
    unset($all_data_rows);

    if ($db_path) {
        try {
            return new SQLite3($db_path, SQLITE3_OPEN_READONLY);
        } catch (\Throwable $e) {
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
            } catch (\Throwable $e) {
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
        } catch (\Throwable $e) {
            error_log("Error loading notes.json: " . $e->getMessage());
        }
    }

    json_response([
        'ok' => true,
        'inventory' => $inventory,
        'notes' => $notes
    ]);
}

// ── Smart Header Detection for Repair ──
// ค้นหาคอลัมน์จาก keyword แทนตำแหน่งตายตัว
// ★ ต้องอยู่ top-level เพราะใช้ทั้งใน validate_gis_file (pre-check) และ repair-data endpoint
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
        if (isset($found['branch']) && count($found) >= 2) {
            return ['header_row' => $r, 'cols' => $found, 'fallback' => false];
        }
    }
    // fallback: ตำแหน่งเดิม
    return ['header_row' => 1, 'cols' => ['branch' => 1, 'closed' => 2, 'complete' => 3, 'score' => 4], 'fallback' => true];
}

// ── Validate file format before upload ──
function validate_gis_file($tmp_path, $category) {
    if (!class_exists('PhpOffice\\PhpSpreadsheet\\IOFactory')) {
        return ['valid' => false, 'message' => 'PhpSpreadsheet ไม่พร้อมใช้งาน — ไม่สามารถตรวจสอบไฟล์ได้'];
    }
    if (!class_exists('ZipArchive')) {
        return ['valid' => false, 'message' => 'PHP zip extension ไม่ได้เปิด — ไม่สามารถอ่านไฟล์ .xlsx ได้ (กรุณาเปิด extension=zip ใน php.ini แล้ว restart Apache)'];
    }
    try {
        // Pending: use lightweight validation (skip full file load to save RAM)
        if ($category === 'pending') {
            $valFilter = new class implements \PhpOffice\PhpSpreadsheet\Reader\IReadFilter {
                public function readCell(string $columnAddress, int $row, string $worksheetName = ''): bool {
                    return $row <= 15;
                }
            };
            $valReader = \PhpOffice\PhpSpreadsheet\IOFactory::createReaderForFile($tmp_path);
            $valReader->setReadFilter($valFilter);
            $valReader->setReadDataOnly(true);
            $valSpreadsheet = $valReader->load($tmp_path);
            $hasData = $valSpreadsheet->getActiveSheet()->getHighestRow() >= 2;
            $valSpreadsheet->disconnectWorksheets();
            unset($valSpreadsheet);

            if (!$hasData) {
                return [
                    'valid' => false,
                    'message' => 'ไฟล์ไม่มีข้อมูล (มีเฉพาะ header หรือว่างเปล่า)'
                ];
            }
            return ['valid' => true, 'message' => ''];
        }

        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($tmp_path);
        // ╔══════════════════════════════════════════════════════════════╗
        // ║ [FILE VALIDATION] SCAN ALL SHEETS                           ║
        // ║                                                              ║
        // ║ ตรวจสอบไฟล์อัปโหลด — สแกนทุกชีท (ไม่ใช่แค่ชีทแรก)          ║
        // ║ เพราะไฟล์อาจมีหลายชีท ข้อมูลอาจไม่อยู่ชีทแรก               ║
        // ║ ข้ามชีทที่ชื่อ สารบัญ/สรุป/ปก/cover/summary/chart/graph    ║
        // ╚══════════════════════════════════════════════════════════════╝
        $gis_skip_sheets = ['สารบัญ', 'สรุป', 'ปก', 'cover', 'summary', 'chart', 'graph', 'index'];

        if ($category === 'repair') {
            $det = null;
            for ($si = 0; $si < $spreadsheet->getSheetCount(); $si++) {
                $ws = $spreadsheet->getSheet($si);
                $wsName = mb_strtolower($ws->getTitle());
                $skip = false;
                foreach ($gis_skip_sheets as $sk) {
                    if (mb_strpos($wsName, $sk) !== false) { $skip = true; break; }
                }
                if ($skip) continue;
                $det = detect_repair_columns($ws);
                if (!$det['fallback']) break;
            }
            $spreadsheet->disconnectWorksheets();
            if ($det === null || $det['fallback']) {
                return [
                    'valid' => false,
                    'message' => 'ไม่พบหัวคอลัมน์ที่คาดหวัง (สาขา, ปิดงาน, สำเร็จ, คะแนน) — รูปแบบไฟล์ไม่ตรงกับที่ระบบรองรับ'
                ];
            }
        } elseif ($category === 'pressure') {
            // ตรวจสอบว่ามี "ปีงบประมาณ" หรือชื่อสาขาในไฟล์ (สแกนทุกชีท)
            $found_fy = false;
            $found_branch = false;
            for ($si = 0; $si < $spreadsheet->getSheetCount(); $si++) {
                $ws = $spreadsheet->getSheet($si);
                $wsName = mb_strtolower($ws->getTitle());
                $skip = false;
                foreach ($gis_skip_sheets as $sk) {
                    if (mb_strpos($wsName, $sk) !== false) { $skip = true; break; }
                }
                if ($skip) continue;
                $maxR = min(10, $ws->getHighestRow());
                $maxC = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($ws->getHighestColumn());
                $maxC = min($maxC, 10);
                for ($r = 1; $r <= $maxR; $r++) {
                    for ($c = 1; $c <= $maxC; $c++) {
                        $v = trim((string)($ws->getCell([$c, $r])->getValue() ?? ''));
                        if (preg_match('/ปีงบประมาณ/', $v)) $found_fy = true;
                        if (preg_match('/สาขา|แรงดัน|pressure/i', $v)) $found_branch = true;
                    }
                }
                if ($found_fy || $found_branch) break;
            }
            $spreadsheet->disconnectWorksheets();
            if (!$found_fy && !$found_branch) {
                return [
                    'valid' => false,
                    'message' => 'ไม่พบข้อมูล "ปีงบประมาณ" หรือ "สาขา/แรงดัน" ในไฟล์ — รูปแบบไฟล์ไม่ตรงกับที่ระบบรองรับ'
                ];
            }
        } else {
            $spreadsheet->disconnectWorksheets();
        }

        return ['valid' => true, 'message' => ''];
    } catch (\Throwable $e) {
        return [
            'valid' => false,
            'message' => 'ไม่สามารถอ่านไฟล์ Excel ได้: ' . $e->getMessage()
        ];
    }
}

// ── Helper: Generate new filename based on category ──
function generate_new_filename($category, $original_filename, $excel_tmp_path = null) {
    $pathinfo = pathinfo($original_filename);
    $name_only = $pathinfo['filename'];
    $ext = isset($pathinfo['extension']) ? '.' . $pathinfo['extension'] : '.xlsx';

    $PREFIX_MAP = ['repair' => 'GIS', 'pressure' => 'PRESSURE', 'pending' => 'PENDING'];
    $prefix = isset($PREFIX_MAP[$category]) ? $PREFIX_MAP[$category] : strtoupper($category);

    if ($category === 'repair') {
        // repair: GIS_YYMMDD.xlsx
        if (preg_match('/(\d{6})/', $name_only, $m)) {
            return $prefix . '_' . $m[1] . $ext;
        } else {
            $today = date('ymd');
            return $prefix . '_' . $today . $ext;
        }
    } elseif ($category === 'pressure') {
        // pressure: PRESSURE_สาขา_ปีงบYY.xlsx
        // Extract branch name from filename — กรองคำที่ไม่ใช่ชื่อสาขาออก
        preg_match_all('/[\x{0e00}-\x{0e7f}]+/u', $name_only, $matches);
        $non_branch_words = ['ปีงบ', 'ปีงบประมาณ', 'แรงดัน', 'แรงดันน้ำ'];
        $branch_candidates = array_filter($matches[0] ?? [], function($w) use ($non_branch_words) {
            foreach ($non_branch_words as $nbw) {
                if (mb_strpos($w, $nbw) !== false) return false;
            }
            return true;
        });
        $branch_name = !empty($branch_candidates) ? reset($branch_candidates) : 'unknown';

        // Read fiscal year: try filename first, then Excel content
        $fiscal_year = '';

        // Try to extract from filename (pattern: _ปีงบYY)
        if (preg_match('/_ปีงบ(\d{2})/', $name_only, $m)) {
            $fiscal_year = '25' . $m[1];
        }

        // If not found in filename, read from Excel content
        if (!$fiscal_year && $excel_tmp_path && file_exists($excel_tmp_path)) {
            try {
                if (class_exists('PhpOffice\\PhpSpreadsheet\\IOFactory')) {
                    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($excel_tmp_path);
                    $worksheet = $spreadsheet->getActiveSheet();

                    // Look for "ปีงบประมาณ XXXX" in first few rows
                    for ($r = 1; $r <= min(6, $worksheet->getHighestRow()); $r++) {
                        for ($c = 1; $c <= min(6, $worksheet->getHighestColumn()); $c++) {
                            $cell_val = (string)($worksheet->getCell([$c, $r])->getValue() ?: '');
                            if (preg_match('/ปีงบประมาณ\s*(\d{4})/', $cell_val, $m)) {
                                $fiscal_year = $m[1];
                                break 2;
                            }
                        }
                    }

                    // If still not found, try from sheet names — majority vote
                    if (!$fiscal_year) {
                        $yc = [];
                        foreach ($spreadsheet->getSheetNames() as $sname) {
                            if (preg_match('/(\d{2})\s*$/', trim($sname), $sm)) {
                                $yy = $sm[1];
                                $yc[$yy] = ($yc[$yy] ?? 0) + 1;
                            }
                        }
                        if (!empty($yc)) {
                            arsort($yc);
                            $fiscal_year = '25' . array_key_first($yc);
                        }
                    }

                    $spreadsheet->disconnectWorksheets();
                }
            } catch (\Throwable $e) {
                // Continue without fiscal year
            }
        }

        $fy_suffix = $fiscal_year ? '_ปีงบ' . substr($fiscal_year, -2) : '';
        return $prefix . '_' . $branch_name . $fy_suffix . $ext;
    } elseif ($category === 'pending') {
        // For pre-check, we can't know the final merged filename without reading all files
        // Return a placeholder that indicates files will be merged
        return null;  // Will be determined during merge
    } else {
        // Fallback
        if (preg_match('/(\d{6})/', $name_only, $m)) {
            return $prefix . '_' . $m[1] . $ext;
        } else {
            $clean = preg_replace('/[^\w\-.]/', '_', $name_only);
            $clean = trim($clean, '_');
            if (strlen($clean) > 30) $clean = substr($clean, 0, 30);
            return $prefix . '_' . $clean . $ext;
        }
    }
}

// ── Helper: Cleanup temp directory ──
function cleanup_temp_dir($batch_id) {
    $temp_dir = RAW_DATA_DIR . DIRECTORY_SEPARATOR . '__tmp_upload' . DIRECTORY_SEPARATOR . $batch_id;
    if (is_dir($temp_dir)) {
        $files = array_diff(scandir($temp_dir), ['.', '..']);
        foreach ($files as $f) {
            $path = $temp_dir . DIRECTORY_SEPARATOR . $f;
            if (is_file($path)) unlink($path);
        }
        rmdir($temp_dir);
    }
}

// Route: POST /api/pre-check/{category}
if ($method === 'POST' && count($path_parts) === 2 && $path_parts[0] === 'pre-check') {
    $category = $path_parts[1];

    if (!isset(CATEGORY_MAP[$category])) {
        json_response([
            'ok' => false,
            'error' => 'ไม่รู้จัก category: ' . $category
        ], 400);
    }

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

    // Generate batch ID
    $batch_id = uniqid('batch_', true);
    $temp_base = RAW_DATA_DIR . DIRECTORY_SEPARATOR . '__tmp_upload';
    $temp_dir = $temp_base . DIRECTORY_SEPARATOR . $batch_id;

    if (!is_dir($temp_dir)) {
        mkdir($temp_dir, 0755, true);
    }

    $folder_path = RAW_DATA_DIR . DIRECTORY_SEPARATOR . CATEGORY_MAP[$category];
    if (!is_dir($folder_path)) {
        mkdir($folder_path, 0755, true);
    }

    $preview = [];
    $errors = [];
    $pending_files_for_merge = [];  // For pending category

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
            // Validate file format
            if (preg_match('/\.xlsx?$/i', $filename)) {
                $validation = validate_gis_file($files['tmp_name'][$i], $category);
                if (!$validation['valid']) {
                    $errors[] = [
                        'filename' => $filename,
                        'error' => $validation['message']
                    ];
                    continue;
                }
            }

            // Save to temp directory
            $temp_path = $temp_dir . DIRECTORY_SEPARATOR . $filename;
            if (!move_uploaded_file($files['tmp_name'][$i], $temp_path)) {
                $errors[] = [
                    'filename' => $filename,
                    'error' => 'Failed to save file to temp directory'
                ];
                continue;
            }
            chmod($temp_path, 0644);

            if ($category === 'pending') {
                // For pending, detect month from data and generate per-file name
                $detected = detect_pending_month($temp_path);
                if ($detected) {
                    $new_name = sprintf('ค้างซ่อม_%02d-%02d.xlsx', $detected['mm'], $detected['yy']);
                } else {
                    // Fallback: use current date
                    $today = new DateTime();
                    $new_name = sprintf('ค้างซ่อม_%02d-%02d.xlsx',
                        (int)$today->format('m'), ((int)$today->format('Y') + 543) % 100);
                }

                $dest_path = $folder_path . DIRECTORY_SEPARATOR . $new_name;
                $will_overwrite = file_exists($dest_path);

                $preview[] = [
                    'original' => $filename,
                    'new_name' => $new_name,
                    'valid' => true,
                    'will_overwrite' => $will_overwrite,
                    'overwrite_file' => $will_overwrite ? basename($dest_path) : null,
                    'detected_month' => $detected ? sprintf('%02d/%02d', $detected['mm'], $detected['yy']) : null
                ];
            } else {
                // For repair/pressure, determine new filename
                $new_name = generate_new_filename($category, $filename, $temp_path);
                if (!$new_name) {
                    $errors[] = [
                        'filename' => $filename,
                        'error' => 'Failed to generate filename'
                    ];
                    continue;
                }

                $dest_path = $folder_path . DIRECTORY_SEPARATOR . $new_name;
                $will_overwrite = file_exists($dest_path);
                $overwrite_file = $will_overwrite ? basename($dest_path) : null;

                $preview[] = [
                    'original' => $filename,
                    'new_name' => $new_name,
                    'valid' => true,
                    'will_overwrite' => $will_overwrite,
                    'overwrite_file' => $overwrite_file
                ];
            }
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
        'preview' => $preview,
        'errors' => $errors
    ]);
}

// Route: POST /api/upload-confirm/{batch_id}
if ($method === 'POST' && count($path_parts) === 2 && $path_parts[0] === 'upload-confirm') {
    $batch_id = $path_parts[1];
    $category = isset($_POST['category']) ? trim($_POST['category']) : '';

    if (!$category || !isset(CATEGORY_MAP[$category])) {
        json_response([
            'ok' => false,
            'error' => 'ไม่ระบุหรือไม่รู้จัก category'
        ], 400);
    }

    $temp_base = RAW_DATA_DIR . DIRECTORY_SEPARATOR . '__tmp_upload';
    $temp_dir = $temp_base . DIRECTORY_SEPARATOR . $batch_id;

    if (!is_dir($temp_dir)) {
        json_response([
            'ok' => false,
            'error' => 'ไม่พบ batch ที่ระบุ'
        ], 400);
    }

    $folder_path = RAW_DATA_DIR . DIRECTORY_SEPARATOR . CATEGORY_MAP[$category];
    if (!is_dir($folder_path)) {
        mkdir($folder_path, 0755, true);
    }

    $results = [];
    $errors = [];

    // Boost memory for merging multiple large .xls files
    ini_set('memory_limit', '2048M');

    try {
        if ($category === 'pending') {
            // Pending: save each file separately with ค้างซ่อม_MM-YY naming
            $files_in_temp = array_diff(scandir($temp_dir), ['.', '..']);

            if (empty($files_in_temp)) {
                $errors[] = ['filename' => 'pending', 'error' => 'ไม่พบไฟล์ในข้อมูลชั่วคราว'];
            } else {
                foreach ($files_in_temp as $fname) {
                    $temp_path = $temp_dir . DIRECTORY_SEPARATOR . $fname;
                    if (!is_file($temp_path)) continue;

                    try {
                        // Detect month from data inside the file
                        $detected = detect_pending_month($temp_path);
                        if ($detected) {
                            $new_name = sprintf('ค้างซ่อม_%02d-%02d.xlsx', $detected['mm'], $detected['yy']);
                        } else {
                            $today = new DateTime();
                            $new_name = sprintf('ค้างซ่อม_%02d-%02d.xlsx',
                                (int)$today->format('m'), ((int)$today->format('Y') + 543) % 100);
                        }

                        $dest_path = $folder_path . DIRECTORY_SEPARATOR . $new_name;
                        $overwritten = file_exists($dest_path);

                        // Move file (keep original Excel, no merge)
                        if (!copy($temp_path, $dest_path)) {
                            throw new Exception('ไม่สามารถบันทึกไฟล์ได้');
                        }
                        chmod($dest_path, 0644);

                        // Clear caches
                        $cache_file = CACHE_DIR . DIRECTORY_SEPARATOR . md5($dest_path) . '.json';
                        if (file_exists($cache_file)) unlink($cache_file);
                        $py_cache = $dest_path . '.cache.json';
                        if (file_exists($py_cache)) unlink($py_cache);

                        // Clear combined FY SQLite so it gets rebuilt with new data
                        $det_mm = $detected ? $detected['mm'] : (int)(new DateTime())->format('m');
                        $det_yy = $detected ? $detected['yy'] : (((int)(new DateTime())->format('Y') + 543) % 100);
                        $det_fy = ($det_mm >= 10) ? (2500 + $det_yy + 1) : (2500 + $det_yy);
                        $fy_sqlite = $folder_path . DIRECTORY_SEPARATOR . "ค้างซ่อม_fy_{$det_fy}.sqlite";
                        if (file_exists($fy_sqlite)) unlink($fy_sqlite);

                        $results[] = [
                            'filename' => $new_name,
                            'original' => $fname,
                            'status' => $overwritten ? 'overwrite' : 'success',
                            'message' => $fname . ' → ' . $new_name
                        ];
                    } catch (\Throwable $e) {
                        $errors[] = [
                            'filename' => $fname,
                            'error' => $e->getMessage()
                        ];
                    }
                }
            }
        } else {
            // repair/pressure: move files from temp to final directory
            $files_in_temp = array_diff(scandir($temp_dir), ['.', '..']);

            foreach ($files_in_temp as $fname) {
                $temp_path = $temp_dir . DIRECTORY_SEPARATOR . $fname;
                if (!is_file($temp_path)) continue;

                try {
                    $new_name = generate_new_filename($category, $fname, $temp_path);
                    if (!$new_name) {
                        $errors[] = [
                            'filename' => $fname,
                            'error' => 'Failed to generate filename'
                        ];
                        continue;
                    }

                    $dest_path = $folder_path . DIRECTORY_SEPARATOR . $new_name;
                    $overwritten = file_exists($dest_path);

                    // Move from temp to final directory
                    if (!rename($temp_path, $dest_path)) {
                        throw new Exception('Failed to move file to final directory');
                    }
                    chmod($dest_path, 0644);

                    // Clear caches
                    $cache_file = CACHE_DIR . DIRECTORY_SEPARATOR . md5($dest_path) . '.json';
                    if (file_exists($cache_file)) unlink($cache_file);
                    $py_cache = $dest_path . '.cache.json';
                    if (file_exists($py_cache)) unlink($py_cache);

                    $results[] = [
                        'filename' => $new_name,
                        'original' => $fname,
                        'status' => $overwritten ? 'overwrite' : 'success',
                        'message' => $fname . ' → ' . $new_name
                    ];
                } catch (\Throwable $e) {
                    $errors[] = [
                        'filename' => $fname,
                        'error' => $e->getMessage()
                    ];
                }
            }
        }
    } catch (\Throwable $e) {
        $errors[] = ['filename' => 'general', 'error' => $e->getMessage()];
    } finally {
        // Cleanup temp directory
        cleanup_temp_dir($batch_id);
    }

    // Write upload log if there are successful uploads
    if (!empty($results)) {
        $log_file = __DIR__ . DIRECTORY_SEPARATOR . 'upload_log.json';
        $log = file_exists($log_file) ? (json_decode(file_get_contents($log_file), true) ?: []) : [];
        $log[] = [
            'time' => date('Y-m-d H:i:s'),
            'category' => $category,
            'files' => array_map(function($r) { return $r['filename']; }, $results),
            'count' => count($results)
        ];
        // Keep last 200 entries
        if (count($log) > 200) $log = array_slice($log, -200);
        file_put_contents($log_file, json_encode($log, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT));
    }

    json_response([
        'ok' => true,
        'category' => $category,
        'results' => $results,
        'errors' => $errors
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

            $prefix = isset($PREFIX_MAP[$category]) ? $PREFIX_MAP[$category] : strtoupper($category);
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
                // Extract branch name from filename — กรองคำที่ไม่ใช่ชื่อสาขาออก
                preg_match_all('/[\x{0e00}-\x{0e7f}]+/u', $name_only, $matches);
                $non_branch_words = ['ปีงบ', 'ปีงบประมาณ', 'แรงดัน', 'แรงดันน้ำ'];
                $branch_candidates = array_filter($matches[0] ?? [], function($w) use ($non_branch_words) {
                    foreach ($non_branch_words as $nbw) {
                        if (mb_strpos($w, $nbw) !== false) return false;
                    }
                    return true;
                });
                $branch_name = !empty($branch_candidates) ? reset($branch_candidates) : 'unknown';

                // Read fiscal year: try filename first, then Excel content
                $fiscal_year = '';

                // Try to extract from filename (pattern: _ปีงบYY)
                if (preg_match('/_ปีงบ(\d{2})/', $name_only, $m)) {
                    $fiscal_year = '25' . $m[1];
                }

                // If not found in filename, read from Excel content
                if (!$fiscal_year) {
                    try {
                        if (class_exists('PhpOffice\\PhpSpreadsheet\\IOFactory')) {
                            $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($files['tmp_name'][$i]);
                            $worksheet = $spreadsheet->getActiveSheet();

                            // Look for "ปีงบประมาณ XXXX" in first few rows
                            for ($r = 1; $r <= min(6, $worksheet->getHighestRow()); $r++) {
                                for ($c = 1; $c <= min(6, $worksheet->getHighestColumn()); $c++) {
                                    $cell_val = (string)($worksheet->getCell([$c, $r])->getValue() ?: '');
                                    if (preg_match('/ปีงบประมาณ\s*(\d{4})/', $cell_val, $m)) {
                                        $fiscal_year = $m[1];
                                        break 2;
                                    }
                                }
                            }

                            // ╔══════════════════════════════════════════════════════════════╗
                            // ║ ⚠️  [FISCAL YEAR FALLBACK] SHEET NAME SCAN                 ║
                            // ║                                                              ║
                            // ║ ตรวจสอบชื่อชีท หากยังไม่พบปีงบประมาณ                       ║
                            // ║ ค้นหา 2 หลักปีจากชื่อ เช่น "ก.พ. 69" → "2569"             ║
                            // ║ ข้ามชีทสรุป/กราฟ — เฉพาะชีทข้อมูลที่มีปีในชื่อ             ║
                            // ║                                                              ║
                            // ║ Sheets to AVOID: "กราฟ", "สรุป", "รวม", "Chart", "Summary" ║
                            // ║ Sheets to PROCESS: Sheets with year suffix (2 digits)       ║
                            // ╚══════════════════════════════════════════════════════════════╝
                            // If still not found, try from sheet names — majority vote
                            if (!$fiscal_year) {
                                $yc = [];
                                foreach ($spreadsheet->getSheetNames() as $sname) {
                                    if (preg_match('/(\d{2})\s*$/', trim($sname), $sm)) {
                                        $yy = $sm[1];
                                        $yc[$yy] = ($yc[$yy] ?? 0) + 1;
                                    }
                                }
                                if (!empty($yc)) {
                                    arsort($yc);
                                    $fiscal_year = '25' . array_key_first($yc);
                                }
                            }

                            $spreadsheet->disconnectWorksheets();
                        }
                    } catch (\Throwable $e) {
                        // Continue without fiscal year
                    }
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
        } catch (\Throwable $e) {
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
                                $row_data[] = $worksheet->getCell([$c, $r])->getValue();
                            }
                            $header_rows[] = $row_data;
                        }
                    }

                    // Capture data rows
                    for ($r = $HEADER_ROWS + 1; $r <= $worksheet->getHighestRow(); $r++) {
                        $row_data = [];
                        for ($c = 1; $c <= $worksheet->getHighestColumn(); $c++) {
                            $row_data[] = $worksheet->getCell([$c, $r])->getValue();
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
        } catch (\Throwable $e) {
            $errors[] = ['filename' => 'pending-merge', 'error' => $e->getMessage()];
        }
    }

    // Write upload log — who uploaded what and when
    if (!empty($results)) {
        $log_file = __DIR__ . DIRECTORY_SEPARATOR . 'upload_log.json';
        $log = file_exists($log_file) ? (json_decode(file_get_contents($log_file), true) ?: []) : [];
        $log[] = [
            'time' => date('Y-m-d H:i:s'),
            'category' => $category,
            'files' => array_map(function($r) { return $r['filename']; }, $results),
            'count' => count($results)
        ];
        // Keep last 200 entries
        if (count($log) > 200) $log = array_slice($log, -200);
        file_put_contents($log_file, json_encode($log, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT));
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
    } catch (\Throwable $e) {
        json_response(['ok' => false, 'error' => $e->getMessage()], 500);
    }
}

// Route: POST /api/notes/<slug>
// Accepts category slugs (e.g. 'repair') and derived keys (e.g. 'repair_source_url')
if ($method === 'POST' && count($path_parts) === 2 && $path_parts[0] === 'notes') {
    $slug = $path_parts[1];

    // Validate: must be a known category OR a derived key like {category}_source_url
    $base_slug = preg_replace('/_source_url$/', '', $slug);
    if (!isset(CATEGORY_MAP[$base_slug])) {
        json_response(['ok' => false, 'error' => 'invalid slug'], 400);
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

    // Parse each file
    foreach ($month_files as $month_key => $info) {
        $fpath = $info[1];
        try {
            $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($fpath);
            // ╔══════════════════════════════════════════════════════════════╗
            // ║ [DATA PARSING] SCAN ALL SHEETS                              ║
            // ║                                                              ║
            // ║ อ่านข้อมูลซ่อมท่อ — สแกนทุกชีท หาชีทที่มี header ถูกต้อง   ║
            // ║ ข้ามชีทสรุป/กราฟ                                            ║
            // ╚══════════════════════════════════════════════════════════════╝
            $gis_skip = ['สารบัญ', 'สรุป', 'ปก', 'cover', 'summary', 'chart', 'graph', 'index'];
            $worksheet = null;
            $det = null;
            for ($si = 0; $si < $spreadsheet->getSheetCount(); $si++) {
                $ws = $spreadsheet->getSheet($si);
                $wsName = mb_strtolower($ws->getTitle());
                $skip = false;
                foreach ($gis_skip as $sk) {
                    if (mb_strpos($wsName, $sk) !== false) { $skip = true; break; }
                }
                if ($skip) continue;
                $det = detect_repair_columns($ws);
                if (!$det['fallback']) { $worksheet = $ws; break; }
            }
            if ($worksheet === null) {
                // Fallback: use first sheet
                $worksheet = $spreadsheet->getSheet(0);
                $det = detect_repair_columns($worksheet);
            }
            $max_row = $worksheet->getHighestRow();
            $month_data = [];
            $hdr = $det['header_row'];
            $cBranch   = $det['cols']['branch']   ?? 1;
            $cClosed   = $det['cols']['closed']    ?? 2;
            $cComplete = $det['cols']['complete']  ?? 3;
            $cScore    = $det['cols']['score']     ?? 4;

            for ($r = $hdr + 1; $r <= $max_row; $r++) {
                $branch = $worksheet->getCell([$cBranch, $r])->getValue();
                if ($branch === null || !is_string($branch)) continue;
                $branch = trim($branch);
                if ($branch === '' || mb_strpos($branch, 'ชื่อสาขา') !== false) continue;

                $closed_v = $worksheet->getCell([$cClosed, $r])->getValue();
                $complete_v = $worksheet->getCell([$cComplete, $r])->getValue();
                $score_v = isset($det['cols']['score']) ? $worksheet->getCell([$cScore, $r])->getValue() : 0;

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
        } catch (\Throwable $e) {
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

// ───────────────────────────────────────────────────────────────────────────
// Route: POST /api/rebuild — run build_dashboard.php to re-embed data into index.html
// ───────────────────────────────────────────────────────────────────────────
if ($method === 'POST' && count($path_parts) === 1 && $path_parts[0] === 'rebuild') {
    $body = json_decode(file_get_contents('php://input'), true) ?: [];
    $only = isset($body['only']) ? preg_replace('/[^a-z0-9]/', '', $body['only']) : '';
    $files_arg = '';
    if (!empty($body['files']) && is_array($body['files'])) {
        $safe_files = [];
        foreach ($body['files'] as $f) {
            $f = basename($f);
            if (preg_match('/^[a-zA-Z0-9_\-\.\x{0E00}-\x{0E7F}]+$/u', $f)) {
                $safe_files[] = $f;
            }
        }
        if (!empty($safe_files)) {
            $files_arg = ' --files=' . implode(',', $safe_files);
        }
    }

    // Clear API cache first so build reads fresh data (selective: only for the category being rebuilt)
    $cache_dir = __DIR__ . DIRECTORY_SEPARATOR . '.cache';
    if (is_dir($cache_dir)) {
        $cache_pattern = !empty($only) ? $only . '_*.json' : '*_*.json';
        foreach (glob($cache_dir . '/' . $cache_pattern) as $cf) { @unlink($cf); }
    }

    $script = __DIR__ . DIRECTORY_SEPARATOR . 'build_dashboard.php';
    if (!file_exists($script)) {
        json_response(['ok' => false, 'message' => 'build_dashboard.php not found'], 500);
    }

    // Find php.exe CLI path
    $php = null;
    $ini = php_ini_loaded_file();
    if ($ini) {
        $candidate = dirname($ini) . DIRECTORY_SEPARATOR . 'php.exe';
        if (file_exists($candidate)) $php = $candidate;
    }
    if (!$php) {
        foreach (['C:\\xampp\\php\\php.exe', 'D:\\xampp\\php\\php.exe', PHP_BINDIR . '\\php.exe'] as $p) {
            if (file_exists($p)) { $php = $p; break; }
        }
    }
    if (!$php) {
        $where_out = [];
        @exec('where php.exe 2>NUL', $where_out);
        if (!empty($where_out) && file_exists(trim($where_out[0]))) {
            $php = trim($where_out[0]);
        }
    }
    if (!$php) {
        json_response(['ok' => false, 'message' => 'php.exe not found'], 500);
    }
    // Ensure required extensions are loaded (zip is needed for .xlsx via PhpSpreadsheet)
    $ext_flags = '';
    if (!extension_loaded('zip')) {
        $ext_dir = dirname($php) . DIRECTORY_SEPARATOR . 'ext';
        if (is_dir($ext_dir)) {
            $ext_flags = ' -d extension_dir="' . $ext_dir . '" -d extension=php_zip.dll';
        }
    }
    $cmd = '"' . $php . '" -d memory_limit=512M' . $ext_flags . ' "' . $script . '"' . ($only ? ' --only=' . $only : '') . $files_arg . ' 2>&1';
    $output = [];
    $exitCode = -1;

    set_time_limit(600);
    ini_set('memory_limit', '512M');
    exec($cmd, $output, $exitCode);

    if ($exitCode === 0) {
        json_response([
            'ok' => true,
            'message' => 'Dashboard rebuilt successfully',
            'log' => implode("\n", $output)
        ]);
    } else {
        json_response([
            'ok' => false,
            'message' => 'Build failed (exit code ' . $exitCode . ')',
            'log' => implode("\n", $output)
        ], 500);
    }
}

// 404 - Route not found
json_response([
    'ok' => false,
    'error' => 'Route not found: ' . $method . ' ' . $path_info
], 404);
