<?php
/**
 * Dashboard PR — XAMPP (Apache + PHP) Backend API
 * ========================================================================
 * PHP equivalent of Flask server.py for Dashboard PR (GUI-019 reports)
 * Supports PR, AON (Always-On), and combined data
 *
 * Features:
 *   - Multi-category support: PR (GUI-019), AON (Always-On), combined
 *   - Complex Excel parsing: GUI-019 format with categories
 *   - Always-On data parsing with month detection
 *   - Auto-detection of month from filenames and file content
 *   - Thai month abbreviation mapping
 *   - Branch name normalization
 *   - Data stored in data.json with specific structure
 *   - Write-back capability to Excel files
 *   - Notes system
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
define('PR_DIR', RAW_DATA_DIR . DIRECTORY_SEPARATOR . 'เรื่องร้องเรียน');
define('AON_DIR', RAW_DATA_DIR . DIRECTORY_SEPARATOR . 'AlwayON');
define('DATA_FILE', RAW_DATA_DIR . DIRECTORY_SEPARATOR . 'data.json');

// Thai month mappings
const THAI_MONTH_ABBR = [
    'ม.ค.' => 1, 'ก.พ.' => 2, 'มี.ค.' => 3, 'เม.ย.' => 4, 'พ.ค.' => 5, 'มิ.ย.' => 6,
    'ก.ค.' => 7, 'ส.ค.' => 8, 'ก.ย.' => 9, 'ต.ค.' => 10, 'พ.ย.' => 11, 'ธ.ค.' => 12,
];

const THAI_MONTH_FULL = [
    'มกราคม' => 1, 'กุมภาพันธ์' => 2, 'มีนาคม' => 3, 'เมษายน' => 4,
    'พฤษภาคม' => 5, 'มิถุนายน' => 6, 'กรกฎาคม' => 7, 'สิงหาคม' => 8,
    'กันยายน' => 9, 'ตุลาคม' => 10, 'พฤศจิกายน' => 11, 'ธันวาคม' => 12,
];

const TH_MONTHS = [
    '', 'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
    'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
];

// Branch name normalization
const BRANCH_NAME_MAP = [
    'สาขาชลบุรี' => 'ชลบุรี',
    'สาขาบ้านบึง' => 'บ้านบึง',
    'สาขาพนัสนิคม' => 'พนัสนิคม',
    'สาขาศรีราชา' => 'ศรีราชา',
    'สาขาแหลมฉบัง' => 'แหลมฉบัง',
    'สาขาพัทยา' => 'พัทยา',
    'สาขาฉะเชิงเทรา' => 'ฉะเชิงเทรา',
    'สาขาบางปะกง' => 'บางปะกง',
    'สาขาบางคล้า' => 'บางคล้า',
    'สาขาพนมสารคาม' => 'พนมสารคาม',
    'สาขาระยอง' => 'ระยอง',
    'สาขาบ้านฉาง' => 'บ้านฉาง',
    'สาขาปากน้ำประแสร์' => 'ปากน้ำประแสร์',
    'สาขาจันทบุรี' => 'จันทบุรี',
    'สาขาขลุง' => 'ขลุง',
    'สาขาตราด' => 'ตราด',
    'สาขาคลองใหญ่' => 'คลองใหญ่',
    'สาขาสระแก้ว' => 'สระแก้ว',
    'สาขาวัฒนานคร' => 'วัฒนานคร',
    'สาขาอรัญประเทศ' => 'อรัญประเทศ',
    'สาขาปราจีนบุรี' => 'ปราจีนบุรี',
    'สาขากบินทร์บุรี' => 'กบินทร์บุรี',
    'ชลบุรี(พ)' => 'ชลบุรี',
    'พัทยา(พ)' => 'พัทยา',
    'ชลบุรี(พิเศษ)' => 'ชลบุรี',
    'พัทยา(พิเศษ)' => 'พัทยา',
];

// ─── Cache Setup ──────────────────────────────────────────────────────────
define('CACHE_DIR', BASE_DIR . DIRECTORY_SEPARATOR . '.cache');
define('CACHE_TTL', 60);
if (!is_dir(CACHE_DIR)) { mkdir(CACHE_DIR, 0755, true); }

function get_folder_mtime_cache($folder_path) {
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
    $folder_mtime = get_folder_mtime_cache($folder_path);
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

// PhpSpreadsheet Loader
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

if (!is_dir(RAW_DATA_DIR)) {
    mkdir(RAW_DATA_DIR, 0755, true);
}
if (!is_dir(PR_DIR)) {
    mkdir(PR_DIR, 0755, true);
}
if (!is_dir(AON_DIR)) {
    mkdir(AON_DIR, 0755, true);
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
            'pr' => [],
            'aon' => [],
            'pr_cat_names' => [],
            'pr_files' => [],
            'aon_files' => [],
            'notes' => []
        ];
    }

    try {
        $content = file_get_contents(DATA_FILE);
        $data = json_decode($content, true);

        if (!isset($data['pr'])) $data['pr'] = [];
        if (!isset($data['aon'])) $data['aon'] = [];
        if (!isset($data['pr_cat_names'])) $data['pr_cat_names'] = [];
        if (!isset($data['pr_files'])) $data['pr_files'] = [];
        if (!isset($data['aon_files'])) $data['aon_files'] = [];
        if (!isset($data['notes'])) $data['notes'] = [];

        return $data;
    } catch (Exception $e) {
        error_log("Error loading data.json: " . $e->getMessage());
        return [
            'pr' => [],
            'aon' => [],
            'pr_cat_names' => [],
            'pr_files' => [],
            'aon_files' => [],
            'notes' => []
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
 * Parse number value
 */
function parse_num($val) {
    if ($val === null) {
        return 0;
    }
    if (is_numeric($val)) {
        return floatval($val);
    }
    $s = str_replace(',', '', trim((string)$val));
    if ($s === '') return 0;
    try {
        return floatval($s);
    } catch (Exception $e) {
        return 0;
    }
}

/**
 * Normalize branch name
 */
function norm_branch($name) {
    $n = trim((string)$name);
    if (isset(BRANCH_NAME_MAP[$n])) {
        return BRANCH_NAME_MAP[$n];
    }
    if (strpos($n, 'สาขา') === 0) {
        return substr($n, strlen('สาขา'));
    }
    return $n;
}

/**
 * Make month key from year and month
 */
function make_month_key($yyyy, $mm) {
    $yy = $yyyy - 2500;
    return sprintf("%02d-%02d", $yy, $mm);
}

/**
 * Detect month from filename (YY-MM pattern)
 */
function detect_month_from_filename($filename) {
    if (preg_match('/(\d{2})-(\d{2})/', $filename, $m)) {
        $yy = intval($m[1]);
        $mm = intval($m[2]);
        if ($mm >= 1 && $mm <= 12) {
            return sprintf("%02d-%02d", $yy, $mm);
        }
    }
    return null;
}

/**
 * Detect month from text content
 */
function detect_month_from_text($text) {
    if (!$text) return null;

    $s = (string)$text;
    $month_num = null;
    $year_num = null;

    // Try full Thai month names
    foreach (THAI_MONTH_FULL as $name => $num) {
        if (strpos($s, $name) !== false) {
            $month_num = $num;
            break;
        }
    }

    // Try Thai month abbreviations
    if ($month_num === null) {
        foreach (THAI_MONTH_ABBR as $abbr => $num) {
            if (strpos($s, $abbr) !== false) {
                $month_num = $num;
                break;
            }
        }
    }

    // Find year (4 digits)
    if (preg_match('/(\d{4})/', $s, $m)) {
        $year_num = intval($m[1]);
    }

    // Find year (2 digits) if not found
    if ($year_num === null && preg_match('/(\d{2})/', $s, $m)) {
        $yy = intval($m[1]);
        $year_num = $yy + 2500;
    }

    if ($month_num && $year_num) {
        return make_month_key($year_num, $month_num);
    }
    return null;
}

/**
 * Sheet name to month key (e.g., 'ต.ค.68' → '68-10')
 */
function sheet_name_to_month_key($name) {
    $s = trim((string)$name);
    foreach (THAI_MONTH_ABBR as $abbr => $mm) {
        if (strpos($s, $abbr) !== false) {
            if (preg_match('/(\d{2})/', $s, $m)) {
                $yy = intval($m[1]);
                return sprintf("%02d-%02d", $yy, $mm);
            }
        }
    }
    return null;
}

/**
 * Parse always-on header
 */
function parse_always_on_header($header) {
    $s = trim(strtolower((string)$header));
    if (strpos($s, 'always on') === false && strpos($s, 'always-on') === false) {
        return null;
    }
    return detect_month_from_text((string)$header);
}

/**
 * Get folder last modified time
 */
function folder_last_modified($folder) {
    $latest = null;
    if (is_dir($folder)) {
        $files = scandir($folder);
        foreach ($files as $f) {
            if ($f[0] === '.') continue;
            $fp = $folder . DIRECTORY_SEPARATOR . $f;
            if (is_file($fp)) {
                $mtime = filemtime($fp);
                if ($latest === null || $mtime > $latest) {
                    $latest = $mtime;
                }
            }
        }
    }
    if ($latest !== null) {
        $dt = new DateTime('@' . $latest);
        $dt->setTimezone(new DateTimeZone('Asia/Bangkok'));
        return $dt->format('d/m/Y H:i');
    }
    return null;
}

// ─── Excel Reading ─────────────────────────────────────────────────────────

/**
 * Read Excel file and return sheets
 */
function read_excel_sheets($filepath) {
    $ext = strtolower(pathinfo($filepath, PATHINFO_EXTENSION));

    if (!in_array($ext, ['xlsx', 'xlsm', 'xls'])) {
        throw new Exception("Unsupported file extension: " . $ext);
    }

    if (!class_exists('PhpOffice\\PhpSpreadsheet\\IOFactory')) {
        throw new Exception("PhpSpreadsheet not available");
    }

    try {
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($filepath);
        $sheets = [];

        foreach ($spreadsheet->getSheetNames() as $name) {
            $worksheet = $spreadsheet->getSheetByName($name);
            $rows = [];

            foreach ($worksheet->getRowIterator() as $row) {
                $cellIterator = $row->getCellIterator();
                $cellIterator->setIterateOnlyExistingCells(false);
                $rowData = [];
                foreach ($cellIterator as $cell) {
                    $rowData[] = $cell->getValue();
                }
                $rows[] = $rowData;
            }

            $sheets[] = ['name' => $name, 'rows' => $rows];
        }

        $spreadsheet->disconnectWorksheets();
        unset($spreadsheet);

        return $sheets;
    } catch (Exception $e) {
        throw new Exception("Failed to read Excel: " . $e->getMessage());
    }
}

// ─── PR Parser (GUI_019 Format) ────────────────────────────────────────────

/**
 * Parse PR file (GUI_019 format)
 */
function parse_pr_file($filepath, $filename, $manual_mk = null) {
    $sheets = read_excel_sheets($filepath);
    if (empty($sheets)) {
        throw new Exception("[{$filename}] ไฟล์ว่างเปล่า");
    }

    $rows = $sheets[0]['rows'];
    if (count($rows) < 7) {
        throw new Exception("[{$filename}] ไฟล์มีข้อมูลไม่เพียงพอ (" . count($rows) . " แถว)");
    }

    // --- Detect month ---
    $mk = $manual_mk;
    if (!$mk) {
        $mk = detect_month_from_filename($filename);
    }
    if (!$mk && count($rows) > 2) {
        foreach ($rows[2] as $cell) {
            $mk = detect_month_from_text($cell);
            if ($mk) break;
        }
    }
    if (!$mk) {
        throw new Exception("[{$filename}] ไม่สามารถตรวจจับเดือนได้");
    }

    // --- Smart Header Detection: ค้นหาแถว header ที่มีหมวดหมู่ (เลข.) ---
    $header_idx = 4; // default
    for ($ri = 0; $ri < min(15, count($rows)); $ri++) {
        if (!isset($rows[$ri])) continue;
        foreach ($rows[$ri] as $cell) {
            $cv = trim((string)($cell ?? ''));
            if (preg_match('/^\d+\.\s*ด้าน/', $cv)) {
                $header_idx = $ri;
                break 2;
            }
        }
    }
    $header_row = isset($rows[$header_idx]) ? $rows[$header_idx] : [];

    // ค้นหาคอลัมน์ สาขา / จำนวนลูกค้า จาก header หรือ sub-header
    $branch_col = 1; // default
    $cust_col = 2;   // default
    $cat_start_col = 4; // default
    for ($ri2 = max(0, $header_idx - 2); $ri2 <= $header_idx; $ri2++) {
        if (!isset($rows[$ri2])) continue;
        for ($ci = 0; $ci < count($rows[$ri2]); $ci++) {
            $hv = mb_strtolower(trim((string)($rows[$ri2][$ci] ?? '')));
            if (mb_strpos($hv, 'สาขา') !== false && mb_strpos($hv, 'รวมสาขา') === false) $branch_col = $ci;
            if (mb_strpos($hv, 'จำนวนลูกค้า') !== false || mb_strpos($hv, 'ลูกค้า') !== false) $cust_col = $ci;
        }
    }

    // Data rows start after header + 1 (skip sub-header row)
    $data_start_idx = $header_idx + 2;

    $cat_groups = [];
    $c = max($cust_col + 1, $cat_start_col);
    while ($c < count($header_row)) {
        $h = trim((string)(isset($header_row[$c]) ? $header_row[$c] : ''));
        if ($h && $h !== 'รวมสาขา') {
            $cat_name = preg_replace('/^\d+\.\s*/', '', $h);
            $cat_name = trim($cat_name);
            $cat_groups[] = ['name' => $cat_name, 'start_col' => $c];
            $c += 3;
        } elseif ($h === 'รวมสาขา') {
            break;
        } else {
            $c += 1;
        }
    }

    // Find รวมสาขา column
    $total_col = -1;
    for ($c = 4; $c < count($header_row); $c++) {
        if (trim((string)(isset($header_row[$c]) ? $header_row[$c] : '')) === 'รวมสาขา') {
            $total_col = $c;
            break;
        }
    }

    // --- Parse data rows (row index 6+) ---
    $branch_data = [];
    $cat_names = array_map(function($cg) { return $cg['name']; }, $cat_groups);

    for ($r = $data_start_idx; $r < count($rows); $r++) {
        $row = $rows[$r];
        if (empty($row) || count($row) < 5) continue;

        $branch_name = trim((string)(isset($row[$branch_col]) ? $row[$branch_col] : ''));
        if (!$branch_name) continue;

        $is_regional = $branch_name === 'รวม เขต 1';
        if (in_array($branch_name, ['รวมทั้งหมด', 'รวม'])) {
            continue;
        }
        if (!$is_regional) {
            $branch_name = norm_branch($branch_name);
        }
        if (!$branch_name) continue;

        $bd = [];
        $bd['จำนวนลูกค้า'] = parse_num(isset($row[$cust_col]) ? $row[$cust_col] : 0);
        $bd['categories'] = [];

        foreach ($cat_groups as $cg) {
            $sc = $cg['start_col'];
            $bd['categories'][$cg['name']] = [
                'รวม' => parse_num(isset($row[$sc]) ? $row[$sc] : 0),
                'ไม่เกิน' => parse_num(isset($row[$sc + 1]) ? $row[$sc + 1] : 0),
                'เกิน' => parse_num(isset($row[$sc + 2]) ? $row[$sc + 2] : 0),
            ];
        }

        // รวมสาขา
        if ($total_col >= 0 && isset($row[$total_col])) {
            $bd['รวมสาขา'] = parse_num($row[$total_col]);
            $bd['รวม_ไม่เกิน'] = parse_num(isset($row[$total_col + 1]) ? $row[$total_col + 1] : 0);
            $bd['รวม_เกิน'] = parse_num(isset($row[$total_col + 2]) ? $row[$total_col + 2] : 0);
        } else {
            $tot = 0;
            $tot_ne = 0;
            $tot_e = 0;
            foreach ($cat_groups as $cg) {
                $tot += parse_num(isset($row[$cg['start_col']]) ? $row[$cg['start_col']] : 0);
                $tot_ne += parse_num(isset($row[$cg['start_col'] + 1]) ? $row[$cg['start_col'] + 1] : 0);
                $tot_e += parse_num(isset($row[$cg['start_col'] + 2]) ? $row[$cg['start_col'] + 2] : 0);
            }
            $bd['รวมสาขา'] = $tot;
            $bd['รวม_ไม่เกิน'] = $tot_ne;
            $bd['รวม_เกิน'] = $tot_e;
        }

        $branch_data[$branch_name] = $bd;
    }

    return [
        'mk' => $mk,
        'data' => $branch_data,
        'count' => count($branch_data),
        'cat_names' => $cat_names,
    ];
}

// ─── AON Parser (Always-On) ────────────────────────────────────────────────

/**
 * Parse AON sheet with specific column
 */
function parse_aon_sheet_with_col($rows, $aon_col, $month_key) {
    $result = [];

    // --- Smart Header Detection: ค้นหาแถว header ที่มี หน่วยงาน/เขต ---
    $name_col = 4;  // default
    $dist_col = -1;
    $header_row_idx = 3; // default
    $kw_name = ['หน่วยงาน', 'สาขา', 'ชื่อสาขา'];
    $kw_dist = ['เขต', 'district'];

    for ($ri = 0; $ri < min(10, count($rows)); $ri++) {
        if (!isset($rows[$ri])) continue;
        for ($c = 0; $c < count($rows[$ri]); $c++) {
            $h = mb_strtolower(trim((string)($rows[$ri][$c] ?? '')));
            if ($h === '') continue;
            foreach ($kw_name as $kw) {
                if (mb_strpos($h, mb_strtolower($kw)) !== false) {
                    $name_col = $c;
                    $header_row_idx = $ri;
                    break;
                }
            }
            foreach ($kw_dist as $kw) {
                if (mb_strpos($h, mb_strtolower($kw)) !== false) {
                    $dist_col = $c;
                    break;
                }
            }
        }
        // Stop if we found name_col in this row
        if ($header_row_idx === $ri) break;
    }

    $data_start = $header_row_idx + 2; // skip header + sub-header

    // Read data rows
    for ($r = $data_start; $r < count($rows); $r++) {
        $row = $rows[$r];
        if (empty($row) || !isset($row[$aon_col])) continue;

        // Filter only เขต 1
        if ($dist_col >= 0) {
            $dist = parse_num(isset($row[$dist_col]) ? $row[$dist_col] : 0);
            if ($dist != 1) continue;
        }

        $raw_name = trim((string)(isset($row[$name_col]) ? $row[$name_col] : ''));
        if (!$raw_name) continue;

        $branch_name = norm_branch($raw_name);
        if (!$branch_name || strpos($branch_name, 'รวม') === 0) continue;

        $val = $row[$aon_col];
        if ($val === null || $val === '') continue;

        try {
            $num_val = floatval($val);
        } catch (Exception $e) {
            continue;
        }

        // Convert 0-1 to percentage
        if ($num_val <= 1.5) {
            $num_val = round($num_val * 10000) / 100;
        }

        $result[$branch_name] = $num_val;
    }

    return $result;
}

/**
 * Parse AON file
 */
function parse_aon_file($filepath, $filename, $mode = 'auto', $manual_mk = null) {
    $sheets = read_excel_sheets($filepath);
    if (empty($sheets)) {
        throw new Exception("[{$filename}] ไฟล์ว่างเปล่า");
    }

    $all_data = [];
    $total_count = 0;
    $processed_months = [];

    if ($mode === 'manual' && $manual_mk) {
        $rows = $sheets[0]['rows'];
        if (count($rows) < 6) {
            throw new Exception("[{$filename}] ข้อมูลไม่เพียงพอ");
        }

        $sub_header = isset($rows[4]) ? $rows[4] : [];
        $aon_col = -1;
        for ($c = 0; $c < count($sub_header); $c++) {
            if (parse_always_on_header(isset($sub_header[$c]) ? $sub_header[$c] : '')) {
                $aon_col = $c;
                break;
            }
        }

        if ($aon_col < 0) {
            throw new Exception("[{$filename}] ไม่พบคอลัมน์ always on");
        }

        $data = parse_aon_sheet_with_col($rows, $aon_col, $manual_mk);
        if (!empty($data)) {
            $all_data[$manual_mk] = $data;
            $total_count += count($data);
            $processed_months[] = $manual_mk;
        }
    } else {
        // Auto mode: scan all sheets
        foreach ($sheets as $sheet) {
            $rows = $sheet['rows'];
            if (count($rows) < 6) continue;

            $sub_header = isset($rows[4]) ? $rows[4] : [];
            $aon_cols = [];
            for ($c = 0; $c < count($sub_header); $c++) {
                $mk2 = parse_always_on_header(isset($sub_header[$c]) ? $sub_header[$c] : '');
                if ($mk2) {
                    $aon_cols[] = ['col' => $c, 'mk' => $mk2];
                }
            }

            if (empty($aon_cols)) continue;

            $sheet_mk = sheet_name_to_month_key($sheet['name']);
            $best_col = null;
            foreach ($aon_cols as $ac) {
                if ($ac['mk'] === $sheet_mk) {
                    $best_col = $ac;
                    break;
                }
            }
            if (!$best_col) {
                $best_col = $aon_cols[0];
            }

            $data = parse_aon_sheet_with_col($rows, $best_col['col'], $best_col['mk']);
            if (!empty($data)) {
                $all_data[$best_col['mk']] = $data;
                $total_count += count($data);
                if (!in_array($best_col['mk'], $processed_months)) {
                    $processed_months[] = $best_col['mk'];
                }
            }
        }
    }

    return [
        'months' => $all_data,
        'count' => $total_count,
        'processed_months' => $processed_months ? array_values($processed_months) : [],
    ];
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
$path_parts = array_values(array_filter(explode('/', $path_info), function($p) { return $p !== ''; }));
if (count($path_parts) > 0 && $path_parts[0] === 'api') { array_shift($path_parts); $path_parts = array_values($path_parts); }

// Route: GET /api/ping
if ($method === 'GET' && count($path_parts) === 1 && $path_parts[0] === 'ping') {
    json_response([
        'ok' => true,
        'version' => '1.0',
        'timestamp' => (new DateTime('now', new DateTimeZone('UTC')))->format('c')
    ]);
}

// Route: GET /api/data
if ($method === 'GET' && count($path_parts) === 1 && $path_parts[0] === 'data') {
    $data = load_data();

    // Load notes
    $notes = isset($data['notes']) ? $data['notes'] : [];

    json_response([
        'ok' => true,
        'pr' => isset($data['pr']) ? $data['pr'] : [],
        'aon' => isset($data['aon']) ? $data['aon'] : [],
        'pr_cat_names' => isset($data['pr_cat_names']) ? $data['pr_cat_names'] : [],
        'pr_files' => isset($data['pr_files']) ? $data['pr_files'] : [],
        'aon_files' => isset($data['aon_files']) ? $data['aon_files'] : [],
        'pr_last_modified' => folder_last_modified(PR_DIR),
        'aon_last_modified' => folder_last_modified(AON_DIR),
        'notes' => $notes,
    ]);
}

// ── Validate file format before upload ──
function validate_pr_file($tmp_path) {
    try {
        $rows = read_excel_first_sheet($tmp_path);
        if (empty($rows) || count($rows) < 3) {
            return ['valid' => false, 'message' => 'ไฟล์ว่างเปล่าหรือมีข้อมูลน้อยเกินไป'];
        }
        // ค้นหา pattern ของ PR: ต้องมี "ด้าน" (เช่น "1. ด้านคุณภาพน้ำ") ในแถวใดแถวหนึ่ง
        $found_category = false;
        $found_branch = false;
        for ($ri = 0; $ri < min(15, count($rows)); $ri++) {
            if (!isset($rows[$ri])) continue;
            foreach ($rows[$ri] as $cell) {
                $cv = trim((string)($cell ?? ''));
                if (preg_match('/^\d+\.\s*ด้าน/', $cv)) $found_category = true;
                if (mb_strpos(mb_strtolower($cv), 'สาขา') !== false) $found_branch = true;
            }
        }
        if (!$found_category) {
            return ['valid' => false, 'message' => 'ไม่พบหมวดหมู่ PR (เช่น "1. ด้านคุณภาพน้ำ") — รูปแบบไฟล์ไม่ตรงกับที่ระบบรองรับ'];
        }
        return ['valid' => true, 'message' => ''];
    } catch (Exception $e) {
        return ['valid' => false, 'message' => 'ไม่สามารถอ่านไฟล์ Excel ได้: ' . $e->getMessage()];
    }
}

function validate_aon_file($tmp_path) {
    try {
        $sheets = read_excel_sheets($tmp_path);
        if (empty($sheets)) {
            return ['valid' => false, 'message' => 'ไม่สามารถอ่าน sheet ในไฟล์ได้'];
        }
        // ค้นหาว่ามี sheet ที่มีคอลัมน์ "หน่วยงาน" หรือ "สาขา" และ "%" หรือ "อันดับ"
        $found_name = false;
        $found_data = false;
        foreach ($sheets as $sname => $rows) {
            for ($ri = 0; $ri < min(10, count($rows)); $ri++) {
                if (!isset($rows[$ri])) continue;
                foreach ($rows[$ri] as $cell) {
                    $cv = mb_strtolower(trim((string)($cell ?? '')));
                    if (mb_strpos($cv, 'หน่วยงาน') !== false || mb_strpos($cv, 'สาขา') !== false) $found_name = true;
                    if (mb_strpos($cv, '%') !== false || mb_strpos($cv, 'อันดับ') !== false || mb_strpos($cv, 'เป้า') !== false) $found_data = true;
                }
            }
            if ($found_name) break;
        }
        if (!$found_name) {
            return ['valid' => false, 'message' => 'ไม่พบคอลัมน์ "หน่วยงาน/สาขา" ในไฟล์ — รูปแบบไฟล์ไม่ตรงกับที่ระบบรองรับ'];
        }
        return ['valid' => true, 'message' => ''];
    } catch (Exception $e) {
        return ['valid' => false, 'message' => 'ไม่สามารถอ่านไฟล์ Excel ได้: ' . $e->getMessage()];
    }
}

// Route: POST /api/upload/pr
if ($method === 'POST' && count($path_parts) === 2 && $path_parts[0] === 'upload' && $path_parts[1] === 'pr') {
    if (!isset($_FILES['files'])) {
        json_response(['ok' => false, 'error' => 'ไม่ได้เลือกไฟล์'], 400);
    }

    $files = $_FILES['files'];
    $mode = isset($_POST['mode']) ? $_POST['mode'] : 'auto';
    $manual_mk = isset($_POST['manualMK']) ? $_POST['manualMK'] : null;

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

    if (!is_dir(PR_DIR)) {
        mkdir(PR_DIR, 0755, true);
    }

    $data = load_data();
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

        $temp_path = RAW_DATA_DIR . DIRECTORY_SEPARATOR . '_temp_' . $filename;
        try {
            if (!move_uploaded_file($files['tmp_name'][$i], $temp_path)) {
                throw new Exception('Failed to move uploaded file');
            }

            // ── ตรวจสอบรูปแบบไฟล์ก่อน parse ──
            $validation = validate_pr_file($temp_path);
            if (!$validation['valid']) {
                if (file_exists($temp_path)) unlink($temp_path);
                $errors[] = [
                    'filename' => $filename,
                    'error' => '⚠️ ' . $validation['message']
                ];
                continue;
            }

            // Parse
            $result = parse_pr_file($temp_path, $filename, $mode === 'manual' ? $manual_mk : null);

            // Auto-rename & store
            $mk = $result['mk'];
            $ext = pathinfo($filename, PATHINFO_EXTENSION) ?: 'xlsx';
            $new_filename = "PR_{$mk}.{$ext}";
            $dest_path = PR_DIR . DIRECTORY_SEPARATOR . $new_filename;

            if (!rename($temp_path, $dest_path)) {
                throw new Exception('Failed to move file to destination');
            }
            chmod($dest_path, 0644);

            // Update data store
            if (!isset($data['pr'][$mk])) {
                $data['pr'][$mk] = [];
            }
            $data['pr'][$mk] = $result['data'];
            $data['pr_files'][$mk] = $new_filename;

            // Update cat_names
            foreach ($result['cat_names'] as $cat) {
                if (!in_array($cat, $data['pr_cat_names'])) {
                    $data['pr_cat_names'][] = $cat;
                }
            }

            $results[] = [
                'mk' => $mk,
                'count' => $result['count'],
                'filename' => $new_filename,
                'cat_names' => $result['cat_names'],
            ];

        } catch (Exception $e) {
            if (file_exists($temp_path)) {
                unlink($temp_path);
            }
            $errors[] = [
                'filename' => $filename,
                'error' => $e->getMessage()
            ];
        }
    }

    save_data($data);

    json_response([
        'ok' => true,
        'results' => $results,
        'errors' => $errors,
        'pr_data' => isset($data['pr']) ? $data['pr'] : [],
        'pr_cat_names' => isset($data['pr_cat_names']) ? $data['pr_cat_names'] : [],
    ]);
}

// Route: POST /api/upload/aon
if ($method === 'POST' && count($path_parts) === 2 && $path_parts[0] === 'upload' && $path_parts[1] === 'aon') {
    if (!isset($_FILES['files'])) {
        json_response(['ok' => false, 'error' => 'ไม่ได้เลือกไฟล์'], 400);
    }

    $files = $_FILES['files'];
    $mode = isset($_POST['mode']) ? $_POST['mode'] : 'auto';
    $manual_mk = isset($_POST['manualMK']) ? $_POST['manualMK'] : null;

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

    if (!is_dir(AON_DIR)) {
        mkdir(AON_DIR, 0755, true);
    }

    $data = load_data();
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

        $temp_path = RAW_DATA_DIR . DIRECTORY_SEPARATOR . '_temp_' . $filename;
        try {
            if (!move_uploaded_file($files['tmp_name'][$i], $temp_path)) {
                throw new Exception('Failed to move uploaded file');
            }

            // ── ตรวจสอบรูปแบบไฟล์ก่อน parse ──
            $validation = validate_aon_file($temp_path);
            if (!$validation['valid']) {
                if (file_exists($temp_path)) unlink($temp_path);
                $errors[] = [
                    'filename' => $filename,
                    'error' => '⚠️ ' . $validation['message']
                ];
                continue;
            }

            // Parse
            $result = parse_aon_file($temp_path, $filename, $mode, $mode === 'manual' ? $manual_mk : null);

            // Auto-rename & store
            $months_str = !empty($result['processed_months']) ? implode('_', $result['processed_months']) : 'unknown';
            $ext = pathinfo($filename, PATHINFO_EXTENSION) ?: 'xlsx';
            $new_filename = "AON_{$months_str}.{$ext}";
            $dest_path = AON_DIR . DIRECTORY_SEPARATOR . $new_filename;

            if (!rename($temp_path, $dest_path)) {
                throw new Exception('Failed to move file to destination');
            }
            chmod($dest_path, 0644);

            // Update data store
            foreach ($result['months'] as $mk => $branch_data) {
                $data['aon'][$mk] = $branch_data;
                $data['aon_files'][$mk] = $new_filename;
            }

            $results[] = [
                'months' => $result['processed_months'],
                'count' => $result['count'],
                'filename' => $new_filename,
            ];

        } catch (Exception $e) {
            if (file_exists($temp_path)) {
                unlink($temp_path);
            }
            $errors[] = [
                'filename' => $filename,
                'error' => $e->getMessage()
            ];
        }
    }

    save_data($data);

    json_response([
        'ok' => true,
        'results' => $results,
        'errors' => $errors,
        'aon_data' => isset($data['aon']) ? $data['aon'] : [],
    ]);
}

// Route: DELETE /api/data/pr/<mk>
if ($method === 'DELETE' && count($path_parts) === 3 && $path_parts[0] === 'data' && $path_parts[1] === 'pr') {
    $mk = $path_parts[2];
    $data = load_data();
    $deleted = false;

    if (isset($data['pr'][$mk])) {
        unset($data['pr'][$mk]);
        $deleted = true;
    }
    if (isset($data['pr_files'][$mk])) {
        $filepath = PR_DIR . DIRECTORY_SEPARATOR . $data['pr_files'][$mk];
        if (file_exists($filepath)) {
            unlink($filepath);
        }
        unset($data['pr_files'][$mk]);
    }

    save_data($data);
    json_response(['ok' => true, 'deleted' => $deleted, 'mk' => $mk]);
}

// Route: DELETE /api/data/aon/<mk>
if ($method === 'DELETE' && count($path_parts) === 3 && $path_parts[0] === 'data' && $path_parts[1] === 'aon') {
    $mk = $path_parts[2];
    $data = load_data();
    $deleted = false;

    if (isset($data['aon'][$mk])) {
        unset($data['aon'][$mk]);
        $deleted = true;
    }
    if (isset($data['aon_files'][$mk])) {
        $filepath = AON_DIR . DIRECTORY_SEPARATOR . $data['aon_files'][$mk];
        if (file_exists($filepath)) {
            unlink($filepath);
        }
        unset($data['aon_files'][$mk]);
    }

    save_data($data);
    json_response(['ok' => true, 'deleted' => $deleted, 'mk' => $mk]);
}

// Route: POST /api/data/edit/pr
if ($method === 'POST' && count($path_parts) === 3 && $path_parts[0] === 'data' && $path_parts[1] === 'edit' && $path_parts[2] === 'pr') {
    $body = json_decode(file_get_contents('php://input'), true) ?: [];
    if (empty($body) || !isset($body['mk']) || !isset($body['data'])) {
        json_response(['ok' => false, 'error' => 'ข้อมูลไม่ครบ (ต้องมี mk และ data)'], 400);
    }

    $mk = $body['mk'];
    $edit_data = $body['data'];

    $data = load_data();
    $data['pr'][$mk] = $edit_data;
    if (!isset($data['pr_files'][$mk])) {
        $data['pr_files'][$mk] = "edited_{$mk}";
    }
    save_data($data);

    json_response(['ok' => true, 'mk' => $mk, 'excel_updated' => false]);
}

// Route: POST /api/data/edit/aon
if ($method === 'POST' && count($path_parts) === 3 && $path_parts[0] === 'data' && $path_parts[1] === 'edit' && $path_parts[2] === 'aon') {
    $body = json_decode(file_get_contents('php://input'), true) ?: [];
    if (empty($body) || !isset($body['mk']) || !isset($body['data'])) {
        json_response(['ok' => false, 'error' => 'ข้อมูลไม่ครบ (ต้องมี mk และ data)'], 400);
    }

    $mk = $body['mk'];
    $edit_data = $body['data'];

    $data = load_data();
    $data['aon'][$mk] = $edit_data;
    if (!isset($data['aon_files'][$mk])) {
        $data['aon_files'][$mk] = "edited_{$mk}";
    }
    save_data($data);

    json_response(['ok' => true, 'mk' => $mk, 'excel_updated' => false]);
}

// Route: POST /api/notes/<slug>
if ($method === 'POST' && count($path_parts) === 2 && $path_parts[0] === 'notes') {
    $slug = $path_parts[1];

    if (!in_array($slug, ['pr', 'aon'])) {
        json_response(['ok' => false, 'error' => 'invalid slug'], 400);
    }

    $body = json_decode(file_get_contents('php://input'), true) ?: [];
    $text = isset($body['text']) ? $body['text'] : '';

    $data = load_data();
    if (!isset($data['notes'])) {
        $data['notes'] = [];
    }
    $data['notes'][$slug] = $text;
    save_data($data);

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

// Route: GET /api/files
if ($method === 'GET' && count($path_parts) === 1 && $path_parts[0] === 'files') {
    $pr_files = is_dir(PR_DIR) ? array_values(array_diff(scandir(PR_DIR), ['.', '..'])) : [];
    $aon_files = is_dir(AON_DIR) ? array_values(array_diff(scandir(AON_DIR), ['.', '..'])) : [];

    // Filter to only include spreadsheet files
    $pr_files = array_filter($pr_files, function($f) {
        $ext = strtolower(pathinfo($f, PATHINFO_EXTENSION));
        return in_array($ext, ['xlsx', 'xls', 'csv']);
    });
    $aon_files = array_filter($aon_files, function($f) {
        $ext = strtolower(pathinfo($f, PATHINFO_EXTENSION));
        return in_array($ext, ['xlsx', 'xls', 'csv']);
    });

    sort($pr_files);
    sort($aon_files);

    $data = load_data();
    json_response([
        'ok' => true,
        'pr_files' => $pr_files,
        'aon_files' => $aon_files,
        'pr_months' => array_keys(isset($data['pr']) ? $data['pr'] : []),
        'aon_months' => array_keys(isset($data['aon']) ? $data['aon'] : []),
    ]);
}

// Route: DELETE /api/data/clear
if ($method === 'DELETE' && count($path_parts) === 2 && $path_parts[0] === 'data' && $path_parts[1] === 'clear') {
    save_data([
        'pr' => [],
        'aon' => [],
        'pr_cat_names' => [],
        'pr_files' => [],
        'aon_files' => [],
        'notes' => []
    ]);

    // Clear files
    foreach ([PR_DIR, AON_DIR] as $folder) {
        if (is_dir($folder)) {
            $files = array_diff(scandir($folder), ['.', '..']);
            foreach ($files as $f) {
                $fp = $folder . DIRECTORY_SEPARATOR . $f;
                if (is_file($fp)) {
                    unlink($fp);
                }
            }
        }
    }

    json_response(['ok' => true, 'message' => 'ล้างข้อมูลทั้งหมดเรียบร้อย']);
}

// ─── Route: GET /api/pr-data (Dual Mode) ──────────────────────────────────
// Parses all PR_YY-MM.xlsx files from ข้อมูลดิบ/เรื่องร้องเรียน/
// Returns same structure as embedded DATA in index.html

if ($method === 'GET' && count($path_parts) === 1 && $path_parts[0] === 'pr-data') {
    if (!$phpSpreadsheet) {
        json_response(['ok' => false, 'error' => 'PhpSpreadsheet not available'], 500);
    }

    if (!is_dir(PR_DIR)) {
        json_response(['ok' => true, 'has_data' => false, 'data' => new stdClass()]);
    }

    $cached = load_cache('pr_data', PR_DIR);
    if ($cached !== null) {
        json_response($cached);
    }

    $files = glob(PR_DIR . DIRECTORY_SEPARATOR . '*.xlsx') ?: [];
    sort($files);

    $all_data = [];
    $branches_order = [];
    $cat_names = [];

    foreach ($files as $file) {
        $filename = basename($file);
        if (strpos($filename, '~$') === 0) continue;

        try {
            $result = parse_pr_file($file, $filename);
            $mk = $result['mk'];
            $all_data[$mk] = $result['data'];
            if (!empty($result['cat_names']) && empty($cat_names)) {
                $cat_names = $result['cat_names'];
            }
            // Collect branch order from first file that has data
            if (empty($branches_order)) {
                foreach ($result['data'] as $bname => $_) {
                    if ($bname !== 'รวม เขต 1' && !in_array($bname, $branches_order)) {
                        $branches_order[] = $bname;
                    }
                }
            }
        } catch (Exception $e) {
            error_log("PR Dual Mode: Cannot parse $filename: " . $e->getMessage());
            continue;
        }
    }

    if (empty($all_data)) {
        json_response(['ok' => true, 'has_data' => false, 'data' => new stdClass()]);
    }

    // Determine 13-month range (same logic as build_dashboard.py)
    $months_sorted = array_keys($all_data);
    sort($months_sorted);
    $latest = end($months_sorted);
    $ly = intval(substr($latest, 0, 2));
    $lm = intval(substr($latest, 3, 2));
    $same_month_ly = sprintf("%02d-%02d", $ly - 1, $lm);
    $months_13 = array_values(array_filter($months_sorted, function($m) use ($same_month_ly) {
        return $m >= $same_month_ly;
    }));

    $data_out = [
        'months' => $months_13,
        'branches' => $branches_order,
        'all_months' => $months_sorted,
        'data' => $all_data,
        'cat_names' => $cat_names,
    ];

    $response = ['ok' => true, 'has_data' => true, 'data' => $data_out];
    save_cache('pr_data', $response);
    json_response($response);
}

// ─── Route: GET /api/aon-data (Dual Mode) ─────────────────────────────────
// Parses all AON_*.xls files from ข้อมูลดิบ/AlwayON/
// Returns AON data: {month_key: {branch: percentage}}

if ($method === 'GET' && count($path_parts) === 1 && $path_parts[0] === 'aon-data') {
    if (!$phpSpreadsheet) {
        json_response(['ok' => false, 'error' => 'PhpSpreadsheet not available'], 500);
    }

    if (!is_dir(AON_DIR)) {
        json_response(['ok' => true, 'has_data' => false, 'data' => new stdClass()]);
    }

    $cached = load_cache('aon_data', AON_DIR);
    if ($cached !== null) {
        json_response($cached);
    }

    $files = array_merge(
        glob(AON_DIR . DIRECTORY_SEPARATOR . '*.xls') ?: [],
        glob(AON_DIR . DIRECTORY_SEPARATOR . '*.xlsx') ?: []
    );
    $files = array_unique($files);
    sort($files);

    $aon_merged = [];

    foreach ($files as $file) {
        $filename = basename($file);
        if (strpos($filename, '~$') === 0) continue;

        try {
            $result = parse_aon_file($file, $filename);
            if (!empty($result['months'])) {
                foreach ($result['months'] as $mk => $branch_data) {
                    // Merge (later files override earlier)
                    if (!isset($aon_merged[$mk])) {
                        $aon_merged[$mk] = [];
                    }
                    foreach ($branch_data as $branch => $val) {
                        $aon_merged[$mk][$branch] = $val;
                    }
                }
            }
        } catch (Exception $e) {
            error_log("AON Dual Mode: Cannot parse $filename: " . $e->getMessage());
            continue;
        }
    }

    $response = ['ok' => true, 'has_data' => !empty($aon_merged), 'data' => $aon_merged];
    save_cache('aon_data', $response);
    json_response($response);
}

// 404 - Route not found
json_response([
    'ok' => false,
    'error' => 'Route not found: ' . $method . ' ' . $path_info
], 404);
