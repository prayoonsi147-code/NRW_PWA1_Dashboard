<?php
/**
 * build_dashboard.php - สร้าง Dashboard แผนที่แนวท่อ (GIS) กปภ.เขต 1
 * PHP CLI equivalent of build_dashboard.py
 * อ่านข้อมูลจาก Excel แล้ว embed ลงใน index.html
 *
 * ข้อมูลที่ฝัง:
 *   TAB 1: const DATA = {...}         — KPI จุดซ่อมท่อ
 *   TAB 2: const PRESSURE_DATA = {...} — แรงดันน้ำ
 *          const PRESSURE_MONTHS = [...]
 *   TAB 3: var PD1_DATA_FALLBACK = {...}  — กราฟค้างซ่อมเปรียบเทียบ
 *          var PD1_MONTHS_FALLBACK = [...]
 *          var PD2_DATA_FALLBACK = {...}  — กราฟค้างซ่อมสะสม
 *          var PD2_MONTHS_FALLBACK = [...]
 *          var PD3_FALLBACK = {...}       — ตารางงานค้างซ่อม
 *          var PENDING_UPDATE_DATE = '...'
 *          var PENDING_FY_LIST = [...]
 *
 * Usage: php.exe build_dashboard.php
 */

// ─── Error Handling ────────────────────────────────────────────────────────
ini_set('display_errors', '0');
error_reporting(E_ALL);
ini_set('log_errors', '1');
ini_set('memory_limit', '1024M');

// ────────────────────────────────────────────────────────────────────────────
// Configuration & Setup
// ────────────────────────────────────────────────────────────────────────────

require __DIR__ . '/../vendor/autoload.php';

$SCRIPT_DIR = __DIR__;
$RAW_DATA_DIR = $SCRIPT_DIR . DIRECTORY_SEPARATOR . 'ข้อมูลดิบ';
$REPAIR_DIR = $RAW_DATA_DIR . DIRECTORY_SEPARATOR . 'ลงข้อมูลซ่อมท่อ';
$PRESSURE_DIR = $RAW_DATA_DIR . DIRECTORY_SEPARATOR . 'แรงดันน้ำ';
$PENDING_DIR = $RAW_DATA_DIR . DIRECTORY_SEPARATOR . 'ซ่อมท่อค้างระบบ';
$HTML_TEMPLATE = $SCRIPT_DIR . DIRECTORY_SEPARATOR . 'index.html';

$MONTH_NAMES = ['', 'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.',
                'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];

$BRANCH_LIST = [
    "ชลบุรี", "พัทยา", "บ้านบึง", "พนัสนิคม", "ศรีราชา", "แหลมฉบัง",
    "ฉะเชิงเทรา", "บางปะกง", "บางคล้า", "พนมสารคาม", "ระยอง", "บ้านฉาง",
    "ปากน้ำประแสร์", "จันทบุรี", "ขลุง", "ตราด", "คลองใหญ่",
    "สระแก้ว", "วัฒนานคร", "อรัญประเทศ", "ปราจีนบุรี", "กบินทร์บุรี"
];

// ────────────────────────────────────────────────────────────────────────────
// Helper Functions
// ────────────────────────────────────────────────────────────────────────────

/**
 * Clean numeric value — handle commas, spaces, non-breaking spaces, dashes
 */
function clean_num($val) {
    if ($val === null) return 0;
    if (is_int($val) || is_float($val)) return $val;

    $s = str_replace([',', "\xa0", ' '], '', (string)$val);
    $s = trim($s);

    if ($s === '' || $s === '-') return 0;

    $num = (float)$s;
    return is_nan($num) ? 0 : $num;
}

/**
 * Parse Thai date (DD/MM/YYYY or DateTime object)
 * Returns: [$dt, $by] or [null, null]
 */
function parse_thai_date($val) {
    if ($val instanceof DateTime) {
        $year = (int)$val->format('Y');
        $by = $year < 2500 ? $year + 543 : $year;
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
 * Smart header detection for repair data
 * Returns: ['header_row' => int, 'cols' => assoc_array, 'fallback' => bool]
 */
function detect_repair_columns($worksheet) {
    $keywords = [
        'branch'   => ['ชื่อสาขา', 'สาขา', 'หน่วยงาน', 'branch'],
        'closed'   => ['ปิดงาน', 'ปิด', 'closed'],
        'complete' => ['สำเร็จ', 'เสร็จสิ้น', 'เสร็จ', 'complete'],
        'score'    => ['คะแนน', 'score', 'ผลคะแนน'],
    ];

    $maxScan = min(10, $worksheet->getHighestRow());
    $maxCol = min(20, \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString(
        $worksheet->getHighestColumn()
    ));

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

    // Fallback: default positions
    return ['header_row' => 1, 'cols' => ['branch' => 1, 'closed' => 2, 'complete' => 3, 'score' => 4], 'fallback' => true];
}

/**
 * Extract year (4-digit) from filename
 * Returns: year string like "2569" or null if not found
 */
function extract_year_from_filename($filename) {
    $base = pathinfo($filename, PATHINFO_FILENAME);
    if (preg_match('/(\d{4})/', $base, $m)) {
        return $m[1];
    }
    return null;
}

/**
 * Extract year from Excel file content as fallback
 * Looks for "ปีงบประมาณ XXXX" or sheet names with 2-digit year suffix like "ก.พ. 69"
 * Returns: year string like "2569" or null if not found
 */
function extract_year_from_excel_content($filepath) {
    if (!class_exists('PhpOffice\\PhpSpreadsheet\\IOFactory')) {
        return null;
    }

    try {
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($filepath);

        // ╔══════════════════════════════════════════════════════════════╗
        // ║ ⚠️  [FISCAL YEAR DETECTION] SHEET FILTER — ข้ามชีทสรุป/กราฟ  ║
        // ║                                                              ║
        // ║ ตรวจสอบเฉพาะ 2 ชีทแรก เพื่อหา "ปีงบประมาณ XXXX"              ║
        // ║ ข้ามชีทสรุป (สรุป, กราฟ, Summary, Chart)                    ║
        // ║                                                              ║
        // ║ Sheets to AVOID: "กราฟ", "สรุป", "รวม", "Chart", "Summary" ║
        // ║ Sheets to PROCESS: Main data sheets with fiscal year info   ║
        // ╚══════════════════════════════════════════════════════════════╝
        // Look for "ปีงบประมาณ XXXX" in first few rows of first sheet
        for ($si = 0; $si < min(2, $spreadsheet->getSheetCount()); $si++) {
            $ws = $spreadsheet->getSheet($si);
            $maxRow = min(6, $ws->getHighestDataRow());
            $maxCol = min(30, \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($ws->getHighestDataColumn()));

            for ($r = 1; $r <= $maxRow; $r++) {
                for ($c = 1; $c <= $maxCol; $c++) {
                    $v = (string)($ws->getCell([$c, $r])->getValue() ?? '');
                    if (preg_match('/ปีงบประมาณ\s*(\d{4})/', $v, $m)) {
                        $spreadsheet->disconnectWorksheets();
                        return $m[1];
                    }
                }
            }
        }

        // ╔══════════════════════════════════════════════════════════════╗
        // ║ ⚠️  [FALLBACK] SHEET NAME SCAN — ค้นหาปีจากชื่อชีท            ║
        // ║                                                              ║
        // ║ ถ้าไม่พบ "ปีงบประมาณ" ในข้อมูล ลองค้นหา 2 หลักปี             ║
        // ║ จากชื่อชีท เช่น "ก.พ. 69" → "2569"                         ║
        // ║ โปรแกรมจะตรวจสอบชีท 1 ชั้น แล้ว break ไม่ต้องตรวจทั้งหมด   ║
        // ║                                                              ║
        // ║ Sheets to AVOID: "กราฟ", "สรุป", "รวม", "Chart", "Summary" ║
        // ║ Sheets to PROCESS: Main data sheets with year in name       ║
        // ╚══════════════════════════════════════════════════════════════╝
        // If not found, try sheet names — majority vote
        $_yc = [];
        foreach ($spreadsheet->getSheetNames() as $sname) {
            if (preg_match('/(\d{2})\s*$/', trim($sname), $m)) {
                $_yy = $m[1];
                $_yc[$_yy] = ($_yc[$_yy] ?? 0) + 1;
            }
        }
        $spreadsheet->disconnectWorksheets();
        if (!empty($_yc)) {
            arsort($_yc);
            return '25' . array_key_first($_yc);
        }
    } catch (\Throwable $e) {
        // Ignore and return null
    }

    return null;
}

// ────────────────────────────────────────────────────────────────────────────
// TAB 1: KPI จุดซ่อมท่อ
// ────────────────────────────────────────────────────────────────────────────

function read_repair_data($REPAIR_DIR, $MONTH_NAMES) {
    if (!is_dir($REPAIR_DIR)) {
        echo "  [SKIP] ไม่พบโฟลเดอร์ ลงข้อมูลซ่อมท่อ/\n";
        return null;
    }

    $files = array_filter(
        scandir($REPAIR_DIR),
        function($f) { return substr($f, -5) === '.xlsx' && substr($f, 0, 2) !== '~$'; }
    );
    $files = array_values($files);
    sort($files);

    echo "  พบไฟล์ข้อมูล: " . count($files) . " ไฟล์\n";

    $file_info = [];
    foreach ($files as $fname) {
        if (preg_match('/(\d{6})/', $fname, $m)) {
            $digits = $m[1];
            $yy = (int)substr($digits, 0, 2);
            $mm = (int)substr($digits, 2, 2);
            $dd = (int)substr($digits, 4, 2);
            $month_key = sprintf("%02d-%02d", $yy, $mm);
            $file_info[] = compact('fname', 'yy', 'mm', 'dd', 'month_key');
        }
    }

    // Pick latest per month
    $month_files = [];
    foreach ($file_info as $fi) {
        $mk = $fi['month_key'];
        if (!isset($month_files[$mk]) || $fi['dd'] > $month_files[$mk]['dd']) {
            $month_files[$mk] = $fi;
        }
    }

    echo "  เดือนที่มีข้อมูล: " . count($month_files) . "\n";
    foreach (array_keys($month_files) as $mk) {
        echo "    {$mk} <- {$month_files[$mk]['fname']}\n";
    }

    $all_data = [];
    $branches_order = [];

    foreach ($month_files as $mk => $fi) {
        $fpath = $REPAIR_DIR . DIRECTORY_SEPARATOR . $fi['fname'];
        try {
            $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($fpath);
            $worksheet = $spreadsheet->getActiveSheet();

            $det = detect_repair_columns($worksheet);
            $hdr = $det['header_row'];
            $cBranch   = $det['cols']['branch']   ?? 1;
            $cClosed   = $det['cols']['closed']    ?? 2;
            $cComplete = $det['cols']['complete']  ?? 3;
            $cScore    = $det['cols']['score']     ?? 4;

            $month_data = [];
            $max_row = $worksheet->getHighestRow();

            for ($r = $hdr + 1; $r <= $max_row; $r++) {
                $branch = $worksheet->getCell([$cBranch, $r])->getValue();
                if (!$branch || !is_string($branch)) continue;

                $branch = trim($branch);
                if ($branch === '' || mb_strpos($branch, 'ชื่อสาขา') !== false) continue;

                $month_data[$branch] = [
                    'closed'   => clean_num($worksheet->getCell([$cClosed, $r])->getValue()),
                    'complete' => clean_num($worksheet->getCell([$cComplete, $r])->getValue()),
                    'score'    => clean_num($worksheet->getCell([$cScore, $r])->getValue())
                ];

                if (!in_array($branch, $branches_order)) {
                    $branches_order[] = $branch;
                }
            }

            $all_data[$mk] = $month_data;
            $spreadsheet->disconnectWorksheets();
        } catch (\Throwable $e) {
            echo "  [WARNING] ข้ามไฟล์เสีย: {$fi['fname']} ({$e->getMessage()})\n";
        }
    }

    $months_sorted = array_keys($all_data);
    sort($months_sorted);

    if ($months_sorted) {
        echo "  ช่วงเดือน: {$months_sorted[0]} - {$months_sorted[count($months_sorted)-1]}\n";
    }
    echo "  จำนวนสาขา: " . count($branches_order) . "\n";

    $month_names_map = [];
    for ($i = 1; $i <= 12; $i++) {
        $month_names_map[sprintf("%02d", $i)] = $MONTH_NAMES[$i];
    }

    return [
        'months' => $months_sorted,
        'branches' => $branches_order,
        'data' => $all_data,
        'month_names' => $month_names_map
    ];
}

// ────────────────────────────────────────────────────────────────────────────
// TAB 2: แรงดันน้ำ
// ────────────────────────────────────────────────────────────────────────────

function read_pressure_data($PRESSURE_DIR) {
    if (!is_dir($PRESSURE_DIR)) {
        echo "  [SKIP] ไม่พบโฟลเดอร์ แรงดันน้ำ/\n";
        return [null, null];
    }

    $files = array_filter(
        scandir($PRESSURE_DIR),
        function($f) {
            return preg_match('/^PRESSURE_.*\.xlsx$/i', $f) && substr($f, 0, 2) !== '~$';
        }
    );

    if (empty($files)) {
        echo "  [SKIP] ไม่พบไฟล์ PRESSURE_*.xlsx\n";
        return [null, null];
    }

    echo "  พบไฟล์แรงดัน: " . count($files) . " ไฟล์\n";

    $thai_month_map = [
        'ม.ค.' => 1, 'ก.พ.' => 2, 'มี.ค.' => 3, 'เม.ย.' => 4, 'พ.ค.' => 5, 'มิ.ย.' => 6,
        'ก.ค.' => 7, 'ส.ค.' => 8, 'ก.ย.' => 9, 'ต.ค.' => 10, 'พ.ย.' => 11, 'ธ.ค.' => 12
    ];

    $pressure_data = [];
    $all_months = [];

    foreach ($files as $fname) {
        // Extract branch name from filename
        $branch_name = '';
        if (preg_match('/PRESSURE_(.+?)_ปีงบ\d+\.xlsx/i', $fname, $m)) {
            $branch_name = $m[1];
        } elseif (preg_match('/PRESSURE_(.+?)\.xlsx/i', $fname, $m)) {
            $branch_name = $m[1];
        }
        if (!$branch_name) continue;

        $fpath = $PRESSURE_DIR . DIRECTORY_SEPARATOR . $fname;
        try {
            $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($fpath);
            $worksheet = $spreadsheet->getActiveSheet();

            // Row 5 has month headers
            $month_cols = [];
            for ($c = 1; $c <= min(50, \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString(
                $worksheet->getHighestColumn()
            )); $c++) {
                $header = trim((string)($worksheet->getCell([$c, 5])->getValue() ?? ''));
                if ($header === '') continue;

                foreach ($thai_month_map as $thai_m => $m_num) {
                    if (mb_strpos($header, $thai_m) !== false) {
                        if (preg_match('/(\d{2})/', $header, $m)) {
                            $yy = (int)$m[1];
                            $mk = sprintf("%02d-%02d", $yy, $m_num);
                            $month_cols[$c] = $mk;
                        }
                        break;
                    }
                }
            }

            // Read data rows (7 to max_row)
            foreach ($month_cols as $mk) {
                if (!in_array($mk, $all_months)) {
                    $all_months[] = $mk;
                }
            }

            foreach ($month_cols as $col_idx => $mk) {
                $total = 0.0;
                $count = 0;
                $max_row = $worksheet->getHighestRow();

                for ($r = 7; $r <= $max_row; $r++) {
                    $v = $worksheet->getCell([$col_idx, $r])->getValue();
                    if (is_numeric($v) && $v > 0) {
                        $total += (float)$v;
                        $count++;
                    }
                }

                if ($count > 0) {
                    $avg = round($total / $count, 2);
                    if (!isset($pressure_data[$mk])) {
                        $pressure_data[$mk] = [];
                    }
                    $pressure_data[$mk][$branch_name] = $avg;
                }
            }

            $spreadsheet->disconnectWorksheets();
        } catch (\Throwable $e) {
            echo "  [WARNING] ข้ามไฟล์: {$fname} ({$e->getMessage()})\n";
        }
    }

    // Filter months with non-zero data
    $months_sorted = [];
    foreach ($all_months as $mk) {
        if (isset($pressure_data[$mk]) && !empty($pressure_data[$mk])) {
            $has_nonzero = false;
            foreach ($pressure_data[$mk] as $v) {
                if ($v > 0) {
                    $has_nonzero = true;
                    break;
                }
            }
            if ($has_nonzero) {
                $months_sorted[] = $mk;
            }
        }
    }
    sort($months_sorted);

    echo "  เดือนแรงดัน: " . json_encode($months_sorted, JSON_UNESCAPED_UNICODE) . "\n";
    echo "  สาขาที่มีข้อมูล: " . count($files) . "\n";

    return [$pressure_data, $months_sorted];
}

// ────────────────────────────────────────────────────────────────────────────
// TAB 3: งานค้างซ่อม
// ────────────────────────────────────────────────────────────────────────────

function read_pending_data($PENDING_DIR, $BRANCH_LIST) {
    if (!is_dir($PENDING_DIR)) {
        echo "  [SKIP] ไม่พบโฟลเดอร์ ซ่อมท่อค้างระบบ/\n";
        return null;
    }

    $files = scandir($PENDING_DIR);
    $fy_files = [];
    $fy_list = [];

    foreach ($files as $fname) {
        if (!preg_match('/\.(xlsx|xls)$/i', $fname) || substr($fname, 0, 2) === '~$' || substr($fname, -5) === '.json') {
            continue;
        }

        $fpath = $PENDING_DIR . DIRECTORY_SEPARATOR . $fname;
        if (preg_match('/(\d{1,2})-(\d{2})_to_(\d{1,2})-(\d{2})/', $fname, $m)) {
            $start_mm = (int)$m[1];
            $start_yy = (int)$m[2];
            $fy_be = 2500 + $start_yy + ($start_mm >= 10 ? 1 : 0);
            $fy_files[$fy_be] = $fpath;
            $fy_list[] = $fy_be;
            echo "    ปีงบฯ {$fy_be} <- {$fname}\n";
        } else {
            $fy_files[0] = $fpath;
            echo "    (ไม่ระบุปี) <- {$fname}\n";
        }
    }

    if (empty($fy_files)) {
        echo "  [SKIP] ไม่พบไฟล์ข้อมูลค้างซ่อม\n";
        return null;
    }

    sort($fy_list);
    $fy_list = array_filter($fy_list, function($k) { return $k > 0; });

    // For files without year in filename (key 0), try to extract from Excel content
    if (isset($fy_files[0])) {
        $fallback_year = extract_year_from_excel_content($fy_files[0]);
        if ($fallback_year) {
            $fallback_fy_be = (int)$fallback_year;
            $fy_files[$fallback_fy_be] = $fy_files[0];
            if (!in_array($fallback_fy_be, $fy_list)) {
                $fy_list[] = $fallback_fy_be;
                echo "    ปีงบฯ {$fallback_fy_be} (ดึงจากเนื้อหา)\n";
            }
            sort($fy_list);
        } else {
            // If can't extract year, use default 2569
            $fy_list[] = 2569;
            $fy_files[2569] = $fy_files[0];
            echo "    ปีงบฯ 2569 (ค่าเริ่มต้น)\n";
        }
        unset($fy_files[0]);
    }

    // Column indices (0-based)
    $col_date = 3;      // วันที่แจ้ง
    $col_finish = 5;    // วันเวลาเสร็จสิ้น
    $col_branch = 19;   // สาขา
    $col_status = 26;   // สถานะ
    $data_start = 8;    // row 9 = index 8

    $all_fy_results = [];

    foreach ($fy_list as $fy) {
        $fpath = isset($fy_files[$fy]) ? $fy_files[$fy] : null;
        if (!$fpath || !is_file($fpath)) {
            if (!empty($fy_files)) {
                $fpath = reset($fy_files);
            } else {
                continue;
            }
        }

        try {
            $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($fpath);
            $worksheet = $spreadsheet->getActiveSheet();
            $max_row = $worksheet->getHighestRow();

            $fy_be = $fy > 0 ? $fy : 2569;
            $fy_ce = $fy_be - 543;
            $count_start = new DateTime("$fy_ce-01-01");

            // Read all records
            $records = [];
            $last_report_dt = null;

            for ($r = $data_start + 1; $r <= $max_row; $r++) {
                $date_val = $worksheet->getCell([$col_date + 1, $r])->getValue();
                if (!$date_val) continue;

                [$dt, $by] = parse_thai_date($date_val);
                if (!$dt) continue;

                if ($last_report_dt === null || $dt > $last_report_dt) {
                    $last_report_dt = $dt;
                }

                $finish_val = $worksheet->getCell([$col_finish + 1, $r])->getValue();
                $finish_dt = null;
                if ($finish_val) {
                    [$fdt, $_] = parse_thai_date($finish_val);
                    if ($fdt) {
                        $finish_dt = $fdt;
                    } elseif (is_string($finish_val) && strlen($finish_val) >= 10) {
                        [$fdt, $_] = parse_thai_date(substr($finish_val, 0, 10));
                        if ($fdt) $finish_dt = $fdt;
                    }
                }

                $status = trim((string)($worksheet->getCell([$col_status + 1, $r])->getValue() ?? ''));
                $branch = trim((string)($worksheet->getCell([$col_branch + 1, $r])->getValue() ?? ''));

                if (!$branch) continue;

                $records[] = [
                    'dt' => $dt,
                    'by' => $by,
                    'finish_dt' => $finish_dt,
                    'status' => $status,
                    'branch' => $branch
                ];
            }

            // Build update_date
            $update_date = '';
            if ($last_report_dt) {
                $by_lrd = (int)$last_report_dt->format('Y') < 2500
                    ? (int)$last_report_dt->format('Y') + 543
                    : (int)$last_report_dt->format('Y');
                $update_date = sprintf("%02d-%02d-%02d",
                    (int)$last_report_dt->format('d'),
                    (int)$last_report_dt->format('m'),
                    $by_lrd % 100
                );
            }

            // --- PD2: ค้างซ่อม ณ สิ้นเดือน ---
            $month_set = [];
            foreach ($records as $rec) {
                if ($rec['dt'] >= $count_start) {
                    $y = (int)$rec['dt']->format('Y');
                    $m = (int)$rec['dt']->format('m');
                    $month_set[] = "$y-$m";
                }
            }
            $month_set = array_unique($month_set);
            sort($month_set);

            $pd2_months = [];
            foreach ($month_set as $ym) {
                [$y, $m] = explode('-', $ym);
                $y = (int)$y;
                $m = (int)$m;
                $yy = $y < 2500 ? ($y + 543) % 100 : $y % 100;
                $pd2_months[] = sprintf("%02d-%02d", $yy, $m);
            }

            $pd2_data = [];
            foreach ($month_set as $ym) {
                [$y, $m] = explode('-', $ym);
                $y = (int)$y;
                $m = (int)$m;

                $end_day = (int)date('t', mktime(0, 0, 0, $m, 1, $y));
                $end_of_month = new DateTime("$y-$m-$end_day 23:59:59");

                $yy = $y < 2500 ? ($y + 543) % 100 : $y % 100;
                $mk = sprintf("%02d-%02d", $yy, $m);

                $branch_counts = [];
                foreach ($records as $rec) {
                    if ($rec['dt'] < $count_start || $rec['dt'] > $end_of_month) {
                        continue;
                    }

                    $is_pending = false;
                    if ($rec['finish_dt'] && $rec['finish_dt'] > $end_of_month) {
                        $is_pending = true;
                    } elseif (mb_strpos($rec['status'], 'ซ่อมไม่เสร็จ') !== false) {
                        $is_pending = true;
                    }

                    if ($is_pending) {
                        if (!isset($branch_counts[$rec['branch']])) {
                            $branch_counts[$rec['branch']] = 0;
                        }
                        $branch_counts[$rec['branch']]++;
                    }
                }

                $pd2_data[$mk] = $branch_counts;
            }

            // --- PD1: เปรียบเทียบเดือน (derived from PD2) ---
            $pd1_data = [];
            foreach ($pd2_months as $i => $mk) {
                $prev_mk = $i > 0 ? $pd2_months[$i - 1] : null;
                $prev_snap = $prev_mk && isset($pd2_data[$prev_mk]) ? $pd2_data[$prev_mk] : [];
                $curr_snap = isset($pd2_data[$mk]) ? $pd2_data[$mk] : [];

                $branch_pairs = [];
                foreach ($BRANCH_LIST as $b) {
                    $pv = isset($prev_snap[$b]) ? $prev_snap[$b] : 0;
                    $cv = isset($curr_snap[$b]) ? $curr_snap[$b] : 0;
                    $branch_pairs[$b] = [$pv, $cv];
                }

                $pd1_data[$mk] = $branch_pairs;
            }

            // --- PD3: ตารางงานซ่อมที่ยังไม่ปิดในระบบ ---
            $fy_yy_start = $fy_be > 0 ? (($fy_be - 2500 - 1) % 100) : 68;
            $fy_yy_end = ($fy_yy_start + 1) % 100;

            $fy_months = [];
            for ($mm = 10; $mm <= 12; $mm++) {
                $fy_months[] = sprintf("%02d-%02d", $fy_yy_start, $mm);
            }
            for ($mm = 1; $mm <= 9; $mm++) {
                $fy_months[] = sprintf("%02d-%02d", $fy_yy_end, $mm);
            }

            $pd3_data = [];
            foreach ($records as $rec) {
                if (mb_strpos($rec['status'], 'ซ่อมไม่เสร็จ') === false) {
                    continue;
                }

                $yy = $rec['by'] % 100;
                $mk = sprintf("%02d-%02d", $yy, (int)$rec['dt']->format('m'));

                if (in_array($mk, $fy_months)) {
                    if (!isset($pd3_data[$rec['branch']])) {
                        $pd3_data[$rec['branch']] = [];
                    }
                    if (!isset($pd3_data[$rec['branch']][$mk])) {
                        $pd3_data[$rec['branch']][$mk] = 0;
                    }
                    $pd3_data[$rec['branch']][$mk]++;
                }
            }

            // Compute col totals and grand total
            $col_totals = [];
            foreach ($fy_months as $mk) {
                $col_totals[$mk] = 0;
            }
            $grand_total = 0;

            foreach ($pd3_data as $branch_data) {
                foreach ($branch_data as $mk => $v) {
                    if (!isset($col_totals[$mk])) {
                        $col_totals[$mk] = 0;
                    }
                    $col_totals[$mk] += $v;
                    $grand_total += $v;
                }
            }

            echo "  ปีงบฯ {$fy_be}: records=" . count($records) . ", pd2_months=" . count($pd2_months) . ", update={$update_date}\n";

            $all_fy_results[(string)$fy_be] = [
                'update_date' => $update_date,
                'pd1_data' => $pd1_data,
                'pd1_months' => $pd2_months,
                'pd2_data' => $pd2_data,
                'pd2_months' => $pd2_months,
                'pd3' => [
                    'months' => $fy_months,
                    'update_date' => $update_date,
                    'data' => $pd3_data,
                    'col_totals' => $col_totals,
                    'grand_total' => $grand_total
                ]
            ];

            $spreadsheet->disconnectWorksheets();
        } catch (\Throwable $e) {
            echo "  [WARNING] อ่านไฟล์ไม่ได้: {$fpath} ({$e->getMessage()})\n";
        }
    }

    return [
        'fy_list' => $fy_list,
        'results' => $all_fy_results
    ];
}

// ────────────────────────────────────────────────────────────────────────────
// Embed into index.html
// ────────────────────────────────────────────────────────────────────────────

/**
 * Replace a JS variable in HTML using regex
 */
function replace_js_var($html, $pattern, $new_value) {
    $new_html = preg_replace($pattern, $new_value, $html, 1);
    return $new_html !== null ? $new_html : $html;
}

/**
 * Embed all data into HTML
 */
function embed_all($html, $repair_data, $pressure_data, $pressure_months, $pending_result) {
    $changes = [];

    // --- TAB 1: KPI ---
    if ($repair_data) {
        $data_json = json_encode($repair_data, JSON_UNESCAPED_UNICODE);

        if (strpos($html, 'GIS_DATA_PLACEHOLDER') !== false) {
            $html = str_replace('GIS_DATA_PLACEHOLDER', $data_json, $html);
            $changes[] = "TAB 1 KPI (placeholder)";
        } else {
            $pat = '/^const DATA = \{.*\};$/m';
            $new_val = 'const DATA = ' . $data_json . ';';
            $html = replace_js_var($html, $pat, $new_val);
            $changes[] = "TAB 1 KPI (const DATA)";
        }
    }

    // --- TAB 2: Pressure ---
    if ($pressure_data && $pressure_months) {
        $pdata_json = json_encode($pressure_data, JSON_UNESCAPED_UNICODE);
        $pmonths_json = json_encode($pressure_months, JSON_UNESCAPED_UNICODE);

        $html = replace_js_var($html,
            '/^const PRESSURE_DATA=\{.*\};$/m',
            'const PRESSURE_DATA=' . $pdata_json . ';');

        $html = replace_js_var($html,
            '/^const PRESSURE_MONTHS=\[.*\];$/m',
            'const PRESSURE_MONTHS=' . $pmonths_json . ';');

        $changes[] = "TAB 2 Pressure (" . count($pressure_months) . " เดือน)";
    }

    // --- TAB 3: Pending ---
    if ($pending_result) {
        $fy_list = $pending_result['fy_list'];
        $results = $pending_result['results'];

        $latest_fy = !empty($fy_list) ? (string)$fy_list[count($fy_list) - 1] : '2569';
        $latest = isset($results[$latest_fy]) ? $results[$latest_fy] : [];

        if ($latest) {
            $update_date = $latest['update_date'];
            $pd1_data = $latest['pd1_data'];
            $pd1_months = $latest['pd1_months'];
            $pd2_data = $latest['pd2_data'];
            $pd2_months = $latest['pd2_months'];

            // PENDING_UPDATE_DATE
            $html = replace_js_var($html,
                '/^(?:const|var) PENDING_UPDATE_DATE=\'[^\']*\';$/m',
                "const PENDING_UPDATE_DATE='{$update_date}';");

            // PENDING_FY_LIST
            $fy_list_json = json_encode(array_map('intval', $fy_list), JSON_UNESCAPED_UNICODE);
            $html = replace_js_var($html,
                '/^var PENDING_FY_LIST=\[.*\];.*$/m',
                'var PENDING_FY_LIST=' . $fy_list_json . '; // fallback');

            // PD1
            $pd1_data_json = json_encode($pd1_data, JSON_UNESCAPED_UNICODE);
            $pd1_months_json = json_encode($pd1_months, JSON_UNESCAPED_UNICODE);

            $html = replace_js_var($html,
                '/^var PD1_DATA_FALLBACK=\{.*\};$/m',
                'var PD1_DATA_FALLBACK=' . $pd1_data_json . ';');

            $html = replace_js_var($html,
                '/^var PD1_MONTHS_FALLBACK=\[.*\];$/m',
                'var PD1_MONTHS_FALLBACK=' . $pd1_months_json . ';');

            // PD2
            $pd2_data_json = json_encode($pd2_data, JSON_UNESCAPED_UNICODE);
            $pd2_months_json = json_encode($pd2_months, JSON_UNESCAPED_UNICODE);

            $html = replace_js_var($html,
                '/^var PD2_DATA_FALLBACK=\{.*\};$/m',
                'var PD2_DATA_FALLBACK=' . $pd2_data_json . ';');

            $html = replace_js_var($html,
                '/^var PD2_MONTHS_FALLBACK=\[.*\];$/m',
                'var PD2_MONTHS_FALLBACK=' . $pd2_months_json . ';');

            // PD3_FALLBACK — embed all FYs
            $pd3_fallback = [];
            foreach ($results as $fy_key => $fy_data) {
                $pd3_fallback[$fy_key] = $fy_data['pd3'];
            }
            $pd3_json = json_encode($pd3_fallback, JSON_UNESCAPED_UNICODE);

            // Try multiline pattern first (DOTALL)
            $html_new = preg_replace(
                '/var PD3_FALLBACK=\{.*?\n\};/s',
                'var PD3_FALLBACK=' . $pd3_json . ';',
                $html,
                1
            );

            if ($html_new !== $html && $html_new !== null) {
                $html = $html_new;
            } else {
                // Try single-line pattern
                $html = replace_js_var($html,
                    '/^var PD3_FALLBACK=\{.*\};$/m',
                    'var PD3_FALLBACK=' . $pd3_json . ';');
            }

            $changes[] = "TAB 3 Pending (ปีงบฯ " . implode(',', $fy_list) . ", update={$update_date})";
        }
    }

    return [$html, $changes];
}

// ────────────────────────────────────────────────────────────────────────────
// Main Build
// ────────────────────────────────────────────────────────────────────────────

/**
 * Parse CLI arguments for incremental build:
 *   --only=repair   → process only repair category (skip pressure, pending)
 *   --files=a.xlsx  → process only these specific files within the category
 */
function parse_cli_args() {
    global $argv;
    $args = ['only' => '', 'files' => []];
    if (!isset($argv)) return $args;
    foreach ($argv as $a) {
        if (strpos($a, '--only=') === 0) {
            $args['only'] = substr($a, 7);
        }
        if (strpos($a, '--files=') === 0) {
            $args['files'] = array_filter(explode(',', substr($a, 8)));
        }
    }
    return $args;
}

function build() {
    global $SCRIPT_DIR, $REPAIR_DIR, $PRESSURE_DIR, $PENDING_DIR, $HTML_TEMPLATE;
    global $MONTH_NAMES, $BRANCH_LIST;

    $args = parse_cli_args();
    $only = $args['only'];       // e.g. 'repair', 'pressure', 'pending', or '' (all)
    $only_files = $args['files'];

    echo "==================================================\n";
    echo "  Build Dashboard แผนที่แนวท่อ (GIS) กปภ.เขต 1\n";
    if ($only) {
        echo "  ⚡ Incremental build: only=$only" . ($only_files ? " files=" . implode(',', $only_files) : '') . "\n";
    }
    echo "==================================================\n";

    // --- 1) Read repair data (TAB 1) ---
    if (!$only || $only === 'repair') {
        echo "\n[1/4] อ่านข้อมูล KPI จุดซ่อมท่อ...\n";
        $repair_data = read_repair_data($REPAIR_DIR, $MONTH_NAMES);
    } else {
        echo "\n⏭️  Repair: ข้าม (ไม่ได้เปลี่ยน)\n";
        $repair_data = [];
    }

    // --- 2) Read pressure data (TAB 2) ---
    if (!$only || $only === 'pressure') {
        echo "\n[2/4] อ่านข้อมูลแรงดันน้ำ...\n";
        [$pressure_data, $pressure_months] = read_pressure_data($PRESSURE_DIR);
    } else {
        echo "⏭️  Pressure: ข้าม (ไม่ได้เปลี่ยน)\n";
        $pressure_data = [];
        $pressure_months = [];
    }

    // --- 3) Read pending data (TAB 3) ---
    if (!$only || $only === 'pending') {
        echo "\n[3/4] อ่านข้อมูลงานค้างซ่อม...\n";
        $pending_result = read_pending_data($PENDING_DIR, $BRANCH_LIST);
    } else {
        echo "⏭️  Pending: ข้าม (ไม่ได้เปลี่ยน)\n";
        $pending_result = null;
    }

    // --- 4) Embed into index.html ---
    echo "\n[4/4] Embed ข้อมูลลงใน index.html...\n";

    if (!is_file($HTML_TEMPLATE)) {
        echo "  [ERROR] ไม่พบ index.html\n";
        exit(1);
    }

    $html = file_get_contents($HTML_TEMPLATE);
    [$html, $changes] = embed_all($html, $repair_data, $pressure_data, $pressure_months, $pending_result);

    foreach ($changes as $c) {
        echo "  ✓ {$c}\n";
    }

    if (empty($changes)) {
        echo "  [WARNING] ไม่มีข้อมูลที่จะ embed!\n";
    }

    $output_path = $SCRIPT_DIR . DIRECTORY_SEPARATOR . "index.html";
    file_put_contents($output_path, $html);

    $size_kb = filesize($output_path) / 1024;
    echo "\n  บันทึก: {$output_path}\n";
    echo "  ขนาดไฟล์: " . number_format($size_kb, 1) . " KB\n";
    echo "\n  เสร็จสิ้น!\n";
}

// Run build
build();
?>
