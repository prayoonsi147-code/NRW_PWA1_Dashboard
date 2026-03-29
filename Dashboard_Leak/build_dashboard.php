<?php
/**
 * build_dashboard.php - สร้าง Dashboard อัตโนมัติจากไฟล์ Excel
 * ============================================================================
 * PHP CLI script that replaces build_dashboard.py
 * Usage: C:\xampp\php\php.exe build_dashboard.php
 *
 * Parses 6 categories of data:
 *   1. D (OIS) - ข้อมูลดิบ/OIS/
 *   2. RL (Real Leak) - ข้อมูลดิบ/Real Leak/
 *   3. EU (หน่วยไฟ) - ข้อมูลดิบ/หน่วยไฟ/
 *   4. MNF - ข้อมูลดิบ/MNF/
 *   5. KPI - ข้อมูลดิบ/เกณฑ์วัดน้ำสูญเสีย/
 *   6. P3 - ข้อมูลดิบ/P3/
 *
 * Embeds data as JavaScript const variables in index.html
 */

// ============================================================================
// Configuration
// ============================================================================

define('SCRIPT_DIR', __DIR__);
define('RAW_DATA_DIR', SCRIPT_DIR . DIRECTORY_SEPARATOR . 'ข้อมูลดิบ');
define('INDEX_HTML', SCRIPT_DIR . DIRECTORY_SEPARATOR . 'index.html');

// Standard branch names for normalization
const STANDARD_BRANCHES = [
    'ชลบุรี(พ)', 'พัทยา(พ)', 'พนัสนิคม', 'บ้านบึง', 'ศรีราชา', 'แหลมฉบัง',
    'ฉะเชิงเทรา', 'บางปะกง', 'บางคล้า', 'พนมสารคาม', 'ระยอง', 'บ้านฉาง',
    'ปากน้ำประแสร์', 'จันทบุรี', 'ขลุง', 'ตราด', 'คลองใหญ่', 'สระแก้ว',
    'วัฒนานคร', 'อรัญประเทศ', 'ปราจีนบุรี', 'กบินทร์บุรี'
];

const BRANCH_ALIASES = [
    'พนัมสารคาม' => 'พนมสารคาม',
];

const MONTH_KEYWORDS = ['ต.ค.', 'พ.ย.', 'ธ.ค.', 'ม.ค.', 'ก.พ.', 'มี.ค.',
                        'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.'];

const MONTH_SHORT = ['ต.ค.', 'พ.ย.', 'ธ.ค.', 'ม.ค.', 'ก.พ.', 'มี.ค.',
                     'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.'];

const MONTH_LONG = ['ตุลาคม', 'พฤศจิกายน', 'ธันวาคม', 'มกราคม', 'กุมภาพันธ์',
                    'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม',
                    'สิงหาคม', 'กันยายน'];

const RL_MONTH_ABBR = [
    'ต.ค.' => 0, 'พ.ย.' => 1, 'ธ.ค.' => 2, 'ม.ค.' => 3, 'ก.พ.' => 4, 'มี.ค.' => 5,
    'เม.ย.' => 6, 'พ.ค.' => 7, 'มิ.ย.' => 8, 'ก.ค.' => 9, 'ส.ค.' => 10, 'ก.ย.' => 11
];

const LABEL_NORMALIZE_MAP = [
    '2.5 อัตราการสูญเสีย (ต่อน้ำผลิตจ่าย)' => '2.5 อัตราน้ำสูญเสีย (ต่อน้ำผลิตจ่าย)',
    '2.2  ปริมาณน้ำจ่ายฟรี + Blowoff' => '2.2  ปริมาณน้ำจ่ายฟรี',
    '4.2 เงินเดือนและค่าจ้างประจำ' => '4.1 เงินเดือนและค่าจ้างประจำ',
    '4.3 ค่าจ้างชั่วคราว' => '4.2 ค่าจ้างชั่วคราว',
    '4.5 วัสดุการผลิต' => '4.4 วัสดุการผลิต',
];

const MNF_ROW_MAP = [
    'MNF เกิดจริง' => 'actual',
    'MNF ที่ยอมรับได้' => 'acceptable',
    'เป้าหมาย MNF' => 'target',
    'น้ำผลิตจ่าย' => 'production',
];

const MNF_SHEET_MAP = [
    '1.ชลบุรี' => 'ชลบุรี(พ)', '2.พัทยา' => 'พัทยา(พ)', '3.บ้านบึง' => 'บ้านบึง',
    '4.พนัสนิคม' => 'พนัสนิคม', '5.ศรีราชา' => 'ศรีราชา', '6.แหลมฉบัง' => 'แหลมฉบัง',
    '7.บางปะกง' => 'บางปะกง', '8.ฉะเชิงเทรา' => 'ฉะเชิงเทรา', '9.บางคล้า' => 'บางคล้า',
    '10.พนมสารคาม' => 'พนมสารคาม', '11.ระยอง' => 'ระยอง', '12.บ้านฉาง' => 'บ้านฉาง',
    '13.ปากน้ำประแสร์' => 'ปากน้ำประแสร์', '14.จันทบุรี' => 'จันทบุรี', '15.ขลุง' => 'ขลุง',
    '16.ตราด' => 'ตราด', '17.คลองใหญ่' => 'คลองใหญ่', '18.สระแก้ว' => 'สระแก้ว',
    '19.วัฒนานคร' => 'วัฒนานคร', '20.อรัญประเทศ' => 'อรัญประเทศ',
    '21.ปราจีนบุรี' => 'ปราจีนบุรี', '22.กบินทร์บุรี' => 'กบินทร์บุรี',
];

// ============================================================================
// Helper Functions
// ============================================================================

function normalize_branch_name($raw_name) {
    if ($raw_name === null || $raw_name === '') return null;
    $name = trim((string)$raw_name);
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

function cellVal($sheet, $col, $row) {
    try {
        return $sheet->getCell([$col, $row])->getValue();
    } catch (Exception $e) {
        return null;
    }
}

function cellCalc($sheet, $col, $row) {
    try {
        return $sheet->getCell([$col, $row])->getCalculatedValue();
    } catch (Exception $e) {
        return null;
    }
}

function extract_year_from_filename($filename) {
    $base = pathinfo($filename, PATHINFO_FILENAME);
    if (preg_match('/(\d{4})/', $base, $m)) {
        return $m[1];
    }
    return null;
}

function find_month_header_row($sheet_data) {
    foreach ($sheet_data as $row_num => $row) {
        $text_cells = [];
        foreach ($row as $v) {
            if ($v !== null) $text_cells[] = (string)$v;
        }
        $row_text = implode(' ', $text_cells);
        $count = 0;
        foreach (MONTH_KEYWORDS as $kw) {
            if (mb_strpos($row_text, $kw) !== false) $count++;
        }
        if ($count >= 6) return $row_num;
    }
    return null;
}

function find_month_columns($sheet_data, $header_row) {
    $row = isset($sheet_data[$header_row]) ? $sheet_data[$header_row] : [];
    $month_cols = array_fill(0, 12, null);

    foreach ($row as $col => $val) {
        if ($val === null) continue;
        $sval = (string)$val;
        for ($mi = 0; $mi < 12; $mi++) {
            if (mb_strpos($sval, MONTH_SHORT[$mi]) !== false ||
                mb_strpos($sval, MONTH_LONG[$mi]) !== false) {
                $month_cols[$mi] = $col;
                break;
            }
        }
    }
    return $month_cols;
}

function find_total_column($sheet_data, $header_row) {
    $row = isset($sheet_data[$header_row - 1]) ? $sheet_data[$header_row - 1] : [];
    foreach ($row as $col => $val) {
        if (is_string($val) && mb_strpos($val, 'รวม') !== false) {
            return $col;
        }
    }
    $row2 = isset($sheet_data[$header_row]) ? $sheet_data[$header_row] : [];
    foreach ($row2 as $col => $val) {
        if (is_string($val) && mb_strpos($val, 'รวม') !== false) {
            return $col;
        }
    }
    return null;
}

function extract_sheet_data($sheet_data, $header_row, $month_cols, $total_col) {
    $rows = [];
    $data_start = $header_row + 1;

    foreach ($sheet_data as $row_num => $row) {
        if ($row_num < $data_start) continue;

        $label = isset($row[0]) ? $row[0] : '';
        if (is_numeric($label)) $label = (string)$label;
        $label = trim((string)$label);
        if (!$label || mb_strpos($label, 'หมายเหตุ') !== false) continue;

        $unit = isset($row[1]) ? $row[1] : '';
        if (is_numeric($unit)) $unit = (string)$unit;
        $unit = trim((string)$unit);

        $monthly = [];
        for ($mi = 0; $mi < 12; $mi++) {
            $mc = $month_cols[$mi];
            if ($mc !== null && isset($row[$mc])) {
                $val = $row[$mc];
                $monthly[] = is_numeric($val) ? $val : null;
            } else {
                $monthly[] = null;
            }
        }

        $total = null;
        if ($total_col !== null && isset($row[$total_col])) {
            $tv = $row[$total_col];
            if (is_numeric($tv)) $total = $tv;
        }

        $target_year = null;
        $target_month = null;
        if (isset($row[2]) && is_numeric($row[2])) $target_year = $row[2];
        if (isset($row[4]) && is_numeric($row[4])) $target_month = $row[4];

        $rows[] = [
            'label' => $label,
            'unit' => $unit,
            'monthly' => $monthly,
            'total' => $total,
            'target_year' => $target_year,
            'target_month' => $target_month,
            'hasData' => count(array_filter($monthly, fn($v) => $v !== null && $v != 0)) > 0
        ];
    }
    return $rows;
}

function normalize_labels(&$all_data) {
    foreach ($all_data as &$sheets) {
        foreach ($sheets as &$sheet_info) {
            foreach ($sheet_info['rows'] as &$row) {
                if (isset(LABEL_NORMALIZE_MAP[$row['label']])) {
                    $row['label'] = LABEL_NORMALIZE_MAP[$row['label']];
                }
            }
        }
    }
}

function fix_trailing_zeros(&$all_data) {
    foreach ($all_data as &$sheets) {
        foreach ($sheets as &$sheet_info) {
            $rows = &$sheet_info['rows'];
            if (empty($rows)) continue;

            $last_real_month = -1;
            for ($mi = 0; $mi < 12; $mi++) {
                $non_zero_count = 0;
                foreach ($rows as $r) {
                    if ($r['monthly'][$mi] !== null && $r['monthly'][$mi] != 0) {
                        $non_zero_count++;
                    }
                }
                if (count($rows) > 0 && $non_zero_count / count($rows) >= 0.30) {
                    $last_real_month = $mi;
                }
            }

            if ($last_real_month < 11) {
                foreach ($rows as &$r) {
                    for ($mi = $last_real_month + 1; $mi < 12; $mi++) {
                        if ($r['monthly'][$mi] === 0) {
                            $r['monthly'][$mi] = null;
                        }
                    }
                    $r['hasData'] = count(array_filter($r['monthly'], fn($v) => $v !== null && $v != 0)) > 0;
                }
            }
        }
    }
}

function load_phsspreadsheet() {
    $composerAutoload = dirname(SCRIPT_DIR) . DIRECTORY_SEPARATOR . 'vendor' . DIRECTORY_SEPARATOR . 'autoload.php';
    if (file_exists($composerAutoload)) {
        try {
            require_once $composerAutoload;
            return class_exists('PhpOffice\\PhpSpreadsheet\\IOFactory');
        } catch (Exception $e) {
            echo "Warning: PhpSpreadsheet not available: " . $e->getMessage() . "\n";
            return false;
        }
    }
    return false;
}

// ============================================================================
// OIS (D) Data Parser
// ============================================================================

function process_ois_files() {
    echo "\n📂 Processing OIS files...\n";
    $ois_dir = RAW_DATA_DIR . DIRECTORY_SEPARATOR . 'OIS';
    $all_data = [];

    if (!is_dir($ois_dir)) {
        echo "   ⚠️  OIS directory not found\n";
        return $all_data;
    }

    if (!class_exists('PhpOffice\\PhpSpreadsheet\\IOFactory')) {
        echo "   ❌ PhpSpreadsheet not available\n";
        return $all_data;
    }

    $files = array_unique(array_merge(
        glob($ois_dir . DIRECTORY_SEPARATOR . 'OIS_*.xls*'),
        glob($ois_dir . DIRECTORY_SEPARATOR . '*_*.xls*')
    ));

    if (empty($files)) {
        echo "   ⚠️  No OIS files found\n";
        return $all_data;
    }

    echo "   Found " . count($files) . " files:\n";
    foreach ($files as $f) {
        echo "      • " . basename($f) . "\n";
    }

    foreach ($files as $filepath) {
        try {
            $year_str = extract_year_from_filename($filepath);
            if (!$year_str) continue;

            $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($filepath);
            if (!isset($all_data[$year_str])) $all_data[$year_str] = [];

            foreach ($spreadsheet->getSheetNames() as $sname) {
                $sheet = $spreadsheet->getSheetByName($sname);
                $highRow = $sheet->getHighestDataRow();
                $highCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($sheet->getHighestDataColumn());
                // Ensure minimum column range (merged cells may report less)
                if ($highCol < 20) $highCol = 20;

                $sheet_data = [];
                for ($r = 1; $r <= $highRow; $r++) {
                    $row = [];
                    for ($c = 1; $c <= $highCol; $c++) {
                        $cell = $sheet->getCell([$c, $r]);
                        // For merged cells, get the value from the master cell
                        $v = $cell->getValue();
                        if ($v === null && $sheet->getMergeCells()) {
                            foreach ($sheet->getMergeCells() as $range) {
                                if ($cell->isInRange($range)) {
                                    $parts = explode(':', $range);
                                    $v = $sheet->getCell($parts[0])->getValue();
                                    break;
                                }
                            }
                        }
                        $row[$c] = $v;
                    }
                    $sheet_data[$r] = $row;
                }

                $header_row = find_month_header_row($sheet_data);
                if ($header_row === null) continue;

                $month_cols = find_month_columns($sheet_data, $header_row);
                if (count(array_filter($month_cols)) === 0) continue;

                $total_col = find_total_column($sheet_data, $header_row);
                $rows = extract_sheet_data($sheet_data, $header_row, $month_cols, $total_col);

                if (!empty($rows)) {
                    $all_data[$year_str][$sname] = ['rows' => $rows];
                }
            }

            $sc = count($all_data[$year_str] ?? []);
            if ($sc > 0) {
                echo "   ✅ ปี $year_str: $sc sheets\n";
            } else {
                // Debug: show sheet names + first few rows preview
                echo "   ⚠️  ปี $year_str: 0 sheets (no month headers found)\n";
                $dbgSheet = $spreadsheet->getSheet(0);
                $dbgHighCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($dbgSheet->getHighestDataColumn());
                echo "      Sheet: " . $dbgSheet->getTitle() . " (cols=$dbgHighCol)\n";
                for ($dr = 1; $dr <= min(5, $dbgSheet->getHighestDataRow()); $dr++) {
                    $dbgRow = '';
                    for ($dc = 1; $dc <= min(16, $dbgHighCol); $dc++) {
                        $dv = $dbgSheet->getCell([$dc, $dr])->getValue();
                        if ($dv !== null) $dbgRow .= "[$dc]" . mb_substr((string)$dv, 0, 12) . ' ';
                    }
                    if ($dbgRow) echo "      R$dr: $dbgRow\n";
                }
            }

            $spreadsheet->disconnectWorksheets();
            unset($spreadsheet);
        } catch (Exception $e) {
            echo "   ❌ " . basename($filepath) . ": " . $e->getMessage() . "\n";
        }
    }

    return $all_data;
}

// ============================================================================
// Real Leak (RL) Data Parser
// ============================================================================

function process_rl_files() {
    echo "\n📂 Processing Real Leak files...\n";
    $rl_dir = RAW_DATA_DIR . DIRECTORY_SEPARATOR . 'Real Leak';
    $rl_data = [];

    if (!is_dir($rl_dir)) {
        echo "   ⚠️  Real Leak directory not found\n";
        return $rl_data;
    }

    $files = array_merge(
        glob($rl_dir . DIRECTORY_SEPARATOR . 'RL*.xlsx'),
        glob($rl_dir . DIRECTORY_SEPARATOR . 'RL*.xls')
    );

    if (empty($files)) {
        echo "   ⚠️  No Real Leak files found\n";
        return $rl_data;
    }

    echo "   Found " . count($files) . " files:\n";
    foreach ($files as $f) {
        echo "      • " . basename($f) . "\n";
    }

    foreach ($files as $filepath) {
        try {
            $year_str = extract_year_from_filename($filepath);
            if (!$year_str) continue;

            $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($filepath);
            $result = [];
            foreach (STANDARD_BRANCHES as $branch) {
                $result[$branch] = [
                    'rate' => array_fill(0, 12, null),
                    'volume' => array_fill(0, 12, null),
                    'production' => array_fill(0, 12, null),
                    'supplied' => array_fill(0, 12, null),
                    'sold' => array_fill(0, 12, null),
                    'blowoff' => array_fill(0, 12, null)
                ];
            }

            foreach ($spreadsheet->getSheetNames() as $sname) {
                $mi = null;
                foreach (RL_MONTH_ABBR as $abbr => $idx) {
                    if (mb_strpos($sname, $abbr) !== false) {
                        $mi = $idx;
                        break;
                    }
                }
                if ($mi === null) continue;

                $sheet = $spreadsheet->getSheetByName($sname);
                $highRow = $sheet->getHighestDataRow();
                $highCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($sheet->getHighestDataColumn());

                $sheet_data = [];
                for ($r = 1; $r <= $highRow; $r++) {
                    $row = [];
                    for ($c = 1; $c <= $highCol; $c++) {
                        $row[$c] = cellVal($sheet, $c, $r);
                    }
                    $sheet_data[$r] = $row;
                }

                $header_row = null;
                $col_branch = 1;
                foreach ($sheet_data as $rn => $row) {
                    foreach ($row as $cn => $val) {
                        if ($val !== null && mb_strpos((string)$val, 'สาขา') !== false) {
                            $header_row = $rn;
                            $col_branch = $cn;
                            break 2;
                        }
                    }
                }
                if ($header_row === null) $header_row = 1;

                $col_map = [];
                $hrow = isset($sheet_data[$header_row]) ? $sheet_data[$header_row] : [];
                $hrow2 = isset($sheet_data[$header_row + 1]) ? $sheet_data[$header_row + 1] : [];
                $hrow3 = isset($sheet_data[$header_row + 2]) ? $sheet_data[$header_row + 2] : [];

                $cols_to_check = array_unique(array_merge(array_keys($hrow), array_keys($hrow2), array_keys($hrow3)));
                foreach ($cols_to_check as $cn) {
                    $h1 = (string)(isset($hrow[$cn]) ? $hrow[$cn] : '');
                    $h2 = (string)(isset($hrow2[$cn]) ? $hrow2[$cn] : '');
                    $h3 = (string)(isset($hrow3[$cn]) ? $hrow3[$cn] : '');
                    $combined = $h1 . ' ' . $h2 . ' ' . $h3;
                    if (mb_strpos($combined, 'น้ำผลิตรวม') !== false) $col_map['production'] = $cn;
                    if (mb_strpos($combined, 'น้ำผลิตจ่ายสุทธิ') !== false && mb_strpos($combined, 'สะสม') === false) $col_map['supplied'] = $cn;
                    if (mb_strpos($combined, 'น้ำจำหน่าย') !== false) $col_map['sold'] = $cn;
                    if (mb_strpos($combined, 'Blow') !== false || mb_strpos($combined, 'blow') !== false) $col_map['blowoff'] = $cn;
                }

                $wl_start_col = null;
                $sub_header_row = null;
                $search_pairs = [[$hrow, $hrow2], [$hrow2, $hrow3]];
                foreach ($search_pairs as $pair) {
                    foreach ($pair[0] as $cn => $val) {
                        if ($val !== null && mb_strpos((string)$val, 'น้ำสูญเสีย') !== false) {
                            $wl_start_col = $cn;
                            $sub_header_row = $pair[1];
                            break 2;
                        }
                    }
                }

                if ($wl_start_col !== null && $sub_header_row !== null) {
                    foreach ($sub_header_row as $cn => $val) {
                        if ($cn < $wl_start_col) continue;
                        $h = (string)$val;
                        if (mb_strpos($h, 'ปริมาณ') !== false && mb_strpos($h, 'สะสม') === false) $col_map['volume'] = $cn;
                        if (mb_strpos($h, 'อัตรา') !== false && mb_strpos($h, 'สะสม') === false) $col_map['rate'] = $cn;
                    }
                }

                // Debug: show what columns were found for first sheet
                if ($mi === 0 || empty($col_map)) {
                    echo "      Sheet '$sname' (month $mi): header_row=$header_row, col_branch=$col_branch\n";
                    echo "      col_map: " . json_encode($col_map) . "\n";
                    if (empty($col_map)) {
                        echo "      ⚠️  No columns mapped! First 3 header rows:\n";
                        for ($dr = $header_row; $dr <= min($header_row + 2, count($sheet_data)); $dr++) {
                            if (isset($sheet_data[$dr])) {
                                $preview = array_slice($sheet_data[$dr], 0, 10, true);
                                echo "        Row $dr: " . json_encode($preview, JSON_UNESCAPED_UNICODE) . "\n";
                            }
                        }
                    }
                }

                $data_start = $header_row + 2;
                foreach ($sheet_data as $rn => $row) {
                    if ($rn <= $header_row) continue;
                    $raw_name = isset($row[$col_branch]) ? $row[$col_branch] : null;
                    if (!is_string($raw_name) && !is_numeric($raw_name)) continue;
                    $raw_name = (string)$raw_name;
                    $branch = normalize_branch_name($raw_name);
                    if ($branch === null) continue;

                    if (isset($col_map['rate']) && isset($row[$col_map['rate']])) {
                        $val = $row[$col_map['rate']];
                        if (is_numeric($val)) $result[$branch]['rate'][$mi] = $val;
                    }
                    if (isset($col_map['volume']) && isset($row[$col_map['volume']])) {
                        $val = $row[$col_map['volume']];
                        if (is_numeric($val)) $result[$branch]['volume'][$mi] = $val;
                    }
                    if (isset($col_map['production']) && isset($row[$col_map['production']])) {
                        $val = $row[$col_map['production']];
                        if (is_numeric($val)) $result[$branch]['production'][$mi] = $val;
                    }
                    if (isset($col_map['supplied']) && isset($row[$col_map['supplied']])) {
                        $val = $row[$col_map['supplied']];
                        if (is_numeric($val)) $result[$branch]['supplied'][$mi] = $val;
                    }
                    if (isset($col_map['sold']) && isset($row[$col_map['sold']])) {
                        $val = $row[$col_map['sold']];
                        if (is_numeric($val)) $result[$branch]['sold'][$mi] = $val;
                    }
                    if (isset($col_map['blowoff']) && isset($row[$col_map['blowoff']])) {
                        $val = $row[$col_map['blowoff']];
                        if (is_numeric($val)) $result[$branch]['blowoff'][$mi] = $val;
                    }
                }

                $rl_data[$year_str] = $result;
            }
            $spreadsheet->disconnectWorksheets();
            unset($spreadsheet);

            $count = 0;
            foreach ($result as $b => $m) {
                if (count(array_filter($m['rate'])) > 0) $count++;
            }
            echo "   ✅ ปี $year_str: $count branches with data\n";
        } catch (Exception $e) {
            echo "   ❌ " . basename($filepath) . ": " . $e->getMessage() . "\n";
        }
    }

    return $rl_data;
}

function build_rl_embedded_data($rl_data) {
    $compact = [];
    foreach ($rl_data as $year_str => $branches) {
        $compact[$year_str] = [];
        foreach ($branches as $branch => $metrics) {
            $compact[$year_str][$branch] = [
                'r' => $metrics['rate'],
                'v' => $metrics['volume'],
                'p' => $metrics['production'],
                's' => $metrics['supplied'],
                'd' => $metrics['sold'],
                'b' => $metrics['blowoff']
            ];
        }
    }
    return json_encode($compact, JSON_UNESCAPED_UNICODE);
}

// ============================================================================
// EU (หน่วยไฟ) Data Parser
// ============================================================================

function process_eu_files() {
    echo "\n📂 Processing EU (หน่วยไฟ) files...\n";
    $eu_dir = RAW_DATA_DIR . DIRECTORY_SEPARATOR . 'หน่วยไฟ';
    $eu_data = [];

    if (!is_dir($eu_dir)) {
        echo "   ⚠️  EU directory not found\n";
        return $eu_data;
    }

    $files = array_merge(
        glob($eu_dir . DIRECTORY_SEPARATOR . 'EU[-_]*.xlsx'),
        glob($eu_dir . DIRECTORY_SEPARATOR . 'EU[-_]*.xls')
    );

    if (empty($files)) {
        echo "   ⚠️  No EU files found\n";
        return $eu_data;
    }

    echo "   Found " . count($files) . " files:\n";
    foreach ($files as $f) {
        echo "      • " . basename($f) . "\n";
    }

    foreach ($files as $filepath) {
        try {
            if (!preg_match('/EU[-_](\d{4})\.xlsx?$/i', basename($filepath), $m)) continue;
            $year_str = $m[1];

            $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($filepath);
            $sheet = $spreadsheet->getSheet(0);
            $highRow = $sheet->getHighestDataRow();
            $highCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($sheet->getHighestDataColumn());

            $eu_branch_col = 2;
            $eu_month_start = 3;
            $eu_data_start = 3;
            $eu_kw_branch = ['สาขา', 'หน่วยงาน', 'ภาพรวม', 'ชื่อสาขา'];
            $eu_kw_month = ['ต.ค.', 'ต.ค', 'พ.ย.', 'ตุลาคม', 'oct'];

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

            $result = [];
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
                $result[$branch_key] = $monthly;
            }

            if (!empty($result)) {
                $eu_data[$year_str] = $result;
                $count = count(array_filter($result, fn($v) => $v !== null));
                echo "   ✅ ปี $year_str: $count branches\n";
            }

            $spreadsheet->disconnectWorksheets();
            unset($spreadsheet);
        } catch (Exception $e) {
            echo "   ❌ " . basename($filepath) . ": " . $e->getMessage() . "\n";
        }
    }

    return $eu_data;
}

function build_eu_embedded_data($eu_data) {
    $compact = [];
    foreach ($eu_data as $year_str => $branches) {
        $compact[$year_str] = $branches;
    }
    return json_encode($compact, JSON_UNESCAPED_UNICODE);
}

// ============================================================================
// MNF Data Parser
// ============================================================================

function process_mnf_files() {
    echo "\n📂 Processing MNF files...\n";
    $mnf_dir = RAW_DATA_DIR . DIRECTORY_SEPARATOR . 'MNF';
    $mnf_data = [];

    if (!is_dir($mnf_dir)) {
        echo "   ⚠️  MNF directory not found\n";
        return $mnf_data;
    }

    $files = array_merge(
        glob($mnf_dir . DIRECTORY_SEPARATOR . 'MNF*.xlsx'),
        glob($mnf_dir . DIRECTORY_SEPARATOR . 'MNF*.xls')
    );

    if (empty($files)) {
        echo "   ⚠️  No MNF files found\n";
        return $mnf_data;
    }

    echo "   Found " . count($files) . " files:\n";
    foreach ($files as $f) {
        echo "      • " . basename($f) . "\n";
    }

    foreach ($files as $filepath) {
        try {
            if (!preg_match('/MNF[-_](\d{4})/', basename($filepath), $m)) continue;
            $year_str = $m[1];

            $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($filepath);
            $result = [];

            foreach ($spreadsheet->getSheetNames() as $sn) {
                if ($sn === 'รวมกราฟสาขา') continue;

                if ($sn === 'ภาพรวมเขต') {
                    $branch_key = '__regional__';
                    $data_start_row = 2;
                } elseif (isset(MNF_SHEET_MAP[$sn])) {
                    $branch_key = MNF_SHEET_MAP[$sn];
                    $data_start_row = 3;
                } else {
                    continue;
                }

                $sheet = $spreadsheet->getSheetByName($sn);
                $highRow = $sheet->getHighestDataRow();

                $metrics = [
                    'actual' => array_fill(0, 12, null),
                    'acceptable' => array_fill(0, 12, null),
                    'target' => array_fill(0, 12, null),
                    'production' => array_fill(0, 12, null),
                ];

                for ($rn = $data_start_row; $rn <= $highRow; $rn++) {
                    $label = cellVal($sheet, 1, $rn) ?? '';
                    if (!is_string($label)) $label = (string)$label;
                    $label = trim($label);

                    $metric_key = null;
                    foreach (MNF_ROW_MAP as $known_label => $key) {
                        if (mb_strpos($label, $known_label) !== false) {
                            $metric_key = $key;
                            break;
                        }
                    }

                    if (!$metric_key) continue;

                    for ($mi = 0; $mi < 12; $mi++) {
                        $col = 2 + $mi;
                        $val = cellCalc($sheet, $col, $rn);
                        if (is_numeric($val) && !is_bool($val)) {
                            if ($metric_key === 'actual' && $val == 0) {
                                $metrics[$metric_key][$mi] = null;
                            } else {
                                $metrics[$metric_key][$mi] = round((float)$val, 4);
                            }
                        }
                    }
                }

                $result[$branch_key] = $metrics;
            }

            if (!empty($result)) {
                $mnf_data[$year_str] = $result;
                echo "   ✅ ปี $year_str: " . count($result) . " branches\n";
            }

            $spreadsheet->disconnectWorksheets();
            unset($spreadsheet);
        } catch (Exception $e) {
            echo "   ❌ " . basename($filepath) . ": " . $e->getMessage() . "\n";
        }
    }

    return $mnf_data;
}

function build_mnf_embedded_data($mnf_data) {
    $compact = [];
    foreach ($mnf_data as $year_str => $branches) {
        $compact[$year_str] = [];
        foreach ($branches as $branch => $metrics) {
            $compact[$year_str][$branch] = [
                'a' => $metrics['actual'],
                'c' => $metrics['acceptable'],
                't' => $metrics['target'],
                'p' => $metrics['production'],
            ];
        }
    }
    return json_encode($compact, JSON_UNESCAPED_UNICODE);
}

// ============================================================================
// KPI Data Parser
// ============================================================================

function process_kpi_files() {
    echo "\n📂 Processing KPI files...\n";
    $kpi_dir = RAW_DATA_DIR . DIRECTORY_SEPARATOR . 'เกณฑ์วัดน้ำสูญเสีย';
    $kpi_dir2 = RAW_DATA_DIR . DIRECTORY_SEPARATOR . 'เกณฑ์ชี้วัด';
    $kpi_data = [];

    $dirs_to_check = [];
    if (is_dir($kpi_dir)) $dirs_to_check[] = $kpi_dir;
    if (is_dir($kpi_dir2)) $dirs_to_check[] = $kpi_dir2;

    if (empty($dirs_to_check)) {
        echo "   ⚠️  KPI directories not found\n";
        return $kpi_data;
    }

    foreach ($dirs_to_check as $kd) {
        $files = array_merge(
            glob($kd . DIRECTORY_SEPARATOR . 'KPI*.xlsx'),
            glob($kd . DIRECTORY_SEPARATOR . 'KPI*.xls')
        );

        if (empty($files)) continue;

        echo "   Found " . count($files) . " files in " . basename($kd) . ":\n";
        foreach ($files as $f) {
            echo "      • " . basename($f) . "\n";
        }

        foreach ($files as $filepath) {
            if (strpos(basename($filepath), '~$') === 0) continue;

            try {
                if (!preg_match('/(\d{4})/', basename($filepath), $m)) continue;
                $year_str = $m[1];

                $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($filepath);
                $sheet = $spreadsheet->getSheet(0);
                $highRow = $sheet->getHighestDataRow();
                $highCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($sheet->getHighestDataColumn());

                $header_row = null;
                for ($r = 1; $r <= min($highRow, 15); $r++) {
                    for ($c = 1; $c <= min($highCol, 10); $c++) {
                        $v = cellVal($sheet, $c, $r);
                        if (is_string($v) && mb_strpos($v, 'สาขา') !== false) {
                            $header_row = $r;
                            break 2;
                        }
                    }
                }
                if ($header_row === null) continue;

                $result = [];
                $data_start = $header_row + 2;
                for ($r = $data_start; $r <= $highRow; $r++) {
                    $branch_raw = cellVal($sheet, 2, $r);
                    if (!$branch_raw) {
                        $c0 = cellVal($sheet, 1, $r);
                        if (is_string($c0) && mb_strpos($c0, 'รวม') !== false) {
                            $branch_raw = $c0;
                        } else {
                            continue;
                        }
                    }

                    $branch_name = normalize_kpi_branch((string)$branch_raw);
                    if (!$branch_name) continue;

                    $target = to_float(cellCalc($sheet, 3, $r));
                    $l1 = to_float(cellCalc($sheet, 4, $r));
                    $l2 = to_float(cellCalc($sheet, 5, $r));
                    $l3 = to_float(cellCalc($sheet, 6, $r));
                    $l4 = to_float(cellCalc($sheet, 7, $r));
                    $l5 = to_float(cellCalc($sheet, 8, $r));
                    $actual = to_float(cellCalc($sheet, 9, $r));

                    if ($target === null && $l1 === null) continue;

                    $result[$branch_name] = [
                        'target' => $target,
                        'levels' => [$l1, $l2, $l3, $l4, $l5],
                        'actual' => $actual
                    ];
                }

                if (!empty($result)) {
                    $kpi_data[$year_str] = $result;
                    echo "   ✅ ปี $year_str: " . count($result) . " branches\n";
                }

                $spreadsheet->disconnectWorksheets();
                unset($spreadsheet);
            } catch (Exception $e) {
                echo "   ❌ " . basename($filepath) . ": " . $e->getMessage() . "\n";
            }
        }
    }

    return $kpi_data;
}

function normalize_kpi_branch($name) {
    $mapping = [
        'ชลบุรี' => 'ชลบุรี(พ)', 'พัทยา' => 'พัทยา(พ)', 'บ้านบึง' => 'บ้านบึง',
        'พนัสนิคม' => 'พนัสนิคม', 'ศรีราชา' => 'ศรีราชา', 'แหลมฉบัง' => 'แหลมฉบัง',
        'ฉะเชิงเทรา' => 'ฉะเชิงเทรา', 'บางปะกง' => 'บางปะกง', 'บางคล้า' => 'บางคล้า',
        'พนมสารคาม' => 'พนมสารคาม', 'ระยอง' => 'ระยอง', 'บ้านฉาง' => 'บ้านฉาง',
        'ปากน้ำประแสร์' => 'ปากน้ำประแสร์', 'จันทบุรี' => 'จันทบุรี', 'ขลุง' => 'ขลุง',
        'ตราด' => 'ตราด', 'คลองใหญ่' => 'คลองใหญ่', 'สระแก้ว' => 'สระแก้ว',
        'วัฒนานคร' => 'วัฒนานคร', 'อรัญประเทศ' => 'อรัญประเทศ',
        'ปราจีนบุรี' => 'ปราจีนบุรี', 'กบินทร์บุรี' => 'กบินทร์บุรี',
    ];
    $name = trim($name);
    if (isset($mapping[$name])) return $mapping[$name];
    if (mb_strpos($name, 'รวม') !== false) return '__regional__';
    foreach ($mapping as $kn => $sn) {
        if (mb_strpos($name, $kn) !== false || mb_strpos($kn, $name) !== false) return $sn;
    }
    return $name;
}

function to_float($val) {
    if ($val === null) return null;
    if (is_numeric($val) && !is_bool($val)) return (float)$val;
    $s = str_replace(',', '', trim((string)$val));
    return is_numeric($s) ? (float)$s : null;
}

function build_kpi_embedded_data($kpi_data) {
    $compact = [];
    foreach ($kpi_data as $year_str => $branches) {
        $compact[$year_str] = [];
        foreach ($branches as $branch => $info) {
            $compact[$year_str][$branch] = [
                't' => $info['target'],
                'l' => $info['levels'],
                'a' => $info['actual']
            ];
        }
    }
    return json_encode($compact, JSON_UNESCAPED_UNICODE);
}

// ============================================================================
// P3 Data Parser
// ============================================================================

function process_p3_files() {
    echo "\n📂 Processing P3 files...\n";
    $p3_dir = RAW_DATA_DIR . DIRECTORY_SEPARATOR . 'P3';
    $result = [];

    if (!is_dir($p3_dir)) {
        echo "   ⚠️  P3 directory not found\n";
        return $result;
    }

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

            $headerRow = null;
            for ($r = 1; $r <= min($highRow, 10); $r++) {
                $v = cellVal($sheet, 1, $r);
                if (is_string($v) && mb_strpos($v, 'พื้นที่') !== false) {
                    $headerRow = $r;
                    break;
                }
            }
            if (!$headerRow) {
                $spreadsheet->disconnectWorksheets();
                return $points;
            }

            for ($r = $headerRow + 1; $r <= $highRow; $r++) {
                $name = cellVal($sheet, 1, $r);
                if (!is_string($name) || mb_strpos($name, 'P3') === false) continue;
                $name = clean_p3_name($name);

                $avg_prev = p3_val(cellCalc($sheet, 2, $r));
                $avg_day = p3_val(cellCalc($sheet, 3, $r));

                $hourly = [];
                for ($col = 4; $col <= 27; $col++) {
                    $hourly[] = p3_val(cellCalc($sheet, $col, $r));
                }

                $points[] = [
                    'n' => $name,
                    'p' => $avg_prev,
                    'a' => $avg_day,
                    'h' => $hourly
                ];
            }

            $spreadsheet->disconnectWorksheets();
        } catch (Exception $e) {
            // Silent skip
        }
        return $points;
    }

    function process_p3_folder($path, $year_key) {
        global $result;
        $files = array_merge(
            glob($path . DIRECTORY_SEPARATOR . 'P3_*.xlsx'),
            glob($path . DIRECTORY_SEPARATOR . '*.xlsx')
        );

        foreach ($files as $fpath) {
            $fname = basename($fpath);
            if (strpos($fname, '~$') === 0) continue;

            $match_flat = preg_match('/^P3_(.+?)_((\d{2})-(\d{2}))\.(xlsx)$/', $fname, $m);
            $match_folder = preg_match('/^(.+?)_((\d{2})-(\d{2}))\.(xlsx)$/', $fname, $m2);

            if (!$match_flat && !$match_folder) continue;

            $match = $match_flat ? $m : $m2;
            $branch = $match[1];
            $month_key = $match[2];
            $yy = intval($match[3]);
            $yk = $year_key ?? (string)(2500 + $yy);

            try {
                $points = parse_p3_xlsx($fpath);
                if (!empty($points)) {
                    if (!isset($result[$yk])) $result[$yk] = [];
                    if (!isset($result[$yk][$month_key])) $result[$yk][$month_key] = [];
                    $result[$yk][$month_key][$branch] = $points;
                }
            } catch (Exception $e) {
                // Silent skip
            }
        }
    }

    // Scan year subfolders
    foreach (scandir($p3_dir) as $year_folder) {
        if ($year_folder[0] === '.') continue;
        $year_path = $p3_dir . DIRECTORY_SEPARATOR . $year_folder;
        if (!is_dir($year_path)) continue;
        process_p3_folder($year_path, $year_folder);
    }

    // Scan flat structure
    process_p3_folder($p3_dir, null);

    if (!empty($result)) {
        $total = 0;
        foreach ($result as $year => $months) {
            foreach ($months as $month => $branches) {
                foreach ($branches as $points) {
                    $total += count($points);
                }
            }
        }
        echo "   ✅ Found " . count($result) . " years, $total P3 points\n";
    } else {
        echo "   ⚠️  No P3 files found or parsed\n";
    }

    return $result;
}

function build_p3_embedded_data($p3_data) {
    if (empty($p3_data)) return '{}';
    $compact = [];
    foreach ($p3_data as $year_str => $months) {
        $compact[$year_str] = [];
        foreach ($months as $month_key => $branches) {
            $compact[$year_str][$month_key] = [];
            foreach ($branches as $branch => $points) {
                $compact[$year_str][$month_key][$branch] = $points;
            }
        }
    }
    return json_encode($compact, JSON_UNESCAPED_UNICODE);
}

// ============================================================================
// Main Dashboard Builder
// ============================================================================

function build_embedded_data($all_data) {
    $compact = [];
    foreach ($all_data as $year_str => $sheets) {
        $compact[$year_str] = [];
        foreach ($sheets as $sname => $sinfo) {
            $compact[$year_str][$sname] = [];
            foreach ($sinfo['rows'] as $r) {
                $compact[$year_str][$sname][] = [
                    'l' => $r['label'],
                    'u' => $r['unit'],
                    'm' => $r['monthly'],
                    't' => $r['total'],
                    'ty' => $r['target_year'] ?? null,
                    'tm' => $r['target_month'] ?? null
                ];
            }
        }
    }
    return json_encode($compact, JSON_UNESCAPED_UNICODE);
}

/**
 * Replace a JavaScript variable declaration in HTML content using strpos/substr.
 * Avoids preg_replace which fails on large content due to PCRE backtrack limits.
 * Finds "var VARNAME={...};" or "const VARNAME={...};" and replaces the entire statement.
 */
function replace_js_var($content, $varName, $replacement) {
    // Try both "var " and "const " prefixes
    $found = false;
    foreach (['var ', 'const '] as $prefix) {
        $needle = $prefix . $varName;
        $offset = 0;
        while (($pos = strpos($content, $needle, $offset)) !== false) {
            // Make sure this is an exact variable name match (next char must be = or whitespace)
            $after_name = $pos + strlen($needle);
            if ($after_name < strlen($content)) {
                $next_ch = $content[$after_name];
                if ($next_ch !== '=' && $next_ch !== ' ' && $next_ch !== "\t" && $next_ch !== "\n" && $next_ch !== "\r") {
                    // Not an exact match (e.g., "var Dashboard" when looking for "var D")
                    $offset = $pos + 1;
                    continue;
                }
            }
            // Find the '=' after the variable name
            $eq_pos = strpos($content, '=', $after_name);
            if ($eq_pos === false) { $offset = $pos + 1; continue; }
            // Find the '{' that starts the object
            $brace_start = strpos($content, '{', $eq_pos);
            if ($brace_start === false) { $offset = $pos + 1; continue; }
            // Count braces to find matching '}'
            $depth = 0;
            $len = strlen($content);
            $i = $brace_start;
            while ($i < $len) {
                $ch = $content[$i];
                if ($ch === '{') $depth++;
                elseif ($ch === '}') {
                    $depth--;
                    if ($depth === 0) break;
                }
                // Skip string literals to avoid counting braces inside strings
                elseif ($ch === "'" || $ch === '"') {
                    $quote = $ch;
                    $i++;
                    while ($i < $len && $content[$i] !== $quote) {
                        if ($content[$i] === '\\') $i++; // skip escaped chars
                        $i++;
                    }
                }
                $i++;
            }
            if ($depth !== 0) { $offset = $pos + 1; continue; }
            // $i now points to the closing '}'
            // Check for semicolon after '}'
            $end_pos = $i + 1;
            if ($end_pos < $len && $content[$end_pos] === ';') {
                $end_pos++;
            }
            // Replace from $pos to $end_pos with the replacement
            $content = substr($content, 0, $pos) . $replacement . substr($content, $end_pos);
            $found = true;
            break 2; // break both while and foreach
        }
    }
    if (!$found) {
        echo "   ⚠️  Variable '$varName' not found in index.html — skipping\n";
    }
    return $content;
}

function build_dashboard($all_data, $rl_data, $eu_data, $mnf_data, $kpi_data, $p3_data) {
    echo "\n🏗️  Building index.html...\n";

    if (!file_exists(INDEX_HTML)) {
        echo "   ❌ index.html not found\n";
        return false;
    }

    $content = file_get_contents(INDEX_HTML);

    // Only replace data in HTML if we actually parsed new data
    // This preserves existing embedded data when source files can't be parsed
    $replacements = [];
    // Check if OIS data actually has sheets (not just empty year arrays)
    $ois_has_data = false;
    foreach ($all_data as $yr => $sheets) {
        if (!empty($sheets)) { $ois_has_data = true; break; }
    }
    if ($ois_has_data) {
        $replacements['D'] = build_embedded_data($all_data);
        echo "   📊 D (OIS): จะอัปเดต\n";
    } else {
        echo "   ⏭️  D (OIS): ไม่มีข้อมูลใหม่ — คงค่าเดิม\n";
    }
    if (!empty($rl_data)) {
        $replacements['RL'] = build_rl_embedded_data($rl_data);
        echo "   📊 RL: จะอัปเดต\n";
    } else {
        echo "   ⏭️  RL: ไม่มีข้อมูลใหม่ — คงค่าเดิม\n";
    }
    if (!empty($eu_data)) {
        $replacements['EU'] = build_eu_embedded_data($eu_data);
        echo "   📊 EU: จะอัปเดต\n";
    } else {
        echo "   ⏭️  EU: ไม่มีข้อมูลใหม่ — คงค่าเดิม\n";
    }
    if (!empty($mnf_data)) {
        $replacements['MNF'] = build_mnf_embedded_data($mnf_data);
        echo "   📊 MNF: จะอัปเดต\n";
    } else {
        echo "   ⏭️  MNF: ไม่มีข้อมูลใหม่ — คงค่าเดิม\n";
    }
    if (!empty($kpi_data)) {
        $replacements['KPI'] = build_kpi_embedded_data($kpi_data);
        echo "   📊 KPI: จะอัปเดต\n";
    } else {
        echo "   ⏭️  KPI: ไม่มีข้อมูลใหม่ — คงค่าเดิม\n";
    }
    if (!empty($p3_data)) {
        $replacements['P3'] = build_p3_embedded_data($p3_data);
        echo "   📊 P3: จะอัปเดต\n";
    } else {
        echo "   ⏭️  P3: ไม่มีข้อมูลใหม่ — คงค่าเดิม\n";
    }

    // Use strpos/substr instead of preg_replace to avoid PCRE backtrack limit on large files
    foreach ($replacements as $varName => $json) {
        $content = replace_js_var($content, $varName, "var $varName=$json;");
        if ($content === false) {
            echo "   ❌ Failed to replace var $varName in index.html\n";
            return false;
        }
    }

    // Update YC (year colors) with all years
    $all_years = array_keys($all_data);
    sort($all_years, SORT_NUMERIC);
    $unique_years = [];
    foreach ($all_years as $y) {
        $unique_years[$y] = true;
        $unique_years[$y - 1] = true;
    }
    if (!empty($all_years)) {
        $last_year = (int)end($all_years);
        for ($i = 1; $i < 4; $i++) {
            $unique_years[$last_year + $i] = true;
        }
    }

    $sorted_years = array_keys($unique_years);
    sort($sorted_years);

    $colors = [
        ['rgba(59,130,246,0.15)', '#3b82f6'],
        ['rgba(239,68,68,0.15)', '#ef4444'],
        ['rgba(34,197,94,0.15)', '#22c55e'],
        ['rgba(168,85,247,0.15)', '#a855f7'],
        ['rgba(249,115,22,0.15)', '#f97316'],
        ['rgba(6,182,212,0.15)', '#06b6d4'],
        ['rgba(236,72,153,0.15)', '#ec4899'],
        ['rgba(202,138,4,0.15)', '#ca8a04'],
        ['rgba(99,102,241,0.15)', '#6366f1'],
        ['rgba(20,184,166,0.15)', '#14b8a6'],
        ['rgba(244,63,94,0.15)', '#f43f5e'],
        ['rgba(139,92,246,0.15)', '#8b5cf6'],
    ];

    $yc_lines = "const YC={\n";
    foreach ($sorted_years as $idx => $yr) {
        $ci = $idx % count($colors);
        list($bg, $border) = $colors[$ci];
        $yc_lines .= "    $yr:{bg:'$bg',border:'$border'},\n";
    }
    $yc_lines .= "};";

    $content = replace_js_var($content, 'YC', $yc_lines);

    if ($content !== false && file_put_contents(INDEX_HTML, $content)) {
        echo "   ✅ index.html updated\n";
        return true;
    } else {
        echo "   ❌ Failed to write index.html\n";
        return false;
    }
}

// ============================================================================
// Main Execution
// ============================================================================

function main() {
    echo str_repeat("=", 60) . "\n";
    echo "  🏗️  Dashboard Builder - PHP CLI Edition\n";
    echo str_repeat("=", 60) . "\n";

    // Check for PhpSpreadsheet
    if (!load_phsspreadsheet()) {
        echo "\n❌ PhpSpreadsheet is not available.\n";
        echo "   Please install it via: composer install\n";
        return;
    }

    // Process each category
    $all_data = process_ois_files();

    if (!empty($all_data)) {
        echo "\n🔧 Normalizing labels...\n";
        normalize_labels($all_data);

        echo "🔧 Fixing trailing zeros...\n";
        fix_trailing_zeros($all_data);
    }

    $rl_data = process_rl_files();
    $eu_data = process_eu_files();
    $mnf_data = process_mnf_files();
    $kpi_data = process_kpi_files();
    $p3_data = process_p3_files();

    // Build dashboard
    build_dashboard($all_data, $rl_data, $eu_data, $mnf_data, $kpi_data, $p3_data);

    // Summary
    echo "\n" . str_repeat("=", 60) . "\n";
    echo "  ✅ Complete!\n";
    if (!empty($all_data)) {
        echo "  📅 OIS years: " . implode(", ", array_keys($all_data)) . "\n";
    }
    if (!empty($rl_data)) {
        echo "  📅 Real Leak years: " . implode(", ", array_keys($rl_data)) . "\n";
    }
    if (!empty($eu_data)) {
        echo "  📅 EU years: " . implode(", ", array_keys($eu_data)) . "\n";
    }
    if (!empty($mnf_data)) {
        echo "  📅 MNF years: " . implode(", ", array_keys($mnf_data)) . "\n";
    }
    if (!empty($kpi_data)) {
        echo "  📅 KPI years: " . implode(", ", array_keys($kpi_data)) . "\n";
    }
    if (!empty($p3_data)) {
        echo "  📅 P3 years: " . implode(", ", array_keys($p3_data)) . "\n";
    }
    echo "  📄 Open index.html in browser to view results\n";
    echo str_repeat("=", 60) . "\n";
}

main();
