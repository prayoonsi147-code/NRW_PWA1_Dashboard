<?php
/**
 * build_dashboard.php - PHP equivalent of build_dashboard.py
 * Reads PR data from Excel files and embeds into index.html
 */

// ─── Error Handling ────────────────────────────────────────────────────────
ini_set('display_errors', '0');
error_reporting(E_ALL);
ini_set('log_errors', '1');

// ─── Setup ─────────────────────────────────────────────────────────────────────

require __DIR__ . '/../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

$SCRIPT_DIR = __DIR__;
$DATA_DIR = $SCRIPT_DIR . DIRECTORY_SEPARATOR . 'uploaded_data' . DIRECTORY_SEPARATOR . 'pr';
$HTML_TEMPLATE = $SCRIPT_DIR . DIRECTORY_SEPARATOR . 'index.html';

// ─── Helper Functions ────────────────────────────────────────────────────────

/**
 * Clean number value from Excel cell
 */
function clean_num($val) {
    if ($val === null) {
        return 0;
    }
    if (is_numeric($val)) {
        return floatval($val);
    }
    $s = str_replace(',', '', str_replace('\xa0', '', trim((string)$val)));
    if ($s === '' || $s === '-') {
        return 0;
    }
    try {
        return floatval($s);
    } catch (\Throwable $e) {
        return 0;
    }
}

/**
 * Read all Excel PR files and return structured data
 */
function read_all_data() {
    global $DATA_DIR;

    $cat_names = [
        'ด้านปริมาณน้ำ',
        'ด้านท่อแตกรั่ว',
        'ด้านคุณภาพน้ำ',
        'ด้านการบริการ',
        'ด้านบุคลากร',
        'การแจ้งเหตุ',
        'ด้านการติดตามเร่งรัดข้อร้องเรียนเดิม',
        'ด้านสอบถามทั่วไป',
        'ความต้องการ ความคาดหวัง และข้อเสนอแนะ',
        'คำชม และอื่นๆ'
    ];

    // Scan for Excel files matching PR_YY-MM.xlsx
    $files = [];
    if (is_dir($DATA_DIR)) {
        $dir_items = scandir($DATA_DIR);
        foreach ($dir_items as $f) {
            if (substr($f, -5) === '.xlsx') {
                $files[] = $f;
            }
        }
    }
    sort($files);

    echo "  พบไฟล์ข้อมูล: " . count($files) . " ไฟล์\n";

    $all_data = [];
    $branches_order = [];

    foreach ($files as $fname) {
        // Extract YY-MM from filename
        $month_key = null;
        if (preg_match('/(\d{2})-(\d{2})/', $fname, $m)) {
            $year_be = intval($m[1]);
            $month = intval($m[2]);
            if ($month >= 1 && $month <= 12) {
                $month_key = sprintf("%02d-%02d", $year_be, $month);
            }
        }

        $fpath = $DATA_DIR . DIRECTORY_SEPARATOR . $fname;

        try {
            $spreadsheet = IOFactory::load($fpath);
        } catch (\Throwable $e) {
            echo "  [WARNING] ข้ามไฟล์เสีย: $fname (" . $e->getMessage() . ")\n";
            continue;
        }

        // If no month_key from filename, try to extract from Excel content
        if (!$month_key) {
            // ╔══════════════════════════════════════════════════════════════╗
            // ║ ⚠️  ACTIVE SHEET ONLY — ทำงานกับชีทแรกเท่านั้น               ║
            // ║                                                              ║
            // ║ getActiveSheet() อ่านชีทแรก/ชีทแอกทีฟเท่านั้น               ║
            // ║ ไม่ท่องชีทอื่น ข้อมูล PR อยู่ในชีทแรก "เรื่องร้องเรียน"    ║
            // ║                                                              ║
            // ║ Sheet: Always first sheet (active/default)                  ║
            // ║ Data source: Month extraction from PR data cells            ║
            // ╚══════════════════════════════════════════════════════════════╝
            $worksheet = $spreadsheet->getActiveSheet();
            // Check first 10 rows for month information
            for ($row_idx = 1; $row_idx <= 10; $row_idx++) {
                if ($month_key) break;
                for ($col_idx = 1; $col_idx <= 15; $col_idx++) {
                    $cell_val = $worksheet->getCellByColumnAndRow($col_idx, $row_idx)->getValue();
                    if (!$cell_val) continue;

                    $cell_str = trim((string)$cell_val);
                    // Try to detect month from Thai month names/abbreviations
                    $thai_months_abbr = [
                        'ม.ค.' => 1, 'ก.พ.' => 2, 'มี.ค.' => 3, 'เม.ย.' => 4, 'พ.ค.' => 5, 'มิ.ย.' => 6,
                        'ก.ค.' => 7, 'ส.ค.' => 8, 'ก.ย.' => 9, 'ต.ค.' => 10, 'พ.ย.' => 11, 'ธ.ค.' => 12,
                    ];
                    $thai_months_full = [
                        'มกราคม' => 1, 'กุมภาพันธ์' => 2, 'มีนาคม' => 3, 'เมษายน' => 4,
                        'พฤษภาคม' => 5, 'มิถุนายน' => 6, 'กรกฎาคม' => 7, 'สิงหาคม' => 8,
                        'กันยายน' => 9, 'ตุลาคม' => 10, 'พฤศจิกายน' => 11, 'ธันวาคม' => 12,
                    ];

                    $month_num = null;
                    foreach ($thai_months_full as $name => $num) {
                        if (mb_strpos($cell_str, $name) !== false) {
                            $month_num = $num;
                            break;
                        }
                    }
                    if (!$month_num) {
                        foreach ($thai_months_abbr as $abbr => $num) {
                            if (strpos($cell_str, $abbr) !== false) {
                                $month_num = $num;
                                break;
                            }
                        }
                    }

                    if ($month_num) {
                        // Try to find year (2 or 4 digits)
                        if (preg_match('/(\d{4})/', $cell_str, $year_match)) {
                            $year_val = intval($year_match[1]);
                            // Convert Buddhist Era (BE) to 2-digit format
                            if ($year_val >= 2500) {
                                $year_be = $year_val - 2500;
                            } else {
                                $year_be = $year_val;
                            }
                            $month_key = sprintf("%02d-%02d", $year_be, $month_num);
                            break 2;
                        } elseif (preg_match('/(\d{2})/', $cell_str, $year_match)) {
                            $year_be = intval($year_match[1]);
                            $month_key = sprintf("%02d-%02d", $year_be, $month_num);
                            break 2;
                        }
                    }
                }
            }
        }

        // Skip file if still no month_key
        if (!$month_key) {
            echo "  [WARNING] ข้ามไฟล์: $fname (ไม่สามารถตรวจจับเดือนได้)\n";
            $spreadsheet->disconnectWorksheets();
            unset($spreadsheet);
            continue;
        }

        // ╔══════════════════════════════════════════════════════════════╗
        // ║ ⚠️  ACTIVE SHEET ONLY — อ่านจากชีทแรก (PR Data)             ║
        // ║                                                              ║
        // ║ getActiveSheet() เข้าถึงชีทแรก "เรื่องร้องเรียน" เท่านั้น   ║
        // ║ ไม่ต้องสแกนชีทอื่น ข้อมูลลูกค้าและหมวดหมู่อยู่ในชีตนี้      ║
        // ║                                                              ║
        // ║ Sheet: First/active sheet (PR category data)                ║
        // ║ Rows processed: 7-29 (branches + regional total)            ║
        // ╚══════════════════════════════════════════════════════════════╝
        $worksheet = $spreadsheet->getActiveSheet();
        $month_data = [];

        // Rows 7-28: branch data (1-indexed in Excel, 0-indexed in PHP code)
        for ($row_idx = 7; $row_idx <= 28; $row_idx++) {
            $branch = $worksheet->getCell("B$row_idx")->getValue();
            if (!$branch) {
                continue;
            }
            $branch = trim((string)$branch);

            $customers = clean_num($worksheet->getCell("C$row_idx")->getValue());
            $categories = [];

            // Columns: 5 (E) + i*3 for i in 0..9
            $col_start = 5;
            foreach ($cat_names as $i => $cat_name) {
                $col = $col_start + $i * 3;
                $col_letter = Coordinate::stringFromColumnIndex($col);
                $col_letter2 = Coordinate::stringFromColumnIndex($col + 1);
                $col_letter3 = Coordinate::stringFromColumnIndex($col + 2);

                $categories[$cat_name] = [
                    'รวม' => clean_num($worksheet->getCell($col_letter . $row_idx)->getValue()),
                    'ไม่เกิน' => clean_num($worksheet->getCell($col_letter2 . $row_idx)->getValue()),
                    'เกิน' => clean_num($worksheet->getCell($col_letter3 . $row_idx)->getValue())
                ];
            }

            // Columns 35, 36, 37 for totals
            $total = clean_num($worksheet->getCell("AI$row_idx")->getValue());
            $total_w = clean_num($worksheet->getCell("AJ$row_idx")->getValue());
            $total_o = clean_num($worksheet->getCell("AK$row_idx")->getValue());

            $month_data[$branch] = [
                'จำนวนลูกค้า' => $customers,
                'categories' => $categories,
                'รวมสาขา' => $total,
                'รวม_ไม่เกิน' => $total_w,
                'รวม_เกิน' => $total_o
            ];

            if (!in_array($branch, $branches_order)) {
                $branches_order[] = $branch;
            }
        }

        // Row 29: Regional total "รวม เขต 1"
        $row_idx = 29;
        $customers = clean_num($worksheet->getCell("C$row_idx")->getValue());
        $categories = [];

        $col_start = 5;
        foreach ($cat_names as $i => $cat_name) {
            $col = $col_start + $i * 3;
            $col_letter = Coordinate::stringFromColumnIndex($col);
            $col_letter2 = Coordinate::stringFromColumnIndex($col + 1);
            $col_letter3 = Coordinate::stringFromColumnIndex($col + 2);

            $categories[$cat_name] = [
                'รวม' => clean_num($worksheet->getCell($col_letter . $row_idx)->getValue()),
                'ไม่เกิน' => clean_num($worksheet->getCell($col_letter2 . $row_idx)->getValue()),
                'เกิน' => clean_num($worksheet->getCell($col_letter3 . $row_idx)->getValue())
            ];
        }

        $total = clean_num($worksheet->getCell("AI$row_idx")->getValue());
        $total_w = clean_num($worksheet->getCell("AJ$row_idx")->getValue());
        $total_o = clean_num($worksheet->getCell("AK$row_idx")->getValue());

        $month_data['รวม เขต 1'] = [
            'จำนวนลูกค้า' => $customers,
            'categories' => $categories,
            'รวมสาขา' => $total,
            'รวม_ไม่เกิน' => $total_w,
            'รวม_เกิน' => $total_o
        ];

        $all_data[$month_key] = $month_data;
        $spreadsheet->disconnectWorksheets();
        unset($spreadsheet);
    }

    // Determine 13-month range: from same-month-last-year to latest month
    $months_sorted = array_keys($all_data);
    sort($months_sorted);

    if (empty($months_sorted)) {
        echo "  [ERROR] ไม่พบข้อมูลเดือนใดๆ!\n";
        exit(1);
    }

    $latest = $months_sorted[count($months_sorted) - 1];
    list($ly_str, $lm_str) = explode('-', $latest);
    $ly = intval($ly_str);
    $lm = intval($lm_str);
    $same_month_ly = sprintf("%02d-%02d", $ly - 1, $lm);

    $months_13 = [];
    foreach ($months_sorted as $m) {
        if ($m >= $same_month_ly) {
            $months_13[] = $m;
        }
    }

    echo "  ช่วงเดือน: " . $months_13[0] . " - " . $months_13[count($months_13) - 1] . " (" . count($months_13) . " เดือน)\n";
    echo "  จำนวนสาขา: " . count($branches_order) . "\n";

    return [
        'months' => $months_13,
        'branches' => $branches_order,
        'all_months' => $months_sorted,
        'data' => $all_data,
        'cat_names' => $cat_names
    ];
}

/**
 * Parse CLI arguments for incremental build:
 *   --only=pr      → process only PR category (skip AON)
 *   --only=aon     → process only AON category (skip PR)
 *   --files=a.xlsx → process only these specific files
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

/**
 * Build dashboard: read data and embed into HTML
 */
function build() {
    global $HTML_TEMPLATE;

    $args = parse_cli_args();
    $only = $args['only'];       // e.g. 'pr', 'aon', or '' (all)
    $only_files = $args['files'];

    echo str_repeat("=", 50) . "\n";
    echo "  Build Dashboard งานลูกค้าสัมพันธ์ กปภ.เขต 1\n";
    if ($only) {
        echo "  ⚡ Incremental build: only=$only" . ($only_files ? " files=" . implode(',', $only_files) : '') . "\n";
    }
    echo str_repeat("=", 50) . "\n";

    echo "\n[1/3] อ่านข้อมูล Excel...\n";
    $data = read_all_data();

    echo "\n[2/3] สร้าง JSON...\n";
    $data_json = json_encode($data, JSON_UNESCAPED_UNICODE);
    echo "  ขนาดข้อมูล: " . number_format(strlen($data_json)) . " bytes\n";

    echo "\n[3/3] Embed ข้อมูลลงใน index.html...\n";

    if (!file_exists($HTML_TEMPLATE)) {
        echo "  [ERROR] ไม่พบไฟล์ $HTML_TEMPLATE\n";
        exit(1);
    }

    $html = file_get_contents($HTML_TEMPLATE);

    // Try placeholder first (fresh template)
    if (strpos($html, 'DASHBOARD_DATA_PLACEHOLDER') !== false) {
        $html = str_replace('DASHBOARD_DATA_PLACEHOLDER', $data_json, $html);
        echo "  (ใช้ placeholder)\n";
    } else {
        // Replace existing embedded DATA (previously built file)
        // DATA is on a single line: const DATA = {...};
        $pattern = '/^const DATA = \{.*\};$/m';
        $new_val = 'const DATA = ' . $data_json . ';';
        $html_new = preg_replace($pattern, $new_val, $html, 1, $count);
        if ($count > 0) {
            $html = $html_new;
            echo "  (แทนที่ const DATA เดิม)\n";
        } else {
            echo "  [WARNING] ไม่พบ DASHBOARD_DATA_PLACEHOLDER หรือ const DATA ใน HTML!\n";
        }
    }

    $output_path = $HTML_TEMPLATE;
    file_put_contents($output_path, $html);

    $size_kb = filesize($output_path) / 1024;
    echo "  บันทึก: $output_path\n";
    echo "  ขนาดไฟล์: " . sprintf("%.1f", $size_kb) . " KB\n";
    echo "\n  เสร็จสิ้น!\n";
}

// ─── Main ──────────────────────────────────────────────────────────────────────

try {
    build();
} catch (\Throwable $e) {
    echo "[ERROR] " . $e->getMessage() . "\n";
    echo $e->getTraceAsString() . "\n";
    exit(1);
}
