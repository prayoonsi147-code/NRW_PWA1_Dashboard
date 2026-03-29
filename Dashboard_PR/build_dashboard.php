<?php
/**
 * build_dashboard.php - PHP equivalent of build_dashboard.py
 * Reads PR data from Excel files and embeds into index.html
 */

// ─── Setup ─────────────────────────────────────────────────────────────────────

require __DIR__ . '/../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

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
    } catch (Exception $e) {
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
        if (!preg_match('/(\d{2})-(\d{2})/', $fname, $m)) {
            continue;
        }
        $year_be = intval($m[1]);
        $month = intval($m[2]);
        $month_key = sprintf("%02d-%02d", $year_be, $month);

        $fpath = $DATA_DIR . DIRECTORY_SEPARATOR . $fname;

        try {
            $spreadsheet = IOFactory::load($fpath);
        } catch (Exception $e) {
            echo "  [WARNING] ข้ามไฟล์เสีย: $fname (" . $e->getMessage() . ")\n";
            continue;
        }

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
                $col_letter = chr(64 + $col); // Convert to column letter
                $col_letter2 = chr(64 + $col + 1);
                $col_letter3 = chr(64 + $col + 2);

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
            $col_letter = chr(64 + $col);
            $col_letter2 = chr(64 + $col + 1);
            $col_letter3 = chr(64 + $col + 2);

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
 * Build dashboard: read data and embed into HTML
 */
function build() {
    global $HTML_TEMPLATE;

    echo str_repeat("=", 50) . "\n";
    echo "  Build Dashboard งานลูกค้าสัมพันธ์ กปภ.เขต 1\n";
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
} catch (Exception $e) {
    echo "[ERROR] " . $e->getMessage() . "\n";
    echo $e->getTraceAsString() . "\n";
    exit(1);
}
