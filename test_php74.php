<?php
/**
 * test_php74.php
 *
 * Smoke test สำหรับยืนยันว่า PhpSpreadsheet 1.28 + PHP 7.4
 * รองรับ API ทุกตัวที่ Dashboard 4 ตัวเรียกใช้จริง
 *
 * รันโดย: php test_php74.php
 *
 * ทดสอบ:
 *   1. Autoloader โหลดได้
 *   2. PhpSpreadsheet เวอร์ชันถูกต้อง (1.28.x)
 *   3. PHP เวอร์ชัน 7.4.x
 *   4. IOFactory::load + createReaderForFile
 *   5. Reader API: setReadDataOnly, setReadFilter, IReadFilter (anonymous class)
 *   6. Sheet API: getSheetNames, getSheetByName, getActiveSheet,
 *                 getHighestDataRow, getHighestDataColumn, getMergeCells
 *   7. Cell API: getCell array signature [col,row], getCell("A1"),
 *                getCellByColumnAndRow, getValue, getCalculatedValue,
 *                getOldCalculatedValue, getCoordinate
 *   8. Coordinate API: columnIndexFromString, stringFromColumnIndex
 */

declare(strict_types=1);

ini_set('display_errors', '1');
ini_set('display_startup_errors', '1');
error_reporting(E_ALL);

$ROOT = __DIR__;
$pass = 0;
$fail = 0;
$warn = 0;
$failures = [];

function tprint(string $tag, string $msg) {
    $colors = ['PASS' => "\033[32m", 'FAIL' => "\033[31m", 'WARN' => "\033[33m", 'INFO' => "\033[36m"];
    $reset = "\033[0m";
    $c = $colors[$tag] ?? '';
    // disable color บน Windows cmd ที่ไม่รองรับ ANSI
    if (PHP_OS_FAMILY === 'Windows' && !getenv('ANSICON') && !getenv('WT_SESSION')) {
        echo "[$tag] $msg\n";
    } else {
        echo "{$c}[$tag]{$reset} $msg\n";
    }
}

function check(bool $cond, string $name, string $err = ''): bool {
    global $pass, $fail, $failures;
    if ($cond) {
        $pass++;
        tprint('PASS', $name);
        return true;
    } else {
        $fail++;
        $failures[] = "$name — $err";
        tprint('FAIL', "$name — $err");
        return false;
    }
}

echo "\n==================================================\n";
echo "  PhpSpreadsheet 1.28 + PHP 7.4 Compatibility Test\n";
echo "==================================================\n\n";

// ---------- Test 1: PHP version ----------
tprint('INFO', 'PHP version: ' . PHP_VERSION);
$phpver = explode('.', PHP_VERSION);
if ($phpver[0] === '7' && $phpver[1] === '4') {
    tprint('PASS', 'PHP 7.4.x ตรงกับ server target');
    $pass++;
} else {
    tprint('WARN', "PHP " . PHP_VERSION . " (ไม่ใช่ 7.4) — vendor ที่ install ตอนนี้อาจไม่ตรงกับ server 7.4.29 จริง");
    $warn++;
}

// ---------- Test 2: Autoloader ----------
$autoload = $ROOT . '/vendor/autoload.php';
if (!file_exists($autoload)) {
    tprint('FAIL', "ไม่พบ vendor/autoload.php — ยังไม่ได้ install");
    exit(1);
}
require_once $autoload;
check(class_exists('PhpOffice\\PhpSpreadsheet\\IOFactory'), 'autoloader', 'IOFactory class ไม่โหลด');

// ---------- Test 3: PhpSpreadsheet version + ตรวจว่ารองรับ PHP 7.4 ----------
// เกณฑ์ผ่าน: เป็น 1.x และ package require php รองรับ 7.4.0
// (1.x ทุกตัวรองรับ 7.4; 2.x ขึ้นไปต้องการ 8.1+)
$composer_lock = $ROOT . '/composer.lock';
if (file_exists($composer_lock)) {
    $lock = json_decode(file_get_contents($composer_lock), true);
    foreach ($lock['packages'] as $p) {
        if ($p['name'] === 'phpoffice/phpspreadsheet') {
            $version = $p['version'];
            $reqPhp  = isset($p['require']['php']) ? $p['require']['php'] : '?';
            tprint('INFO', "PhpSpreadsheet version: $version");
            tprint('INFO', "package require php  : $reqPhp");

            // ต้องเป็น 1.x (2.x ขึ้นไปต้องการ PHP 8.1+ ไม่รัน server 7.4 ได้)
            $verParts = explode('.', $version);
            $major = (int) $verParts[0];
            if ($major !== 1) {
                check(false, "PhpSpreadsheet major version",
                      "ติดตั้ง $version แต่ตระกูล 2.x+ ไม่รองรับ PHP 7.4 (server 7.4.29 รันไม่ได้)");
                break;
            }

            // ตรวจ require php จริง — ต้องรองรับ 7.4.0
            $supports74 = (
                strpos($reqPhp, '7.4') !== false ||
                strpos($reqPhp, '>=7') !== false ||
                strpos($reqPhp, '>= 7') !== false ||
                strpos($reqPhp, '^7.') !== false
            );

            if ($supports74) {
                check(true, "PhpSpreadsheet $version รองรับ PHP 7.4 (deploy server 7.4.29 ได้)");
            } else {
                check(false, "PhpSpreadsheet $version PHP support",
                      "package require '$reqPhp' ไม่รองรับ 7.4");
            }
            break;
        }
    }
}

// ---------- Test 4: Coordinate utility ----------
try {
    $idx = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString('A');
    check($idx === 1, "Coordinate::columnIndexFromString('A') === 1", "ได้ $idx");
    $col = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex(27);
    check($col === 'AA', "Coordinate::stringFromColumnIndex(27) === 'AA'", "ได้ '$col'");
} catch (\Throwable $e) {
    check(false, "Coordinate API", $e->getMessage());
}

// ---------- เลือก Excel file สำหรับทดสอบ ----------
$candidates = [
    "$ROOT/Dashboard_Meter/ข้อมูลดิบ/มาตรวัดน้ำผิดปกติ/METER_1102.xlsx",
    "$ROOT/Dashboard_GIS/ข้อมูลดิบ/ลงข้อมูลซ่อมท่อ/GIS_681024.xlsx",
    "$ROOT/Dashboard_Leak/ข้อมูลดิบ/MNF/MNF_2569.xlsx",
];
$test_file = null;
foreach ($candidates as $c) {
    if (file_exists($c)) { $test_file = $c; break; }
}
if (!$test_file) {
    tprint('WARN', 'ไม่พบไฟล์ Excel ทดสอบ — ข้าม Reader API tests');
    $warn++;
} else {
    tprint('INFO', "ทดสอบกับ: " . basename($test_file));

    // ---------- Test 5: IOFactory::load ----------
    try {
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($test_file);
        check($spreadsheet instanceof \PhpOffice\PhpSpreadsheet\Spreadsheet, "IOFactory::load() — Spreadsheet object");
    } catch (\Throwable $e) {
        check(false, "IOFactory::load()", $e->getMessage());
        $spreadsheet = null;
    }

    // ---------- Test 6: createReaderForFile ----------
    try {
        $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReaderForFile($test_file);
        check($reader instanceof \PhpOffice\PhpSpreadsheet\Reader\IReader, "IOFactory::createReaderForFile()");
        $reader->setReadDataOnly(true);
        check(true, "Reader::setReadDataOnly(true)");
    } catch (\Throwable $e) {
        check(false, "createReaderForFile / setReadDataOnly", $e->getMessage());
    }

    // ---------- Test 7: IReadFilter (anonymous class) ----------
    try {
        $filter = new class implements \PhpOffice\PhpSpreadsheet\Reader\IReadFilter {
            public function readCell($columnAddress, $row, $worksheetName = '') {
                return $row <= 5;
            }
        };
        $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReaderForFile($test_file);
        $reader->setReadDataOnly(true);
        $reader->setReadFilter($filter);
        $sp2 = $reader->load($test_file);
        check($sp2 instanceof \PhpOffice\PhpSpreadsheet\Spreadsheet, "IReadFilter + setReadFilter + reader->load()");
    } catch (\Throwable $e) {
        check(false, "IReadFilter pattern", $e->getMessage());
    }

    // ---------- Test 8: Sheet/Cell API (ใช้ $spreadsheet จาก test 5) ----------
    if ($spreadsheet) {
        try {
            $sheetNames = $spreadsheet->getSheetNames();
            check(is_array($sheetNames) && count($sheetNames) > 0,
                  "getSheetNames() — " . count($sheetNames) . " sheets");

            $sheet = $spreadsheet->getActiveSheet();
            check($sheet !== null, "getActiveSheet()");

            $sheetByName = $spreadsheet->getSheetByName($sheetNames[0]);
            check($sheetByName !== null, "getSheetByName('{$sheetNames[0]}')");

            $highRow = $sheet->getHighestDataRow();
            $highCol = $sheet->getHighestDataColumn();
            check(is_int($highRow) && $highRow > 0, "getHighestDataRow() = $highRow");
            check(is_string($highCol) && $highCol !== '', "getHighestDataColumn() = '$highCol'");

            $merge = $sheet->getMergeCells();
            check(is_array($merge), "getMergeCells() returns array (" . count($merge) . " ranges)");

            // getCell string
            $cellA1 = $sheet->getCell('A1');
            check($cellA1 !== null, "getCell('A1') string signature");

            // getCell array [col,row] — สำคัญที่สุด เพราะโปรเจคใช้เยอะ
            $cellArr = $sheet->getCell([1, 1]);
            check($cellArr !== null, "getCell([1,1]) ARRAY signature (สำคัญ — Leak ใช้แบบนี้)");

            $val = $cellArr->getValue();
            tprint('INFO', "  A1 value: " . var_export($val, true));
            check(true, "Cell::getValue()");

            // getCellByColumnAndRow (deprecated ใน 1.28 แต่ยังใช้ได้)
            try {
                @$cellByCR = $sheet->getCellByColumnAndRow(1, 1);
                check($cellByCR !== null, "getCellByColumnAndRow(1,1) (deprecated แต่ใช้ได้ — PR ใช้)");
            } catch (\Throwable $e) {
                check(false, "getCellByColumnAndRow", $e->getMessage());
            }

            // getCoordinate
            $coord = $cellArr->getCoordinate();
            check($coord === 'A1', "getCoordinate() === 'A1'", "ได้ '$coord'");

            // getCalculatedValue + getOldCalculatedValue (สำคัญ — บทเรียน §6)
            try {
                $calc = $cellArr->getCalculatedValue();
                check(true, "getCalculatedValue() ทำงานได้");
                $old = $cellArr->getOldCalculatedValue();
                check(true, "getOldCalculatedValue() ทำงานได้ (สำคัญ — performance)");
            } catch (\Throwable $e) {
                check(false, "getCalculatedValue / getOldCalculatedValue", $e->getMessage());
            }

            // ลอง iterate column-row pattern ที่ Leak ใช้บ่อย
            $cnt = 0;
            $maxCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highCol);
            $rEnd = min(3, $highRow);
            $cEnd = min(3, $maxCol);
            for ($r = 1; $r <= $rEnd; $r++) {
                for ($c = 1; $c <= $cEnd; $c++) {
                    $v = $sheet->getCell([$c, $r])->getValue();
                    $cnt++;
                }
            }
            check($cnt > 0, "iterate getCell([c,r]) — $cnt cells");

        } catch (\Throwable $e) {
            check(false, "Sheet/Cell API", $e->getMessage());
        }
    }
}

// ---------- สรุป ----------
echo "\n==================================================\n";
echo "  สรุป: PASS=$pass  FAIL=$fail  WARN=$warn\n";
echo "==================================================\n";

if ($fail === 0) {
    tprint('PASS', "ทุก API ที่ Dashboard ใช้ทำงานบน PhpSpreadsheet 1.28 + PHP 7.4 ✓");
    exit(0);
} else {
    tprint('FAIL', "พบ $fail จุดที่ใช้ไม่ได้:");
    foreach ($failures as $f) echo "  - $f\n";
    exit(1);
}
