<?php
/**
 * check_extensions.php — ตรวจ PHP extensions ที่ PhpSpreadsheet 1.28 ต้องการ
 * เรียกจาก install_php74.bat ก่อน composer install
 *
 * Exit code: 0 = ครบทุกตัว, 1 = ขาดอย่างน้อย 1 ตัว
 */

declare(strict_types=1);

// PhpSpreadsheet 1.28 require:
$required = [
    'ctype'      => 'มากับ PHP core ไม่ต้องเปิดเพิ่ม',
    'dom'        => 'มากับ PHP core ไม่ต้องเปิดเพิ่ม',
    'fileinfo'   => 'หาบรรทัด ;extension=fileinfo ใน php.ini → ลบ ; หน้าออก',
    'filter'     => 'มากับ PHP core ไม่ต้องเปิดเพิ่ม',
    'gd'         => 'หาบรรทัด ;extension=gd ใน php.ini → ลบ ; หน้าออก',
    'iconv'      => 'มากับ PHP core (Windows: extension=iconv ใน php.ini)',
    'libxml'     => 'มากับ PHP core',
    'mbstring'   => 'หาบรรทัด ;extension=mbstring ใน php.ini → ลบ ; หน้าออก',
    'simplexml'  => 'มากับ PHP core',
    'xml'        => 'มากับ PHP core',
    'xmlreader'  => 'มากับ PHP core',
    'xmlwriter'  => 'มากับ PHP core',
    'zip'        => 'หาบรรทัด ;extension=zip ใน php.ini → ลบ ; หน้าออก',
    'zlib'       => 'มากับ PHP core',
];

echo "\n";
echo "=== ตรวจ PHP extensions ที่ PhpSpreadsheet 1.28 ต้องการ ===\n";
echo "PHP version : " . PHP_VERSION . "\n";
$ini = php_ini_loaded_file();
echo "php.ini path: " . ($ini ?: '(ไม่พบ)') . "\n";
echo "\n";

$missing = [];
foreach ($required as $ext => $how) {
    if (extension_loaded($ext)) {
        echo "  [OK]  $ext\n";
    } else {
        echo "  [X]   $ext  --  $how\n";
        $missing[] = $ext;
    }
}

echo "\n";
if (empty($missing)) {
    echo "[OK] PHP extensions ครบทุกตัว — install ต่อได้\n";
    exit(0);
}

echo "==============================================================\n";
echo "[X] ขาด " . count($missing) . " extension(s): " . implode(', ', $missing) . "\n";
echo "==============================================================\n\n";
echo "วิธีแก้:\n";
echo "  1. เปิดไฟล์ php.ini ที่อยู่: " . ($ini ?: 'C:\\xampp\\php\\php.ini') . "\n";
echo "     (ใช้ Notepad เปิด — โปรแกรม Run as Administrator ถ้าจำเป็น)\n";
echo "  2. ค้นหาแต่ละบรรทัด \";extension=<ชื่อ>\" ของ extension ที่ขาด\n";
echo "  3. ลบเครื่องหมาย ; หน้าบรรทัดออก\n";
echo "  4. Save php.ini\n";
echo "  5. รัน install_php74.bat อีกครั้ง\n";
echo "\n";
echo "เปิด php.ini ตอนนี้: notepad \"" . ($ini ?: 'C:\\xampp\\php\\php.ini') . "\"\n";
echo "\n";
exit(1);
