<?php
/**
 * verify_install.php — แสดงเวอร์ชัน PhpSpreadsheet ที่ติดตั้ง
 * เรียกจาก install_php74.bat (แยกเป็น PHP เพื่อจัดการ UTF-8 ปลอดภัย)
 */

declare(strict_types=1);

$root = __DIR__;
$lock = $root . '/composer.lock';
$pkgJson = $root . '/vendor/phpoffice/phpspreadsheet/composer.json';

echo "PHP runtime version : " . PHP_VERSION . "\n";

if (!file_exists($lock)) {
    echo "[X] ไม่พบ composer.lock — composer install อาจไม่สำเร็จ\n";
    exit(1);
}

$lockData = json_decode(file_get_contents($lock), true);
if (!is_array($lockData) || !isset($lockData['packages'])) {
    echo "[X] composer.lock เสีย — อ่านไม่ได้\n";
    exit(1);
}

$found = false;
foreach ($lockData['packages'] as $p) {
    if (($p['name'] ?? '') === 'phpoffice/phpspreadsheet') {
        $version = $p['version'] ?? '?';
        $reqPhp  = $p['require']['php'] ?? '?';
        echo "phpoffice/phpspreadsheet : $version\n";
        echo "package require php     : $reqPhp\n";

        if (strpos($version, '1.28') === 0) {
            echo "[OK] เป็น 1.28.x ตามต้องการ\n";
        } elseif (preg_match('/^1\.\d+/', $version)) {
            echo "[!] เป็น 1.x แต่ไม่ใช่ 1.28 — ตรวจ composer.json\n";
        } else {
            echo "[X] ไม่ใช่ 1.x — composer install อาจอ่าน config ผิด\n";
            exit(1);
        }
        $found = true;
        break;
    }
}

if (!$found) {
    echo "[X] ไม่พบ phpoffice/phpspreadsheet ใน composer.lock\n";
    exit(1);
}

if (!file_exists($pkgJson)) {
    echo "[X] ไม่พบ vendor/phpoffice/phpspreadsheet/composer.json\n";
    exit(1);
}

echo "[OK] vendor พร้อมใช้งาน\n";
exit(0);
