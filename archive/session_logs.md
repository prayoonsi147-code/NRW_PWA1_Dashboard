# Session Logs Archive

> Session logs ทั้งหมดย้ายมาที่นี่ — ไม่ต้องอ่านทุก session ใหม่
>
> Log เก่า session 1-9 (ก่อน 2026-04-04) อยู่ใน `quick_context_2026-04-04.txt`
> Log เก่าของ Dashboard_Leak อยู่ใน `Dashboard_Leak_PROJECT_README_old.md`
> Log เก่าของ GIS อยู่ใน `Dashboard_GIS_NOTES.md`

---

## Session 10 — 2026-04-20 (Aong)

**งานที่ทำ:**
1. Audit ปุ่มรีเซ็ตทั้ง 4 Dashboard — พบ 7 variant ของ behavior (ไม่สม่ำเสมอ)
2. สรุปปัญหาการทำงานข้าม session — Claude ไม่มี memory → เอกสารเก่ายาวเกินไป
3. Consolidate เอกสาร 8 ไฟล์ → `START_HERE.md` ไฟล์เดียว (≤400 บรรทัด)
4. ย้ายเอกสารเก่าเข้า `archive/` + ลบ dead files (__pycache__, _migration_scripts, uploaded_data, data.json/data_embed.js เก่า, requirements.txt, New Text Document.txt, test_path.php, composer-setup.php, start_xampp.bat, .sync_test)

**งานค้าง (ดู START_HERE.md §10 Known Issues):**
- PR pr1/pr2 reset hardcoded → ย้ายไป `resetChartSize(this)`
- GIS pd2/pd3 reset แค่ resetZoom → เพิ่ม restore controls
- GIS hardcode height='450px' → ใช้ `data-default-height`
- Leak `resetChartSize` ไม่ restore controls → เพิ่มให้ครบ spec
- PR updateLSUsageBar undefined error
- Meter normalize_size() ใช้ eval() → lookup table
- AON values = 0 (parser bug เดิม)
- PR + GIS build_dashboard.php safety checks ยังไม่ครบ

---

## Session 11 — 2026-04-26/27 (Aong) — PHP 7.4 Migration

**บริบท:** server กปภ. รัน PHP 7.4.29 แต่ vendor ที่ติดตั้งคือ PhpSpreadsheet 5.5 ซึ่งต้องการ PHP 8.1+ → server รันไม่ได้ ต้องลด vendor ลง

**งานที่ทำ:**

1. **แก้ `composer.json`** จาก `phpspreadsheet ^5.5` → `^1.28` + เพิ่ม `config.platform.php = 7.4.29` + ระบุ `require php ^7.4`
2. **สร้าง 3 script ใหม่** สำหรับ deploy:
   - `install_php74.bat` — backup + git untrack vendor + composer install + smoke test (ASCII-only ทั้งหมด)
   - `rollback_php74.bat` — ย้อนกลับจาก `backup_php74_install/` (มี YES confirm)
   - `check_extensions.php` — ตรวจ ext 14 ตัวที่ PhpSpreadsheet ต้องการครั้งเดียว + ระบุ path php.ini
   - `verify_install.php` — ตรวจเวอร์ชันที่ install
   - `test_php74.php` — smoke test API 21 จุด ทดสอบกับ Excel จริงในข้อมูลดิบ
3. **แก้ `.gitignore`** — เพิ่ม `vendor/` (ก่อนหน้า track 732 ไฟล์ ไม่ควร), `backup_php74_install/`, `**/upload_log.json`
4. **ลบ `Dashboard_Leak/อัพเดท Dashboard.bat`** — script ตาย เรียก python ที่ไม่มีแล้ว
5. **ตรวจ + ยืนยัน:** โค้ด PHP 10 ไฟล์ของโปรเจค **ไม่มี syntax PHP 8-only เลย** (ไม่มี match/?->/named args/str_contains/readonly/enum) — compatible กับ 7.4 อยู่แล้ว → ไม่ต้องแก้โค้ด business logic
6. **ตรวจ API ที่ใช้จริง:** IOFactory::load/createReaderForFile, Reader (setReadDataOnly/setReadFilter/IReadFilter), Sheet (getSheetNames/getSheetByName/getActiveSheet/getHighestData*/getMergeCells), Cell (getCell array+string/getValue/getCalculatedValue/getOldCalculatedValue/getCoordinate/getCellByColumnAndRow), Coordinate (columnIndexFromString/stringFromColumnIndex) → **ทุกตัวมีใน PhpSpreadsheet 1.x ครบ**

**ผล:**
- Aong รัน install_php74.bat สำเร็จบนเครื่อง XAMPP (PHP 8.2.12)
- composer ลง PhpSpreadsheet **1.30.4** (ไม่ใช่ 1.28 — เพราะ `^1.28` = >=1.28 <2.0)
- vendor 1.30.4 require php = `>=7.4.0 <8.5.0` → **ใช้กับ server 7.4.29 ได้แน่นอน** ตรวจจาก vendor/phpoffice/phpspreadsheet/composer.json
- smoke test ผ่าน 21/21 API ทำงานครบ — รวม `getCell([col,row])` ARRAY signature ที่ Leak ใช้, `getOldCalculatedValue()`, IReadFilter anonymous class
- ขั้นต่อไปคือ deploy ผ่านวิธีปกติ (push_to_github.bat / FTP / git pull ขึ้น server)

**ปัญหาที่เจอระหว่างทำ + วิธีแก้ (กันลืม + กันซ้ำ):**

| # | ปัญหา | สาเหตุ | วิธีแก้ |
|---|---|---|---|
| 1 | sandbox install PHP 7.4 ไม่ได้ | sudo blocked + apt blocked + static binary domain ไม่อยู่ใน allowlist | ใช้ Node.js `php-parser` ทำ static check + ส่ง script ให้ Aong รันบน XAMPP แทน |
| 2 | install_php74.bat รอบแรกพัง — `'---' is not recognized`, `'Backup' is not recognized` ฯลฯ | Windows cmd parse คอมเมนต์ภาษาไทย + `for /f` + `^^` + inline `php -r` ที่มี `[`,`]`,`'` ซ้อน — ผิดพร้อมกัน | rewrite **batch ให้ ASCII-only 100%** — ข้อความไทยย้ายไปอยู่ใน PHP script (`check_extensions.php`, `verify_install.php`) |
| 3 | install fail ที่ `[X] PHP extension 'gd' is not enabled` | XAMPP default ปิด extension=gd ไว้ใน php.ini | เปิดด้วย Notepad: `;extension=gd` → `extension=gd` save แล้วรันใหม่ — script ใหม่จะตรวจครบ 14 ext ครั้งเดียว |
| 4 | smoke test FAIL = 1 (`PhpSpreadsheet version`) แม้ทุก API ทำงานได้ | ผม pin `^1.28` ใน composer.json (>=1.28 <2.0) → composer ดึง 1.30.4 มา; ผมเขียน check ใน test_php74.php strict ว่าต้องเป็น 1.28.x เป๊ะ | แก้ logic check ให้รับทั้งตระกูล 1.x ที่ require php รองรับ 7.4 (ไม่ pin 1.28 เป๊ะ — 1.30.4 ใหม่กว่า + bug fix มากกว่า) |
| 5 | คิดผิดว่า "1.29 ตัด PHP 7.4" | จำผิด — ที่จริงตระกูล 1.x ทั้งหมดยังรองรับ 7.4; ที่ตัดคือ **2.x** (require 8.1+) | ตรวจ `vendor/phpoffice/phpspreadsheet/composer.json` ของจริงเสมอก่อนตัดสินใจ pin |
| 6 | vendor/ ถูก track ใน git 732 ไฟล์ | `.gitignore` เก่าไม่ได้ ignore vendor → repo บวมโดยใช่เหตุ | เพิ่ม `vendor/` ลง .gitignore + script untrack อัตโนมัติ (`git rm --cached -r vendor`) ใน install_php74.bat |
| 7 | mount cache ของ sandbox bash ไม่ sync กับ file tool | bash อ่านไฟล์เก่าหลัง Edit/Write — server-side cache | ใช้ Read tool (ผ่าน Cowork file tool) ตรวจไฟล์ที่เพิ่งแก้ — ไม่ใช้ bash `cat` ตรวจ |
| 8 | ลบไฟล์ใน mount โดน "Operation not permitted" | sandbox guard | ขออนุญาตผ่าน `mcp__cowork__allow_cowork_file_delete` ก่อนรัน rm |

**ไฟล์ใหม่/แก้ในรอบนี้ (ทั้งหมดอยู่ใน root โปรเจค):**
- ✏️ `composer.json` — pin ^1.28 + platform 7.4.29
- ✏️ `.gitignore` — เพิ่ม vendor/ + backup_php74_install/ + upload_log.json
- ✨ `install_php74.bat` (ASCII-only)
- ✨ `rollback_php74.bat` (ASCII-only)
- ✨ `check_extensions.php` (ตรวจ ext 14 ตัว)
- ✨ `verify_install.php` (ตรวจ version ที่ install)
- ✨ `test_php74.php` (smoke test 21 API)
- 🗑️ `Dashboard_Leak/อัพเดท Dashboard.bat` (เรียก python ที่ไม่มีแล้ว)

**คราวหน้าจะกลับมาทำต่อ — เช็คก่อน:**
1. อ่านไฟล์นี้ (Session 11) ทั้งหมด — รู้บริบท
2. ถ้า Aong เจอปัญหา PHP version อีก: เปิด `vendor/phpoffice/phpspreadsheet/composer.json` ดู `require.php` — ตอบได้ทันทีว่ารองรับ server หรือไม่
3. ถ้า batch script พังบน Windows: ตรวจว่าเป็น ASCII-only ไหม + ไม่มี `for /f` ซับซ้อน + ไม่มี `php -r` ยาวที่มี quote ซ้อน
4. ถ้าจะทำ migration อีก: ใช้ pattern เดียวกัน — backup → script ASCII → smoke test ก่อน push
5. **ห้ามลืม:** rules ใน START_HERE.md ยังเป็น single source of truth — งานค้าง 8 ข้อใน §10 ยังไม่ได้แตะรอบนี้
