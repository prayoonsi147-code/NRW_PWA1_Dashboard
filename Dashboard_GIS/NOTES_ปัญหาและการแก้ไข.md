# Dashboard GIS — บันทึกปัญหาและการแก้ไข

**วันที่**: 5 เมษายน 2569 (2026)

---

## กฎเหล็ก: ห้ามฝังข้อมูลดิบลง index.html

ห้าม hardcode ข้อมูลดิบลงใน index.html โดยตรง ไม่ว่าจะกรณีใด
ข้อมูลทุกชนิดต้องมี: (1) API endpoint, (2) build function ใน build_dashboard.php
ฝังได้เฉพาะตอน Push Git (push_to_github.bat) แล้ว restore กลับทันที
ถ้าพบข้อมูลดิบ hardcode อยู่ใน HTML → ต้องแจ้งผู้ใช้ทันที ห้ามปล่อยผ่าน

---

## ปัญหาที่ 1: ข้อมูล Fallback (GitHub Pages) ไม่ตรงกับ Local

**อาการ**: Dashboard บน GitHub Pages แสดงข้อมูลเก่า ไม่อัปเดตตาม Local

**สาเหตุ**: Fallback data ใน index.html ถูก hardcode ไว้ ไม่มีระบบอัปเดตอัตโนมัติ

**แก้ไข**: สร้าง `build_dashboard.php` ที่อ่าน SQLite โดยตรงแล้วอัปเดต fallback data ใน index.html อัตโนมัติ ทำงานร่วมกับ `push_to_github.bat` (Step 2: BUILD)

**ป้องกัน**: ทุกครั้งที่ push ขึ้น GitHub ให้รัน `push_to_github.bat` ซึ่งจะเรียก `build_dashboard.php` ก่อน push

---

## ปัญหาที่ 2: Cache เก่าทำให้ข้อมูลผิด (สถานะซ่อมไม่ตรง)

**อาการ**: เลขที่งาน R202510004406 แสดงสถานะ "ซ่อมไม่เสร็จ" ทั้งที่ Excel ที่ Upload เข้าไปเป็น "ซ่อมเสร็จ" แล้ว

**สาเหตุ**: ไฟล์ cache เก่า (`ค้างซ่อม_10-68_to_03-69.xlsx.cache.json`) ไม่เคยถูกตรวจสอบว่า Excel ใหม่กว่า cache หรือไม่ ระบบใช้ cache เก่าตลอด แม้ว่า Excel จะถูก Upload ใหม่แล้ว

**แก้ไขใน api.php**:

1. **open_pending_db()**: เพิ่มการตรวจ `filemtime(cache) >= filemtime(excel)` ก่อนใช้ cache ทุกครั้ง ถ้า cache เก่ากว่า Excel จะอ่าน Excel ใหม่ผ่าน PhpSpreadsheet

2. **upload-confirm handler**: เมื่อ Upload ไฟล์ใหม่ จะลบ merged cache (`*_to_*.cache.json`) อัตโนมัติ

**ป้องกัน**: ระบบตรวจสอบความสดของ cache อัตโนมัติแล้ว ไม่ต้องลบ cache ด้วยมือ

---

## ปัญหาที่ 3: SQLite rebuild ได้ 0 records (Dashboard พัง)

**อาการ**: หลังลบ SQLite เก่าเพื่อ rebuild ใหม่ → SQLite ใหม่มี 0 records (45 KB แทนที่จะเป็น ~150 MB)

**สาเหตุหลัก 2 อย่าง**:

### 3a. PHP sqlite3 extension ไม่ได้เปิด

- ไฟล์ `C:\xampp\php\php.ini` มี `;extension=sqlite3` (ถูก comment ไว้)
- `php_sqlite3.dll` มีอยู่ใน `C:\xampp\php\ext\`
- PHP มี PDO SQLite ใช้ได้ แต่ไม่มี class `SQLite3`

**แก้ไข**: เพิ่ม SQLite3 compatibility wrapper ใน api.php (บรรทัด ~99) ที่ใช้ PDO เบื้องหลัง ทำให้โค้ดเดิมทำงานได้โดยไม่ต้องแก้

**ทางเลือกอื่น**: uncomment `extension=sqlite3` ใน php.ini แล้ว restart Apache

### 3b. PhpSpreadsheet หน่วยความจำเกิน 2GB

- `IOFactory::load()` โหลดไฟล์ OLE2 (.xls) พร้อม formatting ทำให้ใช้ RAM เกิน 2GB
- ไฟล์ Excel แต่ละไฟล์ 14-32 MB แต่ PhpSpreadsheet ใช้ RAM 10-50 เท่าของขนาดไฟล์
- Error ถูกจับโดย try/catch → ไม่มี error แสดง → SQLite ว่างเปล่า

**แก้ไข**: เปลี่ยนจาก `IOFactory::load()` เป็น `createReaderForFile()` + `setReadDataOnly(true)` → ลด RAM จาก 2GB+ เหลือ ~324 MB

**ป้องกัน**: ทุกครั้งที่อ่านไฟล์ Excel ขนาดใหญ่ด้วย PhpSpreadsheet **ต้องใช้ `setReadDataOnly(true)` เสมอ**

---

## ปัญหาที่ 4: parse_thai_date() ไม่รองรับ Excel serial date

**อาการ**: ถ้า PhpSpreadsheet คืนค่าวันที่เป็นตัวเลข (Excel serial date) แทน string → ข้อมูลทุกแถวถูกข้าม

**สาเหตุ**: `parse_thai_date()` รองรับแค่ DateTime object กับ string ที่มี `/` (เช่น "01/10/2568")

**แก้ไข**: เพิ่มการรองรับ numeric Excel serial date ใน `parse_thai_date()` โดยใช้ `PhpSpreadsheet\Shared\Date::excelToDateTimeObject()` แล้วแปลง CE → BE (+543)

**หมายเหตุ**: ใน XAMPP ปัจจุบัน ค่าจาก `setReadDataOnly(true)` คืนเป็น string อยู่แล้ว แต่เพิ่มไว้เป็น safety net

---

## สรุปไฟล์ที่แก้ไข

| ไฟล์ | การแก้ไข |
|------|----------|
| `api.php` | SQLite3 wrapper, cache freshness check, setReadDataOnly, parse_thai_date numeric support, upload handler ลบ merged cache |
| `build_dashboard.php` | ไฟล์ใหม่ — auto-update fallback data ใน index.html |
| `index.html` | Fallback data อัปเดตเป็น 05-04-69 |
| `push_to_github.bat` | เพิ่ม Step 2 เรียก build_dashboard.php |

---

## สภาพแวดล้อม

- PHP 8.2.12 (XAMPP)
- PhpSpreadsheet (ผ่าน Composer)
- ไฟล์ Excel เป็น OLE2 (.xls) แต่ตั้งชื่อเป็น .xlsx
- sqlite3 extension: ปิดอยู่ (ใช้ PDO wrapper แทน)
- PDO drivers: mysql, sqlite
