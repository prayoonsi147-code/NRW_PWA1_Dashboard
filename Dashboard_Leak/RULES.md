# กฎและบทเรียนสำคัญ (RULES & LESSONS LEARNED)
## อ่านไฟล์นี้ทุกครั้งที่เริ่มเซสชันใหม่ — ก่อน PROJECT_README.md

> ไฟล์นี้รวบรวมบทเรียนจากการแก้ปัญหาจริง เพื่อไม่ให้ทำผิดซ้ำ

---

## 1. PHP Error Handling — ห้ามใช้ ini_set อย่างเดียว

**ปัญหา:** Upload ไฟล์แล้วได้ JSON error "Unexpected token '<'" เพราะ PHP แสดง HTML error ก่อน JSON output

**สาเหตุ:** บาง PHP error เกิดก่อน script execute (เช่น max_file_uploads exceeded) ทำให้ `ini_set('display_errors','0')` ไม่ทัน

**แก้ไขที่ถูกต้อง:** ต้องตั้งค่าใน `.htaccess` เสมอ:
```apache
php_flag display_errors Off
php_flag html_errors Off
php_value log_errors On
```

**กฎ:** ทุก dashboard ที่มี PHP API ต้องมี `.htaccess` กำกับ — อย่าพึ่ง ini_set() อย่างเดียว

---

## 2. Multi-file Upload — ต้องใช้ files[] (มีวงเล็บเหลี่ยม)

**ปัญหา:** Upload 23 ไฟล์ แต่ PHP ได้รับแค่ไฟล์เดียว

**สาเหตุ:** JavaScript ใช้ `formData.append('files', file)` — ไม่มี `[]` ทำให้ PHP เก็บแค่ไฟล์สุดท้าย

**แก้ไขที่ถูกต้อง:**
```javascript
formData.append('files[]', file);  // ต้องมี [] เสมอ
```

**กฎ:** เวลาเขียน multi-file upload ใน JS → PHP ต้องใช้ `files[]` ทุกครั้ง ตรวจสอบทุก manage.html

---

## 3. PHP max_file_uploads — Default คือ 20

**ปัญหา:** Upload P3 22-23 ไฟล์พร้อมกัน แต่ PHP รับได้แค่ 20 (default)

**แก้ไข:** เพิ่มใน `.htaccess`:
```apache
php_value max_file_uploads 100
php_value upload_max_filesize 50M
php_value post_max_size 200M
php_value max_execution_time 600
php_value memory_limit 512M
```

**กฎ:** ทุก dashboard ต้องตั้ง max_file_uploads ≥ 100 ใน .htaccess

---

## 4. P3 File Naming — ต้องอ่าน Branch จาก Excel Content

**ปัญหา:** ไฟล์ P3 จาก DMAMA มีชื่อเป็น `รายงานแรงดันน้ำ_P1_-_Pn_2026-04-01_HH-MM-SS.xlsx` — ไม่มีชื่อสาขาในชื่อไฟล์

**สาเหตุ:** ทุกไฟล์ถูก rename เป็นชื่อเดียวกัน → เขียนทับกันหมด เหลือไฟล์เดียว

**แก้ไข:**
1. ดึงวันที่จากชื่อไฟล์ DMAMA: `2026-04-01` → Thai year `69-04`
2. อ่านชื่อสาขาจาก Excel content — **ห้ามใช้ regex P3-xxx** เพราะ Row 1 มี "P1 - Pn" ซึ่ง match ก่อนและได้ค่าผิด
3. **ใช้ 2 วิธีที่ถูกต้อง:**
   - วิธี 1: หา "สถานีผลิตน้ำ" + ชื่อที่ตามหลัง (เช่น "สถานีผลิตน้ำสระแก้ว")
   - วิธี 2: Match ชื่อสาขามาตรฐาน 22 สาขาตรงๆ จากเนื้อหา
4. ตั้งชื่อใหม่: `P3_สาขา_YY-MM.xlsx`
5. ใช้ `$used_names[]` ป้องกัน duplicate ในแต่ละ batch

**กฎ:** เวลา rename ไฟล์ upload → ห้าม rename ทุกไฟล์เป็นชื่อเดียวกัน ต้องมี unique identifier (สาขา, เลข, ฯลฯ) ตรวจสอบ `$used_names` ทุกครั้ง

---

## 5. rlcOISMap Cache — ต้อง Invalidate เมื่อ OIS Data เปลี่ยน

**ปัญหา:** Upload ข้อมูลใหม่แล้ว กราฟ OIS comparison ใน tab น้ำสูญเสีย(บริหาร) หายไป

**สาเหตุ:** `rlcOISMap` (mapping RL branch → OIS sheet) ถูก cache ครั้งแรกตอน RL load ก่อน OIS → mapping เป็น null ทั้งหมด พอ OIS มาทีหลัง cache ไม่ถูก reset

**แก้ไข:** ใน OIS callback ของ IIFE:
```javascript
rlcOISMap = null;           // reset cache
rlcInitialized = false;
if (typeof rlcInit === 'function') rlcInit();  // re-init RL tab
```

**กฎ:** เมื่อข้อมูลที่ถูก cache มีการเปลี่ยนแปลง (โดยเฉพาะจาก API callback) ต้อง invalidate cache ที่เกี่ยวข้องทั้งหมด ตรวจสอบว่า callback มี reset cache + re-init ครบ

---

## 6. PhpSpreadsheet Performance

**ปัญหา:** Upload RL ช้ามาก

**สาเหตุ:** PhpSpreadsheet อ่านทุก cell, ใช้ getCalculatedValue() (คำนวณ formula ใหม่ = ช้ามาก)

**แก้ไข:**
- ใช้ `getOldCalculatedValue()` แทน `getCalculatedValue()` — อ่านค่าที่ Excel คำนวณไว้แล้ว (เร็วกว่ามาก)
- ใช้ file-level caching ใน `.cache/` directory กับ mtime-based invalidation
- ตั้ง `memory_limit=512M` และ `max_execution_time=600`

**กฎ:** ใช้ `getOldCalculatedValue()` เสมอ ยกเว้นกรณีที่ต้องการผลลัพธ์ใหม่จริงๆ

---

## 7. API-Only Mode Architecture

**สำคัญ:** index.html ทำงานแบบ API-Only — data vars เป็นค่าว่างตอนเริ่ม ทุกอย่างโหลดจาก api.php

**IIFE ท้ายไฟล์ fetch 6 endpoints:**
- KPI, EU, P3, OIS, MNF, RL

**กฎ:**
- ห้ามแก้ data vars ใน index.html โดยตรง — ข้อมูลมาจาก API เท่านั้น
- เมื่อ API callback ได้ข้อมูลใหม่ ต้อง `rebuildAllData()` + reset cache + re-init tabs ที่เกี่ยวข้อง
- Race condition: ข้อมูลอาจมาไม่พร้อมกัน → callback ต้องจัดการกรณีที่ข้อมูลอื่นยังไม่มา

---

## 8. PhpSpreadsheet Column Index — เป็น 1-based

**ปัญหา:** build_dashboard.php อ่านค่า OIS ผิด column

**สาเหตุ:** PhpSpreadsheet ใช้ 1-based column index (`$row[1]` = column A) แต่โค้ดเขียน 0-based (`$row[0]`)

**กฎ:** PhpSpreadsheet:
- `getCell([1, 1])` = A1 (column 1 = A, row 1 = 1)
- `$row[1]` = column A (ไม่ใช่ `$row[0]`)

---

## 9. Cross-Dashboard Consistency

**กฎ:** เมื่อแก้ bug ใน dashboard หนึ่ง ต้องตรวจสอบและแก้ไขทุก dashboard ที่มีโค้ดเดียวกัน:
- Dashboard_Leak
- Dashboard_PR
- Dashboard_GIS
- Dashboard_Meter

**ไฟล์ที่ต้องตรวจทุก dashboard:**
- `.htaccess` — PHP settings
- `manage.html` — upload form (files[])
- `api.php` — upload handler, error handling
- `build_dashboard.php` — build logic (ถ้ามี)

---

## 10. Browser Cache — แจ้ง User ให้ Hard Refresh

**ปัญหา:** แก้โค้ดแล้วแต่ user บอก "เหมือนเดิม"

**สาเหตุ:** Browser cache ไฟล์ JS/HTML เก่า

**กฎ:** หลังแก้ไฟล์ manage.html หรือ index.html → แจ้ง user ให้กด Ctrl+Shift+R (hard refresh) หรือเปิด Incognito ทดสอบ

---

## 11. สิ่งที่ต้องทำเมื่อ Upload ไฟล์สำเร็จ

ลำดับการทำงานหลัง upload:
1. `performUpload()` → POST `/api/upload/<category>`
2. ได้ response JSON → แสดงผลสำเร็จ/ล้มเหลว
3. `autoRebuildDashboard()` → POST `/api/rebuild` (build ข้อมูลใหม่)
4. Index.html จะ fetch ข้อมูลใหม่จาก API อัตโนมัติเมื่อ reload

---

## อัพเดทล่าสุด: 2026-04-01
