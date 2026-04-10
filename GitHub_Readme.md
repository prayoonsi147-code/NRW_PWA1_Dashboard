# NRW_PWA1_Dashboard — คู่มือระบบ Push Git & Build

> สรุปจาก quick_context.txt, push_to_github.bat, build_dashboard.php ทุก Dashboard
> อัปเดตล่าสุด: 4 เม.ย. 2569 (Session 9)

---

## สารบัญ

1. [ภาพรวมโปรเจค](#1-ภาพรวมโปรเจค)
2. [แนวคิดหลัก (Concept)](#2-แนวคิดหลัก-concept)
3. [push_to_github.bat — 9 ขั้นตอน](#3-push_to_githubbat--9-ขั้นตอน)
4. [build_dashboard.php — 4 Dashboard](#4-build_dashboardphp--4-dashboard)
5. [quick_push.bat — Push แบบเร็ว](#5-quick_pushbat--push-แบบเร็ว)
6. [.gitignore — ไฟล์ที่ไม่ push](#6-gitignore--ไฟล์ที่ไม่-push)
7. [วิธี Push 3 แบบ](#7-วิธี-push-3-แบบ)
8. [ปัญหาที่รู้แล้ว (Known Issues)](#8-ปัญหาที่รู้แล้ว-known-issues)
9. [แก้ปัญหาไฟล์ใหญ่เกิน 100MB](#9-แก้ปัญหาไฟล์ใหญ่เกิน-100mb)
10. [Setup เครื่องใหม่ — Checklist](#10-setup-เครื่องใหม่--checklist)

---

## 1. ภาพรวมโปรเจค

โปรเจคมี 4 Dashboard + 1 Landing Page

| Dashboard | Port | Theme | index.html | สิ่งที่ build ฝัง |
|-----------|------|-------|------------|-------------------|
| Dashboard_PR | 5000 | ส้ม #c43e00 | ~1.6MB | DATA, BRANCHES, PR_CATEGORIES, AON |
| Dashboard_Leak | 5001 | น้ำเงิน #1e3a5f | ~4.7MB | D(OIS), RL, EU, MNF, KPI, P3 (6 หมวด) |
| Dashboard_GIS | 5002 | Teal #004d40 | ~146KB | DATA, PRESSURE_DATA, PD1-PD5 fallback |
| Dashboard_Meter | 5003 | ม่วง #4a148c | ~22KB | DEAD_METER, TOTAL_METERS, DEAD_METER_DATE |

GitHub Pages: https://prayoonsi147-code.github.io/NRW_PWA1_Dashboard/

---

## 2. แนวคิดหลัก (Concept)

```
Local (มี XAMPP)          GitHub Pages (ไม่มี PHP)
┌─────────────────┐       ┌──────────────────────┐
│ index.html      │       │ index.html            │
│ ไม่มี data ฝัง  │       │ มี data ฝังใน HTML    │
│ ใช้ API (PHP)   │       │ ใช้ fallback data     │
│ ข้อมูลสดจาก     │       │ ข้อมูลจาก build       │
│ api.php         │       │ script ล่าสุด         │
└─────────────────┘       └──────────────────────┘
```

**กฎเหล็ก — ห้ามฝังข้อมูลดิบลง index.html:**
- **ห้าม hardcode** ข้อมูลดิบลงใน index.html โดยตรง ไม่ว่าจะกรณีใดก็ตาม
- ข้อมูลใหม่ทุกชนิดต้องมี: (1) API endpoint ใน api.php, (2) build function ใน build_dashboard.php
- Local ใช้ API อย่างเดียว — ห้ามเรียก build_dashboard.php จาก manage.html หรือ api.php/rebuild บน local
- **ยกเว้นเดียว:** ฝังชั่วคราวตอน Push Git (push_to_github.bat Step 2) แล้ว restore กลับทันที (Step 9)
- ถ้าพบข้อมูลดิบ hardcode อยู่ใน HTML → ต้องแจ้งผู้ใช้ทันที ห้ามปล่อยผ่าน
- ทุก manage.html มี comment: "การฝังข้อมูลสำหรับ GitHub Pages ทำผ่าน push_to_github.bat เท่านั้น"

---

## 3. push_to_github.bat — 9 ขั้นตอน

ไฟล์: `push_to_github.bat` (root ของโปรเจค)

```
Step 1: CHECKPOINT
  └─ backup index.html → index.html.checkpoint ทุก 4 Dashboard

Step 2: BUILD
  └─ รัน build_dashboard.php ทุก Dashboard (ฝัง JSON data ลง HTML)
  └─ ต้องมี PHP (XAMPP ที่ C:\xampp\php\php.exe หรือ PATH)

Step 3: VALIDATE
  └─ ตรวจ DOCTYPE ใน index.html ทุกตัว
  └─ ตรวจขนาดไฟล์ ≥ 1KB (ป้องกันไฟล์ถูกตัด)
  └─ ถ้า fail → restore จาก .checkpoint อัตโนมัติ + แจ้ง WARNING

Step 4: GIT INIT
  └─ ถ้ายังไม่มี .git → git init + set remote origin

Step 5: GIT IDENTITY
  └─ git config user.email "prayoonsi147@gmail.com"
  └─ git config user.name "prayoonsi147-code"

Step 6: GIT PULL + SQUASH
  └─ git pull origin main --allow-unrelated-histories
  └─ git reset --soft origin/main  (squash unpushed commits ป้องกันไฟล์ใหญ่ค้าง)

Step 7: GIT STAGE
  └─ git add -A (เคารพ .gitignore)
  └─ git rm --cached *.sqlite *.cache.json *.checkpoint (ลบไฟล์ใหญ่ออก)

Step 8: COMMIT + PUSH
  └─ git commit -m "Update dashboard data"
  └─ git push -u origin main
  └─ ถ้า fail → goto RESTORE

Step 9: RESTORE
  └─ คืน index.html จาก .checkpoint (local กลับเป็นเวอร์ชันไม่มีข้อมูลฝัง)
  └─ ลบไฟล์ .checkpoint
```

**ผลลัพธ์:** GitHub Pages ได้ HTML+data ใหม่ แต่ local กลับเป็นเดิม

---

## 4. build_dashboard.php — 4 Dashboard

ทุกไฟล์อยู่ใน folder ของแต่ละ Dashboard ใช้ PhpSpreadsheet อ่าน Excel

### Dashboard_PR/build_dashboard.php
- อ่าน: `ข้อมูลดิบ/เรื่องร้องเรียน/PR_*.xlsx` + `ข้อมูลดิบ/AlwayON/AON_*.xls`
- ฝัง: `var DATA = {...}`, `var BRANCHES = [...]`, `var PR_CATEGORIES = [...]`, AON data
- วิธี replace: `strpos + brace counting` (ค้น `var DATA` แล้วนับ `{}`)
- memory_limit: 512M, pcre.backtrack_limit: 10000000
- clean_num(): ใช้ preg_replace กำจัด non-breaking space (\xC2\xA0)

### Dashboard_Leak/build_dashboard.php
- อ่าน: 6 หมวด (OIS, RL, EU, MNF, KPI, P3) จาก `ข้อมูลดิบ/`
- ฝัง: `var D = {...}`, `var RL = {...}`, `var EU = {...}`, `var MNF = {...}`, `var KPI = {...}`, `var P3 = {...}`
- RL: รองรับ multi-FY (ไฟล์เดียวมี sheet ข้ามปีงบ) + rate fallback calculation
- P3: process_p3_folder() ส่ง &$result by reference
- memory_limit: 512M
- Known Issue: replace_js_var อาจนับ brace ผิดใน string ซับซ้อน (มี safety checks)

### Dashboard_GIS/build_dashboard.php
- อ่าน: จุดซ่อมท่อ + แรงดันน้ำ + ค้างซ่อม จาก `ข้อมูลดิบ/`
- ฝัง: `const DATA`, `PRESSURE_DATA`, `PRESSURE_MONTHS`, PD1-PD5 fallback data
- วิธี replace: marker comments `/*FALLBACK_PD4_START*/.../*FALLBACK_PD4_END*/`
- read_pending_data(): รองรับหลายไฟล์ต่อปีงบฯ (array)
- memory_limit: 1024M (ค้างซ่อม xlsx ใช้ memory มาก)

### Dashboard_Meter/build_dashboard.php
- อ่าน: `METER_*.xlsx` + OIS จาก `../Dashboard_Leak/ข้อมูลดิบ/OIS/`
- ฝัง: `DEAD_METER`, `TOTAL_METERS`, `DEAD_METER_DATE`
- Safety: backup ก่อนเขียน, DOCTYPE validation, min size 1KB, 30% shrink protection
- 22 branch codes mapped, 10 meter sizes tracked
- Note: normalize_size() ใช้ eval() — ควรเปลี่ยนเป็น lookup table

### Safety Checks ที่มี/ไม่มี

| Dashboard | backup | DOCTYPE check | min size | shrink protection |
|-----------|--------|---------------|----------|-------------------|
| Leak | มี | มี | มี | มี |
| Meter | มี | มี | มี | มี (30%) |
| PR | ยังไม่มี | ยังไม่มี | ยังไม่มี | ยังไม่มี |
| GIS | ยังไม่มี | ยังไม่มี | ยังไม่มี | ยังไม่มี |

---

## 5. quick_push.bat — Push แบบเร็ว

ไฟล์: `quick_push.bat` (root — ถูก .gitignore ไม่ push ขึ้น)

ใช้เมื่อ: แก้แค่ UI/CSS/JS ไม่เกี่ยวกับ data หรือกู้คืน index.html ที่ build พัง

```batch
git add [เฉพาะไฟล์ที่แก้]
git commit -m "Fix: ..."
git push -u origin main
```

ไม่รัน build → ข้อมูล embedded เดิมยังอยู่

---

## 6. .gitignore — ไฟล์ที่ไม่ push

```
*.sqlite          # SQLite databases (generated, อาจเกิน 100MB)
*.cache.json      # Cache files (generated)
*.checkpoint      # Build checkpoint files
*.db / *.mdb      # Database files
*.mp4             # Video files
*.zip / *.rar     # Archives
**/data.json      # Server data (generated)
**/uploaded_data/  # User uploads
__pycache__/      # Python cache (legacy)
prompt_history.txt
quick_push.bat
index.html.bak
```

---

## 7. วิธี Push 3 แบบ

### วิธี A — push_to_github.bat (แนะนำ)
- **เมื่อไหร่:** อัปเดตข้อมูลใหม่ + push ขึ้น GitHub Pages
- **ต้องการ:** XAMPP/PHP บนเครื่อง
- **วิธี:** ดับเบิลคลิก `push_to_github.bat`
- **ผลลัพธ์:** build data ใหม่ + push + local กลับเป็นเดิม

### วิธี B — ผ่าน Cowork/Claude
- **ข้อจำกัด:** Cowork sandbox ไม่มี XAMPP → รัน build ไม่ได้
- **ใช้ได้เมื่อ:** แก้แค่ HTML/CSS/JS ไม่เกี่ยวกับ data
- **ถ้าต้อง push data:** แก้ source แล้วบอก user รัน push_to_github.bat เอง
- **คำเตือน:** push โดยไม่ build → GitHub Pages แสดงกราฟว่าง (ไม่มี fallback data)

### วิธี C — Quick push ไม่รัน build
- **เมื่อไหร่:** แก้ CSS/JS/HTML layout ไม่แก้ data
- **วิธี:**
  ```bash
  git add Dashboard_PR/index.html
  git commit -m "Fix PR chart legend colors"
  git push origin main
  ```
- **ผลลัพธ์:** ข้อมูล embedded เดิมยังอยู่ (ไม่หาย)

---

## 8. ปัญหาที่รู้แล้ว (Known Issues)

### 8.1 build_dashboard.php (Leak) — replace_js_var อาจเขียน index.html เสีย
- **สาเหตุ:** parse `{...}` ด้วย brace counting อาจนับผิดถ้ามี `{` หรือ `}` ใน string literal
- **แก้ชั่วคราว:** safety checks (ตรวจ DOCTYPE, ขนาดไฟล์, backup) — ถ้า content เสียจะไม่เขียนทับ
- **แก้ถาวร (ยังไม่ทำ):** เปลี่ยนเป็น marker comment แบบ GIS

### 8.2 OIS files — format ไม่ตรงกับ parser
- OIS .xls มี branches เป็น columns (เป้าหมาย format) ไม่ใช่ months เป็น columns
- build script จะ skip OIS replacement → คงค่าเดิมที่ฝังใน index.html

### 8.3 AON values ยังเป็น 0
- ปัญหา parse_aon_sheet_with_col เดิม → fallback ยังใช้ได้

### 8.4 PR + GIS build_dashboard.php ยังไม่มี safety checks ครบ
- ตอนนี้มีเฉพาะ Leak + Meter ที่มี backup/DOCTYPE/size validation

### 8.5 Meter build — normalize_size() ใช้ eval()
- ควรเปลี่ยนเป็น lookup table

### 8.6 JS error: updateLSUsageBar is not defined
- เกิดใน rebuilt index.html (line 6758)

---

## 9. แก้ปัญหาไฟล์ใหญ่เกิน 100MB

GitHub ปฏิเสธ push ถ้ามีไฟล์เกิน 100MB ใน commit history (ถึงแม้ลบแล้ว ถ้ายังอยู่ใน commit เก่าก็ push ไม่ได้)

**ป้องกัน (ปัจจุบัน):**
- push_to_github.bat มี `git reset --soft origin/main` ก่อน commit (squash history)
- .gitignore มี: `*.sqlite`, `*.db`, `*.cache.json`, `*.zip`, `*.rar`, `*.7z`, `*.bak`

**วิธี A — ยัง push ไม่สำเร็จ (commits อยู่ local):**
```bash
git reset --soft origin/main
git rm --cached "path/to/largefile"
# เพิ่ม pattern ใน .gitignore
git add -A && git commit -m "Remove large file"
git push
```

**วิธี B — push สำเร็จแล้ว (ไฟล์อยู่ใน remote history):**
```bash
# ใช้ BFG Repo-Cleaner
java -jar bfg.jar --strip-blobs-bigger-than 100M
git reflog expire --expire=now --all
git gc --prune=now --aggressive
git push --force
```

**วิธี C — เริ่มใหม่ (ง่ายสุด):**
```bash
# ลบ .git folder
git init && git remote add origin <url>
git add -A && git commit -m "Fresh start"
git push --force origin main
```

---

## 10. Setup เครื่องใหม่ — Checklist

1. **ติดตั้ง Git for Windows** — https://git-scm.com/download/win

2. **ติดตั้ง XAMPP** — https://www.apachefriends.org/
   - เปิด php.ini → uncomment: `extension=gd` + `extension=zip`
   - path: `C:\xampp\php\php.ini`

3. **ติดตั้ง Composer** (PHP package manager)
   ```bash
   cd [project root]
   C:\xampp\php\php.exe -r "copy('https://getcomposer.org/installer','composer-setup.php');"
   C:\xampp\php\php.exe composer-setup.php
   C:\xampp\php\php.exe composer.phar install
   ```

4. **Clone repo**
   ```bash
   git clone https://github.com/prayoonsi147-code/NRW_PWA1_Dashboard.git
   cd NRW_PWA1_Dashboard
   ```

5. **ตั้งค่า Git identity**
   ```bash
   git config user.email "prayoonsi147@gmail.com"
   git config user.name "prayoonsi147-code"
   ```

6. **ทดสอบ push** — ดับเบิลคลิก `push_to_github.bat` (ครั้งแรกอาจต้อง login GitHub ผ่าน browser)

7. **ตั้งค่า XAMPP สำหรับ Dashboard บนเครื่อง** (ทำครั้งเดียว)
   - XAMPP Control Panel → Config (แถว Apache) → httpd.conf
   - Ctrl+H → Find: `C:/xampp/htdocs` → Replace: `C:/Users/[ชื่อuser]`
   - กด Replace All → Save
   - ทดสอบ: cmd → `"C:\xampp\apache\bin\httpd.exe" -t` → ต้องขึ้น `Syntax OK`
   - Start Apache ใน XAMPP Control Panel
   - เปิด: `http://localhost/Claude Test Cowork/index.html`

8. **ถ้า push ติด "file exceeds 100MB"** → ดูหัวข้อ 9

9. **ถ้า build ทำ index.html เสีย (หน้าว่าง)**
   ```bash
   git checkout HEAD~1 -- Dashboard_Leak/index.html
   ```
   แล้ว push ด้วย quick_push.bat

---

## update_dashboards.bat

ไฟล์: `update_dashboards.bat` (root)

ใช้เมื่อ: ต้องการ build ข้อมูลใหม่ลง index.html โดยไม่ push (preview/test)

ทำงาน:
1. หา PHP (XAMPP)
2. วนลูป Dashboard_* ทุกตัว
3. รัน build_dashboard.php ในแต่ละ folder
4. รายงาน success/failure

ต่างจาก push_to_github.bat: ไม่มี git, ไม่มี backup/restore

---

## GitHub Pages Cache

หลัง push สำเร็จ ต้องรอ 1-3 นาที หรือ Hard Refresh (Ctrl+Shift+R) ถึงจะเห็นข้อมูลใหม่

ตรวจสอบด้วย:
```bash
git show origin/main:Dashboard_PR/index.html | grep -o "69-03.*" | head -1
```

---

## งานค้างปัจจุบัน (Session 9)

- User ต้องรัน push_to_github.bat เพื่อ rebuild + push ขึ้น GitHub Pages
- ทดสอบ pending-chart, pending-table API หลังแก้ multi-file
- เพิ่ม safety checks ให้ PR + GIS build_dashboard.php
- แก้ normalize_size() ใน Meter build (eval → lookup)
- แก้ JS error: updateLSUsageBar
- AON values ยังเป็น 0
