# โปรเจ็ค: Dashboard น้ำสูญเสีย - กปภ.ข.1

> **Claude:** อ่าน `quick_context.txt` ก่อนเสมอ (มีสรุปงาน + สถาปัตยกรรม + คำเตือน)
> ไฟล์นี้เป็น reference เพิ่มเติม — อ่านเฉพาะส่วนที่เกี่ยวข้องกับงานที่ทำ

---

## ข้อตกลงการพัฒนา (Development Conventions)

### กราฟ (Chart)
- **ทุกกราฟ** ต้องมี: Y-axis tick callback (ค่า <100 → 2 ทศนิยม, ≥100 → comma), X-axis maxRotation:60 font 12, beginAtZero:false
- ลงทะเบียนใน `getChartObj()`, `CTX_CHART_TYPE_MAP`, `CTX_CHART_UPDATE_MAP`
- มี export bar ครบ (reset, show values, font, PNG, copy, Excel, PowerPoint)
- ประเภทกราฟเรียงลำดับ: เส้นโค้ง (line, tension:0.3) → เส้นตรง (straight, tension:0) → แท่ง (bar) → พื้นที่ (area)
- **ห้าม reorder `<option>`** เพื่อเปลี่ยน default → ใช้ `selected` attribute แทน

### Controls Layout (ซ้าย→ขวา)
ประเภทกราฟ → แกน X → เกณฑ์ → ตัวเลือกอื่นๆ

### Branch Selector (3 โหมด)
1. สาขาเดียว — dropdown
2. บางสาขา (default) — checkboxes สีตาม accent-color
3. ทุกสาขา — แสดงทั้งหมด
- แสดงเมื่อ xAxis = "ช่วงเวลา" เท่านั้น (ซ่อนเมื่อ xAxis = "สาขา")

---

## หน่วยงาน (22 สาขา)
ชลบุรี(พ), พัทยา(พ), พนัสนิคม, บ้านบึง, ศรีราชา, แหลมฉบัง, ฉะเชิงเทรา, บางปะกง, บางคล้า, พนมสารคาม, ระยอง, บ้านฉาง, ปากน้ำประแสร์, จันทบุรี, ขลุง, ตราด, คลองใหญ่, สระแก้ว, วัฒนานคร, อรัญประเทศ, ปราจีนบุรี, กบินทร์บุรี

**หมายเหตุ:** ชื่อในข้อมูลดิบอาจต่างกัน เช่น "พัทยา (พ)", "พัทยา น.1" → ต้อง normalize

---

## โครงสร้างไฟล์โปรเจค (File Tree)

อัปเดตล่าสุด: 28 มี.ค. 2569

```
Claude Test Cowork/
│
├── index.html                  ← Landing Page หน้าหลัก (เลือก Dashboard)
├── composer.json                ← PHP dependencies (PhpSpreadsheet)
├── composer.lock
├── vendor/                      ← PHP libraries (phpoffice/phpspreadsheet, ...)
│
├── quick_context.txt            ← ★ สรุปงาน+สถาปัตยกรรม (Claude อ่านก่อนเสมอ)
├── PROJECT_README.md            ← Reference เพิ่มเติม (ไฟล์นี้)
├── prompt_history.txt           ← ประวัติ prompt (ยกเลิกการบันทึกแล้ว)
├── .gitignore                   ← git ignore rules (*.cache.json, __pycache__, ...)
│
├── push_to_github.bat           ← ★ Build ทุก Dashboard → git add → commit → push
├── update_dashboards.bat        ← Run build_dashboard ทุกตัว (pushd/popd pattern)
├── start_server.bat             ← [เลิกใช้] เรียก Python Flask
├── start_xampp.bat              ← [ใช้ manual แทน] setup+start XAMPP Apache
│
│
├── Dashboard_PR/                ─── ข้อร้องเรียน+AlwayON (ส้ม #c43e00) ───
│   ├── index.html               ← Dashboard หลัก (~1.6MB, ข้อมูลฝังใน HTML)
│   ├── manage.html              ← หน้าจัดการข้อมูล (upload/delete/notes)
│   ├── api.php                  ← ★ PHP API (upload, data, notes, delete)
│   ├── .htaccess                ← Apache rewrite rules
│   ├── data.json                ← ข้อมูลรวม (build สร้าง)
│   ├── build_dashboard.py       ← [เลิกใช้] Python build script
│   ├── server.py                ← [เลิกใช้] Python Flask server
│   ├── requirements.txt         ← [เลิกใช้] Python dependencies
│   ├── AlwayON_Data_District1.csv  ← ข้อมูล AON (reference)
│   ├── AlwayON_Data_District1.js   ← ข้อมูล AON (JS format)
│   ├── AlwayON_Data_Summary.txt    ← สรุปข้อมูล AON
│   ├── README_AlwayON_Data.txt     ← คำอธิบายข้อมูล AON
│   ├── Excel_Structure_Reference.txt ← อ้างอิงโครงสร้าง Excel
│   ├── Untitled.jpg             ← รูปประกอบ
│   ├── _รายงานข้องร้องเรียน ก.พ.69 (รปก.3).pdf ← ตัวอย่างรายงาน
│   │
│   ├── ข้อมูลดิบ/
│   │   ├── data.json            ← ข้อมูลรวม + notes (API เก็บที่นี่)
│   │   ├── เรื่องร้องเรียน/     ← PR_YY-MM.xlsx (เช่น PR_69-03.xlsx)
│   │   └── AlwayON/             ← AON_YY-MM.xls (เช่น AON_69-01.xls)
│   │
│   └── uploaded_data/           ← สำเนาไฟล์ที่ upload ผ่าน manage.html
│       ├── data.json
│       ├── pr/                  ← PR_YY-MM.xlsx (ปี 66-69)
│       └── aon/                 ← AON_YY-MM.xls
│
│
├── Dashboard_Leak/              ─── น้ำสูญเสีย (น้ำเงิน #1e3a5f) ───
│   ├── index.html               ← Dashboard หลัก (~4.7MB, ข้อมูลฝังใน HTML)
│   ├── manage.html              ← หน้าจัดการข้อมูล
│   ├── api.php                  ← ★ PHP API
│   ├── .htaccess                ← Apache rewrite rules
│   ├── data.json                ← ข้อมูลรวม
│   ├── data_embed.js            ← ข้อมูลฝังสำหรับ build
│   ├── build_dashboard.py       ← [เลิกใช้] Python build script
│   ├── server.py                ← [เลิกใช้] Python Flask server
│   ├── PROJECT_README.md        ← README เฉพาะ Leak
│   ├── อัพเดท Dashboard.bat     ← Shortcut run build
│   │
│   ├── Temp/                    ← ไฟล์ชั่วคราว/อ้างอิง
│   │   ├── กราฟน้ำสูญเสีย.xlsx
│   │   └── รายงานน้ำรับ-น้ำส่ง-69.xlsx
│   │
│   └── ข้อมูลดิบ/
│       ├── OIS/                 ← OIS_YYYY.xls (ปี 2557-2569)
│       ├── Real Leak/           ← RL_YYYY.xlsx (เช่น RL_2569.xlsx)
│       ├── MNF/                 ← MNF_YYYY.xlsx (เช่น MNF_2569.xlsx)
│       ├── P3/                  ← P3_สาขา_YY-MM.xlsx (เช่น P3_ชลบุรี_69-03.xlsx)
│       ├── Activities/          ← ACT_กิจกรรมลดน้ำสูญเสีย.xlsx
│       ├── หน่วยไฟ/             ← EU_YYYY.xlsx (เช่น EU_2569.xlsx)
│       ├── เกณฑ์ชี้วัด/         ← KPI_YYYY.xlsx (เช่น KPI_2569.xlsx)
│       └── เกณฑ์วัดน้ำสูญเสีย/  ← KPI2_YYYY.xlsx (เช่น KPI2_2569.xlsx)
│
│
├── Dashboard_GIS/               ─── จุดซ่อมท่อ+แรงดัน+ค้างซ่อม (Teal #004d40) ───
│   ├── index.html               ← Dashboard หลัก (~146KB + fallback data)
│   ├── manage.html              ← หน้าจัดการข้อมูล
│   ├── api.php                  ← ★ PHP API (data, pending-chart, pending-table, pressure, notes)
│   ├── build_sqlite.php         ← สร้าง .sqlite จาก Excel ค้างซ่อม
│   ├── debug_sqlite.php         ← Debug tool สำหรับ SQLite
│   ├── .htaccess                ← Apache rewrite rules
│   ├── .cache/                  ← Cache files (runtime)
│   ├── build_dashboard.py       ← [เลิกใช้] Python build script
│   ├── server.py                ← [เลิกใช้] Python Flask server
│   │
│   └── ข้อมูลดิบ/
│       ├── ลงข้อมูลซ่อมท่อ/     ← GIS_YYMMDD.xlsx (เช่น GIS_690218.xlsx)
│       ├── แรงดันน้ำ/           ← PRESSURE_สาขา_ปีงบYY.xlsx (22 สาขา)
│       │                          เช่น PRESSURE_ชลบุรี_ปีงบ69.xlsx
│       └── ซ่อมท่อค้างระบบ/     ← ค้างซ่อม_MM-YY_to_MM-YY.xlsx + .sqlite + .cache.json
│                                  เช่น ค้างซ่อม_10-68_to_03-69.xlsx
│
│
└── Dashboard_Meter/             ─── มาตรวัดน้ำผิดปกติ (ม่วง #4a148c) ───
    ├── index.html               ← Dashboard หลัก (~22KB)
    ├── manage.html              ← หน้าจัดการข้อมูล
    ├── api.php                  ← ★ PHP API
    ├── .htaccess                ← Apache rewrite rules
    ├── data.json                ← ข้อมูลรวม
    ├── test_path.php            ← Debug tool
    ├── New Text Document.txt    ← (ว่าง)
    ├── server.py                ← [เลิกใช้] Python Flask server
    ├── update_dead_meter.py     ← [เลิกใช้] Python update script
    │
    └── ข้อมูลดิบ/
        └── มาตรวัดน้ำผิดปกติ/   ← METER_MMYY.xlsx (เช่น METER_1102.xlsx = ก.พ.69)
                                   มี 23 ไฟล์ (METER_1102 ถึง METER_1123)
```

### สรุปไฟล์แต่ละ Dashboard

| ไฟล์ | หน้าที่ | มีทุก Dashboard |
|------|---------|:-:|
| index.html | Dashboard หลัก (แสดงกราฟ/ตาราง) | ✅ |
| manage.html | หน้าจัดการข้อมูล (upload/delete/notes) | ✅ |
| api.php | ★ PHP API หลัก (XAMPP) | ✅ |
| .htaccess | Apache URL rewrite | ✅ |
| data.json | ข้อมูลรวม (บาง Dashboard) | PR, Leak, Meter |
| build_dashboard.py | [เลิกใช้] Python build | PR, Leak, GIS |
| server.py | [เลิกใช้] Python Flask | ✅ |
| ข้อมูลดิบ/ | โฟลเดอร์เก็บไฟล์ Excel ต้นฉบับ | ✅ |

### สถานะไฟล์เก่า (รอลบ)

| ไฟล์ | อยู่ใน | หมายเหตุ |
|------|--------|----------|
| server.py | ทุก Dashboard | Python Flask server → แทนด้วย api.php |
| build_dashboard.py | PR, Leak, GIS | Python build → แทนด้วย PHP build / ฝังข้อมูลตรง |
| update_dead_meter.py | Meter | Python script → แทนด้วย api.php |
| requirements.txt | PR | Python dependencies → ไม่ใช้แล้ว |
| __pycache__/ | ทุก Dashboard | Python cache → ลบได้เลย |
| start_server.bat | root | เรียก Python Flask → ใช้ XAMPP แทน |

### Tab ของแต่ละ Dashboard

**Dashboard_PR** (ส้ม #c43e00):
- Tab 1: ข้อร้องเรียน (กราฟ 1-2)
- Tab 2: รายงานข้อร้องเรียน (กราฟ 3)
- Tab 3: Always-On (กราฟ AON 1-3)

**Dashboard_Leak** (น้ำเงิน #1e3a5f):
- Tab 1: OIS
- Tab 2: น้ำสูญเสีย (Real Leak + WSC-R)
- Tab 3: MNF
- Tab 4: P3
- Tab 5: Custom Chart

**Dashboard_GIS** (Teal #004d40):
- Tab 1: จุดซ่อมท่อ (KPI กราฟ 1-2) — ข้อมูลฝังใน HTML
- Tab 2: แรงดันน้ำ (กราฟ 1) — API สด + fallback
- Tab 3: งานค้างซ่อม (กราฟ pd1-pd2, ตาราง pd3-pd5) — API สด + fallback

**Dashboard_Meter** (ม่วง #4a148c):
- Tab 1: มาตรวัดน้ำผิดปกติ (กราฟ 1)

### Naming Convention ไฟล์ข้อมูลดิบ

| ประเภท | Pattern | ตัวอย่าง |
|--------|---------|----------|
| ข้อร้องเรียน | PR_YY-MM.xlsx | PR_69-03.xlsx |
| AlwayON | AON_YY-MM.xls | AON_69-01.xls |
| OIS | OIS_YYYY.xls | OIS_2569.xls |
| Real Leak | RL_YYYY.xlsx | RL_2569.xlsx |
| MNF | MNF_YYYY.xlsx | MNF_2569.xlsx |
| P3 | P3_สาขา_YY-MM.xlsx | P3_ชลบุรี_69-03.xlsx |
| Activities | ACT_*.xlsx | ACT_กิจกรรมลดน้ำสูญเสีย.xlsx |
| หน่วยไฟ | EU_YYYY.xlsx | EU_2569.xlsx |
| เกณฑ์ชี้วัด | KPI_YYYY.xlsx | KPI_2569.xlsx |
| เกณฑ์วัดน้ำสูญเสีย | KPI2_YYYY.xlsx | KPI2_2569.xlsx |
| จุดซ่อมท่อ | GIS_YYMMDD.xlsx | GIS_690218.xlsx |
| แรงดันน้ำ | PRESSURE_สาขา_ปีงบYY.xlsx | PRESSURE_ชลบุรี_ปีงบ69.xlsx |
| ค้างซ่อม | ค้างซ่อม_MM-YY_to_MM-YY.xlsx | ค้างซ่อม_10-68_to_03-69.xlsx |
| มาตรวัดน้ำ | METER_MMYY.xlsx | METER_1102.xlsx (=ก.พ.69, รหัส 11=Meter) |

---

## รูปแบบข้อมูลดิบ (อ่านเมื่อแก้ parser)

### OIS (ข้อมูลดิบ/OIS/)
- ไฟล์ .xls (BIFF8), ตั้งชื่อตามปี พ.ศ. เช่น `2569.xls`
- แต่ละ Sheet มี 12 เดือนเป็นคอลัมน์, Sheet "เป้าหมาย" → ข้าม
- ปีที่มี: 2558-2569

### Real Leak (ข้อมูลดิบ/Real Leak/)
- ไฟล์ .xlsx, ตั้งชื่อ `RL_YYYY.xlsx`
- แต่ละ Tab = 1 เดือน (ชื่อ "ต.ค.68", "พ.ย. 68"), Tab ที่ไม่ใช่เดือน → ข้าม
- คอลัมน์: B=สาขา, C=น้ำผลิตรวม, D=น้ำผลิตจ่ายสุทธิ, F=น้ำจำหน่าย, H=น้ำสูญเสีย(ปริมาณ), J=อัตราน้ำสูญเสีย(%)
- ข้อมูลเริ่มแถว 4 (แถว 2-3 = header)

### MNF (ข้อมูลดิบ/MNF/)
- ไฟล์ .xlsx, ตั้งชื่อ `MNF_YYYY.xlsx`
- Sheet "ภาพรวมเขต" + Sheet รายสาขา, Sheet "รวมกราฟสาขา" → ข้าม
- 4 รายการ: MNF เกิดจริง, MNF ที่ยอมรับได้, เป้าหมาย MNF, น้ำผลิตจ่าย

### P3 (ข้อมูลดิบ/P3/)
- ไฟล์ P3_สาขา_YY-MM.xlsx (flat) หรือ สาขา_YY-MM.xlsx (ใน subfolder ปี)

---

## กฎการคำนวณค่ารายปี

| ประเภท | เงื่อนไข | วิธีคำนวณ |
|--------|----------|-----------|
| ค่าสะสม (Cumulative) | หมวด 5 ยอดหนี้ค่าน้ำค้างชำระ | ค่าเดือนสุดท้ายของปี |
| อัตราส่วน (Rate) | หน่วยเป็น %, /, หรือ รายได้/รายจ่าย | ค่าเฉลี่ย (avg) |
| ปริมาณ (Volume) | หน่วยอื่นๆ (ลบ.ม., บาท, ราย) | ผลรวม (sum) |

- **OIS:** ไฟล์มีคอลัมน์ "รวม" → ใช้ค่าในตาราง
- **Real Leak, MNF, EU:** ไม่มีค่ารายปี → คำนวณเอง
- ฟังก์ชัน: `isRateMetric()`, `isCumulativeInSheet()`, `getYearlyValue()`, `pickYearlyVal()`

---

## รูปแบบการแสดงผล

| กรณี | รูปแบบ | ตัวอย่าง |
|------|--------|---------|
| เดือน | ตัวย่อ | ม.ค., ก.พ. |
| ปีงบฯ | 2 หลักท้าย | ปีงบฯ 69 |
| ปีปฏิทิน | พ.ศ. 4 หลัก | 2569 |
| เดือน+ปี | ย่อ+2 หลัก | ม.ค.69 |
| ตัวเลข <100 | 2 ทศนิยม | 5.00, 12.50 |
| ตัวเลข ≥100 | comma | 1,234 |
| Label ในกราฟ | ตัดเลขข้อ | "ปริมาณน้ำจำหน่าย" |
| Dropdown | เก็บเลขข้อ | "2.1 ปริมาณน้ำจำหน่าย" |

---

## ข้อตกลงระบบ Upload (อ่านเมื่อแก้ upload)

### สถาปัตยกรรม Dual Mode
- **Server Mode (Primary):** Flask server → openpyxl/xlrd parse → auto-rename → เก็บไฟล์ + data.json
- **Fallback (ไม่มี server):** SheetJS parse ฝั่ง browser + LZ-String compress → localStorage
- ตรวจจับอัตโนมัติด้วย `_checkServer()` → `/api/ping`

### Auto-detect เดือน/ปี (fallback chain)
1. ชื่อไฟล์ (pattern YY-MM)
2. เนื้อหาในไฟล์ (row 3, Sheet name)
3. ค่า manual ที่ผู้ใช้ระบุ
4. แจ้ง error

### หลัง Upload
- `refreshAllAfterUpload()` — sort months, clear+rebuild selects, destroy+re-create charts
- Chart.js: ต้อง `.destroy()` ก่อน re-create เสมอ
- Select options: clear ก่อนแต่เก็บ option แรกที่ hardcoded

### manage.html (ทุก Dashboard)
- ใช้ CATEGORIES loop pattern (buildUI สร้าง tabs/panels แบบ dynamic)
- Drop zone มี file list + remove button (DataTransfer API)
- Textbox บันทึกช่วยจำ + auto-save (debounce 800ms) → API POST /api/notes/<slug>
- Management panel แสดง chips แยกปี

---

## Checklist เมื่อเพิ่มหมวดข้อมูลใหม่
1. ศึกษาไฟล์ดิบ — โครงสร้าง header, data rows, column mapping
2. สร้าง Upload zone (HTML) — สีประจำหมวด, `<details>` คำแนะนำแยก
3. เขียน `processSingle___File()` — auto-detect เดือน + parse + normBranch
4. เขียน `handleUpload___()` — ผ่าน `showUploadConfirm()` + multi-file Promise chain
5. เพิ่ม tracking: `_uploaded___Months[mk]=true`
6. เพิ่ม localStorage key + save/load
7. อัปเดต `refreshAllAfterUpload()`
8. อัปเดต `confirmClearUploadedData()`
9. ทดสอบ: upload → ตรวจกราฟ → refresh → ตรวจข้อมูลอยู่ → ล้าง → ตรวจกลับเป็นเดิม
