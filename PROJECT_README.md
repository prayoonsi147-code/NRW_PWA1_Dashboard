# โปรเจ็ค: Dashboard น้ำสูญเสีย - กปภ.ข.1

## คำสั่งสำหรับ Claude (อ่านทุกครั้งที่เริ่มเซสชันใหม่)
- **ก่อนเริ่มงาน:** อ่านไฟล์นี้ทั้งหมดเพื่อทำความเข้าใจโปรเจ็ค แล้ว**ถาม Aong กลับทุกครั้ง**ว่า:
  1. ต้องการให้ดูโครงสร้างทั้งหมดก่อนมั้ย (grep Tab/กราฟ ใน index.html)
  2. ต้องการให้บันทึกประวัติการสั่ง Prompt มั้ย (prompt_history.txt)
- **ก่อนแก้โค้ด:** ต้อง grep ดูโครงสร้าง Tab และกราฟทั้งหมดใน index.html ก่อนเสมอ (ดูจาก `class="main-tab-btn"` และ `card-header`) แล้วยืนยันกับ Aong ว่าจะแก้ที่ Tab ไหน กราฟไหน canvas id อะไร ก่อนลงมือทำ — ห้ามคิดเอาเอง
- **ก่อนจบเซสชัน:** บันทึกประวัติการทำงานและสรุปการสนทนาลงในส่วน "บันทึกการทำงาน (Changelog)" ด้านล่าง **ทุกครั้ง โดยไม่ต้องรอให้ Aong บอก** ให้บันทึกทั้งสิ่งที่ทำไปแล้ว สิ่งที่คุยกัน และงานที่ยังค้างอยู่
- **บันทึก Checkpoint ทุกครั้งที่ Aong สั่ง Prompt:** เมื่อ Aong สั่งงานใดๆ ให้บันทึก checkpoint ลงใน "Checkpoint ล่าสุด" ด้านล่างทันที ระบุ: (1) คำสั่งที่ Aong ให้ (2) สิ่งที่ทำไปแล้ว (3) งานที่ยังค้าง/กำลังทำอยู่ — เพื่อให้เซสชันหน้าต่องานได้ทันที

## ข้อตกลงการพัฒนา (Development Conventions)

### กราฟ (Chart)
- **ทุกกราฟที่สร้างใหม่** ต้องมีคุณสมบัติ/Option การปรับแต่งเหมือนกับกราฟอื่น ๆ ที่มีอยู่แล้ว ได้แก่:
  - Y-axis tick callback จัดรูปแบบตัวเลข (ค่า <100 → ทศนิยม 2 ตำแหน่ง, ≥100 → คอมม่า)
  - X-axis ticks: `maxRotation:60, font size:12`
  - `beginAtZero:false`
  - ลงทะเบียนใน `getChartObj()` map เพื่อให้ scroll/zoom, double-click แก้ไขแกน, pan แกน Y ทำงานได้
  - ลงทะเบียนใน `CTX_CHART_TYPE_MAP` และ `CTX_CHART_UPDATE_MAP` เพื่อให้ context menu (คลิกขวา) เปลี่ยนประเภทกราฟ/สีพื้นหลัง/สีเส้น ทำงานได้
  - มี export bar ครบ (reset, show values, font, PNG, copy, Excel, PowerPoint)

### ตัวเลือกประเภทกราฟ (Chart Type Options)
- **ทุกกราฟ** ต้องมีตัวเลือกประเภทกราฟเรียงลำดับดังนี้:
  1. กราฟเส้นโค้ง (`value="line"`) — เส้น curve (tension:0.3)
  2. กราฟเส้นตรง (`value="straight"`) — เส้นตรง (tension:0)
  3. กราฟแท่ง (`value="bar"`)
  4. กราฟพื้นที่ (`value="area"`)
- สำหรับชื่อย่อ (เช่น context menu / cc chart): เส้นโค้ง, เส้นตรง, แท่ง, พื้นที่
- เมื่อเลือก "กราฟเส้นตรง" (straight): ใช้ Chart.js type='line' แต่ตั้ง tension:0
- เมื่อเลือก "กราฟเส้นโค้ง" (line): ใช้ Chart.js type='line' ตั้ง tension:0.3
- Context menu (คลิกขวา) ต้องรองรับ straight type ด้วย: ตรวจจับจาก tension===0 และ set tension ตามประเภท

## ภาพรวม
Dashboard น้ำสูญเสียของ กปภ.ข.1 สร้างจากข้อมูลดิบ Excel โดยใช้ Python script สร้าง index.html แบบ standalone (เปิดในเบราว์เซอร์ได้เลย)

## หน่วยงาน
รวม 23 หน่วยงาน = 1 เขต + 22 สาขา

### รายชื่อ 22 สาขา (ชื่อมาตรฐาน)
1. ชลบุรี(พ)
2. พัทยา(พ)
3. พนัสนิคม
4. บ้านบึง
5. ศรีราชา
6. แหลมฉบัง
7. ฉะเชิงเทรา
8. บางปะกง
9. บางคล้า
10. พนมสารคาม
11. ระยอง
12. บ้านฉาง
13. ปากน้ำประแสร์
14. จันทบุรี
15. ขลุง
16. ตราด
17. คลองใหญ่
18. สระแก้ว
19. วัฒนานคร
20. อรัญประเทศ
21. ปราจีนบุรี
22. กบินทร์บุรี

**หมายเหตุ:** ชื่อสาขาในข้อมูลดิบอาจเขียนต่างกัน เช่น "พัทยา (พ)", "พัทยา", "พัทยา น.1" → ทั้งหมดคือ สาขาเดียวกัน ต้องทำ name normalization ให้เป็นชื่อมาตรฐาน

## โครงสร้างโฟลเดอร์

### Dashboard_PR/ (งานลูกค้าสัมพันธ์ — ข้อร้องเรียน + Always-On)
```
Dashboard_PR/
├── index.html              # Dashboard (standalone HTML, ~7000+ lines)
├── server.py               # Flask server port 5000 — upload + auto-rename + parse
├── manage.html             # หน้าจัดการข้อมูล (Theme สีส้ม #c43e00)
└── ข้อมูลดิบ/
    ├── เรื่องร้องเรียน/     # ไฟล์ PR auto-renamed: PR_YY-MM.xlsx
    ├── AlwayON/             # ไฟล์ AON auto-renamed: AON_YY-MM.xls
    └── data.json            # ข้อมูลที่ parse แล้ว (JSON, auto-generated)
```

**โครงสร้าง index.html:**
- Tab 1: ข้อมูลข้อร้องเรียน — กราฟ 1 (เลือกรายการ/สาขา/เดือน), กราฟ 2 (เปรียบเทียบ)
- Tab 2: รายงานข้อร้องเรียน — กราฟ 3 (Chart 3 with X-axis toggle + threshold)
- Tab 3: Always-On — กราฟ AON 1 (รายเดือน), กราฟ AON 2 (เปรียบเทียบสาขา), กราฟ AON 3 (multi-branch)
- ส่วนแหล่งข้อมูล: แสดงข้อมูลที่โหลด + **Upload UI** (PR zone + AON zone, drag & drop, multi-file, auto/manual mode)

**ข้อมูลหลักใน JS:**
- `DATA` (const): ข้อร้องเรียน 22 สาขา × 39 เดือน, 10 หมวดหมู่
- `AON` (var): Always-On % scores 22 สาขา × เดือนที่มีข้อมูล
- `BRANCH_NAME_MAP`: mapping "สาขาXXX" → "XXX" สำหรับ upload normalization

**Upload Feature (เพิ่ม 2026-03-18, ปรับเป็น Server-based 2026-03-18):**
- **Server mode (Primary):** Flask server (`server.py`) รับไฟล์ Excel → Python parse (openpyxl/xlrd) → auto-rename → เก็บใน `uploaded_data/` → ส่ง JSON กลับ
- **Fallback mode:** ถ้า server ไม่ทำงาน → ใช้ SheetJS parse ฝั่ง browser + เก็บใน localStorage (เหมือนเดิม)
- Auto-rename: PR → `PR_YY-MM.xlsx`, AON → `AON_YY-MM.xls`
- `refreshAllAfterUpload()` rebuild ทุก select/chart หลัง upload

### Dashboard_GIS/ (งานแผนที่แนวท่อ)
```
Dashboard_GIS/
├── index.html              # Dashboard (standalone HTML)
├── build_dashboard.py      # สคริปต์สร้าง Dashboard
├── server.py               # Flask server port 5002
├── manage.html             # หน้าจัดการข้อมูล (Theme สี teal #004d40)
└── ข้อมูลดิบ/
    └── ลงข้อมูลซ่อมท่อ/    # ไฟล์ auto-renamed: GIS_YYMMDD.xlsx
```

### Dashboard_Leak/ (งานน้ำสูญเสีย)
```
Dashboard_Leak/
├── index.html              # Dashboard (standalone HTML, ~12,800 lines)
├── build_dashboard.py      # สคริปต์สร้าง Dashboard (Pure Python)
├── server.py               # Flask server port 5001
├── manage.html             # หน้าจัดการข้อมูล (Theme สีน้ำเงิน #1e3a5f)
├── data.json               # ข้อมูลที่ประมวลผลแล้ว
├── data_embed.js           # ข้อมูลแบบ embed ใน JS
└── ข้อมูลดิบ/
    ├── OIS/                # auto-renamed: OIS_2569.xls
    ├── Real Leak/          # auto-renamed: RL_2569.xlsx
    ├── MNF/                # auto-renamed: MNF_2569.xlsx
    ├── P3/                 # auto-renamed: P3_2569_สาขา.xlsx
    ├── Activities/         # auto-renamed: ACT_ชื่อ.xlsx
    ├── หน่วยไฟ/            # auto-renamed: EU_2569.xlsx
    ├── เกณฑ์ชี้วัด/        # auto-renamed: KPI_2569.xlsx
    └── เกณฑ์วัดน้ำสูญเสีย/ # auto-renamed: KPI2_2569.xlsx
```

### Dashboard_Meter/ (งานมาตรวัดน้ำ)
```
Dashboard_Meter/
├── index.html              # Dashboard (standalone HTML)
├── server.py               # Flask server port 5003
├── manage.html             # หน้าจัดการข้อมูล (Theme สีม่วง #4a148c)
└── ข้อมูลดิบ/
    └── มาตรวัดน้ำผิดปกติ/  # auto-renamed: METER_1102.xlsx (22 สาขา)
```

## ข้อมูลดิบ

### 1. OIS (ข้อมูลดิบ/OIS/)
- **ไฟล์:** ตั้งชื่อตามปี พ.ศ. เช่น `2569.xls`
- **รูปแบบ:** .xls (BIFF8/OLE2) บางไฟล์อาจเป็น OOXML ที่นามสกุล .xls
- **โครงสร้าง:** แต่ละ Sheet มี 12 เดือนเป็นคอลัมน์ ข้อมูลเป็นรายการต่างๆ (น้ำผลิต, น้ำสูญเสีย, รายได้ ฯลฯ)
- **Sheet ที่ข้าม:** "เป้าหมาย"
- **ปีที่มี:** 2558-2569

### 2. Real Leak (ข้อมูลดิบ/Real Leak/)
- **ไฟล์:** ตั้งชื่อ `RL-` ตามด้วยปี พ.ศ. เช่น `RL-2569.xlsx`
- **รูปแบบ:** .xlsx
- **โครงสร้าง:** แต่ละ Tab = 1 เดือน ตั้งชื่อเป็นเดือน-ปี เช่น "ต.ค.68", "พ.ย. 68" (อาจมีเว้นวรรคต่างกัน)
- **Tab ที่ไม่เกี่ยว:** Tab ที่ชื่อไม่ใช่เดือน เช่น "อัตรา", "ปริมาณ" → ข้ามไป
- **คอลัมน์หลักในแต่ละ Tab:**
  - B: ชื่อ กปภ.สาขา
  - C: น้ำผลิตรวมจริง (ลบ.ม.)
  - D: น้ำผลิตจ่ายสุทธิจริง (ลบ.ม.)
  - E: น้ำผลิตจ่ายสุทธิจริงสะสม (ลบ.ม.) — มีตั้งแต่เดือนที่ 2
  - F: น้ำจำหน่ายจริง (ลบ.ม.)
  - G: น้ำ Blow off จริง (ลบ.ม.)
  - H: น้ำสูญเสียจริง — ปริมาณ (ลบ.ม.)
  - I: ปริมาณสะสม (ลบ.ม.) — มีตั้งแต่เดือนที่ 2
  - J: อัตราน้ำสูญเสีย (%)
  - K: อัตราสะสม (%) — มีตั้งแต่เดือนที่ 2
- **ข้อมูลเริ่มที่:** แถว 4 (แถว 2-3 เป็น header)
- **หมายเหตุ:** เดือนแรก (ต.ค.) คอลัมน์ต่างจากเดือนอื่นเล็กน้อย (ไม่มีคอลัมน์สะสม, คอลัมน์อาจเลื่อน)
- **ปีที่มี:** 2568, 2569

### 3. หน่วยไฟฟ้า (ข้อมูลดิบ/หน่วยไฟ/)
- **ไฟล์:** ตั้งชื่อ `EU-` ตามด้วยปี พ.ศ. เช่น `EU-2569.xlsx`
- **รูปแบบ:** .xlsx
- **โครงสร้าง:** Sheet เดียว "กราฟ"
  - แถว 1: Header — B="Data", C="ปีงบประมาณ XXXX"
  - แถว 2: ชื่อเดือน (ต.ค. 68, พ.ย. 68, ...) ที่คอลัมน์ C-N (12 เดือน)
  - แถว 3-24: ข้อมูล 22 สาขา — A=ลำดับ, B=ชื่อสาขา, C-N=ค่าหน่วยไฟฟ้า/น้ำจำหน่าย
  - แถว 25: ภาพรวม กปภ.ข.1
- **ค่าข้อมูล:** หน่วยไฟฟ้า(ระบบจำหน่าย) ÷ น้ำจำหน่าย (kWh/ลบ.ม.)
- **ปีที่มี:** 2569

### 4. MNF (ข้อมูลดิบ/MNF/)
- **ไฟล์:** ตั้งชื่อ `MNF-` ตามด้วยปี พ.ศ. เช่น `MNF-2569.xlsx`
- **รูปแบบ:** .xlsx
- **โครงสร้าง:** หลาย Sheet — Sheet "ภาพรวมเขต" + Sheet รายสาขา (1.ชลบุรี, 2.พัทยา, ...)
  - Sheet ภาพรวมเขต: R1=title, R2=เดือน (col 2-13), R3=MNF เกิดจริง, R4=MNF ที่ยอมรับได้, R5=เป้าหมาย MNF, R6=น้ำผลิตจ่าย
  - Sheet สาขา: R2=ชื่อสาขา, R3=เดือน, R4=MNF เกิดจริง, R5=MNF ที่ยอมรับได้, R6=เป้าหมาย MNF, R7=น้ำผลิตจ่าย
  - Sheet "รวมกราฟสาขา" → ข้ามไป
- **ค่าข้อมูล:** MNF (ลบ.ม./ชม.)
- **ปีที่มี:** 2569

## build_dashboard.py
- Pure Python ไม่ต้องติดตั้ง library เพิ่ม
- มี XLSParser (อ่าน .xls BIFF8) และ XLSXParser (อ่าน .xlsx OOXML) ในตัว
- อ่านข้อมูล OIS, Real Leak, หน่วยไฟฟ้า (EU), และ MNF
- สร้าง data.json และอัพเดท index.html โดย embed ข้อมูลลงใน `const D=...`
- มี label normalization สำหรับชื่อรายการที่ต่างกันในแต่ละปี
- มี trailing zeros fix สำหรับปีที่ข้อมูลยังไม่ครบ 12 เดือน

## วิธีใช้งาน
1. วางไฟล์ข้อมูลดิบลงในโฟลเดอร์ที่ถูกต้อง
2. ดับเบิลคลิก `อัพเดท Dashboard.bat` (ต้องมี Python ติดตั้งบนเครื่อง)
3. เปิด `index.html` ในเบราว์เซอร์

---

## กฎการคำนวณค่ารายปี (Yearly Value Rules)

### ข้อตกลงแหล่งที่มาของค่ารายปี (ตกลงเมื่อ 2026-03-12)
- **OIS:** ไฟล์ต้นทางมีคอลัมน์ "รวม" อยู่แล้ว → **ใช้ค่าในตารางเป็นหลัก** (ผ่านการตรวจสอบแล้วว่าตรงกับค่าที่คำนวณเอง 90%+)
  - ปริมาณ (ลบ.ม., บาท, ราย): ค่ารวม = SUM ของ 12 เดือน ✓ ตรงกัน
  - อัตราส่วน (%): ค่ารวม = ค่าเฉลี่ยถ่วงน้ำหนัก (ไม่ใช่ avg เดือน) — ใช้ค่าในตาราง
  - ค่าสะสม (หมวด 5): ค่ารวม = ค่าเดือนสุดท้าย ✓ ตรงกัน
  - ปีที่ข้อมูลไม่ครบ 12 เดือน (เช่น 2569): ค่ารวมอาจเป็น partial sum → ใช้ค่าในตาราง
- **Real Leak, MNF, EU:** ไฟล์ต้นทาง **ไม่มี** ค่ารายปี → **คำนวณเองจากรายเดือน**
- **หลักการ:** หากแหล่งข้อมูลมีค่ารวมอยู่แล้ว ให้ตรวจสอบว่าคำนวณตรงกันหรือไม่ ถ้าตรง → ใช้ค่าในตาราง ถ้าไม่ตรงและไม่เข้าใจ → ถาม Aong ก่อน

### ประเภท metric (ใช้เมื่อต้องคำนวณเอง)

ค่ารายปีถูกคำนวณต่างกันตามประเภทของ metric:

| ประเภท | เงื่อนไข | วิธีคำนวณค่ารายปี | ตัวอย่าง |
|--------|----------|-------------------|---------|
| **ค่าสะสม (Cumulative)** | หมวด 5. ยอดหนี้ค่าน้ำค้างชำระ | ค่า ณ **เดือนสุดท้าย**ของปี | 5.1 ราชการ-จำนวนฉบับ, - จำนวนเงิน, 5.2 เอกชน-จำนวนฉบับ, - จำนวนเงิน |
| **อัตราส่วน (Rate)** | หน่วยเป็น `%`, `/`, หรือชื่อมี `รายได้/รายจ่าย` | **ค่าเฉลี่ย** (avg) ของเดือนที่มีข้อมูล | 2.5 อัตราน้ำสูญเสีย, 2.6 อัตราการใช้น้ำ, 6.1-6.3 |
| **ปริมาณ (Volume)** | หน่วยอื่นๆ (ลบ.ม., บาท, ราย, ฯลฯ) | **ผลรวม** (sum) ของทุกเดือน | 2.1 ปริมาณน้ำจำหน่าย, 3.1 ค่าจำหน่ายน้ำ, ฯลฯ |

### กรณีปีงบประมาณ vs ปีปฏิทิน

| กรณี | ค่าสะสม (เดือนสุดท้าย) | แหล่งข้อมูล |
|------|----------------------|------------|
| **ปีงบประมาณ** | ค่าของ **ก.ย.** (index 11 ใน fiscal array) | ใช้ `total` จากแหล่งข้อมูลก่อน ถ้าไม่มีจึงคำนวณ |
| **ปีปฏิทิน** | ค่าของ **ธ.ค.** (index 11 ใน calendar array) | **คำนวณเองทั้งหมด** (เพราะต้องประกอบจาก 2 ปีงบฯ) |

### ฟังก์ชันที่เกี่ยวข้อง
- `isRateMetric(sheet, metric)` — ตรวจว่าเป็นอัตราส่วน (avg)
- `isCumulativeInSheet(sheet, metric)` — ตรวจว่าอยู่ในหมวด 5 (last)
- `getYearlyValue(monthly)` — คืน `{sum, avg, count, last}` จาก monthly array
- `pickYearlyVal(yv, sheet, metric)` — เลือก sum/avg/last ตามประเภท

---

## รูปแบบการแสดงเวลา (ใช้ทั้ง Dashboard)

| กรณี | รูปแบบ | ตัวอย่าง |
|------|--------|---------|
| เดือน (อย่างเดียว) | ตัวย่อเดือน | ม.ค., ก.พ., มี.ค. |
| ปีงบประมาณ | ปีงบฯ + 2 หลักท้าย | ปีงบฯ 69 |
| ปีปฏิทิน | พ.ศ. เต็ม 4 หลัก | 2569 |
| เดือน+ปี | ตัวย่อเดือน + 2 หลักท้าย | ม.ค.69 |

**ฟังก์ชันที่เกี่ยวข้อง:**
- `fmtYearLabel(y, isCal)` — แสดงปี: ถ้าปฏิทินใช้ "2569", ถ้างบฯ ใช้ "ปีงบฯ 69"
- `fmtMonthYear(m, y)` — แสดงเดือน+ปี: "ม.ค.69"
- ตัวย่อเดือน: `MC = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.']`

---

## รูปแบบตัวเลขแกน Y (ใช้ทั้ง Dashboard)

| เงื่อนไข | รูปแบบ | ตัวอย่าง |
|----------|--------|---------|
| ค่าน้อยกว่า 100 (ไม่เกิน 2 หลัก) | ทศนิยม 2 ตำแหน่งเสมอ | 5.00, 12.50, 99.99 |
| ค่าตั้งแต่ 100 ขึ้นไป | ใช้ comma separator ตามปกติ | 100, 1,234, 50,000 |

---

## รูปแบบข้อความ Label ในกราฟ (ใช้ทั้ง Dashboard)

| ตำแหน่ง | กฎ | ตัวอย่าง |
|---------|-----|---------|
| **Label ในกราฟ** (legend, tooltip, dataset name) | **ตัดเลขข้อหน้าข้อความออก** | "ปริมาณน้ำจำหน่าย" (ไม่ใช่ "2.1 ปริมาณน้ำจำหน่าย") |
| **Listbox / Dropdown ตัวเลือก** | **เก็บเลขข้อไว้เหมือนเดิม** | "2.1 ปริมาณน้ำจำหน่าย" |

**ฟังก์ชันที่เกี่ยวข้อง:** ใช้ regex ตัด prefix เช่น `label.replace(/^\d+[\.\)]\s*/,'')` เฉพาะตอนแสดงในกราฟ

---

## ข้อตกลงระบบ Upload ข้อมูล (Upload Data Convention)

เมื่อต้องเพิ่มหมวดข้อมูลใหม่ หรือทำระบบ Upload ให้ Dashboard อื่น ให้ใช้รูปแบบเดียวกันนี้:

### 1. สถาปัตยกรรม (Architecture) — Dual Mode: Server + Fallback

**Primary: Server Mode (Flask)**
- **`server.py`** — Python Flask server รับ upload ไฟล์ Excel → parse ด้วย openpyxl/xlrd → auto-rename → เก็บใน `uploaded_data/`
- **ไม่ต้องใช้ template** — ใช้ไฟล์ข้อมูลดิบ (raw data) รูปแบบเดียวกับที่มีอยู่ในโฟลเดอร์ ข้อมูลดิบ/
- **ข้อมูลเก็บถาวรเป็นไฟล์** — ไฟล์ Excel ต้นฉบับ (auto-renamed) + `data.json` (parsed data) → ไม่จำกัดขนาดเหมือน localStorage
- **Auto-rename:** PR → `PR_YY-MM.xlsx`, AON → `AON_YY-MM.xls` — ตั้งชื่อตามเดือนที่ตรวจจับได้
- **API Endpoints:**
  - `GET /api/ping` — ตรวจสอบ server ทำงาน
  - `GET /api/data` — ข้อมูลทั้งหมด (JSON)
  - `POST /api/upload/pr` + `POST /api/upload/aon` — รับ upload ไฟล์
  - `DELETE /api/data/pr/<mk>` + `DELETE /api/data/aon/<mk>` — ลบเดือน
  - `POST /api/data/edit/pr` + `POST /api/data/edit/aon` — แก้ไขข้อมูล
  - `DELETE /api/data/clear` — ล้างทั้งหมด
- **เริ่มใช้งาน:** ดับเบิลคลิก `start_server.bat` **(อยู่ที่ Root folder)** → เปิด browser ที่ `http://localhost:5000` อัตโนมัติ
- **ย้ายขึ้น server จริง:** เปลี่ยน host/port ใน `server.py` + ติดตั้ง dependencies จาก `requirements.txt`

**Fallback: localStorage Mode (เมื่อไม่มี server)**
- หน้าเว็บตรวจจับอัตโนมัติ (`_checkServer()` → `/api/ping`) — ถ้า server ไม่ตอบ ใช้ SheetJS parse ฝั่ง browser
- **บีบอัดข้อมูลด้วย LZ-String** (`lz-string.min.js` CDN) ลดขนาดใน localStorage ~60-80%
- **แสดงพื้นที่ localStorage ที่ใช้ไป** — progress bar (สีเขียว <50%, เหลือง 50-80%, แดง >80%)
- **ข้อจำกัด:** localStorage ~5-10 MB ต่อ domain

**ทั้งสองโหมด:**
- **ปุ่ม "ล้างข้อมูลที่ Upload ไว้"** สำหรับลบข้อมูลทั้งหมด (ทั้ง server + localStorage, ต้องยืนยัน → reload)

### 2. UI ของโซน Upload
- **แยกโซนตามหมวดข้อมูล** — แต่ละหมวดมีกรอบ dashed border + สีประจำหมวด
- **แต่ละโซนมี:**
  - หัวข้อ + คำอธิบายสั้น
  - `<details>` คำแนะนำการตั้งชื่อไฟล์ **เฉพาะหมวดนั้น** (แยกกัน ไม่รวม)
  - ตัวเลือกโหมด: **อัตโนมัติ** (default) / **ระบุเดือนเอง** (manual)
  - Manual mode: dropdown ปี พ.ศ. + เดือน (ซ่อนไว้จนกว่าจะเลือก manual)
  - Drop zone: ลากวาง หรือ คลิกเลือกไฟล์ — รองรับ **หลายไฟล์พร้อมกัน** (`multiple`)
  - Status area แสดงผลลัพธ์
- **Confirmation dialog** — ก่อน upload จะขึ้น dialog ยืนยัน แสดงรายชื่อไฟล์ + โหมด ให้กดยืนยัน/ยกเลิก
- **Upload log** (shared) ด้านล่าง — font monospace, max-height:200px, scroll
- **ปุ่มล้างข้อมูล** + **ข้อมูลที่บันทึกไว้ label** แสดงสถานะ localStorage

### 3. Auto-detect เดือน/ปี (สำคัญมาก)
- **ลำดับการตรวจจับ (fallback chain):**
  1. อ่านจาก **ชื่อไฟล์** (เช่น pattern `YY-MM` สำหรับ PR)
  2. อ่านจาก **เนื้อหาในไฟล์** (เช่น row 3 สำหรับ PR, ชื่อ Sheet/header สำหรับ AON)
  3. ถ้าตรวจจับไม่ได้ → ใช้ค่า **manual** ที่ผู้ใช้ระบุ
  4. ถ้า manual ก็ไม่ได้ระบุ → **แจ้ง error** พร้อมชื่อไฟล์
- **ไฟล์แต่ละหมวดอาจมีวิธีตรวจจับต่างกัน** — ให้ศึกษาจากไฟล์ดิบก่อนเขียน parser

### 4. Parser (ตัวอ่านข้อมูล)
- **ฟังก์ชันแยกเป็น `processSingle___File(file, manualMK)`** — อ่าน 1 ไฟล์ return Promise
- **Multi-file:** ใช้ Promise chain (sequential) — ไม่ใช่ Promise.all — เพื่อให้ log แสดงทีละไฟล์
- **ใช้ `readFileAsArrayBuffer(file)`** → `XLSX.read(buf, {type:'array'})` → `sheet_to_json({header:1})`
- **Branch name normalization:** ใช้ `BRANCH_NAME_MAP` + `normBranch()` — ตัด "สาขา" นำหน้า + map ชื่อพิเศษ
- **Month key format:** `"YY-MM"` เช่น `"69-04"` = เมษายน พ.ศ. 2569
- **ค่าที่เป็นสเกล 0-1 ต้องคูณ 100** เป็น % (ตรวจจากค่า ≤ 1.5)

### 5. หลัง Upload สำเร็จ
- **บันทึก localStorage** ทันที — `saveUploadedDataToLS()` เก็บเฉพาะข้อมูลที่ upload (ไม่ซ้ำกับ embed) โดยบีบอัดด้วย `lsSaveCompressed()`
  - Key: `pr_uploaded_data` (PR), `pr_uploaded_aon` (AON)
  - Track ด้วย `window._uploadedPRMonths` / `window._uploadedAONMonths`
- **อัปเดต localStorage usage bar** — `updateLSUsageBar()` แสดงขนาดที่ใช้ + อัตราบีบอัด
- **`refreshAllAfterUpload()`:**
  - **Sort** arrays เดือนทั้งหมด (`.sort()`)
  - **Clear & rebuild** ทุก `<select>` ที่เกี่ยวข้อง — ระวังอย่าลบ option แรกที่เป็น hardcoded (เช่น "รวม เขต 1")
  - **Destroy & re-create** ทุก Chart.js instance ที่ใช้ข้อมูลนี้ — ต้อง `.destroy()` ก่อน
  - **Rebuild checkboxes/toggles** ให้ตรงกับ branch order ใหม่ — ต้อง match สี (`accent-color`) ตาม original style
  - **อัปเดต UI อื่นๆ:** เช่น info panel maxHeight, file list

### 5.1 โหลดข้อมูลจาก localStorage ตอนเปิดหน้า
- **`loadUploadedDataFromLS()`** เรียกใน IIFE ก่อน session restore
- Merge ข้อมูลเข้า DATA/AON → เรียก `refreshAllAfterUpload()` ใน setTimeout(200ms)
- แสดง status "โหลดข้อมูลที่บันทึกไว้: ..." ในแต่ละโซน + update LS info label

### 5.2 จัดการข้อมูลที่ Upload ไว้ (Management Panel)
- **แสดงอัตโนมัติ** เมื่อมีข้อมูลที่ upload ไว้ (ทั้งจาก localStorage หรือ upload ครั้งนี้)
- **แต่ละเดือนเป็น chip** แสดง: ชื่อเดือน + ปุ่ม ✏️ แก้ไข + ปุ่ม ✕ ลบ
- **ลบเฉพาะเดือน:** `deleteUploadedMonth(type, mk)` — confirm → ลบจาก tracking → save LS → reload
- **แก้ไขตัวเลข:** `editUploadedMonth(type, mk)` — เปิด/ปิด editable table
  - **PR:** ตาราง สาขา × (จำนวนลูกค้า + 10 หมวดหมู่ + รวม) — input type=number ทุกช่อง, auto-recalculate รวมสาขา
  - **AON:** ตาราง สาขา × Always-On (%) — input type=number step=0.01
  - ปุ่ม "✓ บันทึก" → อัปเดต DATA/AON object → save LS → refreshAllAfterUpload → alert
- **ล้างทั้งหมด:** `confirmClearUploadedData()` — confirm dialog → `clearUploadedDataFromLS()` → `location.reload()`

### 5.3 แก้ไขข้อมูลย้อนหลัง (Edit Existing/Embedded Data)
- **ส่วน `<details id="editExistingDataWrap">`** — collapsible section สำหรับแก้ไขข้อมูลทุกเดือน (ทั้ง embedded ใน HTML และที่ upload มา)
- **PR:** dropdown เลือกเดือนจาก `DATA.all_months` → กดปุ่ม → เปิด editable table (สาขา × หมวด)
- **AON:** dropdown เลือกเดือนจาก `AON_ALL_MONTHS` → กดปุ่ม → เปิด editable table (สาขา × %)
- **`initEditExistingSelects()`** — populate dropdowns ด้วยเดือนทั้งหมด, default = เดือนล่าสุด; เรียกจาก `refreshAllAfterUpload()` และ page load
- **`openExistingEdit(type)`** — อ่านเดือนที่เลือก → เรียก `renderPREditTable()` / `renderAONEditTable()` เดียวกับ Management Panel
- **หลักการ:** เมื่อแก้ไขข้อมูล embedded แล้วกดบันทึก → mark `_uploadedPRMonths[mk]=true` / `_uploadedAONMonths[mk]=true` → ข้อมูลที่แก้จะถูก save ลง localStorage เป็น "override" ของข้อมูล embedded เดิม

### 6. ข้อควรระวัง
- **`const` object mutation:** `const DATA = {...}` อนุญาตให้ push/add key ได้ แต่ห้าม reassign ตัวแปร
- **`var` reassignment:** ตัวแปรที่ประกาศ `var` (เช่น `AON`, `AON_ALL_MONTHS`) สามารถ reassign ได้เต็มที่
- **Chart.js instance:** ต้อง destroy ก่อน re-create เสมอ ไม่งั้นจะ memory leak + กราฟซ้อนทับ
- **Select options ซ้ำ:** เมื่อ re-init ฟังก์ชันที่ append options ต้อง clear ก่อน แต่เก็บ option แรกที่ hardcoded ไว้
- **Error isolation:** ไฟล์ที่ error ไม่ควรหยุดไฟล์อื่น — ใช้ `.catch()` ในแต่ละ iteration ของ Promise chain
- **Upload log:** แสดงทั้งสำเร็จ (สีเขียว) และ error (สีแดง) พร้อมชื่อไฟล์

### 8. เมื่อเพิ่มหมวดใหม่ (Checklist)
1. ศึกษาไฟล์ดิบในโฟลเดอร์ ข้อมูลดิบ/ — ดูโครงสร้าง header, data rows, column mapping
2. สร้าง Upload zone ใหม่ (HTML) — สีประจำหมวด, `<details>` คำแนะนำแยก
3. เขียน `processSingle___File()` — auto-detect เดือน + parse data + normBranch
4. เขียน `handleUpload___()` — **ผ่าน `showUploadConfirm()` ก่อนเสมอ** + multi-file Promise chain + log
5. เพิ่ม tracking: `window._uploaded___Months[mk]=true` ใน parser
6. เพิ่ม localStorage key + save/load ใน `saveUploadedDataToLS()` / `loadUploadedDataFromLS()`
7. อัปเดต `refreshAllAfterUpload()` — เพิ่ม rebuild ของ select/chart/table ใหม่
8. อัปเดต `confirmClearUploadedData()` — เพิ่ม clear key ของหมวดใหม่
9. ทดสอบ: upload ไฟล์จริง → ตรวจกราฟ → refresh หน้า → ตรวจว่าข้อมูลยังอยู่ → กดล้าง → ตรวจว่ากลับเป็นค่าเดิม

---

## Checkpoint ล่าสุด
> **เซสชัน Dashboard_PR Server-based Upload — 2026-03-18**
> - **คำสั่ง:** เปลี่ยนระบบ upload จาก localStorage เป็น server-based (Flask) พร้อม localStorage fallback
> - **สิ่งที่ทำ:**
>   1. สร้าง `server.py` (Flask) — รับ upload ไฟล์ Excel, parse ด้วย Python (openpyxl/xlrd), auto-rename, เก็บใน `uploaded_data/`
>   2. สร้าง `requirements.txt` (flask, openpyxl, xlrd)
>   3. สร้าง `start_server.bat` **(ย้ายไป Root folder)** — ดับเบิลคลิกเริ่ม server ทุก Dashboard + เปิด browser อัตโนมัติ
>   4. แก้ไข `index.html` — เพิ่ม server API functions (`_checkServer`, `_serverUpload`, `_serverGetData`, `_serverDelete`, `_serverEdit`, `_serverClearAll`)
>   5. แก้ไข `_doUploadPR` + `_doUploadAON` — try server first, fallback to SheetJS
>   6. แก้ไข page load IIFE — try server `/api/data` first, fallback to `loadUploadedDataFromLS()`
>   7. แก้ไข `deleteUploadedMonth`, `savePREdits`, `saveAONEdits`, `confirmClearUploadedData` — sync กับ server
>   8. ทดสอบ: server เริ่มได้ / API ping สำเร็จ / upload PR 22 สาขาถูกต้อง / upload AON 22 สาขาถูกต้อง / auto-rename ทำงาน
>   9. อัปเดต README Architecture section + folder structure + Changelog
> - **งานค้าง:** Aong ทดสอบบนเครื่องจริง (ดับเบิลคลิก start_server.bat)

---

## บันทึกการทำงาน (Changelog)

### 2026-03-18 — เปลี่ยนเป็น Server-based Upload (Flask) + localStorage Fallback
- สร้าง `server.py` (Flask):
  - API: `/api/ping`, `/api/data`, `/api/upload/pr`, `/api/upload/aon`, `/api/data/edit/*`, `/api/data/clear`
  - Python parser สำหรับ PR (GUI_019 format) + AON (Always-On sheets) — replicate จาก SheetJS logic
  - Auto-rename: PR → `PR_YY-MM.xlsx`, AON → `AON_YY-MM.xls`
  - เก็บข้อมูลใน `uploaded_data/data.json` + ไฟล์ต้นฉบับใน `uploaded_data/pr/` + `uploaded_data/aon/`
  - Branch name normalization (BRANCH_NAME_MAP 22+ entries)
  - CORS support, Thai month detection (ทั้งชื่อเต็ม + ย่อ)
- สร้าง `requirements.txt` (flask, openpyxl, xlrd)
- สร้าง `start_server.bat` **(ย้ายจาก Dashboard_PR/ ไป Root folder)** — ตรวจ Python + ติดตั้ง dependencies + เริ่ม server ทุก Dashboard + เปิด browser
- แก้ไข `index.html`:
  - เพิ่ม Server API helpers: `_checkServer()`, `_serverUpload()`, `_serverGetData()`, `_serverDelete()`, `_serverEdit()`, `_serverClearAll()`
  - เพิ่ม Merge functions: `_mergeServerPR()`, `_mergeServerAON()` — merge server data เข้า DATA/AON objects
  - แก้ `_doUploadPR` / `_doUploadAON` → try server → fallback to `_doUploadPRLocal` / `_doUploadAONLocal` (SheetJS)
  - แก้ page load IIFE → try `/api/data` → fallback to `loadUploadedDataFromLS()`
  - แก้ `deleteUploadedMonth` → ลบจาก server ด้วย
  - แก้ `savePREdits` / `saveAONEdits` → sync กับ server ด้วย
  - แก้ `confirmClearUploadedData` → clear ทั้ง server + localStorage

### 2026-03-18 — เพิ่มระบบ Upload + ยืนยัน + localStorage + บีบอัด + แก้ไขข้อมูลย้อนหลัง
- Dashboard_PR/index.html:
  - เพิ่ม SheetJS CDN (`xlsx.full.min.js`) + LZ-String CDN (`lz-string.min.js`)
  - เพิ่ม Upload UI ในส่วน แหล่งข้อมูล: 2 โซน PR (สีส้ม) / AON (สีส้มอ่อน) พร้อม drag & drop
  - รองรับเลือกหลายไฟล์พร้อมกัน (multi-file) ด้วย Promise chain sequential processing
  - **PR Parser (GUI_019 format):**
    - อ่าน header rows 5-6 (merged cells), data rows 7-28
    - Auto-detect เดือนจากชื่อไฟล์ (`YY-MM` pattern) หรือ row 3 (ข้อความ "ถึง เดือน...พ.ศ....")
    - ตัด numbering prefix จาก category names ("1. ด้านปริมาณน้ำ" → "ด้านปริมาณน้ำ")
    - Update DATA.data, DATA.branches, DATA.months, DATA.all_months
  - **AON Parser:**
    - Scan ทุก Sheet หา column ที่มี header "always on ..."
    - อ่านเดือนจากชื่อ Sheet (เช่น "ต.ค.68") หรือ "always on" header
    - ค่า 0-1 scale → คูณ 100 เป็น percentage
    - Update AON object, rebuild AON_ALL_MONTHS, AON_MONTH_LABELS, AON_BRANCH_ORDER
  - เพิ่ม `BRANCH_NAME_MAP` (22 entries) + `normBranch()` สำหรับ normalize "สาขาXXX" → "XXX"
  - เพิ่ม `refreshAllAfterUpload()` — rebuild ทุก select, destroy & re-create charts ทุก tab
  - แยกคำแนะนำการตั้งชื่อไฟล์เป็น 2 `<details>` แยกตามหมวด (PR / AON) อยู่ภายในโซน upload แต่ละอัน
  - เพิ่ม **Confirmation dialog** (`showUploadConfirm()`) — ก่อน upload แสดงรายชื่อไฟล์ + โหมด ให้กดยืนยัน/ยกเลิก
  - เพิ่ม **localStorage persistence:**
    - `saveUploadedDataToLS()` — บันทึกเฉพาะข้อมูลที่ upload (track ด้วย `_uploadedPRMonths`/`_uploadedAONMonths`)
    - `loadUploadedDataFromLS()` — โหลดตอนเปิดหน้า merge เข้า DATA/AON → `refreshAllAfterUpload()`
    - `clearUploadedDataFromLS()` + `confirmClearUploadedData()` — ปุ่มล้าง + confirm + reload
    - Keys: `pr_uploaded_data`, `pr_uploaded_aon`
  - เพิ่มปุ่ม "ล้างข้อมูลที่ Upload ไว้" + label แสดงสถานะข้อมูลที่บันทึก
  - เพิ่ม **Management Panel** (`renderUploadMgmtPanel()`):
    - แสดง chip ทุกเดือนที่ upload ไว้ แยก PR/AON
    - ปุ่ม ✏️ เปิด editable table (PR: สาขา×หมวด / AON: สาขา×%) แก้ตัวเลขแล้วกดบันทึก
    - ปุ่ม ✕ ลบเฉพาะเดือน (confirm → reload)
    - Auto-recalculate รวมสาขา เมื่อแก้ค่าในตาราง PR
  - เพิ่ม **LZ-String compression** สำหรับ localStorage:
    - `lsSaveCompressed(key,obj)` — compress ด้วย `LZString.compressToUTF16()` ก่อน save
    - `lsLoadDecompressed(key)` — ลอง decompress ก่อน, fallback เป็น plain JSON (backward compatible)
    - Validation: `JSON.parse()` หลัง decompress เพื่อป้องกัน garbage data จาก old format
    - ลดขนาดข้อมูลใน localStorage ~60-80%
  - เพิ่ม **localStorage usage indicator** (`updateLSUsageBar()`):
    - Progress bar สีเขียว (<50%) / เหลือง (50-80%) / แดง (>80%) ของ ~5 MB
    - แสดงขนาดบีบอัด + อัตราการบีบอัด "(บีบอัดจาก XXX, ลด XX%)"
    - Warning "⚠️ ใกล้เต็ม!" เมื่อ >80%
    - `getUploadLSUsage()` — คำนวณขนาดจริง (compressed) + ขนาดก่อนบีบอัด (uncompressed)
  - แก้ **bug compression ratio -7012538%:**
    - สาเหตุ: `decompressFromUTF16()` บน plain JSON เก่าคืน garbage string (ไม่ null) ทำให้ uncompressed เล็กกว่า compressed
    - แก้ไข: เพิ่ม `JSON.parse()` validate หลัง decompress + check `json.length > val.length`
    - Ratio text แสดงเฉพาะเมื่อ `uploadRaw > uploadUsed`
  - เพิ่ม **"แก้ไขข้อมูลย้อนหลัง"** (`<details id="editExistingDataWrap">`):
    - PR: `<select>` เดือนจาก `DATA.all_months` → editable table (สาขา × หมวด)
    - AON: `<select>` เดือนจาก `AON_ALL_MONTHS` → editable table (สาขา × %)
    - `initEditExistingSelects()` populate dropdowns ทุกครั้งที่ refresh
    - `openExistingEdit(type)` เปิด editor เดียวกับ Management Panel
    - เมื่อบันทึก → mark `_uploadedXXXMonths[mk]=true` → save ลง localStorage เป็น override ของข้อมูล embedded

### 2026-03-17 — เพิ่มตารางกิจกรรมลดน้ำสูญเสีย (Tab WSC-R Card 2)
- Dashboard_Leak/index.html:
  - เพิ่ม `ACTIVITIES_MONTHS` (13 เดือน) + `ACTIVITIES` (22 สาขา) จาก กิจกรรมลดน้ำสูญเสีย.xlsx Sheet "สรุป"
  - เพิ่ม `<div id="wscRActivitiesWrap">` ใน Card 2 ใต้ flex row (กราฟ 1+2)
  - เพิ่ม `wscRActivitiesUpdate()` สร้างตาราง เชิงกายภาพ/เชิงพาณิชย์/รวม × เดือนที่มีข้อมูล
  - เรียก `wscRActivitiesUpdate()` จาก `wscRSharedUpdate()` ทำให้เปลี่ยนสาขาแล้วตารางอัปเดตตาม
  - Branch name mapping: "ชลบุรี(พิเศษ)" → "ชลบุรี(พ)", "พัทยา(พิเศษ)" → "พัทยา(พ)"

### 2026-03-17 — Dashboard_PR UI/UX ปรับปรุง + Dashboard_Leak/GIS ปุ่มซ่อนตัวเลือก
- Dashboard_PR/index.html:
  - แก้ line/area chart ไม่เริ่มจากจุดเริ่มต้น (x-axis offset dynamic)
  - แก้ AON_MONTH_LABELS ให้ใช้ fmtMonthYear() ตามข้อตกลง (ต.ค.68 แทน ตุลาคม 2568)
  - เปลี่ยน toggle สาขาเป็น .toggle-group + .toggle-btn, checkbox เป็น flex-wrap
  - เพิ่ม กราฟพื้นที่ (area) + เรียงตัวเลือกตามข้อตกลง (เส้นโค้ง → เส้นตรง → แท่ง → พื้นที่)
  - จัดเรียง controls: ช่วงเวลา → ประเภทกราฟ → โหมด(กราฟ/ตาราง)
  - เพิ่ม collapse button (▲/▼) ทุก Card ทั้ง 3 tabs (11 cards)
  - เพิ่มปุ่ม "👁 ซ่อนตัวเลือก" บน Tab Bar (global toggle ด้วย body.hide-controls CSS)
  - เปลี่ยนชื่อกราฟ Tab 3 เป็น dynamic title (aonChartTitle / aon2ChartTitle)
  - แทนที่ fiscal year selector ด้วย month range selector (เดือนที่แสดง + ย้อนหลังถึง)
  - เปลี่ยนจาก MONTHS (13 เดือน) เป็น ALL_MONTHS (39 เดือน: 66-01 ถึง 69-03)
  - เพิ่ม sessionStorage tab persistence + session restore ท้ายสุด script
  - เพิ่ม Chart.resize() ใน switchMainTab แก้กราฟเพี้ยนเมื่อ F5
  - max-width:1400px, tooltip nearest+intersect กราฟ 2 Tab 3
  - สี Header เข้มขึ้น (#c43e00/#e67c00/#8f2800)
- Dashboard_Leak/index.html:
  - max-width:1400px
  - เพิ่มปุ่ม globalCtrlToggle + CSS hide-controls + toggleGlobalControls()
  - ลบปุ่ม "ซ่อนปุ่มตัวเลือก" เดิมใน Tab WSC-R, เพิ่ม .wscr-option ใน hide-controls CSS
- Dashboard_GIS/index.html:
  - เพิ่มปุ่ม globalCtrlToggle + CSS hide-controls + toggleGlobalControls()

### 2026-03-19 — แก้ไขทั่วไป + Dashboard GIS (Tab แรงดันน้ำ + งานค้างซ่อม)
- ทุก Dashboard (PR/Leak/GIS/Meter):
  - ปุ่มจัดการแหล่งข้อมูล: smart onclick ตรวจ protocol/hostname
    - localhost → เปิดปกติใน Tab เดิม
    - file:// → redirect ไป localhost ใน Tab เดิม
    - GitHub Pages / URL อื่น → alert แจ้งใช้ได้เฉพาะ Local
- Dashboard_Leak/index.html:
  - Tab P3 ตาราง "สรุปจุดแรงดันอ่อน" เรียงตาม OIS (RLC_STD_BRANCHES, ตัด "(พ)" เมื่อเทียบ)
  - Tab P3 dropdown หน่วยงาน (p3GetBranches) เรียงตาม OIS
  - หมายเหตุสาขาที่ไม่มีรายงานจุดแรงดันอ่อน (สีเขียว #16a34a)
  - p3HideZero default unchecked (แสดงข้อมูลทั้งหมด)
- Dashboard_GIS — Tab "แรงดันน้ำ" (ใหม่):
  - สร้างโฟลเดอร์ ข้อมูลดิบ/แรงดันน้ำ/ + category 'pressure' (server.py + manage.html)
  - Parse 22 ไฟล์ pwa_pressure_*.xlsx คำนวณเฉลี่ยแรงดัน → embed PRESSURE_DATA (5 เดือน)
  - กราฟแท่ง 22 สาขา + เส้นเกณฑ์แดง (threshold, default 0.85, ปรับได้)
  - แท่งต่ำกว่าเกณฑ์ = สีแดง, ปกติ = teal + ค่าตัวอักษรสีตามแท่ง
  - แสดง "เกณฑ์ X.XX" บนเส้นเกณฑ์, min Y=0, default แสดงค่า
  - ตัวเลือกประเภทกราฟ (เส้นโค้ง/ตรง/แท่ง/พื้นที่) + export bar ครบ
- Dashboard_GIS — Tab "งานค้างซ่อม" (ใหม่):
  - สร้าง category 'pending' (server.py + manage.html สีส้ม #e65100)
  - Parse งานซ่อม_กปภ_เขต1.xlsx → สะสมรายเดือนรายสาขา (547 งาน, จับสาขาจากสถานที่)
  - Embed PENDING_REPAIR_DATA (6 เดือน) นับทุกแถว (รวม "แก้ไข")
  - ตัวเลือกแกน X: ช่วงเวลา (เส้นรายสาขา) / สาขา (แท่งรายเดือน)
  - ตัวเลือกสาขา 3 โหมด: สาขาเดียว (dropdown) / บางสาขา (checkbox) / ทุกสาขา
  - ช่วงเดือน (จาก-ถึง) ใช้ได้ทั้ง 2 โหมดแกน X
  - เส้น "ภาพรวม" สีดำหนา, Legend ด้านล่าง, default แกน X = สาขา
  - Export bar ครบ + Toggle กราฟ/ตาราง
  - หมายเหตุ: 136/547 แถวจับสาขาไม่ได้ → "ไม่ระบุ" (ไฟล์ไม่มีคอลัมน์สาขา)
- Dashboard_GIS — Toggle กราฟ/ตาราง ทุก Card ทุก Tab:
  - Card 1 KPI: gisToggleView (ตารางมีอยู่แล้ว)
  - Card 2 สรุป: sumToggleView + sumBuildTable()
  - Card แรงดันน้ำ: pressureToggleView + pressureBuildTable()
  - Card งานค้างซ่อม: pendingToggleView + pendingBuildTable()
  - ปุ่ม กราฟ/ตาราง ใช้ .ymt/.ymb pattern

### 2026-03-17 — Dashboard_GIS (งานแผนที่แนวท่อ / Pipeline Mapping KPI)
- ปรับปรุง Dashboard_GIS/index.html:
  - เพิ่ม class chart-container ให้ Card 2 charts → right-click menus + font popup ทำงาน
  - เพิ่ม _getFpCharts() ให้ font popup apply ทั้ง sumChart1+sumChart2 พร้อมกัน
  - เพิ่มระบบปรับแกน: scroll zoom (ZOOM_FACTOR=0.1), click-drag pan, double-click axis dialog (Min/Max/Interval/Reset)
  - เพิ่ม CSS: .axis-dlg-overlay, .axis-dlg (รองรับ dark mode)
  - แก้ number formatting ตามข้อตกลง: th-TH locale, ค่า<100→2 ทศนิยม, ≥100→comma
  - แก้ bar spacing: barPercentage:1.0, categoryPercentage ปรับแยกกราฟ (Card 1=0.5, Card 2=0.7)
  - เพิ่มเส้นแกนตั้ง x-axis grid แบ่งสาขา
  - เปลี่ยน max Y ของ Card 2 จาก 110 เป็น 100

### 2026-03-13 — เซสชันที่ 13 (PPTX WSC-R attempts + Dashboard_PR + Landing Page)
- พยายามสร้าง PPTX จาก WSC-R tab ให้เหมือนหน้าเว็บ:
  - แนวทาง A: Server-side rendering (Node.js + @napi-rs/canvas + Chart.js + pptxgenjs + NotoSansThai fonts)
    - สร้าง render_cards.js, wscr_data.json → WSC-R_Dashboard.pptx (361KB)
    - ผู้ใช้ reject: กราฟไม่เหมือนหน้าเว็บ (สร้างใหม่จาก data ไม่ใช่ capture จากเว็บจริง)
  - แนวทาง B: Chrome MCP → standalone card HTML + html2canvas capture
    - สร้าง card1.html - card4.html (Chart.js + embedded data)
    - ติด limitations: base64 blocked, hex blocked, JSON byte array truncated, file:// ไม่รองรับ
  - ยังไม่สำเร็จ — เสนอ 3 ทางเลือกให้ผู้ใช้
- สร้าง Dashboard_PR/ (งานลูกค้าสัมพันธ์) ด้วย build_dashboard.py pipeline + ธีมสีส้ม
- สร้าง index.html (root) เป็น Landing page กองระบบจำหน่าย มี 4 department cards
- ย้าย PROJECT_README.md + prompt_history.txt จาก Dashboard_Leak/ ไป main workspace level

### 2026-03-09 — เซสชันที่ 11 (EU info + Tab MNF)
- เพิ่ม EU (หน่วยไฟฟ้า/น้ำจำหน่าย) ในส่วน "ข้อมูลที่โหลดแล้ว" (สีม่วง) + stats cards
- เพิ่ม MNF reader ใน build_dashboard.py: `process_mnf_file()`, `build_mnf_embedded_data()`
  - อ่านจากโฟลเดอร์ `ข้อมูลดิบ/MNF/MNF-XXXX.xlsx`
  - Sheet ภาพรวมเขต + 22 สาขา (ข้าม "รวมกราฟสาขา")
  - 4 รายการ: MNF เกิดจริง, MNF ที่ยอมรับได้, เป้าหมาย MNF, น้ำผลิตจ่าย
  - Embed เป็น `const MNF={...}` ใน index.html
- สร้าง Tab "MNF" ใหม่ใน dashboard:
  - กราฟที่ 1: MNF ภาพรวมเขต — ดูเฉพาะ __regional__ ไม่มีตัวเลือกสาขา
  - กราฟที่ 2: MNF รายสาขา — checkbox เลือกสาขา (สาขาเดียว/บางสาขา/ทุกสาขา)
  - ทั้ง 2 กราฟ: เลือกรายการ, มุมมองรายเดือน/รายปี, ช่วงปี, ชนิดกราฟ, โหมดต่อเนื่อง/เปรียบเทียบ/ตาราง
  - Export: รีเซ็ต, แสดงค่า, Font, PNG, Copy, Excel, PowerPoint
- เพิ่ม MNF ใน "ข้อมูลที่โหลดแล้ว" (สีส้ม) + stats cards
- อัพเดท README: เพิ่มเอกสารข้อมูล MNF + EU, แก้ build_dashboard.py description

### 2026-03-09 — เซสชันที่ 10 (Multi-tab, Context Menu, Snap Fix)
- Multi-tab Custom Chart: ปุ่ม "＋" เพิ่ม Custom Chart ใหม่เป็น Tab แยก, ปุ่ม ✕ ลบ Tab (ยกเว้น Tab 1)
- State swapping architecture: ใช้ HTML container เดียว swap state เข้า-ออกเมื่อสลับ tab
- localStorage multi-instance: `ccChartMeta_v1` (meta) + `ccChartConfig_v1_N` (per-instance)
- Backward compatible: migrate old `ccChartConfig_v1` → instance 1 อัตโนมัติ
- Label prefix stripping: ใช้ `stripNum()` ตัดเลขหัวข้อจาก label ในกราฟทุก Tab (listbox ไม่ตัด)
- Right-click context menu: คลิกขวาบนเส้น → ชนิดกราฟ **เฉพาะ dataset** (line/bar/area, mixed chart) + แสดงค่า + สี/ชนิดเส้น/ความหนา; คลิกขวาพื้นที่ว่าง → สีพื้นหลังเท่านั้น
- แก้ mouse snap distance: เพิ่ม `getCloseElements()` helper + custom interaction mode `nearbyNearest` กำหนดระยะ 25px สำหรับ highlight/tooltip/contextmenu — ทำทุกกราฟ
- ตั้ง `Chart.defaults.interaction` = `{mode:'nearbyNearest', intersect:false}` เป็นค่าเริ่มต้นทุกกราฟ

### 2026-03-08 — เซสชันที่ 9 (Tab 4 Custom Chart)
- เพิ่ม Tab 4 "Custom Chart" — ผู้ใช้สร้างกราฟเองจากข้อมูล OIS + Real Leak
- ระบบแกน Y หลายแกน (ซ้าย/ขวา) พร้อมเลือกหน่วย + "ไม่มีหน่วย" (`CC_NO_UNIT`)
- Data series dialog กรองรายการตามหน่วยของแกน Y ที่เลือก (ย้าย "ผูกกับแกน Y" ไว้เหนือ "รายการข้อมูล")
- แกน X เป็น dropdown (เวลา/หน่วยงาน) แทนปุ่ม
- Card แยกซ้าย/ขวา แสดงแกน Y + Data series เป็น inline pill tags ประหยัดพื้นที่
- Color palette grid (32 สี) + custom color picker แทน `<input type="color">`
- localStorage persistence (`ccChartConfig_v1`) — จำค่าทั้งหมดเมื่อ F5 + ปุ่ม Clear Chart
- แก้ fiscal month label ordering (ต.ค.-ธ.ค. ใช้ปีก่อนหน้า `displayYear=(i<3)?parseInt(y)-1:y`)
- Y-axis ticks format: <100 → 2 ทศนิยม, ≥100 → comma separator
- Y-axis scroll/pan targeting ใช้ actual pixel bounds (`scaleObj.left/right`) รองรับหลายแกนขวา
- เพิ่มช่องใส่ชื่อ Chart (`ccChartTitle`)
- เปลี่ยนตัวเลือกหน่วยงานจาก checkbox zone (สาขาเดียว/บางสาขา/ทุกสาขา) → dropdown เดียว (ภาพรวมเขต + 22 สาขา)
- บันทึกข้อตกลงลง README: รูปแบบการแสดงเวลา + รูปแบบตัวเลขแกน Y
- Zone ช่วงเวลา มี gray background + title "ช่วงเวลา"

### 2026-03-08 — เซสชันที่ 7-8 (Tab 3 น้ำสูญเสีย(บริหาร))
- กราฟ 2 (rlc2): ย้ายตัวเลือกสาขาไปหลังโหมดกราฟ
- กราฟ 1 (rlc3): เปลี่ยน label ภาพรวม 'กปภ.เขต 1' → 'ภาพรวมเขต' (`RLC3_REGIONAL_LABEL`)
- กราฟ 2 (rlc2): เพิ่ม checkbox เปรียบเทียบค่า OIS — แท่งซ้อนทับ (OIS กว้าง+จาง, ค่าจริงแคบ+เข้ม) ด้วยเทคนิค stack overlap (same stack + stacked x + non-stacked y) + ตาราง OIS 2 แถว
- กราฟ 1 (rlc3): ย้ายแท่ง OIS มาซ้อนค่าจริง (เฉพาะ bar chart, line/area ไม่เปลี่ยน)
- กราฟ 1 (rlc3): ค่า default ช่วงเดือน → 12 เดือนล่าสุด
- กราฟ 2 (rlc2): ค่า default → โหมดรายเดือน + ช่วง 12 เดือนล่าสุด
- กราฟ 3 (rlcChart1): เปลี่ยนชื่อหัวข้อ
- กราฟ 4 (rlc4): แก้ Y-axis ซ้าย-ขวา ปรับ scale แยกกันอิสระ (`getHoveredYScales` เช็ค position)
- ปรับปุ่มรีเซ็ต (`resetChartSize`) ให้ reset Y-axis scale ด้วย (clearAxisSettings + delete min/max/stepSize)
- กราฟ 4 (rlc4): จัดเรียง controls — สาขาแรก, โหมดกราฟท้าย
- กราฟ 3 (rlcChart1): default เดือนรายเดือน → เดือนล่าสุดที่มีข้อมูล (scan RL)
- กราฟ 4 (rlc4): default เดือนรายเดือน → เดือนล่าสุดที่มีข้อมูล + month range 12 เดือนล่าสุด

### 2026-03-08 — เซสชันที่ 5-6
- (เซสชัน 5 — ถูก compact) ปรับปรุง Graph 2 Tab 1 (เปรียบเทียบสาขา):
  - เพิ่ม year range selector (from-to) เมื่อ X=ช่วงเวลา ทั้งรายเดือน+รายปี
  - เพิ่ม month-year selector เมื่อ X=สาขา+รายเดือน
  - บังคับกราฟแท่งเมื่อ X=สาขา, ซ่อนตัวเลือกประเภทกราฟ
  - ตั้ง smart defaults: time+monthly=ปีก่อนถึงปีล่าสุด, time+yearly=5 ปีย้อนหลังถึงปีครบ, branch=ข้อมูลล่าสุด
  - Default line chart เมื่อเลือก X=ช่วงเวลา
  - Filter เฉพาะปีที่ข้อมูลครบ 12 เดือนในโหมดรายปี (ทั้ง Graph 1 และ 2)
- ปรับ UI ทั่วไป: เปลี่ยน label "มุมมอง" → "มุมมองช่วงเวลา", เปลี่ยน "ปปข.+ป หน้า1" → "ภาพรวมเขต"
- แก้ Year Color Map สีซ้ำปี 68/69, แก้สีกราฟแท่งจาง Tab 2
- (เซสชัน 6) แก้ไขวิธีคำนวณค่ารายปีให้ถูกต้อง:
  - หมวด 5 ยอดหนี้ค่าน้ำค้างชำระ → ค่าเดือนสุดท้าย (สะสม)
  - 6.3 รายได้/รายจ่าย → ค่าเฉลี่ย (อัตราส่วน)
  - บันทึกกฎลง README

### 2026-03-07 — เซสชันที่ 3 (ต่อ)
- เพิ่มกราฟที่ 2 "น้ำสูญเสีย (บริหาร)" ใน Tab 4
  - ดูเฉพาะข้อมูลบริหาร (Real Leak) เปรียบเทียบข้ามช่วงเวลา
  - เลือกข้อมูล: อัตราน้ำสูญเสีย (%) หรือ ปริมาณน้ำสูญเสีย (ลบ.ม.)
  - โหมดรายปี: เลือกช่วงปี (จากปีไหนถึงปีไหน) → แต่ละปี = 1 แท่ง/สี
  - โหมดรายเดือน: เลือกช่วงเดือน (ต.ค.68 ถึง ม.ค.69 ฯลฯ) → แต่ละเดือน = 1 แท่ง/สี
  - List เดือนสร้างจากข้อมูลจริงใน RL (แสดงเฉพาะเดือนที่มีข้อมูล)
- แก้ไข Scroll zoom ให้ fix origin แกน Y = 0 เสมอ (ทั้ง zoom, pan, reset)
- กราฟที่ 1 โหมด 2 แกน: เปลี่ยน font แกน Y ทั้งสองเป็นสีดำ
- ลดความกว้างแท่งและเพิ่มช่องว่างระหว่างสาขาทุกโหมด
- ตั้ง default ของกราฟที่ 1 เป็น "ทั้งสองอย่าง (2 แกน)"

### 2026-03-07 — เซสชันที่ 3
- แก้ไขการอ่านคอลัมน์ปริมาณน้ำสูญเสียจาก Real Leak ใน build_dashboard.py
  - ปัญหา: คอลัมน์ "น้ำสูญเสียจริง" ใน header แบ่งเป็น sub-header "ปริมาณ" กับ "อัตรา" ทำให้อ่านค่า volume ไม่ได้
  - แก้ไข: ใช้ two-pass detection — หา "น้ำสูญเสีย" ใน header row ก่อน แล้วค่อยหา sub-columns ใน sub-header
- ปรับปรุง Tab "น้ำสูญเสีย(บริหาร)" ใหม่ทั้งหมด:
  - เปลี่ยนชื่อเป็น "เปรียบเทียบน้ำสูญเสีย (บริหาร)"
  - รวมกราฟ 2 อัน เป็นกราฟเดียวพร้อม selector เลือก 3 โหมด:
    1. อัตราน้ำสูญเสีย (%) — กราฟแท่งเปรียบเทียบ OIS vs บริหาร
    2. ปริมาณน้ำสูญเสีย (ลบ.ม.) — กราฟแท่งเปรียบเทียบ OIS vs บริหาร
    3. ทั้งสองอย่าง (2 แกน) — แท่งซ้อนทับ (overlay): แกนซ้าย=อัตรา แกนขวา=ปริมาณ
       - กลุ่ม OIS (สีน้ำเงิน): แท่งปริมาณ(สีอ่อน)อยู่หลัง + แท่งอัตรา(สีเข้ม)อยู่หน้า
       - กลุ่มบริหาร (สีแดง): แท่งปริมาณ(สีอ่อน)อยู่หลัง + แท่งอัตรา(สีเข้ม)อยู่หน้า
       - ใช้ Chart.js order + barPercentage ต่างกันเพื่อซ้อนทับ
  - ทุกโหมดมีมุมมองรายปี/รายเดือน และ Export ครบ
- หมายเหตุ: น้ำสูญเสียจริง = น้ำสูญเสียบริหาร = Real Leak (สิ่งเดียวกัน)

### 2026-03-07 — เซสชันที่ 2 (ต่อ)
- เพิ่มการอ่านข้อมูล Real Leak ใน build_dashboard.py
  - อ่านไฟล์ RL-XXXX.xlsx จากโฟลเดอร์ ข้อมูลดิบ/Real Leak/
  - สนใจเฉพาะ Tab ที่ชื่อเป็นเดือน (ต.ค.68, พ.ย. 68 ฯลฯ)
  - ข้อมูลถูก embed เป็น `const RL={...}` ใน index.html
  - มี branch name normalization (map ชื่อต่างๆ ให้เป็นชื่อมาตรฐาน 22 สาขา)
- เพิ่ม Tab "น้ำสูญเสีย(บริหาร)" ใน Dashboard (Tab ที่ 4)
  - กราฟที่ 1: เปรียบเทียบอัตราน้ำสูญเสีย OIS vs บริหาร (%)
  - เลือกดูแบบ "รายปี" (เฉลี่ยทั้งปี) หรือ "รายเดือน" (เลือกเดือน)
  - เลือกปีงบประมาณเดียว (ไม่ใช่ช่วงปี)
  - มีปุ่ม Export (PNG, Excel, PowerPoint), Font, รีเซ็ต, แสดงค่า เหมือน Tab อื่น
- ปรับ OIS branch name matching ให้รองรับชื่อ sheet แบบ "ป.ชลบุรี น.3"

### 2026-03-07 — เซสชันที่ 2
- ทบทวนโครงสร้างโปรเจ็คทั้งหมด
- อธิบายโครงสร้างข้อมูล Real Leak (RL-2569.xlsx)
- กำหนดรายชื่อ 22 สาขามาตรฐาน
- สร้างไฟล์ PROJECT_README.md นี้

### ก่อน 2026-03-07 — เซสชันแรก
- สร้างโครงสร้างโฟลเดอร์โปรเจ็ค
- สร้าง build_dashboard.py (อ่านข้อมูล OIS)
- สร้าง index.html
- สร้าง อัพเดท Dashboard.bat
