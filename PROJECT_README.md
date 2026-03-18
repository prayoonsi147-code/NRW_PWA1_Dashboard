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
├── index.html              # Dashboard (standalone HTML, ~5800+ lines)
└── ข้อมูลดิบ/
    ├── เรื่องร้องเรียน/     # ไฟล์ข้อร้องเรียน GUI_019 format
    │   ├── ร้องเรียน_66-01.xlsx
    │   └── ...
    └── AlwayON/             # ไฟล์ PWA Always-on
        ├── 2.ข้อมูล PWA Always-on สะสม เดือนตุลาคม 2568 โดย กลส..xls
        └── ...
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

**Upload Feature (เพิ่ม 2026-03-18):**
- ใช้ SheetJS (XLSX library) อ่านไฟล์ Excel ในเบราว์เซอร์
- PR: อ่าน GUI_019 format, auto-detect เดือนจากชื่อไฟล์ (YY-MM) หรือ row 3
- AON: scan ทุก Sheet หา "always on" column, ค่า 0-1 → ×100
- `refreshAllAfterUpload()` rebuild ทุก select/chart หลัง upload

### Dashboard_GIS/ (งานแผนที่แนวท่อ)
```
Dashboard_GIS/
├── build_dashboard.py      # สคริปต์สร้าง Dashboard
├── index.html              # Dashboard (standalone HTML)
└── ข้อมูลดิบ/              # ไฟล์ข้อมูล GIS KPI
```

**โครงสร้าง index.html:**
- Tab: KPI จุดซ่อมท่อ
  - Card 1: KPI จุดซ่อมท่อ (canvas: gisChart, dual-axis แท่ง+เส้น)
  - Card 2: สรุปภาพรวม KPI (canvas: sumChart1 แถว1, sumChart2 แถว2)
- Chart Customization: right-click menus, font popup, scroll zoom, pan, double-click axis dialog
- Charts registered in getChartObj(): gisChart, sumChart1, sumChart2

### Dashboard_Leak/ (งานน้ำสูญเสีย)
```
Dashboard_Leak/
├── build_dashboard.py      # สคริปต์สร้าง Dashboard (Pure Python ไม่ต้องลง library เพิ่ม)
├── index.html          # Dashboard ที่สร้างแล้ว (เปิดในเบราว์เซอร์)
├── data.json               # ข้อมูลที่ประมวลผลแล้ว
├── data_embed.js           # ข้อมูลแบบ embed ใน JS
├── อัพเดท Dashboard.bat   # ดับเบิลคลิกเพื่อรัน build_dashboard.py
├── PROJECT_README.md       # ไฟล์นี้
├── ข้อมูลดิบ/
│   ├── OIS/                # ข้อมูล OIS รายปี
│   │   ├── 2558.xls
│   │   ├── ...
│   │   └── 2569.xls
│   ├── Real Leak/          # ข้อมูลน้ำสูญเสียตัวจริง รายปี
│   │   ├── RL-2568.xlsx
│   │   └── RL-2569.xlsx
│   ├── หน่วยไฟ/            # ข้อมูลหน่วยไฟฟ้า(ระบบจำหน่าย)/น้ำจำหน่าย รายปี
│   │   └── EU-2569.xlsx
│   └── MNF/                # ข้อมูล Minimum Night Flow รายปี
│       └── MNF-2569.xlsx
├── กราฟ/                   # (ว่าง)
└── รายงาน/                 # (ว่าง)
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

## Checkpoint ล่าสุด
> **เซสชัน Dashboard_PR Upload Feature + Naming Guide — 2026-03-18**
> - **คำสั่ง:** เพิ่มระบบ Upload ข้อมูลใน Dashboard_PR (PR + AON), แยกคำแนะนำการตั้งชื่อไฟล์ตามหมวด
> - **สิ่งที่ทำ:**
>   1. เพิ่ม SheetJS CDN สำหรับอ่านไฟล์ Excel ในเบราว์เซอร์
>   2. สร้าง Upload UI: 2 โซน (PR สีส้ม / AON สีส้มอ่อน) พร้อม drag & drop, multi-file, auto/manual mode
>   3. Parser สำหรับ PR (GUI_019 format): อ่าน header row 5-6, data row 7+, auto-detect เดือนจากชื่อไฟล์ (YY-MM) หรือ row 3
>   4. Parser สำหรับ AON: scan ทุก Sheet หา "always on" column, อ่านเดือนจากชื่อ Sheet/header, ค่า 0-1 → ×100
>   5. BRANCH_NAME_MAP (22 สาขา) สำหรับ normalize ชื่อ "สาขาXXX" → "XXX"
>   6. refreshAllAfterUpload() — rebuild selects, destroy & re-create charts ทุก tab
>   7. คำแนะนำการตั้งชื่อไฟล์ แยกเป็น 2 details (PR / AON) อยู่ในโซน upload แต่ละอัน
> - **งานค้าง:** ผู้ใช้กำลังทดสอบ อาจมี feedback

---

## บันทึกการทำงาน (Changelog)

### 2026-03-18 — เพิ่มระบบ Upload ข้อมูลใน Dashboard_PR + แยกคำแนะนำตามหมวด
- Dashboard_PR/index.html:
  - เพิ่ม SheetJS CDN (`xlsx.full.min.js`) สำหรับอ่านไฟล์ Excel ในเบราว์เซอร์
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
