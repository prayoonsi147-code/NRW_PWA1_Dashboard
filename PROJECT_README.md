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

## โครงสร้างโฟลเดอร์

### Dashboard_PR/ (port 5000, ส้ม #c43e00)
- Tab 1: ข้อร้องเรียน (กราฟ 1-2)
- Tab 2: รายงานข้อร้องเรียน (กราฟ 3)
- Tab 3: Always-On (กราฟ AON 1-3)
- ข้อมูลดิบ/: เรื่องร้องเรียน/ (PR_YY-MM.xlsx), AlwayON/ (AON_YY-MM.xls), data.json

### Dashboard_Leak/ (port 5001, น้ำเงิน #1e3a5f)
- Tab OIS, Tab Real Leak, Tab WSC-R, Tab MNF, Tab P3, Tab Custom Chart
- ข้อมูลดิบ/: OIS/, Real Leak/ (RL_YYYY.xlsx), MNF/ (MNF_YYYY.xlsx), P3/, Activities/, หน่วยไฟ/ (EU_YYYY.xlsx), เกณฑ์ชี้วัด/ (KPI_YYYY.xlsx)

### Dashboard_GIS/ (port 5002, teal #004d40)
- Tab 1: จุดซ่อมท่อ (KPI กราฟ 1-2) — ข้อมูลฝังโดย build_dashboard.py
- Tab 2: แรงดันน้ำ (กราฟ 1-2) — ข้อมูลจาก server.py API
- Tab 3: งานค้างซ่อม (กราฟ 1-2, ตาราง 3-4) — ข้อมูลจาก server.py API + fallback
- ข้อมูลดิบ/: ลงข้อมูลซ่อมท่อ/ (GIS_YYMMDD.xlsx), แรงดันน้ำ/ (PRESSURE_สาขา_ปีงบYY.xlsx), ซ่อมท่อค้างระบบ/

### Dashboard_Meter/ (port 5003, ม่วง #4a148c)
- ข้อมูลดิบ/: มาตรวัดน้ำผิดปกติ/ (METER_MMYY.xlsx)

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
