# กฎการตั้งชื่อไฟล์อัตโนมัติ (Auto-Rename Rules)
> **สำคัญ:** Claude ต้องอ่านไฟล์นี้ทุกครั้งก่อนแก้ไขโค้ดที่เกี่ยวกับ upload / rename
> อัปเดตล่าสุด: 1 เม.ย. 2569

---

## หลักการทั่วไป

1. **Upload เป็น 2-step เสมอ** (ทุก Dashboard, ทุกหมวด)
   - Step 1: pre-check → validate + ตั้งชื่อ + เช็คซ้ำ → แสดง preview
   - Step 2: user confirm → ย้ายไฟล์จาก temp ไปที่จริง
2. **Confirm dialog ต้องบอกชัดเจน** ว่าซ้ำหรือไม่ซ้ำ
   - ไม่ซ้ำ: `✅ ไม่มีชื่อไฟล์ซ้ำกับที่มีอยู่`
   - ซ้ำ: `⚠️ จะเขียนทับไฟล์เดิม X ไฟล์: ...`
3. **เช็คซ้ำ ใช้ stem (ไม่สน extension)** เช่น `OIS_2569.xls` กับ `OIS_2569.xlsx` ถือว่าซ้ำ
4. **ปี fallback ใช้ พ.ศ. เสมอ** → `date('Y') + 543`

---

## Dashboard_Leak (Port 5001)

| หมวด (slug) | Prefix | รูปแบบชื่อ | วิธีหาปี/วันที่ |
|---|---|---|---|
| `ois` | OIS | `OIS_YYYY.xlsx` | 1. ปี 4 หลักจากชื่อไฟล์ → 2. ปีจาก Excel ("ปีงบประมาณ XXXX") → 3. fallback พ.ศ. ปัจจุบัน |
| `rl` | RL | `RL_YYYY.xlsx` | เหมือน ois |
| `mnf` | MNF | `MNF_YYYY.xlsx` | เหมือน ois |
| `eu` | EU | `EU_YYYY.xlsx` | เหมือน ois |
| `kpi2` | KPI2 | `KPI2_YYYY.xlsx` | เหมือน ois |
| `activities` | ACT | `ACT_YYYY.xlsx` | **ดูจาก Excel อย่างเดียว ไม่สนชื่อไฟล์** → 1. ปีจาก Excel ("ปีงบประมาณ XXXX") → 2. ชื่อ Sheet นับปีที่พบมากสุด (majority vote เช่น ธ.ค.68+ม.ค.69...ก.ย.69 → 69 ชนะ → 2569) → 3. fallback พ.ศ. ปัจจุบัน |
| `p3` | P3 | `P3_สาขา_YY-MM.xlsx` | สาขา: จับคู่ BRANCH_ALIASES จากชื่อไฟล์ / วันที่: จากชื่อไฟล์ DD-MM-YY หรือ YY-MM |

### หมายเหตุ Dashboard_Leak
- `activities` อยู่ใน `$parseable` list (ให้อ่าน Excel หาปี)
- PREFIX_MAP อยู่บน api.php: `['ois'=>'OIS', 'rl'=>'RL', 'mnf'=>'MNF', 'p3'=>'P3', 'activities'=>'ACT', 'eu'=>'EU', 'kpi2'=>'KPI2']`

---

## Dashboard_PR (Port 5000)

| หมวด (slug) | Prefix | รูปแบบชื่อ | วิธีหาเดือน |
|---|---|---|---|
| `pr` | PR | `PR_YY-MM.xlsx` | 1. จากชื่อไฟล์ (regex `\d{2}-\d{2}`) → 2. จากเนื้อหา Excel (เดือนภาษาไทย) → 3. Error ถ้าหาไม่เจอ |
| `aon` | AON | `AON_YY-MM[_YY-MM...].xlsx` | จาก Excel: ชื่อ Sheet / Header / คอลัมน์ → สร้างรายชื่อเดือนที่มีข้อมูล |

### หมายเหตุ Dashboard_PR
- ปีเป็น พ.ศ. ย่อ 2 หลัก (YY = YYYY - 2500) เช่น 2569 → 69
- `make_month_key(yyyy, mm)` = `sprintf("%02d-%02d", yyyy-2500, mm)`
- AON อาจมีหลายเดือนในไฟล์เดียว → ชื่อเป็น `AON_69-01_69-02_69-03.xlsx`

---

## Dashboard_GIS (Port 5002)

| หมวด (slug) | Prefix | รูปแบบชื่อ | วิธีหาวันที่ |
|---|---|---|---|
| `repair` | GIS | `GIS_YYMMDD.xlsx` | 1. เลข 6 หลักจากชื่อไฟล์ → 2. fallback วันปัจจุบัน `date('ymd')` |
| `pressure` | PRESSURE | `PRESSURE_สาขา_ปีงบYY.xlsx` | สาขา: ตัวอักษรไทยสุดท้ายในชื่อไฟล์ / ปี: 1. `_ปีงบYY` จากชื่อไฟล์ → 2. Excel ("ปีงบประมาณ XXXX") → 3. ชื่อ Sheet |
| `pending` | (merged) | `ค้างซ่อม_MM-YY_to_MM-YY.xlsx` | ช่วงวันที่จากคอลัมน์วันที่ใน Excel (หลายไฟล์ merge เป็นไฟล์เดียว) |

### หมายเหตุ Dashboard_GIS
- `pending` รวมหลายไฟล์เป็นไฟล์เดียว (batch merge)
- `repair` ใช้ ค.ศ. ย่อ 2 หลัก (yy = CE year % 100)

---

## Dashboard_Meter (Port 5003)

| หมวด (slug) | Prefix | รูปแบบชื่อ | วิธีหาวันที่ |
|---|---|---|---|
| `abnormal` | METER | `METER_รหัสสาขา_YYYYMM.xlsx` | รหัสสาขา: 4 หลักจากชื่อไฟล์ (ตรง BRANCH_CODE_MAP) / วันที่: 1. billing_month จาก Excel → 2. data_date จาก Form |

### หมายเหตุ Dashboard_Meter
- ถ้าไม่เจอรหัสสาขา → ใช้ชื่อไฟล์ที่ sanitize แล้ว (max 30 chars)
- ปีเป็น พ.ศ. 4 หลัก (2569) + เดือน 2 หลัก = `YYYYMM` เช่น `256903`
- ต้องมี `data_date` field ใน form (ส่งเป็น พ.ศ.)

---

## สิ่งที่ห้ามทำ

- **ห้าม** ใช้ `date('Y')` โดยตรงเป็น fallback (ได้ปี ค.ศ. 2026 แทน พ.ศ. 2569)
- **ห้าม** เช็คซ้ำแค่ frontend (ต้องเช็คจาก server เท่านั้น)
- **ห้าม** แสดง confirm แบบ "หว่านแห" (เตือนกว้างๆ) ต้องระบุชัดว่าไฟล์ไหนซ้ำไฟล์ไหน
- **ห้าม** ลืมเพิ่ม `activities` ใน `$parseable` list ถ้าแก้โค้ด validate_leak_file
