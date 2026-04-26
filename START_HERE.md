# START HERE — Dashboard กปภ.ข.1

> **สำหรับ Claude:** อ่านไฟล์นี้ไฟล์เดียวก่อนเริ่มงาน ถ้าเอกสารเก่า (ใน `archive/`) ขัดแย้งกับไฟล์นี้ → **ยึดไฟล์นี้เป็นหลัก**
>
> อัปเดตล่าสุด: 2026-04-20 | Owner: Aong (prayoonsi147@gmail.com)

---

## 1. กฎเหล็ก 5 ข้อ (Non-negotiables)

1. **ห้ามฝังข้อมูลดิบ (hardcode) ลง `index.html` โดยตรง ไม่ว่ากรณีใด**
   - ข้อมูลทุกชนิดต้องมี (a) API endpoint ใน `api.php` อ่านจาก Excel, (b) build function ใน `build_dashboard.php`
   - Local ใช้ API อย่างเดียว — ห้ามเรียก `build_dashboard.php` / `/api/rebuild` จาก `manage.html`
   - ยกเว้นตอน push git (`push_to_github.bat`): ฝังชั่วคราว แล้ว restore กลับใน Step 9
   - พบ hardcode → ต้องแจ้ง Aong ทันที
2. **ห้ามแก้ code โดยไม่ถาม Aong ก่อน** ถ้างานกำกวม — ถามก่อนเสมอ ลงมือเมื่อได้ "ok" ชัดเจน
3. **ห้ามสร้าง wrapper function ใหม่ถ้าของกลางมีอยู่แล้ว** (เช่น `resetChartSize`, `toggleDataLabels`, `openFontPopup`) — ดู §5
4. **ห้าม reorder `<option>` ใน `<select>`** เพื่อเปลี่ยน default → ใช้ `selected` attribute แทน
5. **ห้ามใช้ `date('Y')` เป็น fallback ปี** → ต้อง `date('Y') + 543` (พ.ศ.)

---

## 2. โครงสร้างโปรเจค

**4 Dashboard + 1 Landing** (ทั้งหมดเป็น PHP/XAMPP ไม่ใช้ Python แล้ว)

| Dashboard | Theme | Port (dev) | Tab หลัก |
|---|---|---|---|
| **Dashboard_Leak** (น้ำสูญเสีย) | น้ำเงิน `#1e3a5f` | 5001 | OIS, น้ำสูญเสีย(บริหาร), MNF, P3, Custom |
| **Dashboard_PR** (ข้อร้องเรียน) | ส้ม `#c43e00` | 5000 | ข้อร้องเรียน, รายงาน, Always-On |
| **Dashboard_GIS** (จุดซ่อม+แรงดัน+ค้างซ่อม) | Teal `#004d40` | 5002 | จุดซ่อมท่อ, แรงดันน้ำ, งานค้างซ่อม |
| **Dashboard_Meter** (มาตรวัดผิดปกติ) | ม่วง `#4a148c` | 5003 | มาตรวัดน้ำผิดปกติ |
| `index.html` (ราก) | น้ำเงิน `#0d47a1` | — | Landing เลือก Dashboard |

**22 สาขา (ชื่อมาตรฐาน):** ชลบุรี(พ), พัทยา(พ), พนัสนิคม, บ้านบึง, ศรีราชา, แหลมฉบัง, ฉะเชิงเทรา, บางปะกง, บางคล้า, พนมสารคาม, ระยอง, บ้านฉาง, ปากน้ำประแสร์, จันทบุรี, ขลุง, ตราด, คลองใหญ่, สระแก้ว, วัฒนานคร, อรัญประเทศ, ปราจีนบุรี, กบินทร์บุรี
→ ชื่อในข้อมูลดิบอาจต่างกัน ("พัทยา (พ)", "พัทยา น.1") ต้อง normalize ผ่าน `BRANCH_ALIASES`

**ไฟล์หลักแต่ละ Dashboard:**
- `index.html` — Dashboard แสดงกราฟ (API-only mode)
- `manage.html` — อัพโหลด/ลบ/หมายเหตุ
- `api.php` — GET/POST endpoints (อ่าน Excel → JSON + cache)
- `build_dashboard.php` — ฝังข้อมูลลง `index.html` ตอน push git
- `.htaccess` — PHP settings (ดู §8)
- `ข้อมูลดิบ/` — Excel ต้นฉบับ
- `.cache/` — file cache (TTL 60s, gitignored)

---

## 3. Workflow (ทุก Session)

**เริ่ม session:**
1. อ่าน `START_HERE.md` (ไฟล์นี้)
2. ตอบ Aong ว่าเข้าใจ pattern อะไรบ้าง
3. ห้ามลงมือจนกว่า Aong บอก "ok"/"ทำเลย"

**ก่อนแก้กราฟ/UI:**
- ดู §5 Component Specs
- grep ดูการใช้งานเดิม (pattern search) — ถ้ามี ของกลางอยู่แล้วต้องใช้อันนั้น
- ยืนยันกับ Aong ว่าจะแก้ canvas id อะไร ก่อนลงมือ

**ก่อนแก้ upload/rename:**
- ดู §6 Naming Rules
- ห้าม rename หลายไฟล์เป็นชื่อเดียวกัน (ตรวจ `$used_names[]`)
- ต้องเช็คซ้ำจาก server ไม่ใช่ frontend

**จบ session:**
- บันทึกสรุปสั้นๆ ลง `archive/session_logs.md` (ไม่ใช่ไฟล์นี้)
- ไม่ต้องบันทึก prompt_history (ยกเลิกแล้ว)

---

## 4. การ Deploy (Local + GitHub Pages)

**Local (XAMPP):** `http://localhost/Claude Test Cowork/index.html`
- ต้องแก้ `httpd.conf`: เปลี่ยน DocumentRoot → `C:/Users/<username>` (แก้ครั้งเดียวต่อเครื่อง)
- ต้องเปิด `extension=zip` และ `extension=gd` ใน `php.ini` (PhpSpreadsheet ต้องใช้)

**GitHub Pages:** push ผ่าน `push_to_github.bat` (9-step flow)
- **Step 1:** CHECKPOINT — copy `index.html` → `.checkpoint` ทุก Dashboard
- **Step 2:** BUILD — `php build_dashboard.php` ฝังข้อมูลล่าสุดลง HTML
- **Step 3:** VALIDATE — ตรวจ DOCTYPE + min 1KB, auto-restore ถ้า fail
- **Step 4-7:** Git init + identity + pull+squash (`git reset --soft origin/main`) + stage
- **Step 8:** Commit + push (goto RESTORE ถ้า fail)
- **Step 9:** RESTORE — คืน `index.html` จาก `.checkpoint` (local กลับเป็นเดิม)

**แก้แค่ UI ไม่แตะ data:** ใช้ `quick_push.bat` (ไม่ run build) หรือ `git add/commit/push` ตรง

**หลัง push:** GitHub Pages cache 1-3 นาที — Hard Refresh (Ctrl+Shift+R) ถึงจะเห็นข้อมูลใหม่

---

## 5. Component Specs (UI มาตรฐาน)

### 5.1 Export Bar — โครงสร้างตายตัว (ลำดับ 7 ปุ่ม ห้ามสลับ)

```html
<div class="export-bar">
  <button class="export-btn" onclick="resetChartSize(this)" title="รีเซ็ตขนาดกราฟ">&#8634; รีเซ็ต</button>
  <label class="export-btn" style="cursor:pointer"><input type="checkbox" onchange="toggleDataLabels(this)" style="margin-right:4px">แสดงค่า</label>
  <button class="export-btn" onclick="openFontPopup(this)"><!-- A icon --></button>
  <button class="export-btn" onclick="exportPNG('<canvasId>','<title>')"><!-- image icon --></button>
  <button class="export-btn" onclick="copyToClipboard('<canvasId>',this)"><!-- copy icon --></button>
  <button class="export-btn" onclick="exportExcel('<key>')"><!-- green X icon --></button>
  <button class="export-btn" onclick="exportPPTX('<canvasId>','<title>')"><!-- orange P icon --></button>
</div>
```

- CSS: `display:flex; gap:6px; margin-top:12px; justify-content:flex-end; flex-wrap:wrap`
- ต้องอยู่**หลัง** `.chart-container` (reset ใช้ `previousElementSibling` หา)
- SVG ขนาด 18×18 (font/PNG/copy) หรือ 20×20 (Excel/PPTX)
- สำหรับ **ตาราง**: ไม่มี "แสดงค่า" + tooltip = "รีเซ็ตขนาดตาราง"

### 5.2 ปุ่ม "↺ รีเซ็ต" — ต้องทำ 4 ขั้นตอน (pre-baseline: `Dashboard_PR/index.html` ~4606)

**Step 1 — คืนขนาด container:**
```js
el.style.width = '';
el.style.height = el.getAttribute('data-default-height') || '400px';
```

**Step 2 — ล้าง axis zoom/pan:**
```js
clearAxisSettings(canvas.id);
const c = getChartObj(canvas.id);
['y','yLeft','yRight'].forEach(k=>{
  if(c?.options.scales[k]){
    delete c.options.scales[k].min;
    delete c.options.scales[k].max;
    if(c.options.scales[k].ticks) delete c.options.scales[k].ticks.stepSize;
  }
});
```

**Step 3 — Restore controls กลับค่า default** (อ่านจาก `data-dsi`, `data-dck`, `data-default-active`):
```js
const card = bar.closest('.chart-card');
if(card && card.getAttribute('data-defaults-stored')){
  card.querySelectorAll('.controls select').forEach(sel=>{
    const di = sel.getAttribute('data-dsi');
    if(di !== null) sel.selectedIndex = parseInt(di);
  });
  card.querySelectorAll('.controls input[type=checkbox]').forEach(cb=>{
    const dc = cb.getAttribute('data-dck');
    if(dc !== null) cb.checked = (dc === '1');
  });
  card.querySelectorAll('.controls select, .controls input[type=checkbox]').forEach(e=>e.dispatchEvent(new Event('change')));
  card.querySelectorAll('.ymt').forEach(grp=>{
    const defBtn = grp.querySelector('[data-default-active]');
    if(defBtn && !defBtn.classList.contains('active')) defBtn.click();
  });
}
```

**Step 4 — Reset "แสดงค่า" + re-render:**
```js
const dlCb = bar.querySelector('input[type=checkbox]');
if(dlCb){
  const dck = dlCb.getAttribute('data-dck');
  dlCb.checked = (dck !== null) ? (dck === '1') : dlCb.defaultChecked;
  toggleDataLabels(dlCb);
}
const renderFn = getChartRenderFn(canvas.id);
if(renderFn) renderFn();   // destroy + new Chart
else setTimeout(()=> c.resize(), 50);
```

**ตอน init กราฟ — เก็บ defaults (ครั้งเดียว):**
```js
if(!card.getAttribute('data-defaults-stored')){
  card.querySelectorAll('.controls select').forEach(s=> s.setAttribute('data-dsi', s.selectedIndex));
  card.querySelectorAll('.controls input[type=checkbox]').forEach(c=> c.setAttribute('data-dck', c.checked?'1':'0'));
  card.querySelectorAll('.ymt .ymb.active').forEach(b=> b.setAttribute('data-default-active','1'));
  const dl = card.querySelector('.export-bar input[type=checkbox]');
  if(dl) dl.setAttribute('data-dck', dl.checked?'1':'0');
  card.setAttribute('data-defaults-stored','1');
}
```

### 5.3 Controls Layout (ซ้าย→ขวา)
1. ประเภทกราฟ → 2. แกน X (เวลา/สาขา) → 3. เกณฑ์/metric → 4. ตัวเลือกอื่น

**HTML classes:** `<div class="card chart-card">` > `<div class="controls">` + `<div class="chart-container" style="height:320px">` + `<div class="ymt">` (toggle buttons)

**Branch Selector 3 โหมด** (แสดงเฉพาะเมื่อ xAxis = "ช่วงเวลา"):
- 1 สาขา (dropdown) / บางสาขา (checkboxes, default) / ทุกสาขา

### 5.4 Axis Y Interaction (ทุกกราฟ)

- **Scroll เม้าส์:** Origin=0 → ปรับ Max เท่านั้น; Origin≠0 → zoom รอบกึ่งกลาง
- **Pan (drag):** Origin=0 → Min=0 คงที่; Origin≠0 → อิสระ
- **Double-click:** dialog ตั้ง Min/Max/Interval
- **Right-click:** context menu (Min/Max/Interval + สี + grid style)
- โค้ด: `axisSettings{}`, `axisEnforcer` plugin, `handleWheel/MouseMove`, `showAxisDialog`, `showAxisCtxMenu`

**Paired charts (ต้อง sync Y-axis):** mapping ใน `PAIRED_CHARTS` → `syncPairedChart()` hook เข้า scroll/drag/dialog
- Leak: `wscRChart3↔4`, `wscRChart3V↔4V`
- PR: `waterQty1↔2, pipe1↔2, quality1↔2, service1↔2, staff1↔2`

### 5.5 Chart Type Options (ลำดับตายตัว)
```html
<option value="line" selected>กราฟเส้นโค้ง</option>  <!-- tension 0.3 -->
<option value="straight">กราฟเส้นตรง</option>        <!-- tension 0 -->
<option value="bar">กราฟแท่ง</option>
<option value="area">กราฟพื้นที่</option>           <!-- line + fill:true -->
```
- Context menu ตรวจ `tension === 0` เพื่อแยก straight ออกจาก line

### 5.6 Number Formatting (ทุกกราฟ)
- **Y-axis tick:** `< 100` → 2 ทศนิยม, `≥ 100` → comma, locale `th-TH`
- **X-axis:** `maxRotation: 60, font.size: 12`
- **beginAtZero:** `false`
- **เดือน:** ย่อ ("ม.ค.", "ก.พ.69"), **ปีงบ:** 2 หลัก ("ปีงบฯ 69"), **ปีปฏิทิน:** 4 หลัก (2569)

### 5.7 Note Textarea (ทุกกราฟ ทุก Dashboard)
```html
<div class="card-note-box">
  <label>📝 หมายเหตุ</label>
  <textarea id="note-<key>" oninput="saveCardNote('<key>')" placeholder="เพิ่มหมายเหตุ..."></textarea>
</div>
```
- Auto-save debounce 800ms → API `POST /api/notes/<slug>` (เก็บใน `data.json` หรือ `notes.json`)
- ห้ามเก็บแค่ localStorage (ย้ายเครื่องแล้วหาย)

### 5.8 Chart Registration (Dashboard_Leak ใช้เต็ม pattern)

เมื่อเพิ่มกราฟใหม่ ต้องลงทะเบียน 4 ที่:
```js
var myChartObj = null;
// 1. getChartObj mapping
function getChartObj(id){ return {'myChart': myChartObj, ...}[id] || null; }
// 2. getChartRenderFn mapping (สำหรับ reset)
function getChartRenderFn(id){ return {'myChart': ()=> myUpdate(), ...}[id] || null; }
// 3+4. CTX_CHART_TYPE_MAP, CTX_CHART_UPDATE_MAP (สำหรับ context menu)
```

---

## 6. Naming Rules (Upload/Rename)

**หลักการทั่วไป:**
- Upload เป็น **2-step** เสมอ (pre-check → confirm)
- Confirm dialog ต้องบอกชัด: `✅ ไม่ซ้ำ` หรือ `⚠️ จะเขียนทับ X ไฟล์: ...`
- เช็คซ้ำใช้ **stem** (ไม่สน extension) — `OIS_2569.xls` ≡ `OIS_2569.xlsx`
- ปี fallback: `date('Y') + 543` เสมอ (ห้าม `date('Y')` ตรงๆ)

**รูปแบบชื่อไฟล์:**

| Dashboard | หมวด | Prefix | รูปแบบ | หาปี/วันที่ |
|---|---|---|---|---|
| Leak | ois, rl, mnf, eu, kpi2 | OIS/RL/MNF/EU/KPI2 | `<P>_YYYY.xlsx` | ชื่อไฟล์ → Excel "ปีงบประมาณ XXXX" → fallback พ.ศ. |
| Leak | activities | ACT | `ACT_YYYY.xlsx` | Excel → majority vote จากชื่อ Sheet → fallback |
| Leak | p3 | P3 | `P3_สาขา_YY-MM.xlsx` | สาขา: BRANCH_ALIASES / เดือน: DD-MM-YY หรือ YY-MM |
| PR | pr | PR | `PR_YY-MM.xlsx` | regex `\d{2}-\d{2}` → เนื้อ Excel → error |
| PR | aon | AON | `AON_YY-MM[_YY-MM...].xlsx` | Excel (Sheet/header/col) — หลายเดือนในไฟล์เดียวได้ |
| GIS | repair | GIS | `GIS_YYMMDD.xlsx` | เลข 6 หลักจากชื่อ → fallback `date('ymd')` |
| GIS | pressure | PRESSURE | `PRESSURE_สาขา_ปีงบYY.xlsx` | `_ปีงบYY` → Excel → Sheet |
| GIS | pending | — | `ค้างซ่อม_MM-YY.xlsx` | `detect_pending_month` (ReadFilter 50 แถวแรก) |
| Meter | abnormal | METER | `METER_รหัสสาขา_YYYYMM.xlsx` | billing_month จาก Excel → data_date จาก form |

**ปีเป็น พ.ศ.:** ย่อ 2 หลัก YY = YYYY - 2500 (2569 → 69)

---

## 7. บทเรียน 12 ข้อ (สำคัญมาก — เคยเจ็บมาแล้ว)

1. **PHP Error Handling** → ตั้งค่าใน `.htaccess` เสมอ (`php_flag display_errors Off`) ไม่พึ่ง `ini_set()` อย่างเดียว
2. **Multi-file Upload** → JS ใช้ `formData.append('files[]', file)` (ต้องมี `[]`)
3. **`max_file_uploads`** → `.htaccess` ต้องตั้ง ≥ 100 (+ `upload_max_filesize 50M`, `post_max_size 200M`, `memory_limit 512M`, `max_execution_time 600`)
4. **P3 File Naming** → อ่านสาขาจาก Excel content (หา "สถานีผลิตน้ำ" หรือ match 22 สาขา) ห้าม regex `P3-xxx`
5. **Cache Invalidation** → ถ้าข้อมูลที่ cache เปลี่ยน (เช่น `rlcOISMap` เมื่อ OIS reload) ต้อง reset cache + re-init
6. **PhpSpreadsheet Performance** → ใช้ `getOldCalculatedValue()` (เร็วกว่า `getCalculatedValue()` มาก) + file cache `mtime`
7. **API-Only Mode** → `index.html` data vars เริ่มเป็นค่าว่าง โหลดทุกอย่างจาก API → callback ต้อง `rebuildAllData()` + re-init tabs
8. **PhpSpreadsheet Column Index** → **1-based** (A=1, row A1=`getCell([1,1])`, `$row[1]`=col A)
9. **Cross-Dashboard Consistency** → แก้ใน Dashboard หนึ่งแล้วต้องตรวจอีก 3 ตัว (`.htaccess`, `manage.html`, `api.php`, `build_dashboard.php`)
10. **Browser Cache** → หลังแก้โค้ด → แจ้ง user Hard Refresh (Ctrl+Shift+R) หรือ Incognito
11. **Non-breaking space (`\xC2\xA0`)** → PHP `is_numeric("\xa01")` = false → ใช้ `preg_replace('/[\xC2\xA0\s,]+/','',...)` ไม่ใช่ `trim()`
12. **`build_dashboard.php` replace logic** → ใช้ `strpos` + brace counting ไม่ใช่ `preg_replace` (PCRE backtrack fail บน line ยาว) + `ini_set('memory_limit','1024M')` สำหรับ GIS

---

## 8. ไฟล์ `.htaccess` มาตรฐาน (ทุก Dashboard)

```apache
php_flag display_errors Off
php_flag html_errors Off
php_value log_errors On
php_value max_file_uploads 100
php_value upload_max_filesize 50M
php_value post_max_size 200M
php_value max_execution_time 600
php_value memory_limit 512M
# (GIS ตั้ง memory_limit 1024M เพราะ pending files ใหญ่)
```

---

## 9. Checklist: เพิ่มกราฟใหม่

- [ ] HTML: `card.chart-card` > `.controls` + `.chart-container` (data-default-height ถ้า ≠ 400px) + `.export-bar` (7 ปุ่มครบ) + note textarea
- [ ] JS: ตัวแปร `xxxChartObj`, ฟังก์ชัน `xxxUpdate/Init`
- [ ] JS: ลงทะเบียน 4 ที่ — `getChartObj`, `getChartRenderFn`, `CTX_CHART_TYPE_MAP`, `CTX_CHART_UPDATE_MAP`
- [ ] JS: ตอนจบ init → store defaults (`data-dsi`, `data-dck`, `data-default-active`, `data-defaults-stored='1'`)
- [ ] Test: ปรับ controls → รีเซ็ต → กลับ default ครบ
- [ ] Test: scroll/pan แกน Y → รีเซ็ต → scale เดิม
- [ ] Test: เปิด "แสดงค่า" → รีเซ็ต → ปิด
- [ ] Y-axis tick: `<100` → 2 decimal, `≥100` → comma
- [ ] Chart type dropdown: line → straight → bar → area

---

## 10. Known Issues (ยังค้าง)

1. **PR กราฟ Tab 1 (`pr1`, `pr2`)** ใช้ `pr1Reset()`, `pr2Reset()` แบบ hardcoded → ควรย้ายไป `resetChartSize(this)` (pattern PR 5 ด้าน)
2. **GIS `pd2ResetChart`, `pd3ResetChart`** แค่ `resetZoom()` → ขาด restore controls + คืนขนาด container
3. **GIS `gisResetChart`** hardcode height='450px' → ควรใช้ `data-default-height`
4. **Leak `resetChartSize`** ไม่ restore controls (ข้ามข้อ 3) → ควรเพิ่มให้ครบ spec
5. **PR `updateLSUsageBar` is not defined** (rebuilt index.html) — JS error ตอน rebuild
6. **Meter `normalize_size()`** ยังใช้ `eval()` — ควรเปลี่ยนเป็น lookup table
7. **AON values ยังเป็น 0** — ปัญหา `parse_aon_sheet_with_col` เดิม (fallback ใช้ได้)
8. **PR + GIS `build_dashboard.php`** ยังไม่มี safety checks ครบ (backup, DOCTYPE, size validation) — Leak + Meter มีแล้ว

---

## 11. Reference Files (Deep Dive เท่านั้น)

อ่านเฉพาะเมื่อต้องลงรายละเอียด — ไม่ใช่ required reading:

| เรื่อง | ไฟล์ | บรรทัด |
|---|---|---|
| `resetChartSize` baseline (best practice) | `Dashboard_PR/index.html` | ~4606 |
| `resetChartSize` (F5 + render map) | `Dashboard_Leak/index.html` | ~8432 |
| `getChartObj`, `getChartRenderFn` | `Dashboard_Leak/index.html` | ~8476, ~8504 |
| `toggleDataLabels`, `openFontPopup` | `Dashboard_Leak/index.html` | ~2132, ~2151 |
| `axisSettings`, `axisEnforcer` | `Dashboard_Leak/index.html` | ~2082 |
| Export bar HTML template | `Dashboard_Leak/index.html` | ~288 |
| PR paired-chart Y-sync | `Dashboard_PR/index.html` | search `PAIRED_CHARTS` |
| Leak RL multi-FY parsing | `Dashboard_Leak/build_dashboard.php` | ~737-950 |
| GIS pending multi-file merge | `Dashboard_GIS/api.php` | `find_pending_files`, `build_pending_sqlite` |

---

## 12. Archive (เก็บแต่ไม่บังคับอ่าน)

`archive/` เก็บเอกสารเก่าและ session logs ทุกเวอร์ชันที่ผ่านมา — อ่านเฉพาะเมื่อ Aong อ้างถึง หรือต้องดู history เฉพาะเรื่อง

- `archive/quick_context_2026-04-04.txt` — รายละเอียด session 1-9
- `archive/PROJECT_README_old.md`
- `archive/NAMING_RULES_old.md`
- `archive/Dashboard_Leak_RULES_old.md`
- `archive/Dashboard_Leak_PROJECT_README_old.md`
- `archive/Dashboard_GIS_NOTES.md`
- `archive/GitHub_Readme_old.md`
- `archive/COMPONENT_SPECS_old.md`
- `archive/session_logs.md` (ของ session ใหม่จะเพิ่มเข้ามา)

---

## 13. Setup เครื่องใหม่ (Checklist)

1. Git for Windows
2. XAMPP (C:\xampp) — เปิด `extension=zip`, `extension=gd` ใน `php.ini`
3. แก้ `httpd.conf`: Find `C:/xampp/htdocs` → Replace `C:/Users/<username>` → Save (ครั้งเดียว)
4. ติดตั้ง Composer: `C:\xampp\php\php.exe composer.phar install`
5. Clone repo: `git clone https://github.com/prayoonsi147-code/NRW_PWA1_Dashboard.git`
6. `git config user.email "prayoonsi147@gmail.com"` + `git config user.name "prayoonsi147-code"`
7. ทดสอบ: ดับเบิลคลิก `push_to_github.bat`

**ถ้าไฟล์ใหญ่เกิน 100MB บล็อก push:**
- `git reset --soft origin/main` → `git rm --cached <path>` → เพิ่มใน `.gitignore` → commit + push
- หนักกว่านั้น: BFG Repo-Cleaner หรือลบ `.git` แล้ว `push --force`

**`.gitignore` ต้องมี:** `*.sqlite`, `*.db`, `*.cache.json`, `*.zip`, `*.rar`, `*.7z`, `*.bak`, `*.checkpoint`, `__pycache__/`

---

## 14. ประวัติการแก้ไขไฟล์นี้

- **2026-04-20 (Session 10):** สร้างครั้งแรก — consolidate จาก 8 ไฟล์ (quick_context, PROJECT_README×2, RULES, NAMING_RULES, COMPONENT_SPECS, GitHub_Readme, NOTES_GIS) → ไฟล์เดียวจบ
