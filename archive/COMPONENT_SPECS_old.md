# COMPONENT_SPECS — สเปคพฤติกรรม UI ของทุก Dashboard

> **วัตถุประสงค์:** เอกสารนี้เป็น "สัญญา" (contract) ระหว่าง Aong กับ Claude ว่า UI component ทุกตัว (ปุ่มรีเซ็ต, export bar, controls, แกน Y ฯลฯ) **ต้องทำงานแบบไหน** ไม่ว่าจะอยู่ Dashboard ไหน
>
> **สำหรับ Claude:** อ่านไฟล์นี้**ทุกครั้งก่อนแก้หรือเพิ่ม component UI** ถ้า code ปัจจุบันไม่ตรงกับ spec → ยึด spec เป็นหลัก ไม่ยึด code เก่า
>
> อัปเดตล่าสุด: 2026-04-20 (Session 10)
> Baseline: Dashboard_Leak (ถือเป็น reference implementation ที่สะอาดสุด)

---

## 0. กฎเหล็ก (Non-negotiables)

1. **ห้ามสร้าง wrapper function แยกสำหรับแต่ละกราฟ** ถ้าของกลางมีอยู่แล้ว
   - มี `resetChartSize(btn)` แล้ว → ห้ามสร้าง `pd2ResetChart()`, `gisResetChart()`, `pr1Reset()` ใหม่
   - มี `openFontPopup(btn)`, `toggleDataLabels(cb)` แล้ว → ใช้ตัวเดิม
2. **ทุกกราฟใหม่ต้องลงทะเบียนใน 3 ที่:**
   - `getChartObj(canvasId)` → คืน Chart.js instance
   - `getChartRenderFn(canvasId)` → คืนฟังก์ชัน render ต้นฉบับ (สำหรับ reset)
   - `CTX_CHART_TYPE_MAP`, `CTX_CHART_UPDATE_MAP` → สำหรับ context menu เปลี่ยนประเภทกราฟ
3. **HTML export-bar ต้องตรง template** (ดูข้อ 2) — ห้ามเปลี่ยนลำดับปุ่ม ห้าม omit ปุ่ม
4. **Controls ทุกตัวต้องมี `data-dsi` (selects) และ `data-dck` (checkboxes)** เพื่อให้ reset รู้ค่า default
5. **Card container ต้องมี attribute `data-defaults-stored="1"`** หลังโหลดเสร็จ (signal ว่า default ถูกเก็บแล้ว)

---

## 1. ลำดับการอ่านเอกสาร (สำหรับ Claude session ใหม่)

1. `quick_context.txt` — สรุปงาน + กฎเหล็กฝังข้อมูล
2. `NAMING_RULES.md` — ถ้าแก้ upload/rename
3. `Dashboard_Leak/RULES.md` — บทเรียน 12 ข้อ
4. **`COMPONENT_SPECS.md` (ไฟล์นี้)** — ก่อนแตะ UI component ใดๆ
5. `PROJECT_README.md` — reference เพิ่มเติม

---

## 2. Export Bar — โครงสร้าง HTML ตายตัว

ทุกกราฟต้องมี export-bar แบบนี้ (ลำดับ 7 ปุ่ม ห้ามสลับ):

```html
<div class="export-bar">
  <!-- 1. รีเซ็ต -->
  <button class="export-btn" onclick="resetChartSize(this)" title="รีเซ็ตขนาดกราฟ">&#8634; รีเซ็ต</button>

  <!-- 2. แสดงค่า (checkbox in label) -->
  <label class="export-btn" style="cursor:pointer">
    <input type="checkbox" onchange="toggleDataLabels(this)" style="margin-right:4px">แสดงค่า
  </label>

  <!-- 3. Font popup (icon A) -->
  <button class="export-btn" onclick="openFontPopup(this)">
    <svg><!-- font A icon --></svg>
  </button>

  <!-- 4. PNG -->
  <button class="export-btn" onclick="exportPNG('<canvasId>','<title>')">
    <svg><!-- image icon --></svg>
  </button>

  <!-- 5. Copy -->
  <button class="export-btn" onclick="copyToClipboard('<canvasId>',this)">
    <svg><!-- copy icon --></svg>
  </button>

  <!-- 6. Excel -->
  <button class="export-btn" onclick="exportExcel('<chartKey>')">
    <svg><!-- green X icon --></svg>
  </button>

  <!-- 7. PowerPoint -->
  <button class="export-btn" onclick="exportPPTX('<canvasId>','<title>')">
    <svg><!-- orange P icon --></svg>
  </button>
</div>
```

**กฎ:**
- CSS: `.export-bar{display:flex;gap:6px;margin-top:12px;justify-content:flex-end;flex-wrap:wrap}`
- ปุ่มต้องอยู่**หลัง** `.chart-container` (หรือ table wrap) เพราะ `resetChartSize` ใช้ `previousElementSibling` หา container
- title attribute ต้องมีทุกปุ่ม (tooltip ภาษาไทย)
- SVG icon ขนาด 18×18 (font/PNG/copy) หรือ 20×20 (Excel/PPTX)
- **สำหรับตาราง:** ไม่มีปุ่ม "แสดงค่า" + ไม่มี Font popup (เฉพาะบางตาราง) — tooltip รีเซ็ตเป็น "รีเซ็ตขนาดตาราง"

---

## 3. ปุ่มรีเซ็ต (↺ รีเซ็ต) — สเปคสำคัญ

### 3.1 พฤติกรรมที่ถูกต้อง (baseline: Dashboard_Leak)

กดปุ่มรีเซ็ต = **"F5 เฉพาะกราฟนั้น"** ต้องทำ **4 ขั้นตอนเสมอ** ในลำดับนี้:

**Step 1: คืนขนาด container**
```js
el.style.width = '';
el.style.height = el.getAttribute('data-default-height') || '400px';
```
- ถ้ากราฟมีขนาดพิเศษ → set attribute `data-default-height="450px"` บน `.chart-container`
- **ห้าม hardcode height** ในฟังก์ชัน reset (Leak ปัจจุบัน hardcode '400px' — ถือเป็น known deviation ที่ควรปรับ)

**Step 2: ล้าง Axis zoom/pan**
```js
clearAxisSettings(canvas.id);           // ลบจาก axisSettings{}
const c = getChartObj(canvas.id);
if(c){
  ['y','yLeft','yRight'].forEach(k=>{
    if(c.options.scales[k]){
      delete c.options.scales[k].min;
      delete c.options.scales[k].max;
      if(c.options.scales[k].ticks) delete c.options.scales[k].ticks.stepSize;
    }
  });
}
```

**Step 3: Restore controls กลับค่า default**
```js
const card = bar.closest('.chart-card') || el.closest('.card');
if(card && card.getAttribute('data-defaults-stored')){
  // Selects
  card.querySelectorAll('.controls select').forEach(sel=>{
    const di = sel.getAttribute('data-dsi');
    if(di !== null){
      const idx = parseInt(di);
      if(idx >= 0 && idx < sel.options.length) sel.selectedIndex = idx;
    }
  });
  // Checkboxes
  card.querySelectorAll('.controls input[type=checkbox]').forEach(cb=>{
    const dc = cb.getAttribute('data-dck');
    if(dc !== null) cb.checked = (dc === '1');
  });
  // Fire change events → sync JS state
  card.querySelectorAll('.controls select').forEach(sel=> sel.dispatchEvent(new Event('change')));
  card.querySelectorAll('.controls input[type=checkbox]').forEach(cb=> cb.dispatchEvent(new Event('change')));
  // Button groups (.ymt) — คลิกปุ่ม default กลับ
  card.querySelectorAll('.ymt').forEach(grp=>{
    const defBtn = grp.querySelector('[data-default-active]');
    if(defBtn && !defBtn.classList.contains('active')) defBtn.click();
  });
}
```

**Step 4: Reset "แสดงค่า" + re-render + resize**
```js
// Reset "แสดงค่า" checkbox ใน export-bar
const dlCb = bar.querySelector('input[type=checkbox]');
if(dlCb){
  const dck = dlCb.getAttribute('data-dck');
  const defState = (dck !== null) ? (dck === '1') : dlCb.defaultChecked;
  if(dlCb.checked !== defState) dlCb.checked = defState;
  toggleDataLabels(dlCb);
}
// Re-render (destroy + new Chart) ผ่าน getChartRenderFn
const renderFn = getChartRenderFn(canvas.id);
if(renderFn) renderFn();
else { c.update(); setTimeout(()=> c.resize(), 50); }  // fallback
```

### 3.2 สิ่งที่ต้องเก็บไว้ตอนโหลด (เพื่อให้ reset ใช้งานได้)

เมื่อ render กราฟครั้งแรก (ใน init function ของแต่ละกราฟ) ให้เก็บ default:
```js
// ที่ท้ายของ initXxxChart()
const card = document.getElementById('xxxCard');
if(card && !card.getAttribute('data-defaults-stored')){
  card.querySelectorAll('.controls select').forEach(sel=>{
    sel.setAttribute('data-dsi', sel.selectedIndex);
  });
  card.querySelectorAll('.controls input[type=checkbox]').forEach(cb=>{
    cb.setAttribute('data-dck', cb.checked ? '1' : '0');
  });
  card.querySelectorAll('.ymt .ymb.active').forEach(btn=>{
    btn.setAttribute('data-default-active', '1');
  });
  // "แสดงค่า" ใน export-bar
  const dlCb = card.querySelector('.export-bar input[type=checkbox]');
  if(dlCb) dlCb.setAttribute('data-dck', dlCb.checked ? '1' : '0');
  card.setAttribute('data-defaults-stored', '1');
}
```

### 3.3 Violations ปัจจุบัน (ต้องแก้ทีหลัง — ยังไม่ต้องรีบ)

| Dashboard | กราฟ | ปัญหา |
|---|---|---|
| Leak | ทุกกราฟ | ✅ ตรง spec 90% — ไม่ restore controls (ข้อ 3) |
| PR | pr1, pr2 | ❌ ใช้ฟังก์ชันแยก hardcoded — ควรย้ายไป `resetChartSize(this)` |
| PR | 5 ด้าน × 2 | ✅ ตรง spec — ต้นแบบที่ดีที่สุด |
| GIS | pd2, pd3 | ❌ แค่ `resetZoom()` — ขาด step 1, 3, 4 |
| GIS | gis, sum, pressure, press2, pd1 | ⚠️ hardcode height='450px' + ไม่ `clearAxisSettings` |
| Meter | dm1 | ⚠️ แค่ลบ width/height — ok สำหรับตาราง แต่ควรมี Step 3 ถ้ามี controls |

---

## 4. Data Labels (แสดงค่า) checkbox

**HTML:**
```html
<label class="export-btn" style="cursor:pointer">
  <input type="checkbox" onchange="toggleDataLabels(this)" style="margin-right:4px">แสดงค่า
</label>
```

**ฟังก์ชัน (ของกลาง):**
```js
function toggleDataLabels(cb){
  const bar = cb.closest('.export-bar');
  let el = bar.previousElementSibling;
  while(el && !el.classList.contains('chart-container')) el = el.previousElementSibling;
  if(!el) return;
  const canvas = el.querySelector('canvas');
  if(!canvas) return;
  const chart = getChartObj(canvas.id);
  if(!chart) return;
  chart.options.plugins.datalabels = dlConfig(cb.checked);
  chart.update();
}
```

**ต้องมี `dlConfig(show)` helper** คืน config ของ chartjs-plugin-datalabels — ดูตัวอย่าง Dashboard_Leak

---

## 5. Font Popup — `openFontPopup(btn)`

เปิด popup ปรับขนาด font ของ 4 ส่วน:
- **X axis ticks** (default 10)
- **Y axis ticks** (default 12)
- **Data labels** (default 10)
- **Legend labels** (default 11)
- **X axis rotation** (0-90°)

Popup ต้อง:
- อ่านค่าปัจจุบันจาก chart options (ไม่ใช่ค่า default คงที่)
- มีปุ่ม "ตกลง", "ยกเลิก", "รีเซ็ตค่า" (กลับเป็น default ที่บันทึกตอน init)
- Toggle: คลิกปุ่ม font ซ้ำอีกครั้ง → ปิด popup
- ปิด popup เมื่อคลิกข้างนอก

---

## 6. Controls Layout (ภายใน card)

**ลำดับ controls (ซ้าย→ขวา):**
1. ประเภทกราฟ (`<select>` chart type)
2. แกน X (ช่วงเวลา / สาขา) ถ้ามี
3. เกณฑ์/ตัวแปร (เลือก metric)
4. ตัวเลือกอื่นๆ (ปี, เดือน, checkboxes)

**Branch Selector (3 โหมด) — สำหรับกราฟที่มีสาขา:**
- โหมด 1: สาขาเดียว (dropdown)
- โหมด 2: บางสาขา (checkboxes สีตาม accent color ของ Dashboard) — **default**
- โหมด 3: ทุกสาขา (แสดงทั้งหมด)
- **แสดงเฉพาะเมื่อ** `xAxis === "ช่วงเวลา"` — ซ่อนเมื่อ `xAxis === "สาขา"`

**HTML class ที่ใช้:**
- Card wrapper: `<div class="card chart-card">` (chart-card = มี reset-able content)
- Controls row: `<div class="controls">` (สำคัญ — resetChartSize หาผ่าน `.controls`)
- Chart container: `<div class="chart-container" style="height:320px">` (หรือ 400/450)
- Display toggle: `<div class="ymt">` พร้อม `<button class="ymb">` แต่ละปุ่ม

---

## 7. Axis Y Interaction — ทุกกราฟ ทุก Dashboard

**Scroll เม้าส์บนแกน Y:**
- ถ้า Origin เดิม = 0 → ยึด Origin ไว้ที่ 0 (ปรับเฉพาะ Max)
- ถ้า Origin เดิม ≠ 0 → Zoom รอบจุดกึ่งกลาง (ปรับทั้ง Min และ Max)

**ลากเม้าส์ Pan แกน Y:**
- ถ้า Origin = 0 → ยึด Min=0, ปรับ Max ตาม
- ถ้า Origin ≠ 0 → Pan อิสระ ทั้ง Min และ Max

**Double-click บนแกน Y:** เปิด dialog ตั้งค่า Min / Max / Interval (stepSize)

**Right-click บนแกน Y:** เปิด context menu ปรับ:
- Min / Max / Interval
- สีแกน (tick labels, title)
- สีเส้น Grid, รูปแบบ (Solid/Dashed/Dotted), ซ่อน/แสดง

**โค้ดอยู่ใน (baseline Leak):**
- `axisSettings{}` — เก็บ min/max/stepSize ต่อ canvasId
- `axisEnforcer` plugin — apply ทุกครั้งที่ `chart.update()`
- `handleWheel()`, `handleMouseMove()` — scroll + pan
- `showAxisDialog()`, `showAxisCtxMenu()`, `ctxAxisApply()`, `ctxAxisReset()` — dialog + menu

---

## 8. ตัวเลือกประเภทกราฟ (Chart Type Dropdown)

**ลำดับ option ที่ถูกต้อง (ห้ามสลับ):**
```html
<select onchange="xxxSetType(this.value)">
  <option value="line" selected>กราฟเส้นโค้ง</option>
  <option value="straight">กราฟเส้นตรง</option>
  <option value="bar">กราฟแท่ง</option>
  <option value="area">กราฟพื้นที่</option>
</select>
```

**กฎ:**
- เส้นโค้ง = Chart.js type `line` + `tension: 0.3`
- เส้นตรง = Chart.js type `line` + `tension: 0`
- แท่ง = Chart.js type `bar`
- พื้นที่ = Chart.js type `line` + `fill: true` + `tension: 0.3`
- Context menu ตรวจ type ด้วย `tension === 0` เพื่อแยก straight ออกจาก line
- **ห้าม reorder options** เพื่อเปลี่ยน default → ใช้ `selected` attribute แทน

---

## 9. การจัดรูปแบบตัวเลข (ทุกกราฟ)

- **Y-axis tick callback:**
  - ค่า `< 100` → 2 ทศนิยม (เช่น `5.00`, `12.50`)
  - ค่า `>= 100` → comma (เช่น `1,234`)
  - Locale: `th-TH`
- **X-axis ticks:** `maxRotation: 60`, `font.size: 12`
- **beginAtZero:** `false` (ให้ Chart.js คำนวณเอง)

---

## 10. หมายเหตุ (Note Textarea)

**HTML pattern:**
```html
<div class="card-note-box">
  <label>📝 หมายเหตุ</label>
  <textarea id="note-<chartKey>" placeholder="เพิ่มหมายเหตุ..."
    oninput="saveCardNote('<chartKey>')"
    style="background:<pastelColor>;border-color:<borderColor>"></textarea>
</div>
```

**กฎ:**
- ทุกกราฟต้องมี textarea หมายเหตุ
- บันทึกผ่าน API (PR/Leak/GIS/Meter เก็บใน `data.json` หรือ `notes.json`)
- Auto-save debounce 800ms ผ่าน `saveCardNote()`
- ห้ามเก็บใน localStorage อย่างเดียว (สูญหายเมื่อเปลี่ยนเครื่อง)
- Placeholder เป็นภาษาไทยเสมอ

---

## 11. Card Header

```html
<div class="card-header-row">
  <h3><span>📊</span> <span id="xxxChartTitle">ชื่อกราฟ</span></h3>
  <button class="card-collapse-btn" onclick="xxxToggleBody(this)" title="ย่อ/ขยาย">&#9650;</button>
</div>
<div class="xxx-card-body">
  <!-- controls, chart, export-bar, note -->
</div>
```

**กฎ:**
- Title ต้องเป็น `<span>` ที่มี id → รองรับการเปลี่ยนชื่อแบบ dynamic
- ปุ่มย่อ/ขยาย ใช้ `&#9650;` (▲) default, toggle เป็น `&#9660;` (▼)
- Body มี class suffix `-card-body` (pr-card-body, leak-card-body ฯลฯ) สำหรับ collapse

---

## 12. Chart Registration — ต้องเพิ่มใน 3 ที่ (Dashboard_Leak)

เมื่อเพิ่มกราฟใหม่ใน Leak (หรือ Dashboard อื่นที่มี reset ผ่าน `getChartRenderFn`):

```js
// 1. ตัวแปรเก็บ Chart.js instance
var myNewChartObj = null;

// 2. getChartObj mapping
function getChartObj(canvasId){
  const map = {
    'myNewChart': myNewChartObj,
    // ...
  };
  return map[canvasId] || null;
}

// 3. getChartRenderFn mapping (สำหรับ reset)
function getChartRenderFn(canvasId){
  const map = {
    'myNewChart': function(){ myNewUpdate(); },
    // ...
  };
  return map[canvasId] || null;
}

// 4. CTX_CHART_TYPE_MAP + CTX_CHART_UPDATE_MAP (สำหรับ context menu)
CTX_CHART_TYPE_MAP['myNewChart'] = function(type){ /* change type logic */ };
CTX_CHART_UPDATE_MAP['myNewChart'] = function(){ myNewUpdate(); };
```

---

## 13. Checklist: เพิ่มกราฟใหม่

- [ ] HTML: card-header, controls, chart-container (`data-default-height` ถ้าไม่ใช่ 400px)
- [ ] HTML: export-bar ครบ 7 ปุ่ม (ดูข้อ 2)
- [ ] HTML: note textarea (ดูข้อ 10)
- [ ] JS: ตัวแปร `xxxChartObj`, ฟังก์ชัน `xxxUpdate()`, `xxxInit()`
- [ ] JS: ลงทะเบียนใน `getChartObj`, `getChartRenderFn`, `CTX_CHART_TYPE_MAP`, `CTX_CHART_UPDATE_MAP`
- [ ] JS: ตอนจบ init → store defaults (`data-dsi`, `data-dck`, `data-defaults-stored`)
- [ ] ทดสอบ: ปรับ controls → กดรีเซ็ต → ต้องกลับเป็น default ครบทุกอย่าง
- [ ] ทดสอบ: scroll/pan แกน Y → กดรีเซ็ต → กลับเป็น scale เดิม
- [ ] ทดสอบ: เปลี่ยนขนาด container (ลากมุม) → กดรีเซ็ต → กลับเป็นขนาดเดิม
- [ ] ทดสอบ: เปิด/ปิด "แสดงค่า" → กดรีเซ็ต → กลับเป็นสถานะเริ่มต้น
- [ ] Y-axis tick callback: ค่า < 100 → 2 ทศนิยม, ≥ 100 → comma
- [ ] Chart type dropdown: ลำดับ line → straight → bar → area

---

## 14. Checklist: refactor กราฟเก่าให้ตรง spec

ใช้เมื่อแก้กราฟที่ violate spec (ดูตาราง §3.3):

1. ลบ wrapper function เฉพาะตัว (เช่น `pd2ResetChart`, `pr1Reset`) — แทนด้วย `resetChartSize(this)` ใน onclick
2. ย้ายค่า default → `data-dsi`, `data-dck` attributes บน HTML element (ไม่ hardcode ใน JS)
3. ย้าย card ให้มี class `chart-card` + set `data-defaults-stored="1"` หลัง init
4. เพิ่ม `data-default-height` ถ้า container ไม่ใช่ 400px
5. ลงทะเบียนใน `getChartRenderFn` (ถ้ายังไม่มี)
6. ทดสอบตาม checklist §13

---

## 15. Known Good Reference Files

| Component | ไฟล์ | บรรทัด |
|---|---|---|
| `resetChartSize(btn)` baseline | `Dashboard_PR/index.html` | ~4606 |
| `resetChartSize(btn)` (F5 version) | `Dashboard_Leak/index.html` | ~8432 |
| `getChartRenderFn`, `getChartObj` | `Dashboard_Leak/index.html` | ~8476, ~8504 |
| `toggleDataLabels(cb)` | `Dashboard_Leak/index.html` | ~2132 |
| `openFontPopup(btn)` | `Dashboard_Leak/index.html` | ~2151 |
| `clearAxisSettings`, `axisEnforcer` | `Dashboard_Leak/index.html` | ~2082 |
| export-bar HTML template | `Dashboard_Leak/index.html` | ~288 |
| PR card-based reset (ต้นแบบที่ดีสุด) | `Dashboard_PR/index.html` | ~4606 |

> **หมายเหตุ:** "ต้นแบบที่ดีสุด" สำหรับ reset คือ `Dashboard_PR/index.html` เวอร์ชัน `resetChartSize(this)` (4 ขั้นตอนครบ) ส่วน Leak เวอร์ชันสะอาดกว่าแต่ขาด Step 3 (restore controls)
>
> **ทางออกระยะยาว:** รวม 2 เวอร์ชันเป็นตัวเดียว (Leak's render approach + PR's controls restore) ใส่ไว้ใน shared JS file → include ทุก Dashboard

---

## 16. ประวัติการแก้ไข spec นี้

- **2026-04-20 (Session 10):** ฉบับแรก — ร่างจาก audit ปุ่มรีเซ็ต 4 Dashboard (Aong)
