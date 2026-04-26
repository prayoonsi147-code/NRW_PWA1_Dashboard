# Session Logs Archive

> Session logs ทั้งหมดย้ายมาที่นี่ — ไม่ต้องอ่านทุก session ใหม่
>
> Log เก่า session 1-9 (ก่อน 2026-04-04) อยู่ใน `quick_context_2026-04-04.txt`
> Log เก่าของ Dashboard_Leak อยู่ใน `Dashboard_Leak_PROJECT_README_old.md`
> Log เก่าของ GIS อยู่ใน `Dashboard_GIS_NOTES.md`

---

## Session 10 — 2026-04-20 (Aong)

**งานที่ทำ:**
1. Audit ปุ่มรีเซ็ตทั้ง 4 Dashboard — พบ 7 variant ของ behavior (ไม่สม่ำเสมอ)
2. สรุปปัญหาการทำงานข้าม session — Claude ไม่มี memory → เอกสารเก่ายาวเกินไป
3. Consolidate เอกสาร 8 ไฟล์ → `START_HERE.md` ไฟล์เดียว (≤400 บรรทัด)
4. ย้ายเอกสารเก่าเข้า `archive/` + ลบ dead files (__pycache__, _migration_scripts, uploaded_data, data.json/data_embed.js เก่า, requirements.txt, New Text Document.txt, test_path.php, composer-setup.php, start_xampp.bat, .sync_test)

**งานค้าง (ดู START_HERE.md §10 Known Issues):**
- PR pr1/pr2 reset hardcoded → ย้ายไป `resetChartSize(this)`
- GIS pd2/pd3 reset แค่ resetZoom → เพิ่ม restore controls
- GIS hardcode height='450px' → ใช้ `data-default-height`
- Leak `resetChartSize` ไม่ restore controls → เพิ่มให้ครบ spec
- PR updateLSUsageBar undefined error
- Meter normalize_size() ใช้ eval() → lookup table
- AON values = 0 (parser bug เดิม)
- PR + GIS build_dashboard.php safety checks ยังไม่ครบ
