# -*- coding: utf-8 -*-
"""
Dashboard GIS — Local Development Server
=========================================
Flask server สำหรับ Dashboard แผนที่แนวท่อ
รับ upload ไฟล์ Excel → จัดเก็บลงโฟลเดอร์ ข้อมูลดิบ/ → ไม่ parse (ให้ build_dashboard.py จัดการ)

Usage:
    python server.py
    หรือ ดับเบิลคลิก start_server.bat

API Endpoints:
    GET  /                     → serve index.html หรือ manage.html
    GET  /api/ping             → health check
    GET  /api/data             → inventory ของไฟล์ข้อมูลดิบในแต่ละ category
    POST /api/upload/<category> → upload ไฟล์ไปยัง category folder
    DELETE /api/data/<category>/<filename> → ลบไฟล์เฉพาะ
    POST /api/open-folder      → เปิดโฟลเดอร์ ข้อมูลดิบ/ ใน File Explorer
    POST /api/open-main        → เปิด ../index.html (หน้า Landing)
    POST /api/rebuild          → รัน build_dashboard.py
"""

from flask import Flask, request, jsonify, send_from_directory
import os, sys, json, shutil, traceback, subprocess
from datetime import datetime
import platform

# ─── Configuration ────────────────────────────────────────────────────────────

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RAW_DATA_DIR = os.path.join(BASE_DIR, 'ข้อมูลดิบ')
PORT = 5002

# Category mapping: URL slug → Thai folder name
CATEGORY_MAP = {
    'repair': 'ลงข้อมูลซ่อมท่อ',
    'pressure': 'แรงดันน้ำ',
    'pending': 'ซ่อมท่อค้างระบบ',
}

# Create directories
os.makedirs(RAW_DATA_DIR, exist_ok=True)
for folder in CATEGORY_MAP.values():
    os.makedirs(os.path.join(RAW_DATA_DIR, folder), exist_ok=True)

# ─── Flask App ────────────────────────────────────────────────────────────────

app = Flask(__name__, static_folder=BASE_DIR)

# ─── CORS Middleware ──────────────────────────────────────────────────────────

@app.after_request
def add_cors_headers(response):
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Methods'] = 'GET, POST, DELETE, OPTIONS'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type'
    return response

# ─── Serve Static Files ───────────────────────────────────────────────────────

@app.route('/')
def serve_index():
    return send_from_directory(BASE_DIR, 'index.html')

@app.route('/manage.html')
def serve_manage():
    return send_from_directory(BASE_DIR, 'manage.html')

# ─── API Endpoints ────────────────────────────────────────────────────────────

@app.route('/api/ping')
def api_ping():
    """ตรวจสอบว่า server ทำงานอยู่"""
    return jsonify({'ok': True, 'version': '1.0', 'timestamp': datetime.now().isoformat()})

@app.route('/api/data')
def api_get_data():
    """คืนรายชื่อไฟล์ในแต่ละ category folder"""
    inventory = {}

    for slug, thai_name in CATEGORY_MAP.items():
        folder_path = os.path.join(RAW_DATA_DIR, thai_name)
        files = []
        last_modified = None

        if os.path.exists(folder_path):
            try:
                for f in os.listdir(folder_path):
                    fp = os.path.join(folder_path, f)
                    if os.path.isfile(fp) and not f.startswith('.'):
                        files.append(f)
                        mtime = os.path.getmtime(fp)
                        if last_modified is None or mtime > last_modified:
                            last_modified = mtime
                files.sort()
            except Exception as e:
                pass

        inventory[slug] = {
            'thai_name': thai_name,
            'files': files,
            'count': len(files),
            'last_modified': datetime.fromtimestamp(last_modified).strftime('%d/%m/%Y %H:%M') if last_modified else None
        }

    # Load saved notes
    notes_file = os.path.join(RAW_DATA_DIR, 'notes.json')
    notes = {}
    if os.path.exists(notes_file):
        try:
            with open(notes_file, 'r', encoding='utf-8') as f:
                notes = json.load(f)
        except (json.JSONDecodeError, IOError):
            pass

    return jsonify({'ok': True, 'inventory': inventory, 'notes': notes})

@app.route('/api/notes/<slug>', methods=['POST'])
def api_save_note(slug):
    """บันทึก note ของแต่ละ category"""
    if slug not in CATEGORY_MAP:
        return jsonify({'ok': False, 'error': 'invalid slug'}), 400
    body = request.get_json(force=True)
    text = body.get('text', '')
    notes_file = os.path.join(RAW_DATA_DIR, 'notes.json')
    notes = {}
    if os.path.exists(notes_file):
        try:
            with open(notes_file, 'r', encoding='utf-8') as f:
                notes = json.load(f)
        except (json.JSONDecodeError, IOError):
            pass
    notes[slug] = text
    with open(notes_file, 'w', encoding='utf-8') as f:
        json.dump(notes, f, ensure_ascii=False, indent=2)
    return jsonify({'ok': True})

@app.route('/api/upload/<category>', methods=['POST', 'OPTIONS'])
def api_upload(category):
    """รับ upload ไฟล์ไปยัง category folder"""
    if request.method == 'OPTIONS':
        return '', 204

    if category not in CATEGORY_MAP:
        return jsonify({'ok': False, 'error': f'ไม่รู้จัก category: {category}'}), 400

    files = request.files.getlist('files')
    if not files:
        return jsonify({'ok': False, 'error': 'ไม่ได้เลือกไฟล์'}), 400

    folder_path = os.path.join(RAW_DATA_DIR, CATEGORY_MAP[category])
    os.makedirs(folder_path, exist_ok=True)

    PREFIX_MAP = { 'repair': 'GIS', 'pressure': 'PRESSURE', 'pending': 'PENDING' }
    import re

    results = []
    errors = []
    _pending_batch = []  # สำหรับ pending: เก็บไฟล์ทั้งหมดไว้รวมทีหลัง

    for f in files:
        filename = f.filename.strip()
        if not filename:
            continue

        try:
            prefix = PREFIX_MAP.get(category, category.upper())
            name_only = os.path.splitext(filename)[0]
            ext = os.path.splitext(filename)[1] or '.xlsx'

            # ── Category-specific auto-rename ──
            if category == 'repair':
                # repair: GIS_YYMMDD.xlsx — ดึงตัวเลข 6 หลัก (YYMMDD)
                m = re.search(r'(\d{6})', name_only)
                if m:
                    new_name = f"{prefix}_{m.group(1)}{ext}"
                else:
                    # fallback: ใช้วันที่ upload
                    today = datetime.now().strftime('%y%m%d')
                    new_name = f"{prefix}_{today}{ext}"

            elif category == 'pressure':
                # pressure: PRESSURE_สาขา_ปีงบYYYY.xlsx
                # ดึงชื่อสาขาจากชื่อไฟล์ + ปีงบประมาณจาก Row 3 ในไฟล์
                thai_parts = re.findall(r'[\u0e00-\u0e7f]+', name_only)
                branch_name = thai_parts[-1] if thai_parts else 'unknown'

                # อ่านปีงบประมาณจากเนื้อไฟล์ (Row 3: "ประจำปีงบประมาณ 25XX")
                fiscal_year = ''
                try:
                    import io
                    raw_bytes = f.read()
                    f.seek(0)
                    if ext.lower() in ('.xlsx',):
                        wb_tmp = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True)
                        ws_tmp = wb_tmp.active
                        for r in range(1, min(6, ws_tmp.max_row + 1)):
                            for c in range(1, min(6, ws_tmp.max_column + 1)):
                                cell_val = str(ws_tmp.cell(r, c).value or '')
                                fy_match = re.search(r'ปีงบประมาณ\s*(\d{4})', cell_val)
                                if fy_match:
                                    fiscal_year = fy_match.group(1)
                                    break
                            if fiscal_year:
                                break
                        wb_tmp.close()
                    elif ext.lower() == '.xls':
                        if xlrd:
                            wb_tmp = xlrd.open_workbook(file_contents=raw_bytes)
                            ws_tmp = wb_tmp.sheet_by_index(0)
                            for r in range(min(5, ws_tmp.nrows)):
                                for c in range(min(5, ws_tmp.ncols)):
                                    cell_val = str(ws_tmp.cell_value(r, c) or '')
                                    fy_match = re.search(r'ปีงบประมาณ\s*(\d{4})', cell_val)
                                    if fy_match:
                                        fiscal_year = fy_match.group(1)
                                        break
                                if fiscal_year:
                                    break
                except Exception:
                    pass

                fy_suffix = f'_ปีงบ{fiscal_year[-2:]}' if fiscal_year else ''
                new_name = f"{prefix}_{branch_name}{fy_suffix}{ext}"

            elif category == 'pending':
                # pending: จะถูกจัดการแบบ batch ด้านล่าง (รวมหลายไฟล์เป็น 1)
                raw_bytes = f.read()
                _pending_batch.append({'filename': filename, 'ext': ext, 'raw_bytes': raw_bytes})
                continue  # ไม่ save ทีละไฟล์ — รอรวมด้านล่าง

            else:
                # fallback ทั่วไป
                m = re.search(r'(\d{6})', name_only)
                if m:
                    new_name = f"{prefix}_{m.group(1)}{ext}"
                else:
                    clean = re.sub(r'[^\w\-.]', '_', name_only).strip('_')[:30]
                    new_name = f"{prefix}_{clean}{ext}"

            dest_path = os.path.join(folder_path, new_name)

            overwritten = os.path.exists(dest_path)
            f.save(dest_path)

            msg = f'{filename} → {new_name}' + (' (เขียนทับ)' if overwritten else '')
            results.append({
                'filename': new_name, 'original': filename,
                'status': 'success', 'message': msg
            })
        except Exception as e:
            errors.append({
                'filename': filename,
                'error': str(e)
            })

    # ── Pending batch merge: รวมหลายไฟล์เป็น 1 ไฟล์ ──
    if category == 'pending' and _pending_batch:
        import io
        try:
            HEADER_ROWS = 8  # Row 0-7 = header area, Row 8+ = data
            DATE_COL = 3     # Col 3 = วันที่แจ้ง (dd/mm/yyyy พ.ศ.)

            all_data_rows = []
            header_rows_cache = None

            for item in _pending_batch:
                raw_bytes = item['raw_bytes']
                ext_lower = item['ext'].lower()

                if ext_lower == '.xls':
                    try:
                        import xlrd
                        wb = xlrd.open_workbook(file_contents=raw_bytes)
                        ws = wb.sheet_by_index(0)
                        # เก็บ header จากไฟล์แรก
                        if header_rows_cache is None:
                            header_rows_cache = []
                            for r in range(min(HEADER_ROWS, ws.nrows)):
                                header_rows_cache.append([ws.cell_value(r, c) for c in range(ws.ncols)])
                        # เก็บ data rows
                        for r in range(HEADER_ROWS, ws.nrows):
                            row_data = [ws.cell_value(r, c) for c in range(ws.ncols)]
                            # ข้ามแถวว่าง (ไม่มีเลขแจ้ง)
                            if not str(row_data[2]).strip():
                                continue
                            all_data_rows.append(row_data)
                    except ImportError:
                        errors.append({'filename': item['filename'], 'error': 'xlrd not installed'})

                elif ext_lower == '.xlsx':
                    try:
                        import openpyxl
                        wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), read_only=True, data_only=True)
                        ws = wb.active
                        rows_list = list(ws.iter_rows(values_only=True))
                        if header_rows_cache is None:
                            header_rows_cache = [list(r) for r in rows_list[:HEADER_ROWS]]
                        for row in rows_list[HEADER_ROWS:]:
                            row_data = list(row)
                            if not str(row_data[2] if len(row_data) > 2 else '').strip():
                                continue
                            all_data_rows.append(row_data)
                        wb.close()
                    except ImportError:
                        errors.append({'filename': item['filename'], 'error': 'openpyxl not installed'})

            if not all_data_rows:
                errors.append({'filename': 'pending', 'error': 'ไม่พบข้อมูลในไฟล์ที่อัปโหลด'})
            else:
                # ── Parse วันที่แจ้ง เพื่อ sort ──
                def parse_date_for_sort(row):
                    """แปลง dd/mm/yyyy (พ.ศ.) → sortable tuple (yyyy_ce, mm, dd)"""
                    try:
                        val = str(row[DATE_COL]).strip()
                        if '/' in val:
                            parts = val.split('/')
                            dd, mm, yyyy = int(parts[0]), int(parts[1]), int(parts[2])
                            yyyy_ce = yyyy - 543 if yyyy > 2500 else yyyy
                            return (yyyy_ce, mm, dd)
                    except:
                        pass
                    return (9999, 12, 31)  # ข้อมูลที่ parse ไม่ได้ → ไว้ท้ายสุด

                all_data_rows.sort(key=parse_date_for_sort)

                # ── หาช่วงเดือน (min/max) เพื่อตั้งชื่อ ──
                min_date = (9999, 12)
                max_date = (0, 0)
                for row in all_data_rows:
                    try:
                        val = str(row[DATE_COL]).strip()
                        if '/' in val:
                            parts = val.split('/')
                            mm, yyyy = int(parts[1]), int(parts[2])
                            yy = yyyy - 2500 if yyyy > 2500 else yyyy
                            if (yy, mm) < min_date:
                                min_date = (yy, mm)
                            if (yy, mm) > max_date:
                                max_date = (yy, mm)
                    except:
                        pass

                # ตั้งชื่อ: ค้างซ่อม_MM-YY_to_MM-YY.xlsx
                if min_date[0] < 9999 and max_date[0] > 0:
                    new_name = f"ค้างซ่อม_{min_date[1]:02d}-{min_date[0]:02d}_to_{max_date[1]:02d}-{max_date[0]:02d}.xlsx"
                else:
                    today = datetime.now()
                    yy = today.year - 2543
                    new_name = f"ค้างซ่อม_{today.month:02d}-{yy:02d}.xlsx"

                # ── เขียนไฟล์ .xlsx ──
                import openpyxl
                _ILLEGAL_XML_RE = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]')
                wb_out = openpyxl.Workbook()
                ws_out = wb_out.active
                ws_out.title = 'ค้างซ่อม'

                def safe_val(v):
                    """ลบ illegal XML characters ที่ .xlsx ไม่รองรับ"""
                    if isinstance(v, str):
                        return _ILLEGAL_XML_RE.sub('', v)
                    return v

                # เขียน header
                if header_rows_cache:
                    for r_idx, row in enumerate(header_rows_cache):
                        for c_idx, val in enumerate(row):
                            ws_out.cell(row=r_idx + 1, column=c_idx + 1, value=safe_val(val))

                # เขียน data (เรียงตามวันที่แจ้งแล้ว)
                start_row = HEADER_ROWS + 1
                for r_idx, row in enumerate(all_data_rows):
                    for c_idx, val in enumerate(row):
                        ws_out.cell(row=start_row + r_idx, column=c_idx + 1, value=safe_val(val))

                dest_path = os.path.join(folder_path, new_name)
                overwritten = os.path.exists(dest_path)
                wb_out.save(dest_path)

                orig_names = ', '.join(item['filename'] for item in _pending_batch)
                if overwritten:
                    results.append({
                        'filename': new_name, 'original': orig_names,
                        'status': 'overwrite', 'message': f'เขียนทับ {new_name}'
                    })
                results.append({
                    'filename': new_name, 'original': orig_names,
                    'status': 'success',
                    'message': f'รวม {len(_pending_batch)} ไฟล์ ({len(all_data_rows):,} แถว) → {new_name}'
                })

        except Exception as e:
            errors.append({'filename': 'pending-merge', 'error': str(e)})

    return jsonify({
        'ok': True,
        'category': category,
        'thai_name': CATEGORY_MAP[category],
        'results': results,
        'errors': errors
    })

@app.route('/api/data/<category>/<filename>', methods=['DELETE', 'OPTIONS'])
def api_delete_file(category, filename):
    """ลบไฟล์เฉพาะ"""
    if request.method == 'OPTIONS':
        return '', 204

    if category not in CATEGORY_MAP:
        return jsonify({'ok': False, 'error': f'ไม่รู้จัก category: {category}'}), 400

    folder_path = os.path.join(RAW_DATA_DIR, CATEGORY_MAP[category])
    file_path = os.path.join(folder_path, filename)

    # Safety check: ensure file is within folder
    if not os.path.abspath(file_path).startswith(os.path.abspath(folder_path)):
        return jsonify({'ok': False, 'error': 'ไม่อนุญาต'}), 403

    try:
        if os.path.exists(file_path):
            os.remove(file_path)
            return jsonify({'ok': True, 'filename': filename, 'deleted': True})
        else:
            return jsonify({'ok': False, 'error': 'ไม่พบไฟล์'}), 404
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500

@app.route('/api/open-folder', methods=['POST', 'OPTIONS'])
def api_open_folder():
    """เปิดโฟลเดอร์ ข้อมูลดิบ/ ใน File Explorer"""
    if request.method == 'OPTIONS':
        return '', 204

    try:
        if platform.system() == 'Windows':
            os.startfile(RAW_DATA_DIR)
        elif platform.system() == 'Darwin':
            subprocess.Popen(['open', RAW_DATA_DIR])
        else:
            subprocess.Popen(['xdg-open', RAW_DATA_DIR])

        return jsonify({'ok': True, 'path': RAW_DATA_DIR})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500

@app.route('/api/open-main', methods=['POST', 'OPTIONS'])
def api_open_main():
    """เปิดหน้าหลัก (Landing Page) ใน browser"""
    if request.method == 'OPTIONS':
        return '', 204

    try:
        parent_dir = os.path.dirname(BASE_DIR)
        main_file = os.path.join(parent_dir, 'index.html')

        if platform.system() == 'Windows':
            os.startfile(main_file)
        else:
            subprocess.Popen(['xdg-open', main_file])

        return jsonify({'ok': True})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500

@app.route('/api/rebuild', methods=['POST', 'OPTIONS'])
def api_rebuild():
    """รัน build_dashboard.py เพื่อสร้าง dashboard ใหม่"""
    if request.method == 'OPTIONS':
        return '', 204

    try:
        build_script = os.path.join(BASE_DIR, 'build_dashboard.py')

        if not os.path.exists(build_script):
            return jsonify({'ok': False, 'error': 'ไม่พบไฟล์ build_dashboard.py'}), 404

        # Run build script
        result = subprocess.run(
            [sys.executable, build_script],
            cwd=BASE_DIR,
            capture_output=True,
            text=True,
            timeout=120
        )

        return jsonify({
            'ok': result.returncode == 0,
            'message': 'สร้าง Dashboard สำเร็จ' if result.returncode == 0 else 'เกิดข้อผิดพลาด',
            'stdout': result.stdout[-500:] if result.stdout else '',  # Last 500 chars
            'stderr': result.stderr[-500:] if result.stderr else '',
            'returncode': result.returncode
        })
    except subprocess.TimeoutExpired:
        return jsonify({'ok': False, 'error': 'Timeout — build script ใช้เวลานานเกินไป'}), 500
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500

# ─── Catch-all Static Files (must be AFTER all API routes) ────────────────────

@app.route('/<path:path>')
def serve_static(path):
    full_path = os.path.join(BASE_DIR, path)
    if os.path.isfile(full_path):
        return send_from_directory(BASE_DIR, path)
    return send_from_directory(BASE_DIR, 'index.html')

# ─── Main ─────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    print("=" * 60)
    print("  Dashboard GIS — Development Server")
    print(f"  http://localhost:{PORT}")
    print("=" * 60)
    print(f"  Base directory : {BASE_DIR}")
    print(f"  Raw data dir   : {RAW_DATA_DIR}")
    print()

    app.run(host='0.0.0.0', port=PORT, debug=True)
