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
                # pending: Repair_MM-YY.ext — ตรวจสอบข้อมูลในไฟล์เพื่อหาเดือน
                # อ่านคอลัมน์ วันที่แจ้ง (col 3) ซึ่งเป็น dd/mm/yyyy (พ.ศ.)
                # แล้วหาเดือนที่มีจำนวนมากที่สุด
                detected_month = None
                try:
                    import tempfile, io
                    raw_bytes = f.read()
                    f.seek(0)  # reset file pointer for later save

                    # Try xlrd for .xls files
                    if ext.lower() == '.xls':
                        try:
                            import xlrd
                            wb = xlrd.open_workbook(file_contents=raw_bytes)
                            ws = wb.sheet_by_index(0)
                            from collections import Counter
                            month_counter = Counter()
                            for row_idx in range(1, min(ws.nrows, 5000)):
                                try:
                                    cell_val = str(ws.cell_value(row_idx, 3)).strip()
                                    if '/' in cell_val:
                                        parts = cell_val.split('/')
                                        if len(parts) >= 3:
                                            mm = int(parts[1])
                                            yy = int(parts[2])
                                            # Convert Buddhist era to 2-digit year
                                            if yy > 2500:
                                                yy = yy - 2500
                                            month_counter[(mm, yy)] += 1
                                except:
                                    pass
                            if month_counter:
                                top_month = month_counter.most_common(1)[0][0]
                                detected_month = top_month  # (mm, yy)
                        except ImportError:
                            pass

                    # Try openpyxl for .xlsx files
                    if detected_month is None and ext.lower() == '.xlsx':
                        try:
                            import openpyxl
                            wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), read_only=True, data_only=True)
                            ws = wb.active
                            from collections import Counter
                            month_counter = Counter()
                            row_count = 0
                            for row in ws.iter_rows(min_row=2, max_col=4, values_only=True):
                                if row_count >= 5000:
                                    break
                                row_count += 1
                                try:
                                    cell_val = str(row[3]).strip() if len(row) > 3 and row[3] else ''
                                    if '/' in cell_val:
                                        parts = cell_val.split('/')
                                        if len(parts) >= 3:
                                            mm = int(parts[1])
                                            yy = int(parts[2])
                                            if yy > 2500:
                                                yy = yy - 2500
                                            month_counter[(mm, yy)] += 1
                                except:
                                    pass
                            wb.close()
                            if month_counter:
                                top_month = month_counter.most_common(1)[0][0]
                                detected_month = top_month
                        except ImportError:
                            pass

                except Exception as detect_err:
                    # If detection fails, fall through to fallback
                    pass

                if detected_month:
                    mm, yy = detected_month
                    new_name = f"Repair_{mm:02d}-{yy:02d}{ext}"
                else:
                    # fallback: ใช้วันที่ upload
                    today = datetime.now()
                    mm = today.month
                    yy = today.year - 2543  # Convert CE to BE 2-digit
                    new_name = f"Repair_{mm:02d}-{yy:02d}{ext}"

            else:
                # fallback ทั่วไป
                m = re.search(r'(\d{6})', name_only)
                if m:
                    new_name = f"{prefix}_{m.group(1)}{ext}"
                else:
                    clean = re.sub(r'[^\w\-.]', '_', name_only).strip('_')[:30]
                    new_name = f"{prefix}_{clean}{ext}"

            dest_path = os.path.join(folder_path, new_name)

            if os.path.exists(dest_path):
                results.append({
                    'filename': new_name, 'original': filename,
                    'status': 'overwrite', 'message': f'เขียนทับ {new_name}'
                })

            f.save(dest_path)

            results.append({
                'filename': new_name, 'original': filename,
                'status': 'success', 'message': f'{filename} → {new_name}'
            })
        except Exception as e:
            errors.append({
                'filename': filename,
                'error': str(e)
            })

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
