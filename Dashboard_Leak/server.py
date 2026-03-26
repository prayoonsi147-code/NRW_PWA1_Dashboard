# -*- coding: utf-8 -*-
"""
Dashboard Leak — Local Development Server
============================================
Flask server สำหรับ Dashboard น้ำสูญเสีย
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
PORT = 5001

# Category mapping: URL slug → Thai folder name
CATEGORY_MAP = {
    'ois': 'OIS',
    'rl': 'Real Leak',
    'mnf': 'MNF',
    'p3': 'P3',
    'activities': 'Activities',
    'eu': 'หน่วยไฟ',
    'kpi': 'เกณฑ์ชี้วัด',
    'kpi2': 'เกณฑ์วัดน้ำสูญเสีย',
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

    # Auto-rename prefix map
    PREFIX_MAP = {
        'ois': 'OIS', 'rl': 'RL', 'mnf': 'MNF', 'p3': 'P3',
        'activities': 'ACT', 'eu': 'EU', 'kpi': 'KPI', 'kpi2': 'KPI2'
    }

    results = []
    errors = []

    for f in files:
        filename = f.filename.strip()
        if not filename:
            continue

        try:
            # Auto-rename: PREFIX_xxxxx.ext
            prefix = PREFIX_MAP.get(category, category.upper())
            name_only = os.path.splitext(filename)[0]
            ext = os.path.splitext(filename)[1] or '.xlsx'
            import re

            if category == 'p3':
                # P3: extract branch name + YY-MM → P3_สาขา_YY-MM.xlsx
                # Try to find branch name (Thai) and date code
                branch_name = None
                date_code = None
                # Extract YY-MM pattern
                dm = re.search(r'(\d{2}-\d{2})', name_only)
                if dm:
                    date_code = dm.group(1)
                # Extract branch: try known branches or Thai text between underscores
                parts = re.split(r'[_\-]', name_only)
                for p in parts:
                    p = p.strip()
                    if p and not re.match(r'^(P3|p3|\d+)$', p) and not re.match(r'^\d{2}-\d{2}$', p):
                        branch_name = p
                        break
                if branch_name and date_code:
                    new_name = f"{prefix}_{branch_name}_{date_code}{ext}"
                elif branch_name:
                    new_name = f"{prefix}_{branch_name}{ext}"
                elif date_code:
                    new_name = f"{prefix}_{date_code}{ext}"
                else:
                    clean = re.sub(r'[^\w\-.]', '_', name_only).strip('_')[:30]
                    new_name = f"{prefix}_{clean}{ext}"
            else:
                # Other categories: extract numbers/dates
                m = re.search(r'(\d{6})', name_only)  # date code YYMMDD
                if m:
                    new_name = f"{prefix}_{m.group(1)}{ext}"
                else:
                    m2 = re.search(r'(\d{4})', name_only)  # year like 2569
                    if m2:
                        new_name = f"{prefix}_{m2.group(1)}{ext}"
                    else:
                        m3 = re.search(r'(\d{3,4})', name_only)
                        if m3:
                            new_name = f"{prefix}_{m3.group(1)}{ext}"
                        else:
                            clean = re.sub(r'[^\w\-.]', '_', name_only).strip('_')[:30]
                            new_name = f"{prefix}_{clean}{ext}"

            dest_path = os.path.join(folder_path, new_name)

            overwritten = os.path.exists(dest_path)
            f.save(dest_path)

            msg = f'{filename} → {new_name}' + (' (เขียนทับ)' if overwritten else '')
            results.append({
                'filename': new_name,
                'original': filename,
                'status': 'success',
                'message': msg
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
    print("  Dashboard Leak — Development Server")
    print(f"  http://localhost:{PORT}")
    print("=" * 60)
    print(f"  Base directory : {BASE_DIR}")
    print(f"  Raw data dir   : {RAW_DATA_DIR}")
    print()

    app.run(host='0.0.0.0', port=PORT, debug=True)
