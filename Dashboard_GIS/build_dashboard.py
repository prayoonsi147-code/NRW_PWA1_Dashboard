#!/usr/bin/env python3
"""
build_dashboard.py - สร้าง Dashboard แผนที่แนวท่อ (GIS) กปภ.เขต 1
อ่านข้อมูลจาก Excel แล้ว embed ลงใน index.html
"""
import openpyxl
import json
import os
import re

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(SCRIPT_DIR, "ข้อมูลดิบ", "ลงข้อมูลซ่อมท่อ")
HTML_TEMPLATE = os.path.join(SCRIPT_DIR, "index.html")

MONTH_NAMES = ['','ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.',
               'ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.']

def clean_num(val):
    if val is None: return 0
    if isinstance(val, (int, float)): return val
    s = str(val).replace(',','').replace('\xa0','').replace(' ','').strip()
    if s == '' or s == '-': return 0
    try: return float(s)
    except: return 0

def read_all_data():
    """Read all xlsx files, group by month, pick latest per month"""
    files = sorted([f for f in os.listdir(DATA_DIR) if f.endswith('.xlsx')])
    print(f"  พบไฟล์ข้อมูล: {len(files)} ไฟล์")

    # Parse filename to extract date: "สรุปข้อมูลจำนวนข้อร้องเรียนงานซ่อม YYMMDD.xlsx"
    file_info = []
    for fname in files:
        m = re.search(r'(\d{6})', fname)
        if not m: continue
        digits = m.group(1)
        yy = int(digits[:2])
        mm = int(digits[2:4])
        dd = int(digits[4:6])
        month_key = f"{yy:02d}-{mm:02d}"
        file_info.append({'fname': fname, 'yy': yy, 'mm': mm, 'dd': dd, 'month_key': month_key})

    # Group by month, pick latest date per month
    month_files = {}
    for fi in file_info:
        mk = fi['month_key']
        if mk not in month_files or fi['dd'] > month_files[mk]['dd']:
            month_files[mk] = fi

    print(f"  เดือนที่มีข้อมูล: {len(month_files)}")
    for mk in sorted(month_files.keys()):
        fi = month_files[mk]
        print(f"    {mk} <- {fi['fname']}")

    # Read data for each month
    all_data = {}
    branches_order = []

    for mk in sorted(month_files.keys()):
        fi = month_files[mk]
        fpath = os.path.join(DATA_DIR, fi['fname'])
        try:
            wb = openpyxl.load_workbook(fpath, data_only=True)
        except Exception as e:
            print(f"  [WARNING] ข้ามไฟล์เสีย: {fi['fname']} ({e})")
            continue
        ws = wb[wb.sheetnames[0]]

        month_data = {}
        for r in range(2, ws.max_row + 1):
            branch = ws.cell(row=r, column=1).value
            if not branch or not isinstance(branch, str): continue
            branch = branch.strip()
            if branch in ('', 'ชื่อสาขา'): continue

            closed = clean_num(ws.cell(row=r, column=2).value)
            complete = clean_num(ws.cell(row=r, column=3).value)
            score = clean_num(ws.cell(row=r, column=4).value)

            month_data[branch] = {
                'closed': closed,
                'complete': complete,
                'score': score
            }
            if branch not in branches_order:
                branches_order.append(branch)

        all_data[mk] = month_data
        wb.close()

    months_sorted = sorted(all_data.keys())
    print(f"  ช่วงเดือน: {months_sorted[0]} - {months_sorted[-1]}")
    print(f"  จำนวนสาขา: {len(branches_order)}")

    return {
        'months': months_sorted,
        'branches': branches_order,
        'data': all_data,
        'month_names': {f"{i:02d}": MONTH_NAMES[i] for i in range(1, 13)}
    }

def build():
    print("=" * 50)
    print("  Build Dashboard แผนที่แนวท่อ (GIS) กปภ.เขต 1")
    print("=" * 50)

    print("\n[1/3] อ่านข้อมูล Excel...")
    data = read_all_data()

    print("\n[2/3] สร้าง JSON...")
    data_json = json.dumps(data, ensure_ascii=False)
    print(f"  ขนาดข้อมูล: {len(data_json):,} bytes")

    print("\n[3/3] Embed ข้อมูลลงใน index.html...")
    with open(HTML_TEMPLATE, 'r', encoding='utf-8') as f:
        html = f.read()

    if 'GIS_DATA_PLACEHOLDER' in html:
        html = html.replace('GIS_DATA_PLACEHOLDER', data_json)
        print("  (ใช้ placeholder)")
    else:
        pattern = r'^const DATA = \{.*\};$'
        new_val = 'const DATA = ' + data_json + ';'
        html_new, count = re.subn(pattern, new_val, html, count=1, flags=re.MULTILINE)
        if count > 0:
            html = html_new
            print("  (แทนที่ const DATA เดิม)")
        else:
            print("  [WARNING] ไม่พบ placeholder หรือ const DATA ใน HTML!")

    output_path = os.path.join(SCRIPT_DIR, "index.html")
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)

    size_kb = os.path.getsize(output_path) / 1024
    print(f"  บันทึก: {output_path}")
    print(f"  ขนาดไฟล์: {size_kb:.1f} KB")
    print("\n  เสร็จสิ้น!")

if __name__ == '__main__':
    build()
