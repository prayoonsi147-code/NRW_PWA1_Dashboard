#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_dashboard.py - สร้าง Dashboard อัตโนมัติจากไฟล์ .xls
==========================================
วิธีใช้: ดับเบิลคลิก หรือรัน python build_dashboard.py
สคริปต์จะ:
  1. scan หาไฟล์ .xls ทั้งหมดในโฟลเดอร์ ข้อมูลดิบ/OIS/
  2. อ่านข้อมูลจากทุกไฟล์
  3. สร้าง data.json และ index.html ใหม่
"""

import struct
import json
import os
import sys
import glob
import re
import zipfile
import xml.etree.ElementTree as ET

# ============================================================
# BIFF8 / OLE2 XLS Parser (Pure Python, no external libs)
# ============================================================

class XLSParser:
    """Parse BIFF8 .xls files (OLE2 compound document format)"""

    def __init__(self, filename):
        with open(filename, 'rb') as f:
            self.raw = f.read()
        self.wb = None
        self.sst_strings = []
        self.sheet_names = []
        self.sheets_data = {}  # {sheet_idx: {row: {col: value}}}

    def parse(self):
        self._read_ole2()
        self._parse_workbook()

    def _read_ole2(self):
        d = self.raw
        if d[:8] != b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
            raise ValueError("Not an OLE2 file")
        sec_shift = struct.unpack_from('<H', d, 30)[0]
        sec_size = 1 << sec_shift

        # Build FAT
        fat = []
        for i in range(109):
            s = struct.unpack_from('<I', d, 76 + i * 4)[0]
            if s >= 0xFFFFFFFE:
                break
            off = (s + 1) * sec_size
            for j in range(sec_size // 4):
                fat.append(struct.unpack_from('<I', d, off + j * 4)[0])

        def chain(start):
            secs, s = [], start
            while s < 0xFFFFFFFE and len(secs) < 50000:
                secs.append(s)
                s = fat[s] if s < len(fat) else 0xFFFFFFFE
            return secs

        # Read directory
        first_dir = struct.unpack_from('<I', d, 48)[0]
        dd = b''
        for s in chain(first_dir):
            dd += d[(s + 1) * sec_size:(s + 2) * sec_size]

        # Find Workbook stream
        for i in range(0, len(dd), 128):
            e = dd[i:i + 128]
            nl = struct.unpack_from('<H', e, 64)[0]
            nm = e[:nl].decode('utf-16-le', errors='replace').rstrip('\x00')
            if nm.lower() in ('workbook', 'book'):
                start = struct.unpack_from('<I', e, 116)[0]
                size = struct.unpack_from('<I', e, 120)[0]
                wb = b''
                for s in chain(start):
                    wb += d[(s + 1) * sec_size:(s + 2) * sec_size]
                self.wb = wb[:size]
                return
        raise ValueError("No Workbook stream found")

    def _parse_sst_string(self, data, offset):
        """Parse one Unicode string from SST data"""
        if offset + 3 > len(data):
            return None, offset
        cc = struct.unpack_from('<H', data, offset)[0]
        flags = data[offset + 2]
        is_utf16 = bool(flags & 0x01)
        has_rich = bool(flags & 0x08)
        has_phonetic = bool(flags & 0x04)

        pos = offset + 3
        rich_runs = phonetic_size = 0
        if has_rich:
            if pos + 2 > len(data):
                return None, pos
            rich_runs = struct.unpack_from('<H', data, pos)[0]
            pos += 2
        if has_phonetic:
            if pos + 4 > len(data):
                return None, pos
            phonetic_size = struct.unpack_from('<I', data, pos)[0]
            pos += 4

        byte_len = cc * 2 if is_utf16 else cc
        if pos + byte_len > len(data):
            return None, pos
        if is_utf16:
            s = data[pos:pos + byte_len].decode('utf-16-le', errors='replace')
        else:
            s = data[pos:pos + byte_len].decode('latin-1', errors='replace')
        pos += byte_len + rich_runs * 4 + phonetic_size

        return s, pos

    def _decode_rk(self, rk_val):
        """Decode an RK value to float"""
        if rk_val & 2:
            val = float(rk_val >> 2) if not (rk_val & 0x80000000) else float(struct.unpack('<i', struct.pack('<I', rk_val))[0] >> 2)
        else:
            # IEEE 754 with top 30 bits
            raw = (rk_val & 0xFFFFFFFC) << 32
            val = struct.unpack('<d', struct.pack('<Q', raw))[0]
        if rk_val & 1:
            val /= 100.0
        return val

    def _parse_workbook(self):
        wb = self.wb
        pos = 0
        sheet_names = []
        sst_raw = b''
        unique_count = 0

        # === Pass 1: Global records (SST, BOUNDSHEET) ===
        global_done = False
        while pos + 4 <= len(wb) and not global_done:
            op = struct.unpack_from('<H', wb, pos)[0]
            ln = struct.unpack_from('<H', wb, pos + 2)[0]
            rd = wb[pos + 4:pos + 4 + ln]

            if op == 0x0085 and ln >= 8:  # BOUNDSHEET8
                vis = rd[4]
                stype = rd[5]
                str_len = rd[6]
                flag = rd[7]
                if flag & 0x01:
                    name = rd[8:8 + str_len * 2].decode('utf-16-le', errors='replace')
                else:
                    name = rd[8:8 + str_len].decode('latin-1', errors='replace')
                sheet_names.append(name)

            elif op == 0x00FC:  # SST
                unique_count = struct.unpack_from('<I', rd, 4)[0]
                sst_raw = rd[8:]
                # Collect CONTINUE records
                npos = pos + 4 + ln
                while npos + 4 <= len(wb):
                    nop = struct.unpack_from('<H', wb, npos)[0]
                    nln = struct.unpack_from('<H', wb, npos + 2)[0]
                    if nop == 0x003C:
                        sst_raw += wb[npos + 4:npos + 4 + nln]
                        npos += 4 + nln
                    else:
                        break

            pos += 4 + ln

        # Parse SST strings
        self.sst_strings = []
        offset = 0
        for i in range(unique_count):
            s, offset = self._parse_sst_string(sst_raw, offset)
            if s is None:
                break
            self.sst_strings.append(s)

        self.sheet_names = sheet_names

        # === Pass 2: Sheet data ===
        pos = 0
        bof_count = 0
        current_sheet = -1
        pending_formula_cell = None  # (row, col) for STRING record after FORMULA

        while pos + 4 <= len(wb):
            op = struct.unpack_from('<H', wb, pos)[0]
            ln = struct.unpack_from('<H', wb, pos + 2)[0]
            rd = wb[pos + 4:pos + 4 + ln]

            if op == 0x0809:  # BOF
                bof_count += 1
                if bof_count >= 2:  # Skip global BOF
                    current_sheet = bof_count - 2
                    self.sheets_data[current_sheet] = {}
                pending_formula_cell = None

            elif op == 0x000A:  # EOF
                pending_formula_cell = None

            elif current_sheet >= 0:
                sd = self.sheets_data[current_sheet]

                if op == 0x00FD and ln >= 10:  # LABELSST
                    row = struct.unpack_from('<H', rd, 0)[0]
                    col = struct.unpack_from('<H', rd, 2)[0]
                    sst_idx = struct.unpack_from('<I', rd, 6)[0]
                    if sst_idx < len(self.sst_strings):
                        sd.setdefault(row, {})[col] = self.sst_strings[sst_idx]
                    pending_formula_cell = None

                elif op == 0x0203 and ln >= 14:  # NUMBER
                    row = struct.unpack_from('<H', rd, 0)[0]
                    col = struct.unpack_from('<H', rd, 2)[0]
                    val = struct.unpack_from('<d', rd, 6)[0]
                    sd.setdefault(row, {})[col] = val
                    pending_formula_cell = None

                elif op == 0x027E and ln >= 10:  # RK
                    row = struct.unpack_from('<H', rd, 0)[0]
                    col = struct.unpack_from('<H', rd, 2)[0]
                    rk = struct.unpack_from('<I', rd, 6)[0]
                    sd.setdefault(row, {})[col] = self._decode_rk(rk)
                    pending_formula_cell = None

                elif op == 0x00BD and ln >= 6:  # MULRK
                    row = struct.unpack_from('<H', rd, 0)[0]
                    fc = struct.unpack_from('<H', rd, 2)[0]
                    i = 4
                    col = fc
                    while i + 6 <= ln:
                        rk = struct.unpack_from('<I', rd, i + 2)[0]
                        sd.setdefault(row, {})[col] = self._decode_rk(rk)
                        col += 1
                        i += 6
                    pending_formula_cell = None

                elif op == 0x0006 and ln >= 20:  # FORMULA
                    row = struct.unpack_from('<H', rd, 0)[0]
                    col = struct.unpack_from('<H', rd, 2)[0]
                    result_bytes = rd[6:14]
                    if result_bytes[6] == 0xFF and result_bytes[7] == 0xFF:
                        # String result - will follow in STRING record
                        pending_formula_cell = (row, col)
                    else:
                        val = struct.unpack_from('<d', result_bytes, 0)[0]
                        sd.setdefault(row, {})[col] = val
                        pending_formula_cell = None

                elif op == 0x0207 and pending_formula_cell:  # STRING (formula string result)
                    if ln >= 3:
                        cc = struct.unpack_from('<H', rd, 0)[0]
                        flag = rd[2]
                        if flag & 0x01:
                            s = rd[3:3 + cc * 2].decode('utf-16-le', errors='replace')
                        else:
                            s = rd[3:3 + cc].decode('latin-1', errors='replace')
                        r, c = pending_formula_cell
                        sd.setdefault(r, {})[c] = s
                    pending_formula_cell = None
                else:
                    if op not in (0x003C, 0x0207):
                        pending_formula_cell = None

            pos += 4 + ln

    def get_sheet_count(self):
        return len(self.sheet_names)

    def get_sheet_name(self, idx):
        return self.sheet_names[idx] if idx < len(self.sheet_names) else f"Sheet{idx}"

    def get_cell(self, sheet_idx, row, col):
        return self.sheets_data.get(sheet_idx, {}).get(row, {}).get(col)

    def get_sheet_data(self, sheet_idx):
        return self.sheets_data.get(sheet_idx, {})


# ============================================================
# XLSX Parser (for .xls files that are actually OOXML/ZIP)
# ============================================================

class XLSXParser:
    """Parse OOXML .xlsx files (ZIP-based format)"""

    def __init__(self, filename):
        self.filename = filename
        self.sheet_names = []
        self.sheets_data = {}
        self.shared_strings = []

    def parse(self):
        with zipfile.ZipFile(self.filename, 'r') as zf:
            # Read shared strings
            if 'xl/sharedStrings.xml' in zf.namelist():
                self._parse_shared_strings(zf.read('xl/sharedStrings.xml'))

            # Read workbook for sheet names
            self._parse_workbook(zf.read('xl/workbook.xml'))

            # Read each sheet
            for si, sname in enumerate(self.sheet_names):
                sheet_file = f'xl/worksheets/sheet{si + 1}.xml'
                if sheet_file in zf.namelist():
                    self.sheets_data[si] = self._parse_sheet(zf.read(sheet_file))

    def _ns(self, tag):
        """Add spreadsheetml namespace"""
        return f'{{http://schemas.openxmlformats.org/spreadsheetml/2006/main}}{tag}'

    def _parse_shared_strings(self, xml_bytes):
        root = ET.fromstring(xml_bytes)
        ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        for si in root.findall(f'{{{ns}}}si'):
            texts = []
            for t in si.iter(f'{{{ns}}}t'):
                if t.text:
                    texts.append(t.text)
            self.shared_strings.append(''.join(texts))

    def _parse_workbook(self, xml_bytes):
        root = ET.fromstring(xml_bytes)
        ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        for sheet in root.iter(f'{{{ns}}}sheet'):
            self.sheet_names.append(sheet.get('name', ''))

    def _col_to_idx(self, col_str):
        """Convert column letter (A, B, AA) to 0-based index"""
        result = 0
        for c in col_str:
            result = result * 26 + (ord(c) - ord('A') + 1)
        return result - 1

    def _parse_cell_ref(self, ref):
        """Parse cell reference like 'A1' into (row, col)"""
        match = re.match(r'([A-Z]+)(\d+)', ref)
        if not match:
            return None, None
        col = self._col_to_idx(match.group(1))
        row = int(match.group(2)) - 1  # 0-based
        return row, col

    def _parse_sheet(self, xml_bytes):
        root = ET.fromstring(xml_bytes)
        ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        data = {}

        for row_el in root.iter(f'{{{ns}}}row'):
            for cell in row_el.findall(f'{{{ns}}}c'):
                ref = cell.get('r', '')
                row, col = self._parse_cell_ref(ref)
                if row is None:
                    continue

                cell_type = cell.get('t', '')
                val_el = cell.find(f'{{{ns}}}v')

                if val_el is not None and val_el.text is not None:
                    if cell_type == 's':
                        # Shared string reference
                        idx = int(val_el.text)
                        if idx < len(self.shared_strings):
                            data.setdefault(row, {})[col] = self.shared_strings[idx]
                    elif cell_type == 'str' or cell_type == 'inlineStr':
                        data.setdefault(row, {})[col] = val_el.text
                    else:
                        try:
                            val = float(val_el.text)
                            data.setdefault(row, {})[col] = val
                        except ValueError:
                            data.setdefault(row, {})[col] = val_el.text
                elif cell_type == 'inlineStr':
                    is_el = cell.find(f'{{{ns}}}is')
                    if is_el is not None:
                        texts = []
                        for t in is_el.iter(f'{{{ns}}}t'):
                            if t.text:
                                texts.append(t.text)
                        if texts:
                            data.setdefault(row, {})[col] = ''.join(texts)

        return data

    def get_sheet_count(self):
        return len(self.sheet_names)

    def get_sheet_name(self, idx):
        return self.sheet_names[idx] if idx < len(self.sheet_names) else f"Sheet{idx}"

    def get_cell(self, sheet_idx, row, col):
        return self.sheets_data.get(sheet_idx, {}).get(row, {}).get(col)

    def get_sheet_data(self, sheet_idx):
        return self.sheets_data.get(sheet_idx, {})


# ============================================================
# Data Extraction Logic
# ============================================================

MONTH_KEYWORDS = ['ต.ค.', 'พ.ย.', 'ธ.ค.', 'ม.ค.', 'ก.พ.', 'มี.ค.',
                  'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.',
                  'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม', 'มกราคม', 'กุมภาพันธ์',
                  'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม',
                  'สิงหาคม', 'กันยายน']


def find_month_header_row(sheet_data):
    """Find the row that contains month column headers (requires 6+ month keywords)"""
    for row_num in sorted(sheet_data.keys()):
        row = sheet_data[row_num]
        text_cells = [str(v) for v in row.values() if isinstance(v, str)]
        row_text = ' '.join(text_cells)
        count = sum(1 for kw in MONTH_KEYWORDS if kw in row_text)
        if count >= 6:
            return row_num
    return None


def find_month_columns(sheet_data, header_row):
    """Find the column indices for each of the 12 months"""
    row = sheet_data.get(header_row, {})
    month_cols = [None] * 12  # index 0=ต.ค., 1=พ.ย., ..., 11=ก.ย.

    month_short = ['ต.ค.', 'พ.ย.', 'ธ.ค.', 'ม.ค.', 'ก.พ.', 'มี.ค.',
                   'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.']
    month_long = ['ตุลาคม', 'พฤศจิกายน', 'ธันวาคม', 'มกราคม', 'กุมภาพันธ์',
                  'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม',
                  'สิงหาคม', 'กันยายน']

    for col, val in row.items():
        if not isinstance(val, str):
            continue
        for mi in range(12):
            if month_short[mi] in val or month_long[mi] in val:
                month_cols[mi] = col
                break

    return month_cols


def find_total_column(sheet_data, header_row):
    """Find the 'รวม' (total) column"""
    row = sheet_data.get(header_row - 1, {})  # Usually one row above month headers
    for col, val in row.items():
        if isinstance(val, str) and 'รวม' in val:
            return col
    # Also check the header row itself
    row2 = sheet_data.get(header_row, {})
    for col, val in row2.items():
        if isinstance(val, str) and 'รวม' in val:
            return col
    return None


def extract_sheet_data(sheet_data, header_row, month_cols, total_col):
    """Extract data rows from a sheet"""
    rows = []
    data_start = header_row + 1

    # Collect all data rows
    for row_num in sorted(sheet_data.keys()):
        if row_num < data_start:
            continue
        row = sheet_data.get(row_num, {})

        # Get label (col 0)
        label = row.get(0, '')
        if isinstance(label, (int, float)):
            label = str(label)
        label = label.strip()

        if not label:
            continue

        # Skip header-like rows and "หมายเหตุ"
        if 'หมายเหตุ' in label:
            continue

        # Get unit (col 1)
        unit = row.get(1, '')
        if isinstance(unit, (int, float)):
            unit = str(unit)
        unit = unit.strip()

        # Get 12 monthly values
        monthly = []
        for mi in range(12):
            mc = month_cols[mi]
            if mc is not None:
                val = row.get(mc)
                if isinstance(val, (int, float)):
                    monthly.append(val)
                else:
                    monthly.append(None)
            else:
                monthly.append(None)

        # Get total
        total = None
        if total_col is not None:
            tv = row.get(total_col)
            if isinstance(tv, (int, float)):
                total = tv

        # Get target values: Col C (index 2) = yearly target, Col E (index 4) = monthly target
        target_year = None
        target_month = None
        tv_y = row.get(2)
        tv_m = row.get(4)
        if isinstance(tv_y, (int, float)):
            target_year = tv_y
        if isinstance(tv_m, (int, float)):
            target_month = tv_m

        rows.append({
            'label': label,
            'unit': unit,
            'monthly': monthly,
            'total': total,
            'target_year': target_year,
            'target_month': target_month,
            'hasData': any(v is not None and v != 0 for v in monthly)
        })

    return rows


# ============================================================
# Label Normalization
# ============================================================
# บางปีใช้ชื่อรายการต่างกันเล็กน้อย เช่น
# 2557: "อัตราการสูญเสีย" vs 2558+: "อัตราน้ำสูญเสีย"
# 2562: "น้ำจ่ายฟรี + Blowoff" vs อื่นๆ: "น้ำจ่ายฟรี"
# 2562-64: ลำดับหัวข้อ 4.x เปลี่ยน

LABEL_NORMALIZE_MAP = {
    '2.5 อัตราการสูญเสีย (ต่อน้ำผลิตจ่าย)': '2.5 อัตราน้ำสูญเสีย (ต่อน้ำผลิตจ่าย)',
    '2.2  ปริมาณน้ำจ่ายฟรี + Blowoff': '2.2  ปริมาณน้ำจ่ายฟรี',
    '4.2 เงินเดือนและค่าจ้างประจำ': '4.1 เงินเดือนและค่าจ้างประจำ',
    '4.3 ค่าจ้างชั่วคราว': '4.2 ค่าจ้างชั่วคราว',
    '4.5 วัสดุการผลิต': '4.4 วัสดุการผลิต',
}


def normalize_labels(all_data):
    """Normalize label names so the same metric matches across years"""
    for year_str, sheets in all_data.items():
        for sheet_name, sheet_info in sheets.items():
            for row in sheet_info['rows']:
                canonical = LABEL_NORMALIZE_MAP.get(row['label'])
                if canonical:
                    row['label'] = canonical


def fix_trailing_zeros(all_data):
    """
    For incomplete fiscal years, convert trailing 0s to null.
    Uses 30% threshold: if fewer than 30% of rows have non-zero values
    for a given month, those values are likely not real data.
    """
    for year_str, sheets in all_data.items():
        for sheet_name, sheet_info in sheets.items():
            rows = sheet_info['rows']
            if not rows:
                continue

            # Find last month with "real" data
            last_real_month = -1
            for mi in range(12):
                non_zero_count = sum(1 for r in rows if r['monthly'][mi] is not None and r['monthly'][mi] != 0)
                total_rows = len([r for r in rows if r['monthly'][mi] is not None])
                if total_rows > 0 and non_zero_count / max(len(rows), 1) >= 0.30:
                    last_real_month = mi

            # Convert trailing zeros to null
            if last_real_month < 11:
                for r in rows:
                    for mi in range(last_real_month + 1, 12):
                        if r['monthly'][mi] == 0:
                            r['monthly'][mi] = None
                    # Update hasData
                    r['hasData'] = any(v is not None and v != 0 for v in r['monthly'])


def extract_year_from_filename(filename):
    """Extract Buddhist Era year from filename like '2565.xls'"""
    base = os.path.splitext(os.path.basename(filename))[0]
    match = re.search(r'(\d{4})', base)
    if match:
        return match.group(1)
    return None


# ============================================================
# Real Leak Data Extraction
# ============================================================

# Standard branch names for normalization
STANDARD_BRANCHES = [
    'ชลบุรี(พ)', 'พัทยา(พ)', 'พนัสนิคม', 'บ้านบึง', 'ศรีราชา', 'แหลมฉบัง',
    'ฉะเชิงเทรา', 'บางปะกง', 'บางคล้า', 'พนมสารคาม', 'ระยอง', 'บ้านฉาง',
    'ปากน้ำประแสร์', 'จันทบุรี', 'ขลุง', 'ตราด', 'คลองใหญ่', 'สระแก้ว',
    'วัฒนานคร', 'อรัญประเทศ', 'ปราจีนบุรี', 'กบินทร์บุรี'
]

# Month abbreviations for matching tab names
RL_MONTH_ABBR = {
    'ต.ค.': 0, 'พ.ย.': 1, 'ธ.ค.': 2, 'ม.ค.': 3, 'ก.พ.': 4, 'มี.ค.': 5,
    'เม.ย.': 6, 'พ.ค.': 7, 'มิ.ย.': 8, 'ก.ค.': 9, 'ส.ค.': 10, 'ก.ย.': 11
}


# Known typo aliases in Excel source files
BRANCH_ALIASES = {
    'พนัมสารคาม': 'พนมสารคาม',
}

def normalize_branch_name(raw_name):
    """Normalize branch name from Excel to standard name"""
    if not raw_name or not isinstance(raw_name, str):
        return None
    name = raw_name.strip()
    # Remove leading numbers and dots like "1.", "01."
    name = re.sub(r'^\d+\.?\s*', '', name)
    # Check alias table first (known typos in source files)
    for alias, std in BRANCH_ALIASES.items():
        if alias in name:
            return std
    # Try exact match first
    for std in STANDARD_BRANCHES:
        if name == std:
            return std
    # Try matching core name (without parenthetical suffixes in source)
    for std in STANDARD_BRANCHES:
        # Extract core of standard name (e.g., "ชลบุรี" from "ชลบุรี(พ)")
        core = re.sub(r'\(.*?\)', '', std).strip()
        # Check if raw name starts with or contains the core
        raw_core = re.sub(r'\(.*?\)', '', name).strip()
        raw_core = re.sub(r'\s*(พ|น\.\d+)$', '', raw_core).strip()
        if raw_core == core:
            return std
    return None


def parse_rl_month_tab(tab_name, file_year_str=None):
    """Parse tab name like 'ต.ค.68' or 'พ.ย. 68' to (month_index, fiscal_year_str).
    Returns (mi, fy_str) or (None, None).
    Fiscal year: ต.ค.-ธ.ค. of year Y belongs to FY (2500+Y+1),
                 ม.ค.-ก.ย. of year Y belongs to FY (2500+Y).
    """
    tab_name = tab_name.strip()
    mi = None
    for abbr, idx in RL_MONTH_ABBR.items():
        if abbr in tab_name:
            mi = idx
            break
    if mi is None:
        return None, None

    # Extract 2-digit calendar year from tab name (e.g., '68' from 'ต.ค.68')
    year_match = re.search(r'(\d{2})\s*$', tab_name)
    if year_match:
        cal_year_short = int(year_match.group(1))
        cal_year = 2500 + cal_year_short  # e.g., 68 -> 2568
        # ต.ค.(0), พ.ย.(1), ธ.ค.(2) → fiscal year = cal_year + 1
        # ม.ค.(3) - ก.ย.(11) → fiscal year = cal_year
        if mi <= 2:  # ต.ค., พ.ย., ธ.ค.
            fy = cal_year + 1
        else:  # ม.ค. - ก.ย.
            fy = cal_year
        return mi, str(fy)

    # Fallback: use file year if no year in tab name
    return mi, file_year_str


def process_rl_file(filepath):
    """Process one Real Leak .xlsx file.
    Returns: {year_str: {branch_std: {months: [12 values or None for rate %]}}}
    """
    print(f"  กำลังอ่าน Real Leak: {os.path.basename(filepath)} ...")

    parser = XLSXParser(filepath)
    parser.parse()

    year_str = extract_year_from_filename(filepath)
    if not year_str:
        return None, None

    # result[branch_std] = { 'rate': [12 x float or None],
    #                        'volume': [12 x float or None],
    #                        'production': [12 x float or None],
    #                        'supplied': [12 x float or None],
    #                        'sold': [12 x float or None],
    #                        'blowoff': [12 x float or None] }
    result = {}
    for std in STANDARD_BRANCHES:
        result[std] = {
            'rate': [None]*12,
            'volume': [None]*12,
            'production': [None]*12,
            'supplied': [None]*12,
            'sold': [None]*12,
            'blowoff': [None]*12
        }

    for si in range(parser.get_sheet_count()):
        sname = parser.get_sheet_name(si)
        mi, fy_str = parse_rl_month_tab(sname, year_str)
        if mi is None:
            continue
        # Only process sheets that belong to this file's fiscal year
        if fy_str != year_str:
            continue

        sd = parser.get_sheet_data(si)
        if not sd:
            continue

        # Find header row (row with "กปภ.สาขา" or "สาขา")
        header_row = None
        col_branch = 1  # default B
        for rn in sorted(sd.keys()):
            row = sd[rn]
            for cn, val in row.items():
                if isinstance(val, str) and ('สาขา' in val):
                    header_row = rn
                    col_branch = cn
                    break
            if header_row is not None:
                break

        if header_row is None:
            header_row = 1  # fallback

        # Determine column mapping
        # Layout varies: ต.ค. has fewer columns, later months have สะสม columns
        # Strategy: find "น้ำสูญเสีย" in header row, then check sub-header for ปริมาณ/อัตรา
        col_map = {}
        hrow = sd.get(header_row, {})
        hrow2 = sd.get(header_row + 1, {})

        # Also read one more row below (some sheets have 3 header rows)
        hrow3 = sd.get(header_row + 2, {})

        # First pass: simple keyword matching for non-สูญเสีย columns
        # Search across hrow, hrow2, and hrow3 for maximum compatibility
        for cn in set(list(hrow.keys()) + list(hrow2.keys()) + list(hrow3.keys())):
            h1 = str(hrow.get(cn, ''))
            h2 = str(hrow2.get(cn, ''))
            h3 = str(hrow3.get(cn, ''))
            combined = h1 + ' ' + h2 + ' ' + h3
            if 'น้ำผลิตรวม' in combined:
                col_map['production'] = cn
            elif 'น้ำผลิตจ่ายสุทธิ' in combined and 'สะสม' not in combined:
                col_map['supplied'] = cn
            elif 'น้ำจำหน่าย' in combined:
                col_map['sold'] = cn
            elif 'Blow' in combined or 'blow' in combined:
                col_map['blowoff'] = cn

        # Second pass: find "น้ำสูญเสีย" header — search hrow AND hrow2
        wl_start_col = None
        sub_header_row = None
        for search_row, next_row in [(hrow, hrow2), (hrow2, hrow3)]:
            for cn in sorted(search_row.keys()):
                h = str(search_row.get(cn, ''))
                if 'น้ำสูญเสีย' in h:
                    wl_start_col = cn
                    sub_header_row = next_row
                    break
            if wl_start_col is not None:
                break

        if wl_start_col is not None and sub_header_row is not None:
            # Scan sub-header starting from wl_start_col
            for cn in sorted(sub_header_row.keys()):
                if cn < wl_start_col:
                    continue
                h = str(sub_header_row.get(cn, ''))
                if 'ปริมาณ' in h and 'สะสม' not in h:
                    col_map['volume'] = cn
                elif 'อัตรา' in h and 'สะสม' not in h:
                    col_map['rate'] = cn

        # Scan data rows — find first row that contains a branch name
        data_start = header_row + 2  # default
        for rn in sorted(sd.keys()):
            if rn <= header_row:
                continue
            row = sd[rn]
            raw = row.get(col_branch)
            if isinstance(raw, str) and normalize_branch_name(raw):
                data_start = rn
                break
        for rn in sorted(sd.keys()):
            if rn < data_start:
                continue
            row = sd[rn]
            raw_name = row.get(col_branch)
            if not raw_name or not isinstance(raw_name, str):
                continue

            branch = normalize_branch_name(raw_name)
            if branch is None:
                continue

            # Extract rate (%)
            if 'rate' in col_map:
                val = row.get(col_map['rate'])
                if isinstance(val, (int, float)):
                    result[branch]['rate'][mi] = val

            # Extract volume
            if 'volume' in col_map:
                val = row.get(col_map['volume'])
                if isinstance(val, (int, float)):
                    result[branch]['volume'][mi] = val

            # Extract production
            if 'production' in col_map:
                val = row.get(col_map['production'])
                if isinstance(val, (int, float)):
                    result[branch]['production'][mi] = val

            # Extract supplied
            if 'supplied' in col_map:
                val = row.get(col_map['supplied'])
                if isinstance(val, (int, float)):
                    result[branch]['supplied'][mi] = val

            # Extract sold
            if 'sold' in col_map:
                val = row.get(col_map['sold'])
                if isinstance(val, (int, float)):
                    result[branch]['sold'][mi] = val

            # Extract blowoff
            if 'blowoff' in col_map:
                val = row.get(col_map['blowoff'])
                if isinstance(val, (int, float)):
                    result[branch]['blowoff'][mi] = val

    return year_str, result


def build_rl_embedded_data(rl_data):
    """Convert Real Leak data to compact JSON for embedding in HTML"""
    compact = {}
    for year_str in sorted(rl_data.keys()):
        branches = rl_data[year_str]
        compact[year_str] = {}
        for branch, metrics in branches.items():
            compact[year_str][branch] = {
                'r': metrics['rate'],       # rate %
                'v': metrics['volume'],     # volume ลบ.ม.
                'p': metrics['production'], # production
                's': metrics['supplied'],   # supplied
                'd': metrics['sold'],       # sold (demand)
                'b': metrics['blowoff']     # blowoff
            }
    return json.dumps(compact, ensure_ascii=False, separators=(',', ':'))


# ============================================================
# Electric Unit (EU) per Water Sold - Parser
# ============================================================

def process_eu_file(filepath):
    """Process one EU .xlsx file (หน่วยไฟฟ้า/น้ำจำหน่าย).
    File has single sheet 'กราฟ' with:
      Row 1: header with fiscal year label
      Row 2: month headers (ต.ค. - ก.ย.) in columns C-N (3-14)
      Rows 3-24: 22 branches (col B = name, cols C-N = monthly values)
      Row 25: ภาพรวม กปภ.ข.1
    Returns (year_str, {branch_name: [12 monthly values]})
    """
    basename = os.path.basename(filepath)
    # Extract fiscal year from filename: EU-2569.xlsx -> '2569'
    m = re.search(r'EU-(\d{4})', basename)
    if not m:
        print(f"  ⚠️  ไม่พบปี พ.ศ. ในชื่อไฟล์ EU: {basename}")
        return None, None
    year_str = m.group(1)

    parser = XLSXParser(filepath)
    parser.parse()

    if parser.get_sheet_count() == 0:
        return None, None

    # Use first sheet (กราฟ)
    sd = parser.sheets_data.get(0, {})
    if not sd:
        return None, None

    result = {}
    # XLSXParser uses 0-based indexing:
    # Row 0 = header (ปีงบประมาณ), Row 1 = month headers
    # Rows 2-23 = 22 branches (col 0=number, col 1=name, cols 2-13=monthly values)
    # Row 24 = ภาพรวม กปภ.ข.1
    for rn in sorted(sd.keys()):
        if rn < 2:
            continue
        row = sd[rn]
        raw_name = row.get(1, '')  # col 1 (B) = branch name
        if not isinstance(raw_name, str) or not raw_name.strip():
            continue

        # Check for regional aggregate row
        name = raw_name.strip()
        is_regional = 'ภาพรวม' in name

        if is_regional:
            branch_key = '__regional__'
        else:
            branch_key = normalize_branch_name(name)
            if not branch_key:
                continue

        monthly = [None] * 12
        for mi in range(12):
            col = 2 + mi  # col 2 (C) = ต.ค. (mi=0) through col 13 (N) = ก.ย. (mi=11)
            val = row.get(col)
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                monthly[mi] = round(val, 4)
            elif isinstance(val, str):
                # Skip #REF! or other error strings
                pass
        result[branch_key] = monthly

    return year_str, result


def build_eu_embedded_data(eu_data):
    """Convert EU data to compact JSON for embedding in HTML.
    Output: {"2569":{"ชลบุรี(พ)":[0.5,0.4,...], "__regional__":[0.51,0.46,...]}}
    """
    compact = {}
    for year_str in sorted(eu_data.keys()):
        branches = eu_data[year_str]
        compact[year_str] = branches
    return json.dumps(compact, ensure_ascii=False, separators=(',', ':'))


# ============================================================
# MNF (Minimum Night Flow) - Parser
# ============================================================

# Known row labels in MNF sheets (row name -> key)
MNF_ROW_MAP = {
    'MNF เกิดจริง': 'actual',
    'MNF ที่ยอมรับได้': 'acceptable',
    'เป้าหมาย MNF': 'target',
    'น้ำผลิตจ่าย': 'production',
}

# Map sheet names like "1.ชลบุรี" to standard branch names
MNF_SHEET_MAP = {
    '1.ชลบุรี': 'ชลบุรี(พ)',
    '2.พัทยา': 'พัทยา(พ)',
    '3.บ้านบึง': 'บ้านบึง',
    '4.พนัสนิคม': 'พนัสนิคม',
    '5.ศรีราชา': 'ศรีราชา',
    '6.แหลมฉบัง': 'แหลมฉบัง',
    '7.บางปะกง': 'บางปะกง',
    '8.ฉะเชิงเทรา': 'ฉะเชิงเทรา',
    '9.บางคล้า': 'บางคล้า',
    '10.พนมสารคาม': 'พนมสารคาม',
    '11.ระยอง': 'ระยอง',
    '12.บ้านฉาง': 'บ้านฉาง',
    '13.ปากน้ำประแสร์': 'ปากน้ำประแสร์',
    '14.จันทบุรี': 'จันทบุรี',
    '15.ขลุง': 'ขลุง',
    '16.ตราด': 'ตราด',
    '17.คลองใหญ่': 'คลองใหญ่',
    '18.สระแก้ว': 'สระแก้ว',
    '19.วัฒนานคร': 'วัฒนานคร',
    '20.อรัญประเทศ': 'อรัญประเทศ',
    '21.ปราจีนบุรี': 'ปราจีนบุรี',
    '22.กบินทร์บุรี': 'กบินทร์บุรี',
}


def process_mnf_file(filepath):
    """Process one MNF .xlsx file (Minimum Night Flow).
    Structure:
      - Sheet 'ภาพรวมเขต': regional summary
        R1: title, R2: month headers (col 2-13),
        R3: MNF เกิดจริง, R4: MNF ที่ยอมรับได้, R5: เป้าหมาย MNF, R6: น้ำผลิตจ่าย
      - Sheet '1.ชลบุรี' .. '22.กบินทร์บุรี': per-branch
        R2: branch name, R3: month headers (col 2-13),
        R4: MNF เกิดจริง, R5: MNF ที่ยอมรับได้, R6: เป้าหมาย MNF, R7: น้ำผลิตจ่าย
    Returns (year_str, {branch_or_regional: {actual:[12], acceptable:[12], target:[12], production:[12]}})
    """
    basename = os.path.basename(filepath)
    m = re.search(r'MNF-(\d{4})', basename)
    if not m:
        print(f"  ⚠️  ไม่พบปี พ.ศ. ในชื่อไฟล์ MNF: {basename}")
        return None, None
    year_str = m.group(1)

    print(f"  กำลังอ่าน MNF: {basename} ...")
    parser = XLSXParser(filepath)
    parser.parse()

    result = {}

    for si in range(parser.get_sheet_count()):
        sn = parser.sheet_names[si]
        sd = parser.sheets_data.get(si, {})
        if not sd:
            continue

        # Determine branch key
        if sn == 'ภาพรวมเขต':
            branch_key = '__regional__'
            # Regional: data rows start at row index 2 (R3)
            data_start_row = 2
        elif sn in MNF_SHEET_MAP:
            branch_key = MNF_SHEET_MAP[sn]
            # Branch: data rows start at row index 3 (R4)
            data_start_row = 3
        elif sn == 'รวมกราฟสาขา':
            continue  # Skip summary chart sheet
        else:
            continue

        metrics = {
            'actual': [None] * 12,
            'acceptable': [None] * 12,
            'target': [None] * 12,
            'production': [None] * 12,
        }

        # Read data rows
        for rn in sorted(sd.keys()):
            if rn < data_start_row:
                continue
            row = sd[rn]
            label = row.get(0, '')
            if not isinstance(label, str):
                label = str(label).strip() if label is not None else ''
            else:
                label = label.strip()

            # Match label to metric key
            metric_key = None
            for known_label, key in MNF_ROW_MAP.items():
                if known_label in label:
                    metric_key = key
                    break

            if not metric_key:
                continue

            # Read 12 monthly values from col 1-12 (0-based: B-M)
            for mi in range(12):
                col = 1 + mi  # col 1 (B) = ต.ค., col 12 (M) = ก.ย.
                val = row.get(col)
                if isinstance(val, (int, float)) and not isinstance(val, bool):
                    # Treat 0 as no-data if actual MNF (could be unfilled month)
                    if metric_key == 'actual' and val == 0:
                        metrics[metric_key][mi] = None
                    else:
                        metrics[metric_key][mi] = round(float(val), 4)

        result[branch_key] = metrics

    return year_str, result


def build_mnf_embedded_data(mnf_data):
    """Convert MNF data to compact JSON for embedding in HTML.
    Output: {"2569":{"__regional__":{"a":[...],"c":[...],"t":[...],"p":[...]}, "ชลบุรี(พ)":{...}}}
    Keys: a=actual, c=acceptable, t=target, p=production
    """
    compact = {}
    for year_str in sorted(mnf_data.keys()):
        branches = mnf_data[year_str]
        compact[year_str] = {}
        for branch, metrics in branches.items():
            compact[year_str][branch] = {
                'a': metrics['actual'],
                'c': metrics['acceptable'],
                't': metrics['target'],
                'p': metrics['production'],
            }
    return json.dumps(compact, ensure_ascii=False, separators=(',', ':'))


# ============================================================
# KPI Leak (เกณฑ์วัดน้ำสูญเสีย) - Parser
# ============================================================

def process_kpi_file(filepath):
    """Process one KPI_Leak-XXXX.xlsx file.
    Returns (year_str, {branch_name: {target: float, levels: [5 floats], actual: float}})
    """
    fname = os.path.basename(filepath)
    # Extract year from filename: KPI_Leak-2569.xlsx or KPI_Leal-2569.xlsx
    m = re.search(r'(\d{4})', fname)
    if not m:
        return None, None
    year_str = m.group(1)

    parser = XLSXParser(filepath)
    parser.parse()

    result = {}
    for si in range(parser.get_sheet_count()):
        sdata = parser.get_sheet_data(si)
        if not sdata:
            continue
        # Find header row with 'กปภ.สาขา' or 'สาขา'
        header_row = None
        for r in sorted(sdata.keys()):
            for c in sdata[r]:
                val = sdata[r][c]
                if isinstance(val, str) and ('สาขา' in val):
                    header_row = r
                    break
            if header_row is not None:
                break
        if header_row is None:
            continue

        # Read data rows after header (skip sub-header rows for level labels)
        data_start = header_row + 2  # skip level number row
        for r in sorted(sdata.keys()):
            if r < data_start:
                continue
            row = sdata[r]
            # Column 1 = ลำดับ or 'รวม', Column 2 = branch name
            branch_raw = row.get(1, None)  # 0-indexed: col B = index 1
            if branch_raw is None:
                # Check if col 0 has 'รวม'
                c0 = row.get(0, '')
                if isinstance(c0, str) and 'รวม' in c0:
                    branch_raw = c0
                else:
                    continue

            branch_name = str(branch_raw).strip()
            if not branch_name:
                continue

            # Normalize KPI branch name to match RL standard names
            branch_name = normalize_kpi_branch(branch_name)

            # Read values
            target_ois = _to_float(row.get(2, None))   # col C
            level1 = _to_float(row.get(3, None))       # col D
            level2 = _to_float(row.get(4, None))       # col E
            level3 = _to_float(row.get(5, None))       # col F
            level4 = _to_float(row.get(6, None))       # col G
            level5 = _to_float(row.get(7, None))       # col H
            actual_target = _to_float(row.get(8, None)) # col I

            if target_ois is None and level1 is None:
                continue

            result[branch_name] = {
                'target': target_ois,
                'levels': [level1, level2, level3, level4, level5],
                'actual': actual_target
            }

    return year_str, result if result else None


def _to_float(val):
    """Convert value to float, return None if not numeric"""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    try:
        return float(str(val).strip().replace(',', ''))
    except (ValueError, TypeError):
        return None


def normalize_kpi_branch(name):
    """Map KPI branch names to standard RL branch names"""
    mapping = {
        'ชลบุรี': 'ชลบุรี(พ)',
        'พัทยา': 'พัทยา(พ)',
        'บ้านบึง': 'บ้านบึง',
        'พนัสนิคม': 'พนัสนิคม',
        'ศรีราชา': 'ศรีราชา',
        'แหลมฉบัง': 'แหลมฉบัง',
        'ฉะเชิงเทรา': 'ฉะเชิงเทรา',
        'บางปะกง': 'บางปะกง',
        'บางคล้า': 'บางคล้า',
        'พนมสารคาม': 'พนมสารคาม',
        'ระยอง': 'ระยอง',
        'บ้านฉาง': 'บ้านฉาง',
        'ปากน้ำประแสร์': 'ปากน้ำประแสร์',
        'จันทบุรี': 'จันทบุรี',
        'ขลุง': 'ขลุง',
        'ตราด': 'ตราด',
        'คลองใหญ่': 'คลองใหญ่',
        'สระแก้ว': 'สระแก้ว',
        'วัฒนานคร': 'วัฒนานคร',
        'อรัญประเทศ': 'อรัญประเทศ',
        'ปราจีนบุรี': 'ปราจีนบุรี',
        'กบินทร์บุรี': 'กบินทร์บุรี',
    }
    # Direct mapping
    if name in mapping:
        return mapping[name]
    # Check if it's a regional total
    if 'รวม' in name:
        return '__regional__'
    # Fuzzy match
    for kpi_name, std_name in mapping.items():
        if kpi_name in name or name in kpi_name:
            return std_name
    return name


def build_kpi_embedded_data(kpi_data):
    """Convert KPI data to compact JSON for embedding in HTML.
    Output: {year: {branch: {t: target, l: [5 levels], a: actual_target}}}
    """
    compact = {}
    for year_str in sorted(kpi_data.keys()):
        branches = kpi_data[year_str]
        compact[year_str] = {}
        for branch, info in branches.items():
            compact[year_str][branch] = {
                't': info['target'],
                'l': info['levels'],
                'a': info['actual']
            }
    return json.dumps(compact, ensure_ascii=False, separators=(',', ':'))


# ============================================================
# P3 (Pressure Data) - Parser
# ============================================================

import csv

def clean_p3_name(name):
    """Remove tree characters and whitespace from P3 point name"""
    if not isinstance(name, str):
        return name
    # Remove tree characters: ├ └ │ ─
    name = name.replace('├', '').replace('└', '').replace('│', '').replace('─', '')
    return name.strip()


def process_p3_files(p3_dir):
    """Process all P3 pressure data files.
    Returns: {year: {month_key: {branch: [list of P3 points]}}}
    month_key format: "YY-MM" e.g. "69-03"
    Each P3 point: {n: name, p: avgPrev, a: avgDay, h: [24 hourly values]}
    """
    result = {}

    if not os.path.isdir(p3_dir):
        return result

    # Scan year subfolders
    for year_folder in sorted(os.listdir(p3_dir)):
        year_path = os.path.join(p3_dir, year_folder)
        if not os.path.isdir(year_path):
            continue

        # Process all .xlsx and .csv files in year folder
        files = sorted(glob.glob(os.path.join(year_path, '*.xlsx')) +
                      glob.glob(os.path.join(year_path, '*.csv')))

        for filepath in files:
            fname = os.path.basename(filepath)

            # Parse filename: {branch}_{YY}-{MM}.xlsx or .csv
            # e.g. "ชลบุรี_69-03.xlsx" -> branch="ชลบุรี", month="69-03"
            match = re.match(r'^(.+?)_((\d{2})-(\d{2}))\.(xlsx|csv)$', fname)
            if not match:
                continue

            branch = match.group(1)
            month_key = match.group(2)  # "69-03"
            file_ext = match.group(5).lower()  # "xlsx" or "csv"

            try:
                if file_ext == 'xlsx':
                    p3_points = _parse_p3_xlsx(filepath)
                else:  # csv
                    p3_points = _parse_p3_csv(filepath)

                if p3_points:
                    # Initialize nested structure
                    if year_folder not in result:
                        result[year_folder] = {}
                    if month_key not in result[year_folder]:
                        result[year_folder][month_key] = {}

                    result[year_folder][month_key][branch] = p3_points
            except Exception as e:
                # Silently skip files with parsing errors
                pass

    return result


def _parse_p3_xlsx(filepath):
    """Parse P3 data from .xlsx file.
    Returns: [list of P3 points with {n, p, a, h}]
    """
    parser = XLSXParser(filepath)
    parser.parse()

    p3_points = []

    # Process first sheet (usually only one)
    for si in range(parser.get_sheet_count()):
        sdata = parser.get_sheet_data(si)
        if not sdata:
            continue

        # Find header row: row with "พื้นที่", "เฉลี่ยเดือน ก.พ.", etc.
        header_row = None
        for r in sorted(sdata.keys()):
            row = sdata[r]
            if 0 in row and isinstance(row[0], str) and 'พื้นที่' in row[0]:
                header_row = r
                break

        if header_row is None:
            continue

        # Extract P3 data rows (rows after header where col 0 contains "P3")
        for r in sorted(sdata.keys()):
            if r <= header_row:
                continue

            row = sdata[r]
            if 0 not in row:
                continue

            name = row.get(0, '')
            if not isinstance(name, str) or 'P3' not in name:
                continue

            # Clean name
            name = clean_p3_name(name)

            # Extract values
            avg_prev = row.get(1)
            avg_day = row.get(2)

            # Convert "-" or non-numeric to None
            avg_prev = _convert_p3_value(avg_prev)
            avg_day = _convert_p3_value(avg_day)

            # Extract 24 hourly values (columns 3-26)
            hourly = []
            for col in range(3, 27):
                val = _convert_p3_value(row.get(col))
                hourly.append(val)

            p3_points.append({
                'n': name,
                'p': avg_prev,
                'a': avg_day,
                'h': hourly
            })

    return p3_points


def _parse_p3_csv(filepath):
    """Parse P3 data from .csv file.
    Returns: [list of P3 points with {n, p, a, h}]
    """
    p3_points = []

    try:
        with open(filepath, 'r', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            rows = list(reader)

        if not rows:
            return p3_points

        # First row is header (skip)
        # Find P3 rows (where col 0 contains "P3")
        for row in rows[1:]:
            if not row or len(row) < 3:
                continue

            name = row[0].strip() if row[0] else ''
            if 'P3' not in name:
                continue

            # Clean name
            name = clean_p3_name(name)

            # Extract values: col 1 = avg_prev, col 2 = avg_day, col 3-26 = hourly
            avg_prev = _convert_p3_value(row[1] if len(row) > 1 else None)
            avg_day = _convert_p3_value(row[2] if len(row) > 2 else None)

            # Extract 24 hourly values
            hourly = []
            for col_idx in range(3, 27):
                val = _convert_p3_value(row[col_idx] if len(row) > col_idx else None)
                hourly.append(val)

            p3_points.append({
                'n': name,
                'p': avg_prev,
                'a': avg_day,
                'h': hourly
            })

    except Exception:
        pass

    return p3_points


def _convert_p3_value(val):
    """Convert P3 value to float or None.
    "-" or empty string -> None
    numbers -> float
    """
    if val is None or val == '' or val == '-':
        return None

    if isinstance(val, (int, float)):
        return float(val)

    try:
        return float(str(val).strip())
    except (ValueError, AttributeError):
        return None


def build_p3_embedded_data(p3_data):
    """Convert P3 data to compact JSON for embedding.
    Returns JSON string for: const P3={...}
    Input structure: {year: {month_key: {branch: [P3 points]}}}
    Output: Compact JSON preserving the structure
    """
    if not p3_data:
        return '{}'

    compact = {}
    for year_str in sorted(p3_data.keys()):
        compact[year_str] = {}
        months = p3_data[year_str]
        for month_key in sorted(months.keys()):
            compact[year_str][month_key] = {}
            branches = months[month_key]
            for branch, points in branches.items():
                # Compact each point: {n, p, a, h}
                compact[year_str][month_key][branch] = points

    return json.dumps(compact, ensure_ascii=False, separators=(',', ':'))


def process_xls_file(filepath):
    """Process one .xls file and return {sheet_name: {rows: [...]}}"""
    print(f"  กำลังอ่าน: {os.path.basename(filepath)} ...")

    # Detect format: OLE2 (.xls) vs OOXML (.xlsx in disguise)
    with open(filepath, 'rb') as f:
        magic = f.read(4)

    if magic[:2] == b'PK':
        # ZIP-based OOXML format
        parser = XLSXParser(filepath)
    else:
        # OLE2 BIFF8 format
        parser = XLSParser(filepath)
    parser.parse()

    result = {}
    # Skip first sheet if it's "เป้าหมาย" (target/goal sheet, not data)
    skip_sheets = {'เป้าหมาย'}

    for si in range(parser.get_sheet_count()):
        sname = parser.get_sheet_name(si)
        if sname in skip_sheets:
            continue

        sd = parser.get_sheet_data(si)
        if not sd:
            continue

        header_row = find_month_header_row(sd)
        if header_row is None:
            continue

        month_cols = find_month_columns(sd, header_row)
        if all(c is None for c in month_cols):
            continue

        total_col = find_total_column(sd, header_row)
        rows = extract_sheet_data(sd, header_row, month_cols, total_col)

        if rows:
            # Rename sheet to match expected format
            result[sname] = {'rows': rows}

    return result


# ============================================================
# Dashboard HTML Builder
# ============================================================

def build_embedded_data(all_data):
    """Convert data to compact JSON for embedding in HTML"""
    compact = {}
    for year_str in sorted(all_data.keys()):
        sheets = all_data[year_str]
        compact[year_str] = {}
        for sname, sinfo in sheets.items():
            compact[year_str][sname] = []
            for r in sinfo['rows']:
                compact[year_str][sname].append({
                    'l': r['label'],
                    'u': r['unit'],
                    'm': r['monthly'],
                    't': r['total'],
                    'ty': r.get('target_year'),
                    'tm': r.get('target_month')
                })
    return json.dumps(compact, ensure_ascii=False, separators=(',', ':'))


def build_dashboard(all_data, template_path, output_path, rl_data=None, eu_data=None, mnf_data=None, kpi_data=None, p3_data=None):
    """Build index.html by replacing embedded data in template"""
    with open(template_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    # Find the line with "const D="
    data_line_idx = None
    for i, line in enumerate(lines):
        if line.strip().startswith('const D='):
            data_line_idx = i
            break

    if data_line_idx is None:
        print("  ❌ ไม่พบ 'const D=' ใน template!")
        return False

    # Build new data line
    compact_json = build_embedded_data(all_data)
    new_data_line = f'const D={compact_json};\n'

    # Build Real Leak data line
    rl_line_idx = None
    for i, line in enumerate(lines):
        if line.strip().startswith('const RL='):
            rl_line_idx = i
            break

    rl_json = build_rl_embedded_data(rl_data) if rl_data else '{}'
    new_rl_line = f'const RL={rl_json};\n'

    # Build EU data line
    eu_line_idx = None
    for i, line in enumerate(lines):
        if line.strip().startswith('const EU='):
            eu_line_idx = i
            break

    eu_json = build_eu_embedded_data(eu_data) if eu_data else '{}'
    new_eu_line = f'const EU={eu_json};\n'

    # Build MNF data line
    mnf_line_idx = None
    for i, line in enumerate(lines):
        if line.strip().startswith('const MNF='):
            mnf_line_idx = i
            break

    mnf_json = build_mnf_embedded_data(mnf_data) if mnf_data else '{}'
    new_mnf_line = f'const MNF={mnf_json};\n'

    # Build KPI data line
    kpi_line_idx = None
    for i, line in enumerate(lines):
        if line.strip().startswith('const KPI='):
            kpi_line_idx = i
            break

    kpi_json = build_kpi_embedded_data(kpi_data) if kpi_data else '{}'
    new_kpi_line = f'const KPI={kpi_json};\n'

    # Build P3 data line
    p3_line_idx = None
    for i, line in enumerate(lines):
        if line.strip().startswith('const P3='):
            p3_line_idx = i
            break

    p3_json = build_p3_embedded_data(p3_data) if p3_data else '{}'
    new_p3_line = f'const P3={p3_json};\n'

    # Also update YC (year colors) to include all years
    yc_start = None
    yc_end = None
    for i, line in enumerate(lines):
        if 'const YC={' in line:
            yc_start = i
        if yc_start is not None and '};' in line and i > yc_start:
            yc_end = i
            break

    # Generate year colors - collect unique years first to avoid duplicate keys
    all_years = sorted(all_data.keys(), key=int)
    colors = [
        ('rgba(59,130,246,0.15)', '#3b82f6'),    # blue
        ('rgba(239,68,68,0.15)', '#ef4444'),      # red
        ('rgba(34,197,94,0.15)', '#22c55e'),      # green
        ('rgba(168,85,247,0.15)', '#a855f7'),     # purple
        ('rgba(249,115,22,0.15)', '#f97316'),     # orange
        ('rgba(6,182,212,0.15)', '#06b6d4'),      # cyan
        ('rgba(236,72,153,0.15)', '#ec4899'),     # pink
        ('rgba(202,138,4,0.15)', '#ca8a04'),      # gold
        ('rgba(99,102,241,0.15)', '#6366f1'),     # indigo
        ('rgba(20,184,166,0.15)', '#14b8a6'),     # teal
        ('rgba(244,63,94,0.15)', '#f43f5e'),      # rose
        ('rgba(139,92,246,0.15)', '#8b5cf6'),     # violet
    ]

    # Collect all unique years (fiscal + calendar + future)
    unique_years = set()
    for year_str in all_years:
        unique_years.add(int(year_str))
        unique_years.add(int(year_str) - 1)  # calendar year
    for extra in range(1, 4):
        unique_years.add(int(all_years[-1]) + extra)
    sorted_years = sorted(unique_years)

    yc_lines = ['const YC={\n']
    for idx, yr in enumerate(sorted_years):
        ci = idx % len(colors)
        bg, border = colors[ci]
        yc_lines.append(f"    {yr}:{{bg:'{bg}',border:'{border}'}},\n")
    yc_lines.append('};\n')

    # Rebuild file
    new_lines = []
    # Lines before data
    new_lines.extend(lines[:data_line_idx])
    # New data line
    new_lines.append(new_data_line)

    # Handle RL data line
    remaining_start = data_line_idx + 1
    if rl_line_idx is not None:
        # Replace existing RL line
        new_lines.extend(lines[data_line_idx + 1:rl_line_idx])
        new_lines.append(new_rl_line)
        remaining_start = rl_line_idx + 1
    else:
        # Insert RL line right after D line
        new_lines.append(new_rl_line)

    # Handle EU data line
    if eu_line_idx is not None:
        new_lines.extend(lines[remaining_start:eu_line_idx])
        new_lines.append(new_eu_line)
        remaining_start = eu_line_idx + 1
    else:
        # Insert EU line right after RL line
        new_lines.append(new_eu_line)

    # Handle MNF data line
    if mnf_line_idx is not None:
        new_lines.extend(lines[remaining_start:mnf_line_idx])
        new_lines.append(new_mnf_line)
        remaining_start = mnf_line_idx + 1
    else:
        # Insert MNF line right after EU line
        new_lines.append(new_mnf_line)

    # Handle KPI data line
    if kpi_line_idx is not None:
        new_lines.extend(lines[remaining_start:kpi_line_idx])
        new_lines.append(new_kpi_line)
        remaining_start = kpi_line_idx + 1
    else:
        # Insert KPI line right after MNF line
        new_lines.append(new_kpi_line)

    # Handle P3 data line
    if p3_line_idx is not None:
        new_lines.extend(lines[remaining_start:p3_line_idx])
        new_lines.append(new_p3_line)
        remaining_start = p3_line_idx + 1
    else:
        # Insert P3 line right after KPI line
        new_lines.append(new_p3_line)

    # Lines between data/RL/EU/MNF/KPI/P3 and YC (or after data if no YC change needed)
    if yc_start is not None and yc_end is not None:
        new_lines.extend(lines[remaining_start:yc_start])
        new_lines.extend(yc_lines)
        new_lines.extend(lines[yc_end + 1:])
    else:
        new_lines.extend(lines[remaining_start:])

    with open(output_path, 'w', encoding='utf-8') as f:
        f.writelines(new_lines)

    return True


# ============================================================
# Main
# ============================================================

def main():
    # Determine paths
    script_dir = os.path.dirname(os.path.abspath(__file__))
    xls_dir = os.path.join(script_dir, 'ข้อมูลดิบ', 'OIS')
    dashboard_path = os.path.join(script_dir, 'index.html')
    data_json_path = os.path.join(script_dir, 'data.json')

    print("=" * 60)
    print("  🔨 สร้าง Dashboard น้ำสูญเสีย - กปภ.ข.1")
    print("=" * 60)

    # Find all .xls files
    if not os.path.isdir(xls_dir):
        print(f"\n❌ ไม่พบโฟลเดอร์: {xls_dir}")
        print("  กรุณาวางไฟล์ .xls ไว้ในโฟลเดอร์ ข้อมูลดิบ/OIS/")
        input("\nกด Enter เพื่อปิด...")
        return

    xls_files = sorted(glob.glob(os.path.join(xls_dir, '*.xls')) +
                       glob.glob(os.path.join(xls_dir, '*.xlsx')))
    if not xls_files:
        print(f"\n❌ ไม่พบไฟล์ .xls ในโฟลเดอร์: {xls_dir}")
        input("\nกด Enter เพื่อปิด...")
        return

    print(f"\n📂 พบไฟล์ .xls {len(xls_files)} ไฟล์:")
    for f in xls_files:
        print(f"   • {os.path.basename(f)}")

    # Check template exists
    if not os.path.isfile(dashboard_path):
        print(f"\n❌ ไม่พบ index.html template ที่: {dashboard_path}")
        input("\nกด Enter เพื่อปิด...")
        return

    # Process each file
    print(f"\n📊 กำลังอ่านข้อมูล...")
    all_data = {}
    for xls_file in xls_files:
        year_str = extract_year_from_filename(xls_file)
        if year_str is None:
            print(f"  ⚠️  ข้าม {os.path.basename(xls_file)} (ไม่พบปี พ.ศ. ในชื่อไฟล์)")
            continue

        try:
            sheet_data = process_xls_file(xls_file)
            if sheet_data:
                all_data[year_str] = sheet_data
                print(f"     ✅ ปี {year_str}: {len(sheet_data)} sheets")
            else:
                print(f"     ⚠️  ปี {year_str}: ไม่พบข้อมูล")
        except Exception as e:
            print(f"     ❌ ปี {year_str}: เกิดข้อผิดพลาด - {e}")

    if not all_data:
        print("\n❌ ไม่สามารถอ่านข้อมูลได้เลย")
        input("\nกด Enter เพื่อปิด...")
        return

    # Normalize labels across years
    print(f"\n🔧 ปรับชื่อรายการให้ตรงกันทุกปี...")
    normalize_labels(all_data)

    # Fix trailing zeros
    print(f"\n🔧 แก้ไขข้อมูลเดือนที่ยังไม่มีข้อมูล...")
    fix_trailing_zeros(all_data)

    # ============================================================
    # Process Real Leak files
    # ============================================================
    rl_dir = os.path.join(script_dir, 'ข้อมูลดิบ', 'Real Leak')
    rl_data = {}
    if os.path.isdir(rl_dir):
        rl_files = sorted(glob.glob(os.path.join(rl_dir, 'RL-*.xlsx')) +
                          glob.glob(os.path.join(rl_dir, 'RL-*.xls')))
        if rl_files:
            print(f"\n📂 พบไฟล์ Real Leak {len(rl_files)} ไฟล์:")
            for f in rl_files:
                print(f"   • {os.path.basename(f)}")
            print(f"\n📊 กำลังอ่านข้อมูล Real Leak...")
            for rl_file in rl_files:
                try:
                    year_str, branch_data = process_rl_file(rl_file)
                    if year_str and branch_data:
                        rl_data[year_str] = branch_data
                        # Count branches with data
                        count = sum(1 for b, m in branch_data.items()
                                    if any(v is not None for v in m['rate']))
                        print(f"     ✅ ปี {year_str}: {count} สาขามีข้อมูล")
                    else:
                        print(f"     ⚠️  {os.path.basename(rl_file)}: ไม่พบข้อมูล")
                except Exception as e:
                    print(f"     ❌ {os.path.basename(rl_file)}: เกิดข้อผิดพลาด - {e}")
        else:
            print(f"\n📂 ไม่พบไฟล์ Real Leak ในโฟลเดอร์: {rl_dir}")
    else:
        print(f"\n📂 ไม่พบโฟลเดอร์ Real Leak (ข้ามไป)")

    # ============================================================
    # Process EU (Electric Unit / Water Sold) files
    # ============================================================
    eu_dir = os.path.join(script_dir, 'ข้อมูลดิบ', 'หน่วยไฟ')
    eu_data = {}
    if os.path.isdir(eu_dir):
        eu_files = sorted(glob.glob(os.path.join(eu_dir, 'EU-*.xlsx')) +
                          glob.glob(os.path.join(eu_dir, 'EU-*.xls')))
        if eu_files:
            print(f"\n📂 พบไฟล์หน่วยไฟฟ้า {len(eu_files)} ไฟล์:")
            for f in eu_files:
                print(f"   • {os.path.basename(f)}")
            print(f"\n📊 กำลังอ่านข้อมูลหน่วยไฟฟ้า...")
            for eu_file in eu_files:
                try:
                    year_str, branch_data = process_eu_file(eu_file)
                    if year_str and branch_data:
                        eu_data[year_str] = branch_data
                        count = sum(1 for b, vals in branch_data.items()
                                    if b != '__regional__' and any(v is not None for v in vals))
                        print(f"     ✅ ปี {year_str}: {count} สาขามีข้อมูล")
                    else:
                        print(f"     ⚠️  {os.path.basename(eu_file)}: ไม่พบข้อมูล")
                except Exception as e:
                    print(f"     ❌ {os.path.basename(eu_file)}: เกิดข้อผิดพลาด - {e}")
        else:
            print(f"\n📂 ไม่พบไฟล์หน่วยไฟฟ้าในโฟลเดอร์: {eu_dir}")
    else:
        print(f"\n📂 ไม่พบโฟลเดอร์หน่วยไฟฟ้า (ข้ามไป)")

    # ============================================================
    # Process MNF (Minimum Night Flow) files
    # ============================================================
    mnf_dir = os.path.join(script_dir, 'ข้อมูลดิบ', 'MNF')
    mnf_data = {}
    if os.path.isdir(mnf_dir):
        mnf_files = sorted(glob.glob(os.path.join(mnf_dir, 'MNF-*.xlsx')) +
                           glob.glob(os.path.join(mnf_dir, 'MNF-*.xls')))
        if mnf_files:
            print(f"\n📂 พบไฟล์ MNF {len(mnf_files)} ไฟล์:")
            for f in mnf_files:
                print(f"   • {os.path.basename(f)}")
            print(f"\n📊 กำลังอ่านข้อมูล MNF...")
            for mnf_file in mnf_files:
                try:
                    year_str, branch_data = process_mnf_file(mnf_file)
                    if year_str and branch_data:
                        mnf_data[year_str] = branch_data
                        count = sum(1 for b, m in branch_data.items()
                                    if b != '__regional__' and any(v is not None for v in m['actual']))
                        print(f"     ✅ ปี {year_str}: {count} สาขามีข้อมูล")
                    else:
                        print(f"     ⚠️  {os.path.basename(mnf_file)}: ไม่พบข้อมูล")
                except Exception as e:
                    print(f"     ❌ {os.path.basename(mnf_file)}: เกิดข้อผิดพลาด - {e}")
        else:
            print(f"\n📂 ไม่พบไฟล์ MNF ในโฟลเดอร์: {mnf_dir}")
    else:
        print(f"\n📂 ไม่พบโฟลเดอร์ MNF (ข้ามไป)")

    # ============================================================
    # Process KPI Leak (เกณฑ์วัดน้ำสูญเสีย) files
    # ============================================================
    kpi_dir = os.path.join(script_dir, 'ข้อมูลดิบ', 'เกณฑ์วัดน้ำสูญเสีย')
    kpi_dir2 = os.path.join(script_dir, 'ข้อมูลดิบ', 'เกณฑ์ชี้วัด')
    kpi_data = {}
    for kd in [kpi_dir, kpi_dir2]:
        if os.path.isdir(kd):
            kpi_files = sorted(glob.glob(os.path.join(kd, 'KPI_Lea*-*.xlsx')) +
                               glob.glob(os.path.join(kd, 'KPI_Lea*-*.xls')))
            if kpi_files:
                print(f"\n📂 พบไฟล์ KPI Leak {len(kpi_files)} ไฟล์:")
                for f in kpi_files:
                    print(f"   • {os.path.basename(f)}")
                print(f"\n📊 กำลังอ่านข้อมูล KPI Leak...")
                for kpi_file in kpi_files:
                    # Skip temp files
                    if os.path.basename(kpi_file).startswith('~$'):
                        continue
                    try:
                        year_str, branch_data = process_kpi_file(kpi_file)
                        if year_str and branch_data:
                            kpi_data[year_str] = branch_data
                            count = sum(1 for b in branch_data if b != '__regional__')
                            print(f"     ✅ ปี {year_str}: {count} สาขา")
                        else:
                            print(f"     ⚠️  {os.path.basename(kpi_file)}: ไม่พบข้อมูล")
                    except Exception as e:
                        print(f"     ❌ {os.path.basename(kpi_file)}: เกิดข้อผิดพลาด - {e}")

    # ============================================================
    # Process P3 (Pressure Data) files
    # ============================================================
    p3_dir = os.path.join(script_dir, 'ข้อมูลดิบ', 'P3')
    p3_data = {}
    if os.path.isdir(p3_dir):
        # Just process all files directly via process_p3_files
        p3_data = process_p3_files(p3_dir)
        if p3_data:
            total_points = sum(sum(len(points) for points in branch_dict.values())
                              for month_dict in p3_data.values()
                              for branch_dict in month_dict.values())
            print(f"\n📂 พบโฟลเดอร์ P3:")
            for year_str in sorted(p3_data.keys()):
                months = p3_data[year_str]
                month_count = len(months)
                points_count = sum(len(points) for month_dict in months.values()
                                  for points in month_dict.values())
                print(f"     ✅ ปี {year_str}: {month_count} เดือน, {points_count} จุด P3")
        else:
            print(f"\n📂 โฟลเดอร์ P3 ไม่มีไฟล์ที่สามารถอ่านได้")
    else:
        print(f"\n📂 ไม่พบโฟลเดอร์ P3 (ข้ามไป)")

    # Save data.json
    print(f"\n💾 บันทึก data.json...")
    full_data = {}
    for year_str, sheets in all_data.items():
        full_data[year_str] = sheets
    with open(data_json_path, 'w', encoding='utf-8') as f:
        json.dump(full_data, f, ensure_ascii=False, indent=2)
    fsize = os.path.getsize(data_json_path)
    print(f"   ✅ data.json ({fsize:,} bytes)")

    # Build dashboard
    print(f"\n🏗️  สร้าง index.html...")
    ok = build_dashboard(all_data, dashboard_path, dashboard_path, rl_data=rl_data, eu_data=eu_data, mnf_data=mnf_data, kpi_data=kpi_data, p3_data=p3_data)
    if ok:
        fsize = os.path.getsize(dashboard_path)
        print(f"   ✅ index.html ({fsize:,} bytes)")
    else:
        print(f"   ❌ สร้าง index.html ไม่สำเร็จ")

    # Summary
    print(f"\n{'=' * 60}")
    print(f"  ✅ เสร็จสิ้น!")
    print(f"  📅 ปีข้อมูล OIS: {', '.join(sorted(all_data.keys()))}")
    if rl_data:
        print(f"  📅 ปีข้อมูล Real Leak: {', '.join(sorted(rl_data.keys()))}")
    if eu_data:
        print(f"  📅 ปีข้อมูลหน่วยไฟฟ้า: {', '.join(sorted(eu_data.keys()))}")
    if mnf_data:
        print(f"  📅 ปีข้อมูล MNF: {', '.join(sorted(mnf_data.keys()))}")
    if kpi_data:
        print(f"  📅 ปีข้อมูล KPI Leak: {', '.join(sorted(kpi_data.keys()))}")
    if p3_data:
        print(f"  📅 ปีข้อมูล P3: {', '.join(sorted(p3_data.keys()))}")
    total_sheets = sum(len(s) for s in all_data.values())
    print(f"  📋 รวม {total_sheets} sheets จาก {len(all_data)} ปี")
    print(f"  📄 เปิด index.html ในเบราว์เซอร์เพื่อดูผลลัพธ์")
    print(f"{'=' * 60}")

    # On Windows, pause before closing
    if sys.platform == 'win32':
        input("\nกด Enter เพื่อปิด...")


if __name__ == '__main__':
    main()
