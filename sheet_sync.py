#!/usr/bin/env python3
"""
Sheet Sync - แทรกข้อมูล orders จาก FB ลง Google Sheet เรียงตามวันที่
Usage:
    python sheet_sync.py --json output/orders_xxx.json --sheet "มี.ค.69"
    python sheet_sync.py --json output/orders_xxx.json --sheet "ก.พ.69" --dry-run
"""
import argparse
import json
import os
import gspread
from google.oauth2.service_account import Credentials
from pathlib import Path
from datetime import datetime
import re
import time

# ============================================================
# Config
# ============================================================
CREDENTIALS_FILE = Path(__file__).parent / "credentials.json"
SPREADSHEET_NAME = "บัญชี HDG 69"
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

# Header rows (ไม่แทรกข้อมูลก่อน row นี้)
DATA_START_ROW = 8  # ข้อมูลเริ่มที่ row 8

# ============================================================
# Helpers
# ============================================================

def parse_date(date_str):
    """Parse date string เช่น '7/3/2026' เป็น datetime"""
    try:
        parts = date_str.strip().split('/')
        if len(parts) == 3:
            d, m, y = int(parts[0]), int(parts[1]), int(parts[2])
            return datetime(y, m, d)
    except:
        pass
    return None


def connect_sheet(sheet_name):
    """เชื่อมต่อ Google Sheet — รองรับทั้ง local file และ env var credentials"""
    creds_json = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    if creds_json:
        creds_info = json.loads(creds_json)
        creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file(str(CREDENTIALS_FILE), scopes=SCOPES)
    
    gc = gspread.authorize(creds)
    sh = gc.open(SPREADSHEET_NAME)
    ws = sh.worksheet(sheet_name)
    return ws


def get_existing_dates(ws):
    """อ่านวันที่ที่มีอยู่ใน column B+C เริ่มจาก Row 8 (ต่อเนื่อง)"""
    col_b = ws.col_values(2)  # Column B = วัน/เดือน/ปี
    col_c = ws.col_values(3)  # Column C = รายการ
    
    # คำที่บอกว่าเป็น footer section (ไม่ใช่ข้อมูล orders)
    FOOTER_KEYWORDS = ['รวม', 'สรุป', 'คำอธิบาย', 'รายรับ', 'รายจ่าย', 'กำไร', 'หมายเหตุ', 'ช่อง']
    
    date_rows = []  # [(row_number, date_str, datetime_obj)]
    last_data_row = DATA_START_ROW - 1  # เริ่มก่อน row 8
    empty_count = 0
    
    for i, val in enumerate(col_b):
        row_num = i + 1
        if row_num < DATA_START_ROW:
            continue
        
        b_text = val.strip()
        c_text = col_c[i].strip() if i < len(col_c) else ''
        
        # ตรวจ footer keywords — "รวม" ต้อง exact match (ไม่ match "รวมแชท")
        if b_text == 'รวม' or c_text == 'รวม':
            break
        if any(b_text.startswith(kw) for kw in FOOTER_KEYWORDS if kw != 'รวม'):
            break
        
        has_b = b_text != ''
        has_c = c_text != ''
        
        if has_b or has_c:
            empty_count = 0
            last_data_row = row_num
            
            date_obj = parse_date(b_text)
            if date_obj:
                date_rows.append((row_num, val.strip(), date_obj))
        else:
            empty_count += 1
            # ถ้าเจอ empty rows ต่อกัน > 1 row = ถือว่าจบข้อมูล
            if empty_count > 1:
                break
    
    return date_rows, last_data_row


def find_insert_position(date_rows, order_date, last_data_row):
    """
    หาตำแหน่ง row ที่ควรแทรก order ใหม่ (เรียงวันที่จากน้อยไปมาก)
    Returns: row number to insert BEFORE
    """
    if not date_rows:
        return DATA_START_ROW
    
    for row_num, date_str, date_obj in date_rows:
        if order_date < date_obj:
            return row_num
    
    # ถ้าวันที่ใหม่มากกว่าทุก row = แทรกหลังข้อมูลสุดท้าย
    return last_data_row + 1


def clean_product_name(raw_name):
    """
    Normalize ชื่อสินค้าให้เป็น format: [รุ่น] ลาย [ลาย]
    ตัวอย่าง:
      "rf24-105 f2.8 ลาย mamba red" → "rf24-105 f2.8 ลาย mamba red" (OK)
      "Sony A6600 ลาย cmd ลด 700 บาท" → "Sony A6600 ลาย cmd"
      "Ptbk. และในส่วนที่เป็นกริบจับใช้เป็นLtbk.ครับ Nikon Z6iiiครับ" → "Nikon Z6iii ลาย ptbk/ltbk"
    """
    name = raw_name.strip()
    
    # ลบ suffix ที่ไม่จำเป็น (เฉพาะคำลงท้าย)
    name = re.sub(r'ครับ\s*$', '', name)
    name = re.sub(r'ค่ะ\s*$', '', name)
    name = name.strip()
    
    return name


def split_multi_items(product_name, total_price=''):
    """
    แยกรายการที่มี 'กับ' หรือ 'และ' เป็นหลายชิ้น
    ตัวอย่าง:
      "กล้อง Sony a7iii กับ เลนส์ Sony24-70 2.8 II ลาย mbbk"
      → ["Sony a7iii ลาย mbbk", "Sony 24-70 2.8 II ลาย mbbk"]
    Returns: list of (name, price) tuples
    """
    # มองหา pattern แยกด้วย "กับ" หรือ "และ"
    # Pattern: "กล้อง X กับ เลนส์ Y ลาย Z"
    split_match = re.match(
        r'(?:กล้อง\s*)?(.+?)\s+(?:กับ|และ)\s+(?:เลนส์\s*)?(.+)',
        product_name, re.IGNORECASE
    )
    
    if split_match:
        part1 = split_match.group(1).strip()
        part2 = split_match.group(2).strip()
        
        # หา "ลาย xxx" ที่อยู่ท้ายสุด (ใช้กับทั้ง 2 ชิ้น)
        pattern_match = re.search(r'ลาย\s+(\S+)', part2)
        pattern_name = pattern_match.group(1) if pattern_match else ''
        
        # ถ้า part1 ไม่มี "ลาย" ให้เพิ่ม pattern จาก part2
        if 'ลาย' not in part1 and pattern_name:
            part1 = f"{part1} ลาย {pattern_name}"
        
        # ราคารวม → ไม่แยกได้ ให้ user แก้เอง
        return [(part1, ''), (part2, '')]
    
    return [(product_name, total_price)]


def prepare_order_rows(order):
    """เตรียมข้อมูล rows สำหรับ order 1 รายการ"""
    rows = []
    products = order.get('products', [])
    date_str = order.get('date', '')
    customer = order.get('customer', '')
    fb_name = f"fb.{customer}" if customer else ''
    
    # Row template: A=empty, B=date, C=item, D=income, E-N=empty, O=customer
    # Column O = index 14 (0-based), need 15 elements
    
    all_items = []  # [(name, price)]
    
    if order.get('split_manually') and products:
        # ✂️ Order ที่ user แยกแล้วจาก Split Modal — ใช้ products โดยตรง (ไม่ split auto ซ้ำ)
        for product in products:
            name = clean_product_name(product.get('name', ''))
            price = product.get('price', '')
            all_items.append((name, price))
    elif products:
        for product in products:
            raw_name = product.get('name', '')
            price = product.get('price', '')
            
            # Clean ชื่อสินค้า
            cleaned = clean_product_name(raw_name)
            
            # แยกรายการ 2+ ชิ้น
            split_items = split_multi_items(cleaned, price)
            all_items.extend(split_items)
    else:
        price = order.get('price', '')
        all_items.append(('(ไม่ทราบชื่อสินค้า)', price))
    
    # สร้าง rows
    for i, (name, price) in enumerate(all_items):
        row = [''] * 15  # A through O
        row[1] = date_str if i == 0 else ''  # B - วันที่ (เฉพาะแถวแรก)
        row[2] = f"Skin {name}"  # C - รายการ
        row[3] = f"  {price} " if price else ''  # D - รายรับ
        row[14] = fb_name if i == 0 else ''  # O - ชื่อลูกค้า (เฉพาะแถวแรก)
        rows.append(row)
    
    # เพิ่มแถว "ค่าส่งพัสดุ" (ให้ user ไปใส่ราคาเอง)
    shipping_row = [''] * 15
    shipping_row[2] = 'ค่าส่งพัสดุ'  # C
    rows.append(shipping_row)
    
    return rows


# ============================================================
# Main
# ============================================================

def sync_orders_to_sheet(json_file, sheet_name, dry_run=False):
    """แทรกข้อมูล orders ลง Google Sheet"""
    print("\n" + "=" * 60)
    print("📊 Sheet Sync - แทรก orders ลง Google Sheet")
    print("=" * 60)
    
    # อ่าน orders JSON
    with open(json_file, 'r', encoding='utf-8') as f:
        orders = json.load(f)
    
    print(f"\n📦 อ่านได้ {len(orders)} orders จาก {json_file}")
    
    # เชื่อมต่อ Sheet
    print(f"📄 กำลังเชื่อมต่อ Sheet: {SPREADSHEET_NAME} → {sheet_name}")
    ws = connect_sheet(sheet_name)
    
    # อ่านวันที่ที่มีอยู่
    date_rows, last_data_row = get_existing_dates(ws)
    print(f"   พบวันที่ {len(date_rows)} รายการ (ข้อมูลสุดท้ายที่ row {last_data_row})")
    
    if date_rows:
        print(f"   วันที่แรก: {date_rows[0][1]} (row {date_rows[0][0]})")
        print(f"   วันที่สุดท้าย: {date_rows[-1][1]} (row {date_rows[-1][0]})")
    
    # ตรวจสอบ orders ที่ยังไม่ได้ลง (เช็คจากชื่อสินค้า)
    existing_col_c = ws.col_values(3)  # Column C = รายการ
    existing_items = set()
    for item in existing_col_c:
        if item:
            existing_items.add(item.strip().lower())
    
    # เตรียมข้อมูลสำหรับแทรก
    new_orders = []
    skipped = []
    
    for order in orders:
        products = order.get('products', [])
        date_str = order.get('date', '')
        
        if products:
            # เช็คว่าสินค้าแรกมีอยู่ใน sheet แล้วหรือไม่
            first_product = f"Skin {products[0].get('name', '')}".strip().lower()
            if first_product in existing_items:
                skipped.append(order)
                continue
        
        new_orders.append(order)
    
    print(f"\n📋 สรุป:")
    print(f"   ✅ orders ใหม่ที่ต้องแทรก: {len(new_orders)}")
    print(f"   ⏭️ มีอยู่แล้ว (skip): {len(skipped)}")
    
    if not new_orders:
        print("\n✅ ไม่มี orders ใหม่ที่ต้องแทรก!")
        return
    
    # จัดเรียงตามวันที่
    def sort_key(o):
        d = parse_date(o.get('date', ''))
        return d if d else datetime.min
    
    new_orders.sort(key=sort_key)
    
    # แสดงข้อมูลที่จะแทรก
    print(f"\n📝 ข้อมูลที่จะแทรก:")
    for o in new_orders:
        products = o.get('products', [])
        names = [p.get('name', '') for p in products] if products else ['(ไม่ทราบ)']
        print(f"   {o.get('date', 'N/A')} | {', '.join(names)} | {o.get('customer', '')}")
    
    if dry_run:
        print(f"\n⚠️ DRY RUN - ไม่ได้แทรกจริง")
        return
    
    # แทรกข้อมูลทีละ order (จากวันที่มากไปน้อย เพื่อไม่ให้ row shift กระทบ)
    new_orders_reversed = list(reversed(new_orders))
    
    print(f"\n🔄 กำลังแทรก {len(new_orders)} orders...")
    
    # Re-read date positions before each insert
    for idx, order in enumerate(new_orders_reversed):
        order_date = parse_date(order.get('date', ''))
        if not order_date:
            print(f"   ⚠️ Skip (ไม่มีวันที่): {order.get('customer', '')}")
            continue
        
        # อ่านวันที่ใหม่ทุกรอบ (เพราะ row shift)
        date_rows, last_data_row = get_existing_dates(ws)
        insert_row = find_insert_position(date_rows, order_date, last_data_row)
        
        # เตรียม rows
        order_rows = prepare_order_rows(order)
        
        products = order.get('products', [])
        first_name = products[0].get('name', '') if products else '(ไม่ทราบ)'
        split_tag = ' ✂️ [split manually]' if order.get('split_manually') else ''
        
        print(f"   [{idx+1}/{len(new_orders)}] Row {insert_row}: {order.get('date', '')} | {first_name}{split_tag}")
        
        # Insert rows
        for i, row_data in enumerate(order_rows):
            ws.insert_row(row_data, index=insert_row + i, value_input_option='USER_ENTERED')
            time.sleep(0.5)  # Rate limit
        
        # เติมสีเขียวอ่อนที่ช่องวันที่ (Column B) เพื่อให้ตรวจสอบได้ง่าย
        try:
            date_cell = gspread.utils.rowcol_to_a1(insert_row, 2)  # Column B
            ws.format(date_cell, {
                "backgroundColor": {
                    "red": 0.85,
                    "green": 1.0,
                    "blue": 0.85
                }
            })
        except Exception as e:
            print(f"      ⚠️ ไม่สามารถเติมสี: {e}")
        
        time.sleep(1)  # Extra pause between orders
    
    print(f"\n✅ แทรกเสร็จ! {len(new_orders)} orders (เซลล์วันที่เติมสีเขียวอ่อนไว้แล้ว)")
    print(f"📄 ตรวจสอบผลลัพธ์ที่ Sheet: {sheet_name}")


def main():
    parser = argparse.ArgumentParser(description='Sheet Sync - แทรก orders ลง Google Sheet')
    parser.add_argument('--json', required=True, help='Path ไฟล์ JSON ที่ export จาก bot')
    parser.add_argument('--sheet', required=True, help='ชื่อ worksheet เช่น "มี.ค.69"')
    parser.add_argument('--dry-run', action='store_true', help='แสดงผลแต่ไม่แทรกจริง')
    
    args = parser.parse_args()
    
    if not Path(args.json).exists():
        print(f"❌ ไม่เจอไฟล์: {args.json}")
        return
    
    sync_orders_to_sheet(args.json, args.sheet, dry_run=args.dry_run)


if __name__ == '__main__':
    main()
