#!/usr/bin/env python3
"""
HDG Orders Dashboard — Web UI สำหรับจัดการ orders
"""
import json
import os
import re
import subprocess
import sys
import threading
import time
from datetime import datetime
from pathlib import Path
from flask import Flask, render_template_string, jsonify, request

app = Flask(__name__)

BASE_DIR = Path(__file__).parent
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Detect cloud deployment (Render sets RENDER=true)
IS_CLOUD = bool(os.environ.get('RENDER') or os.environ.get('GOOGLE_CREDENTIALS_JSON'))


def get_gspread_client():
    """Get authenticated gspread client — supports both local file and env var credentials"""
    import gspread
    from google.oauth2.service_account import Credentials
    
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    
    # Try environment variable first (for cloud deployment)
    creds_json = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    if creds_json:
        import json as _json
        creds_info = _json.loads(creds_json)
        creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    else:
        # Fall back to local file
        CREDENTIALS_FILE = BASE_DIR / 'credentials.json'
        creds = Credentials.from_service_account_file(str(CREDENTIALS_FILE), scopes=SCOPES)
    
    return gspread.authorize(creds)


def save_orders_to_sheet(orders):
    """Save orders data to Google Sheets for persistent storage on cloud"""
    try:
        gc = get_gspread_client()
        spreadsheet = gc.open('บัญชี HDG 69')
        
        # Get or create orders_cache worksheet
        try:
            ws = spreadsheet.worksheet('orders_cache')
        except:
            ws = spreadsheet.add_worksheet(title='orders_cache', rows=5, cols=2)
        
        # Store as JSON string in cell A1, timestamp in B1
        orders_json = json.dumps(orders, ensure_ascii=False)
        ws.update('A1', [[orders_json, datetime.now().strftime('%Y-%m-%d %H:%M:%S')]])
        return True
    except Exception as e:
        print(f'❌ Error saving to sheet cache: {e}')
        return False


def load_orders_from_sheet():
    """Load orders data from Google Sheets cache"""
    try:
        gc = get_gspread_client()
        spreadsheet = gc.open('บัญชี HDG 69')
        ws = spreadsheet.worksheet('orders_cache')
        
        cell_value = ws.acell('A1').value
        if cell_value:
            return json.loads(cell_value)
        return []
    except Exception as e:
        print(f'❌ Error loading from sheet cache: {e}')
        return []

# Task status tracking
tasks = {}

THAI_MONTHS = {
    1: 'ม.ค.', 2: 'ก.พ.', 3: 'มี.ค.', 4: 'เม.ย.',
    5: 'พ.ค.', 6: 'มิ.ย.', 7: 'ก.ค.', 8: 'ส.ค.',
    9: 'ก.ย.', 10: 'ต.ค.', 11: 'พ.ย.', 12: 'ธ.ค.'
}

SHEET_NAMES = {
    1: 'ม.ค.69', 2: 'ก.พ.69', 3: 'มี.ค.69', 4: 'เม.ย.69',
    5: 'พ.ค.69', 6: 'มิ.ย.69', 7: 'ก.ค.69', 8: 'ส.ค.69',
    9: 'ก.ย.69', 10: 'ต.ค.69', 11: 'พ.ย.69', 12: 'ธ.ค.69'
}


def run_task(task_id, cmd, cwd=None):
    """Run a command in background and track status"""
    tasks[task_id] = {'status': 'running', 'output': '', 'started': datetime.now().isoformat()}
    try:
        env = os.environ.copy()
        env['PYTHONIOENCODING'] = 'utf-8'
        proc = subprocess.Popen(
            cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding='utf-8', errors='replace',
            cwd=cwd or str(BASE_DIR), env=env
        )
        output_lines = []
        for line in proc.stdout:
            output_lines.append(line.rstrip())
            tasks[task_id]['output'] = '\n'.join(output_lines[-100:])
        
        proc.wait()
        tasks[task_id]['status'] = 'done' if proc.returncode == 0 else 'error'
        tasks[task_id]['exit_code'] = proc.returncode
    except Exception as e:
        tasks[task_id]['status'] = 'error'
        tasks[task_id]['output'] += f'\n❌ Error: {e}'


def get_latest_json(month=None):
    """Find latest orders JSON file that contains orders for the specified month"""
    json_files = sorted(OUTPUT_DIR.glob("orders_*.json"), reverse=True)
    if not json_files:
        return None
    
    if month:
        # หาไฟล์ที่มี orders ตรงเดือนที่เลือก
        for jf in json_files[:10]:
            try:
                orders = load_orders(jf)
                if any(o.get('month') == month for o in orders):
                    return jf
            except:
                continue
    
    # Fallback: return most recent
    return json_files[0]


def load_orders(json_path):
    """Load orders from JSON"""
    if not json_path or not json_path.exists():
        return []
    with open(json_path, 'r', encoding='utf-8') as f:
        return json.load(f)


@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route('/api/orders')
def api_orders():
    """Get orders list"""
    month = request.args.get('month', type=int)
    
    # Try local JSON files first
    json_files = sorted(OUTPUT_DIR.glob("orders_*.json"), reverse=True)
    
    all_orders = []
    for jf in json_files[:5]:
        orders = load_orders(jf)
        for o in orders:
            if month and o.get('month') != month:
                continue
            o['_source'] = jf.name
            all_orders.append(o)
    
    # If no local files and on cloud, try Google Sheets cache
    if not all_orders and IS_CLOUD:
        try:
            cached = load_orders_from_sheet()
            for o in cached:
                if month and o.get('month') != month:
                    continue
                o['_source'] = 'sheets_cache'
                all_orders.append(o)
        except Exception as e:
            print(f'Error loading from sheets: {e}')
    
    # Dedup by order_number
    seen = set()
    unique = []
    for o in all_orders:
        key = o.get('order_number', '')
        if key and key not in seen:
            seen.add(key)
            unique.append(o)
    
    # Sort by date ascending (เก่า → ใหม่) ใช้ date parsing จริง
    def parse_sort_date(o):
        d = o.get('date', '')
        try:
            parts = d.split('/')
            if len(parts) == 3:
                return (int(parts[2]), int(parts[1]), int(parts[0]))
        except:
            pass
        return (0, 0, 0)
    
    unique.sort(key=parse_sort_date)
    
    # Flag orders ที่ราคา > 890 และยังไม่ได้แยก + ยังไม่ได้ dismiss
    for o in unique:
        try:
            price = float(o.get('price', '0').replace(',', ''))
            products = o.get('products', [])
            has_individual_prices = all(p.get('price') for p in products) if products else False
            o['needs_split'] = (
                price > 890 
                and not o.get('split_manually') 
                and not o.get('dismiss_split')
                and (len(products) <= 1 or not has_individual_prices)
            )
        except:
            o['needs_split'] = False
    
    return jsonify({
        'orders': unique,
        'count': len(unique),
        'json_files': [f.name for f in json_files[:10]]
    })


@app.route('/api/split-order', methods=['POST'])
def api_split_order():
    """แยก order เป็นหลายรายการ — update JSON file"""
    data = request.json
    order_number = data.get('order_number', '')
    new_products = data.get('products', [])  # [{name, price}, ...]
    
    if not order_number or not new_products:
        return jsonify({'error': 'ต้องระบุ order_number และ products'}), 400
    
    # หา JSON file ที่มี order นี้
    json_files = sorted(OUTPUT_DIR.glob("orders_*.json"), reverse=True)
    updated = False
    
    for jf in json_files[:10]:
        orders = load_orders(jf)
        for o in orders:
            if o.get('order_number', '').replace('#', '') == order_number.replace('#', ''):
                # Update products
                o['products'] = new_products
                o['split_manually'] = True
                
                # Save back
                with open(jf, 'w', encoding='utf-8') as f:
                    json.dump(orders, f, ensure_ascii=False, indent=2)
                
                updated = True
                print(f"✂️ แยก order #{order_number} เป็น {len(new_products)} รายการ → {jf.name}")
                break
        if updated:
            break
    
    if not updated:
        return jsonify({'error': f'ไม่พบ order #{order_number}'}), 404
    
    return jsonify({
        'success': True,
        'order_number': order_number,
        'products': new_products
    })


@app.route('/api/dismiss-split', methods=['POST'])
def api_dismiss_split():
    """Mark order ว่าไม่ต้องแยก (dismiss false positive)"""
    data = request.json
    order_number = data.get('order_number', '')
    
    if not order_number:
        return jsonify({'error': 'ต้องระบุ order_number'}), 400
    
    # Update local JSON files
    json_files = sorted(OUTPUT_DIR.glob("orders_*.json"), reverse=True)
    updated = False
    
    for jf in json_files[:10]:
        orders = load_orders(jf)
        for o in orders:
            if o.get('order_number', '').replace('#', '') == order_number.replace('#', ''):
                o['dismiss_split'] = True
                with open(jf, 'w', encoding='utf-8') as f:
                    json.dump(orders, f, ensure_ascii=False, indent=2)
                updated = True
                print(f"🚫 Dismiss split for order #{order_number} → {jf.name}")
                # Also update Google Sheets cache
                if IS_CLOUD:
                    save_orders_to_sheet(orders)
                break
        if updated:
            break
    
    # If no local file found (cloud), try updating sheets cache directly
    if not updated and IS_CLOUD:
        try:
            cached = load_orders_from_sheet()
            for o in cached:
                if o.get('order_number', '').replace('#', '') == order_number.replace('#', ''):
                    o['dismiss_split'] = True
                    updated = True
                    break
            if updated:
                save_orders_to_sheet(cached)
                print(f"🚫 Dismiss split for order #{order_number} → sheets_cache")
        except Exception as e:
            print(f"Error updating sheets cache: {e}")
    
    if not updated:
        return jsonify({'error': f'ไม่พบ order #{order_number}'}), 404
    
    return jsonify({'success': True, 'order_number': order_number})


@app.route('/api/cleanup-json', methods=['POST'])
def api_cleanup_json():
    """ลบ JSON files เก่าที่ซ้ำซ้อน เหลือแค่ latest per month + named files"""
    json_files = sorted(OUTPUT_DIR.glob("orders_*.json"), reverse=True)
    
    # Keep: named files (e.g. orders_march2026_all.json) + latest 2 timestamped per month
    keep = set()
    timestamped_by_month = {}  # month -> [files]
    
    for jf in json_files:
        name = jf.name
        # Named files (not timestamped) — always keep
        if not re.match(r'orders_\d{8}_\d{6}\.json', name):
            keep.add(jf)
            continue
        
        # Timestamped files — group by month content
        try:
            orders = load_orders(jf)
            months = set(o.get('month', 0) for o in orders)
            month_key = tuple(sorted(months))
        except:
            month_key = (0,)
        
        if month_key not in timestamped_by_month:
            timestamped_by_month[month_key] = []
        timestamped_by_month[month_key].append(jf)
    
    # Keep latest 2 per month group
    for month_key, files in timestamped_by_month.items():
        for f in files[:2]:  # Already sorted newest first
            keep.add(f)
    
    # Delete the rest
    deleted = []
    for jf in json_files:
        if jf not in keep:
            deleted.append(jf.name)
            jf.unlink()
    
    return jsonify({
        'success': True,
        'kept': len(keep),
        'deleted': len(deleted),
        'deleted_files': deleted
    })


@app.route('/api/fetch-orders', methods=['POST'])
def api_fetch_orders():
    """Start fetching orders from FB"""
    if IS_CLOUD:
        return jsonify({'error': '⚠️ ฟีเจอร์นี้ใช้ได้เฉพาะบนเครื่อง local เท่านั้น (ต้องใช้ browser สำหรับ Facebook)'}), 400
    
    month = request.json.get('month', 3)
    task_id = f'fetch_{int(time.time())}'
    
    cmd = f'{sys.executable} fb_orders_bot.py --month {month} --format both --details'
    thread = threading.Thread(target=run_task, args=(task_id, cmd))
    thread.start()
    
    return jsonify({'task_id': task_id, 'message': f'กำลังดึง orders เดือน {THAI_MONTHS.get(month, month)}...'})


@app.route('/api/upload-orders', methods=['POST'])
def api_upload_orders():
    """Upload orders JSON file (for cloud deployment)"""
    orders = None
    
    if 'file' not in request.files:
        data = request.json
        if data and isinstance(data, list):
            orders = data
        else:
            return jsonify({'error': 'ไม่พบไฟล์'}), 400
    else:
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'ไม่ได้เลือกไฟล์'}), 400
        if not file.filename.endswith('.json'):
            return jsonify({'error': 'รองรับเฉพาะไฟล์ .json'}), 400
        try:
            content = file.read().decode('utf-8')
            orders = json.loads(content)
        except Exception as e:
            return jsonify({'error': f'ไฟล์ JSON ไม่ถูกต้อง: {str(e)}'}), 400
    
    if orders is None:
        return jsonify({'error': 'ไม่พบข้อมูล orders'}), 400
    
    try:
        # On cloud: delete ALL old order files first to prevent stale data
        if IS_CLOUD:
            old_files = list(OUTPUT_DIR.glob("orders_*.json"))
            for f in old_files:
                f.unlink()
        
        # Save new file
        filename = f"orders_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        filepath = OUTPUT_DIR / filename
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(orders, f, ensure_ascii=False, indent=2)
        
        # Save to Google Sheets for persistence
        if IS_CLOUD:
            save_orders_to_sheet(orders if isinstance(orders, list) else [])
        
        count = len(orders) if isinstance(orders, list) else 0
        return jsonify({'success': True, 'filename': filename, 'count': count})
    except Exception as e:
        return jsonify({'error': f'เกิดข้อผิดพลาด: {str(e)}'}), 500


@app.route('/api/sync-sheet', methods=['POST'])
def api_sync_sheet():
    """Sync orders to Google Sheet"""
    month = request.json.get('month', 3)
    dry_run = request.json.get('dry_run', False)
    json_file = request.json.get('json_file', '')
    
    if not json_file:
        # Find latest JSON
        jf = get_latest_json(month)
        if not jf:
            return jsonify({'error': 'ไม่พบไฟล์ JSON — กรุณาดึง orders ก่อน'}), 400
        json_file = str(jf)
    
    sheet_name = SHEET_NAMES.get(month, f'{THAI_MONTHS.get(month, month)}69')
    task_id = f'sync_{int(time.time())}'
    
    dry_flag = ' --dry-run' if dry_run else ''
    cmd = f'{sys.executable} sheet_sync.py --json "{json_file}" --sheet "{sheet_name}"{dry_flag}'
    thread = threading.Thread(target=run_task, args=(task_id, cmd))
    thread.start()
    
    mode = 'Dry Run' if dry_run else 'Sync จริง'
    return jsonify({'task_id': task_id, 'message': f'{mode} → Sheet {sheet_name}...'})


@app.route('/api/fetch-bills', methods=['POST'])
def api_fetch_bills():
    """Fetch shipping bills from chat"""
    if IS_CLOUD:
        return jsonify({'error': '⚠️ ฟีเจอร์นี้ใช้ได้เฉพาะบนเครื่อง local เท่านั้น (ต้องใช้ browser สำหรับ Facebook)'}), 400
    
    json_file = request.json.get('json_file', '')
    
    if not json_file:
        jf = get_latest_json()
        if not jf:
            return jsonify({'error': 'ไม่พบไฟล์ JSON'}), 400
        json_file = str(jf)
    
    task_id = f'bills_{int(time.time())}'
    cmd = f'{sys.executable} fetch_shipping.py --json "{json_file}"'
    thread = threading.Thread(target=run_task, args=(task_id, cmd))
    thread.start()
    
    return jsonify({'task_id': task_id, 'message': 'กำลังดึงบิลขนส่ง...'})


@app.route('/api/task/<task_id>')
def api_task_status(task_id):
    """Check task status"""
    task = tasks.get(task_id, {'status': 'not_found'})
    return jsonify(task)


@app.route('/api/files')
def api_files():
    """List output files"""
    files = []
    for f in sorted(OUTPUT_DIR.glob("*"), reverse=True):
        if f.is_file():
            files.append({
                'name': f.name,
                'size': f.stat().st_size,
                'modified': datetime.fromtimestamp(f.stat().st_mtime).strftime('%Y-%m-%d %H:%M')
            })
    return jsonify({'files': files[:30]})


@app.route('/api/update-shipping', methods=['POST'])
def api_update_shipping():
    """Update shipping date/cost in Google Sheet"""
    try:
        data = request.json
        month = data.get('month', 3)
        updates = data.get('updates', [])  # [{customer, ship_date, ship_cost}]
        
        if not updates:
            return jsonify({'error': 'ไม่มีข้อมูลให้อัปเดต'}), 400
        
        gc = get_gspread_client()
        
        sheet_name = SHEET_NAMES.get(month, f'{THAI_MONTHS.get(month, month)}69')
        ws = gc.open('บัญชี HDG 69').worksheet(sheet_name)
        
        col_c = ws.col_values(3)   # C = รายการ
        col_o = ws.col_values(15)  # O = ลูกค้า
        col_b = ws.col_values(2)   # B = วันที่
        
        results = []
        import time as _time
        
        for upd in updates:
            customer = upd.get('customer', '')
            ship_date = upd.get('ship_date', '').strip()
            ship_cost = upd.get('ship_cost', '').strip()
            order_date = upd.get('order_date', '')
            
            if not ship_date and not ship_cost:
                continue
            
            # หา row ค่าส่งพัสดุ ของลูกค้านี้
            found = False
            for i in range(len(col_c)):
                row_num = i + 1
                c_val = col_c[i].strip() if i < len(col_c) else ''
                
                if c_val == 'ค่าส่งพัสดุ' and row_num >= 8:
                    # ค้นหาชื่อลูกค้าจาก rows ก่อนหน้า (ย้อนขึ้นไปสูงสุด 5 rows)
                    # เพราะ order อาจมีหลายสินค้า เช่น Nasri มี 2 items
                    owner_customer = ''
                    owner_date = ''
                    for back in range(1, 6):
                        check_idx = i - back
                        if check_idx < 0:
                            break
                        check_o = col_o[check_idx].replace('fb.', '').strip() if check_idx < len(col_o) else ''
                        if check_o:
                            owner_customer = check_o
                            owner_date = col_b[check_idx] if check_idx < len(col_b) else ''
                            break
                        # ถ้าเจอ row ค่าส่งพัสดุ อีก แสดงว่าเลยเขต order นี้แล้ว
                        check_c = col_c[check_idx].strip() if check_idx < len(col_c) else ''
                        if check_c == 'ค่าส่งพัสดุ':
                            break
                    
                    if not owner_customer:
                        continue
                    
                    if customer.lower() in owner_customer.lower() or owner_customer.lower() in customer.lower():
                        # ตรวจ order_date ด้วยถ้ามี (กรณี Watchara มี 2 orders)
                        if order_date and owner_date and order_date != owner_date:
                            continue
                        
                        if ship_date:
                            ws.update_cell(row_num, 2, ship_date)  # Col B
                            _time.sleep(0.5)
                        if ship_cost:
                            ws.update_cell(row_num, 14, f'  {ship_cost} ')  # Col N
                            _time.sleep(0.5)
                        
                        results.append({
                            'customer': customer,
                            'row': row_num,
                            'ship_date': ship_date,
                            'ship_cost': ship_cost,
                            'status': 'ok'
                        })
                        found = True
                        break
            
            if not found:
                results.append({'customer': customer, 'status': 'not_found'})
        
        return jsonify({'results': results, 'count': len([r for r in results if r['status'] == 'ok'])})
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/mark-shipped', methods=['POST'])
def api_mark_shipped():
    """Mark order as shipped on Facebook"""
    if IS_CLOUD:
        return jsonify({'error': '⚠️ ฟีเจอร์นี้ใช้ได้เฉพาะบนเครื่อง local เท่านั้น (ต้องใช้ browser สำหรับ Facebook)'}), 400
    
    data = request.json
    order_number = data.get('order_number', '')
    customer = data.get('customer', '')
    
    if not order_number:
        return jsonify({'error': 'ไม่มี order number'}), 400
    
    task_id = f'ship_{int(time.time())}'
    cmd = f'{sys.executable} mark_shipped.py --order "{order_number}" --customer "{customer}"'
    thread = threading.Thread(target=run_task, args=(task_id, cmd))
    thread.start()
    
    return jsonify({'task_id': task_id, 'message': f'กำลังทำเครื่องหมาย order {order_number}...'})


@app.route('/api/sort-sheet', methods=['POST'])
def api_sort_sheet():
    """เรียงข้อมูลใน Google Sheet ตามวันที่ (เก่า→ใหม่)"""
    try:
        import time as _time
        
        data = request.json
        month = data.get('month', 3)
        sort_col = data.get('sort_column', 'B').upper()  # B=วันที่สั่ง, N=วันที่ส่ง
        dry_run = data.get('dry_run', False)
        
        COL_MAP = {'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5, 'F': 6, 'G': 7,
                   'H': 8, 'I': 9, 'J': 10, 'K': 11, 'L': 12, 'M': 13, 'N': 14, 'O': 15}
        sort_col_idx = COL_MAP.get(sort_col, 2) - 1  # 0-based index
        
        gc = get_gspread_client()
        
        sheet_name = SHEET_NAMES.get(month, f'{THAI_MONTHS.get(month, month)}69')
        ws = gc.open('บัญชี HDG 69').worksheet(sheet_name)
        
        # อ่านข้อมูลทั้งหมด
        all_values = ws.get_all_values()
        
        DATA_START_ROW = 8  # row 8 (1-indexed)
        FOOTER_KEYWORDS = ['รวม', 'สรุป', 'คำอธิบาย', 'รายรับ', 'รายจ่าย', 'กำไร', 'หมายเหตุ', 'ช่อง']
        
        # หาขอบเขตข้อมูล (row 8 ถึงก่อน footer)
        data_rows = []  # [(row_index_0based, row_values)]
        last_data_idx = DATA_START_ROW - 2  # 0-based
        empty_streak = 0
        
        for i in range(DATA_START_ROW - 1, len(all_values)):
            row = all_values[i]
            b_text = row[1].strip() if len(row) > 1 else ''
            c_text = row[2].strip() if len(row) > 2 else ''
            
            # ตรวจ footer
            if b_text == 'รวม' or c_text == 'รวม':
                break
            if any(b_text.startswith(kw) for kw in FOOTER_KEYWORDS if kw != 'รวม'):
                break
            
            has_content = b_text or c_text
            if has_content:
                empty_streak = 0
                data_rows.append((i, row))
                last_data_idx = i
            else:
                empty_streak += 1
                if empty_streak > 1:
                    break
                data_rows.append((i, row))  # เก็บ empty row ไว้ด้วย (อาจเป็น spacer)
        
        if not data_rows:
            return jsonify({'error': 'ไม่พบข้อมูลใน Sheet'}), 400
        
        # จัดกรุป rows เป็น "order blocks" (แต่ละ order = สินค้า + ค่าส่งพัสดุ)
        # Order block เริ่มจาก row ที่มีวันที่ (col B) ไปจนถึง row ก่อนหน้าวันที่ถัดไป
        def parse_date_val(date_str):
            try:
                parts = date_str.strip().split('/')
                if len(parts) == 3:
                    d, m, y = int(parts[0]), int(parts[1]), int(parts[2])
                    return y * 10000 + m * 100 + d
            except:
                pass
            return 0
        
        blocks = []  # [{date_val, date_str, rows: [row_values]}]
        current_block = None
        
        for _, row in data_rows:
            b_val = row[1].strip() if len(row) > 1 else ''
            c_val = row[2].strip() if len(row) > 2 else ''
            
            # ตรวจวันที่จาก column ที่เลือก
            sort_val_text = row[sort_col_idx].strip() if len(row) > sort_col_idx else ''
            date_val = parse_date_val(b_val) if b_val else 0
            
            if date_val > 0:
                # เริ่ม block ใหม่
                if current_block:
                    blocks.append(current_block)
                
                # หาค่า sort key ตาม column ที่เลือก
                if sort_col == 'B':
                    sort_key = date_val
                else:
                    sort_key = parse_date_val(sort_val_text) if sort_val_text else 0
                
                current_block = {
                    'sort_key': sort_key,
                    'date_str': b_val,
                    'rows': [row]
                }
            elif current_block:
                # rows ต่อเนื่อง (สินค้าเพิ่มเติม, ค่าส่ง, rent, รายจ่าย)
                current_block['rows'].append(row)
                
                # ถ้า sort ตาม column อื่น และ row นี้มีค่าใน sort col — ใช้เป็น sort key
                if sort_col != 'B' and sort_val_text and current_block['sort_key'] == 0:
                    sv = parse_date_val(sort_val_text)
                    if sv > 0:
                        current_block['sort_key'] = sv
            else:
                # orphan row ก่อน block แรก = เป็น expense/rent ที่ไม่มีวันที่ใน B
                # สร้าง block เดี่ยวๆ ไว้
                current_block = {
                    'sort_key': 0,
                    'date_str': '',
                    'rows': [row]
                }
        
        if current_block:
            blocks.append(current_block)
        
        # Sort blocks ตาม sort_key (0 = ไม่มีวันที่ → ไว้ท้ายสุด)
        blocks.sort(key=lambda b: (b['sort_key'] == 0, b['sort_key']))
        
        # Flatten blocks กลับเป็น rows
        sorted_rows = []
        for block in blocks:
            sorted_rows.extend(block['rows'])
        
        result_info = {
            'total_blocks': len(blocks),
            'total_rows': len(sorted_rows),
            'sort_column': sort_col,
            'sheet': sheet_name,
            'blocks': [{'date': b['date_str'], 'rows': len(b['rows'])} for b in blocks[:20]]
        }
        
        if dry_run:
            return jsonify({
                'success': True,
                'dry_run': True,
                'message': f'Dry Run: {len(blocks)} blocks, {len(sorted_rows)} rows จะถูก sort ตาม column {sort_col}',
                **result_info
            })
        
        # เขียนกลับลง Sheet
        # ใช้ batch update เพื่อประหยัด API calls
        start_row = DATA_START_ROW
        end_row = start_row + len(sorted_rows) - 1
        num_cols = max(len(r) for r in sorted_rows) if sorted_rows else 15
        
        # Pad rows ให้ยาวเท่ากัน
        padded = []
        for r in sorted_rows:
            padded_row = list(r) + [''] * (num_cols - len(r))
            padded.append(padded_row[:num_cols])
        
        range_str = f'A{start_row}:{chr(64+num_cols)}{end_row}'
        ws.update(range_str, padded, value_input_option='USER_ENTERED')
        
        return jsonify({
            'success': True,
            'message': f'✅ Sort เสร็จ! {len(blocks)} blocks, {len(sorted_rows)} rows เรียงตาม column {sort_col}',
            **result_info
        })
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


# ============================================================
# HTML Template
# ============================================================
HTML_TEMPLATE = r"""
<!DOCTYPE html>
<html lang="th">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>HDG Orders Dashboard</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    <style>
        :root {
            --bg-primary: #0f0f1a;
            --bg-secondary: #1a1a2e;
            --bg-card: #16213e;
            --bg-card-hover: #1a2747;
            --accent: #e94560;
            --accent-glow: rgba(233, 69, 96, 0.3);
            --green: #00d4aa;
            --green-glow: rgba(0, 212, 170, 0.3);
            --blue: #5c7cfa;
            --blue-glow: rgba(92, 124, 250, 0.3);
            --orange: #ff922b;
            --orange-glow: rgba(255, 146, 43, 0.3);
            --text-primary: #e8e8f0;
            --text-secondary: #8888aa;
            --text-muted: #555570;
            --border: #2a2a45;
            --glass: rgba(255,255,255,0.05);
        }
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Inter', sans-serif;
            background: var(--bg-primary);
            color: var(--text-primary);
            min-height: 100vh;
        }
        
        /* Header */
        .header {
            background: linear-gradient(135deg, var(--bg-secondary) 0%, var(--bg-card) 100%);
            border-bottom: 1px solid var(--border);
            padding: 20px 32px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            position: sticky;
            top: 0;
            z-index: 100;
            backdrop-filter: blur(20px);
        }
        .header h1 {
            font-size: 22px;
            font-weight: 700;
            background: linear-gradient(135deg, var(--accent), #ff6b8a);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }
        .header-sub {
            font-size: 12px;
            color: var(--text-secondary);
            margin-top: 2px;
        }
        .header-right {
            display: flex;
            align-items: center;
            gap: 12px;
        }
        
        /* Month Selector */
        .month-selector {
            display: flex;
            align-items: center;
            gap: 8px;
            background: var(--glass);
            border: 1px solid var(--border);
            border-radius: 12px;
            padding: 6px 12px;
        }
        .month-selector label {
            font-size: 13px;
            color: var(--text-secondary);
            font-weight: 500;
        }
        .month-selector select {
            background: var(--bg-card);
            border: 1px solid var(--border);
            color: var(--text-primary);
            padding: 6px 12px;
            border-radius: 8px;
            font-family: 'Inter', sans-serif;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            outline: none;
        }
        .month-selector select:focus {
            border-color: var(--accent);
            box-shadow: 0 0 0 3px var(--accent-glow);
        }
        
        /* Main Layout */
        .main {
            max-width: 1400px;
            margin: 0 auto;
            padding: 24px 32px;
        }
        
        /* Action Cards */
        .actions {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: 16px;
            margin-bottom: 24px;
        }
        .action-card {
            background: var(--bg-card);
            border: 1px solid var(--border);
            border-radius: 16px;
            padding: 24px;
            cursor: pointer;
            transition: all 0.3s ease;
            position: relative;
            overflow: hidden;
        }
        .action-card::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 3px;
            border-radius: 16px 16px 0 0;
        }
        .action-card:hover {
            transform: translateY(-3px);
            border-color: transparent;
        }
        .action-card.fetch::before { background: linear-gradient(90deg, var(--accent), #ff6b8a); }
        .action-card.fetch:hover { box-shadow: 0 8px 30px var(--accent-glow); border-color: var(--accent); }
        .action-card.sync::before { background: linear-gradient(90deg, var(--green), #4aedc4); }
        .action-card.sync:hover { box-shadow: 0 8px 30px var(--green-glow); border-color: var(--green); }
        .action-card.bills::before { background: linear-gradient(90deg, var(--blue), #8da2fb); }
        .action-card.bills:hover { box-shadow: 0 8px 30px var(--blue-glow); border-color: var(--blue); }
        
        .action-icon {
            font-size: 28px;
            margin-bottom: 12px;
        }
        .action-title {
            font-size: 16px;
            font-weight: 700;
            margin-bottom: 6px;
        }
        .action-desc {
            font-size: 12px;
            color: var(--text-secondary);
            line-height: 1.5;
        }
        .action-btns {
            display: flex;
            gap: 8px;
            margin-top: 16px;
        }
        
        /* Buttons */
        .btn {
            padding: 8px 18px;
            border: none;
            border-radius: 8px;
            font-family: 'Inter', sans-serif;
            font-size: 13px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
            display: inline-flex;
            align-items: center;
            gap: 6px;
        }
        .btn:active { transform: scale(0.97); }
        .btn-primary {
            background: linear-gradient(135deg, var(--accent), #c73e54);
            color: #fff;
        }
        .btn-primary:hover { box-shadow: 0 4px 15px var(--accent-glow); }
        .btn-green {
            background: linear-gradient(135deg, var(--green), #00b894);
            color: #fff;
        }
        .btn-green:hover { box-shadow: 0 4px 15px var(--green-glow); }
        .btn-blue {
            background: linear-gradient(135deg, var(--blue), #4a6cf7);
            color: #fff;
        }
        .btn-blue:hover { box-shadow: 0 4px 15px var(--blue-glow); }
        .btn-outline {
            background: transparent;
            border: 1px solid var(--border);
            color: var(--text-secondary);
        }
        .btn-outline:hover { border-color: var(--text-secondary); color: var(--text-primary); }
        .btn:disabled {
            opacity: 0.5;
            cursor: not-allowed;
        }
        
        /* Status Log */
        .status-panel {
            background: var(--bg-card);
            border: 1px solid var(--border);
            border-radius: 16px;
            margin-bottom: 24px;
            overflow: hidden;
            display: none;
        }
        .status-panel.active { display: block; }
        .status-header {
            padding: 14px 20px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            border-bottom: 1px solid var(--border);
        }
        .status-header h3 {
            font-size: 14px;
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .status-dot {
            width: 8px;
            height: 8px;
            border-radius: 50%;
            background: var(--orange);
            animation: pulse 1.5s infinite;
        }
        .status-dot.done { background: var(--green); animation: none; }
        .status-dot.error { background: var(--accent); animation: none; }
        @keyframes pulse {
            0%, 100% { opacity: 1; }
            50% { opacity: 0.3; }
        }
        .status-log {
            padding: 16px 20px;
            max-height: 300px;
            overflow-y: auto;
            font-family: 'JetBrains Mono', 'Fira Code', monospace;
            font-size: 12px;
            line-height: 1.8;
            color: var(--text-secondary);
            white-space: pre-wrap;
            word-break: break-all;
        }
        
        /* Orders Table */
        .table-container {
            background: var(--bg-card);
            border: 1px solid var(--border);
            border-radius: 16px;
            overflow: hidden;
        }
        .table-header {
            padding: 16px 20px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            border-bottom: 1px solid var(--border);
        }
        .table-header h3 {
            font-size: 15px;
            font-weight: 600;
        }
        .table-count {
            background: var(--glass);
            border: 1px solid var(--border);
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 12px;
            color: var(--text-secondary);
        }
        table {
            width: 100%;
            border-collapse: collapse;
        }
        th {
            padding: 10px 16px;
            text-align: left;
            font-size: 11px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            color: var(--text-muted);
            background: var(--bg-secondary);
            border-bottom: 1px solid var(--border);
            position: sticky;
            top: 0;
            z-index: 2;
        }
        th.sortable {
            cursor: pointer;
            user-select: none;
            transition: color 0.2s, background 0.2s;
            position: relative;
            padding-right: 24px;
        }
        th.sortable:hover {
            color: var(--text-primary);
            background: rgba(255,255,255,0.06);
        }
        th.sortable.active {
            color: var(--blue);
        }
        th.sortable::after {
            content: '⇅';
            position: absolute;
            right: 8px;
            top: 50%;
            transform: translateY(-50%);
            font-size: 10px;
            opacity: 0.3;
            transition: opacity 0.2s;
        }
        th.sortable:hover::after {
            opacity: 0.6;
        }
        th.sortable.asc::after {
            content: '▲';
            opacity: 1;
            color: var(--blue);
        }
        th.sortable.desc::after {
            content: '▼';
            opacity: 1;
            color: var(--blue);
        }
        td {
            padding: 12px 16px;
            font-size: 13px;
            border-bottom: 1px solid rgba(255,255,255,0.03);
        }
        tr:hover td {
            background: var(--glass);
        }
        .badge {
            padding: 3px 10px;
            border-radius: 20px;
            font-size: 11px;
            font-weight: 600;
        }
        .badge-green { background: rgba(0,212,170,0.15); color: var(--green); }
        .badge-red { background: rgba(233,69,96,0.15); color: var(--accent); }
        .badge-yellow { background: rgba(255,146,43,0.15); color: var(--orange); }
        
        .empty-state {
            text-align: center;
            padding: 60px 20px;
            color: var(--text-muted);
        }
        .empty-state .icon { font-size: 48px; margin-bottom: 16px; }
        .empty-state p { font-size: 14px; }
        
        .price { font-weight: 600; color: var(--green); }
        .product-name { 
            max-width: 250px; 
            overflow: hidden; 
            text-overflow: ellipsis; 
            white-space: nowrap; 
        }
        .product-lines {
            max-width: 300px;
        }
        .product-line {
            font-size: 12px;
            padding: 3px 0;
            display: flex;
            justify-content: space-between;
            gap: 8px;
            border-bottom: 1px solid rgba(255,255,255,0.03);
        }
        .product-line:last-child { border-bottom: none; }
        .product-line .pl-name {
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }
        .product-line .pl-price {
            color: var(--green);
            font-weight: 600;
            white-space: nowrap;
            font-size: 11px;
        }
        .item-count {
            background: var(--glass);
            border: 1px solid var(--border);
            color: var(--text-secondary);
            padding: 1px 7px;
            border-radius: 10px;
            font-size: 10px;
            font-weight: 700;
            margin-bottom: 4px;
            display: inline-block;
        }
        .btn-refresh {
            background: var(--glass);
            border: 1px solid var(--border);
            color: var(--text-secondary);
            width: 36px;
            height: 36px;
            border-radius: 10px;
            font-size: 16px;
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
            transition: all 0.3s;
        }
        .btn-refresh:hover {
            border-color: var(--blue);
            color: var(--blue);
            transform: rotate(180deg);
        }
        
        .scroll-table {
            max-height: 600px;
            overflow-y: auto;
        }
        
        /* Shipping Inputs */
        .ship-input {
            background: var(--bg-secondary);
            border: 1px solid var(--border);
            color: var(--text-primary);
            padding: 6px 10px;
            border-radius: 6px;
            font-family: 'Inter', sans-serif;
            font-size: 12px;
            width: 110px;
            outline: none;
            transition: all 0.2s;
        }
        .ship-input:focus {
            border-color: var(--blue);
            box-shadow: 0 0 0 2px var(--blue-glow);
        }
        .ship-input.cost {
            width: 70px;
            text-align: right;
        }
        .ship-input::placeholder {
            color: var(--text-muted);
            font-size: 11px;
        }
        .btn-save-row {
            background: var(--glass);
            border: 1px solid var(--border);
            color: var(--text-secondary);
            padding: 4px 10px;
            border-radius: 6px;
            font-size: 11px;
            cursor: pointer;
            transition: all 0.2s;
        }
        .btn-save-row:hover {
            border-color: var(--green);
            color: var(--green);
            background: rgba(0,212,170,0.1);
        }
        .btn-save-row.saved {
            border-color: var(--green);
            color: var(--green);
        }
        .save-all-bar {
            padding: 12px 20px;
            border-top: 1px solid var(--border);
            display: flex;
            align-items: center;
            justify-content: flex-end;
            gap: 12px;
            background: var(--bg-secondary);
        }
        .save-msg {
            font-size: 12px;
            color: var(--green);
            display: none;
        }
        .save-msg.active { display: inline; }
        
        /* Mark Shipped Button */
        .btn-mark-ship {
            background: linear-gradient(135deg, var(--orange), #e67e22);
            border: none;
            color: #fff;
            padding: 4px 10px;
            border-radius: 6px;
            font-size: 11px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
            white-space: nowrap;
        }
        .btn-mark-ship:hover {
            box-shadow: 0 3px 12px var(--orange-glow);
            transform: translateY(-1px);
        }
        .btn-mark-ship:disabled {
            opacity: 0.5;
            cursor: not-allowed;
        }
        .btn-mark-ship.done {
            background: linear-gradient(135deg, var(--green), #00b894);
        }
        
        /* Dismiss Split Button */
        .btn-dismiss {
            background: var(--glass);
            border: 1px solid var(--border);
            color: var(--text-muted);
            padding: 4px 8px;
            border-radius: 6px;
            font-size: 11px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
            white-space: nowrap;
        }
        .btn-dismiss:hover {
            border-color: var(--green);
            color: var(--green);
            background: rgba(0,212,170,0.1);
        }
        
        /* Split Warning Banner */
        .split-warning {
            background: linear-gradient(135deg, rgba(255,146,43,0.12), rgba(233,69,96,0.10));
            border: 1px solid rgba(255,146,43,0.35);
            border-radius: 14px;
            padding: 16px 22px;
            margin-bottom: 20px;
            display: none;
            align-items: flex-start;
            gap: 14px;
            animation: slideDown 0.35s ease;
        }
        .split-warning.active { display: flex; }
        @keyframes slideDown {
            from { opacity: 0; transform: translateY(-10px); }
            to   { opacity: 1; transform: translateY(0); }
        }
        .split-warning-icon {
            font-size: 26px;
            flex-shrink: 0;
            margin-top: 2px;
        }
        .split-warning-body {
            flex: 1;
        }
        .split-warning-title {
            font-size: 14px;
            font-weight: 700;
            color: var(--orange);
            margin-bottom: 4px;
        }
        .split-warning-desc {
            font-size: 12px;
            color: var(--text-secondary);
            line-height: 1.6;
        }
        .split-warning-list {
            margin-top: 8px;
            display: flex;
            flex-wrap: wrap;
            gap: 6px;
        }
        .split-warning-chip {
            background: rgba(255,146,43,0.18);
            border: 1px solid rgba(255,146,43,0.3);
            color: var(--orange);
            padding: 3px 10px;
            border-radius: 20px;
            font-size: 11px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
        }
        .split-warning-chip:hover {
            background: rgba(255,146,43,0.3);
            transform: scale(1.05);
        }
        
        /* Split Badge on rows */
        .badge-split {
            background: rgba(255,146,43,0.15);
            color: var(--orange);
            padding: 3px 10px;
            border-radius: 20px;
            font-size: 10px;
            font-weight: 700;
            margin-left: 6px;
            animation: pulseGlow 2s infinite;
            display: inline-flex;
            align-items: center;
            gap: 4px;
        }
        @keyframes pulseGlow {
            0%, 100% { box-shadow: 0 0 0 0 rgba(255,146,43,0.3); }
            50%      { box-shadow: 0 0 0 6px rgba(255,146,43,0); }
        }
        .badge-split-done {
            background: rgba(0,212,170,0.15);
            color: var(--green);
            animation: none;
        }
        
        /* Split Button */
        .btn-split {
            background: linear-gradient(135deg, var(--orange), #e67e22);
            border: none;
            color: #fff;
            padding: 4px 10px;
            border-radius: 6px;
            font-size: 11px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
            white-space: nowrap;
        }
        .btn-split:hover {
            box-shadow: 0 3px 12px var(--orange-glow);
            transform: translateY(-1px);
        }
        .btn-split.done {
            background: linear-gradient(135deg, var(--green), #00b894);
        }
        
        /* Split Modal */
        .modal-overlay {
            position: fixed;
            inset: 0;
            background: rgba(0,0,0,0.65);
            z-index: 1000;
            display: none;
            align-items: center;
            justify-content: center;
            backdrop-filter: blur(4px);
            animation: fadeIn 0.2s ease;
        }
        .modal-overlay.active { display: flex; }
        @keyframes fadeIn {
            from { opacity: 0; }
            to   { opacity: 1; }
        }
        .modal {
            background: var(--bg-card);
            border: 1px solid var(--border);
            border-radius: 20px;
            width: 540px;
            max-width: 95vw;
            max-height: 85vh;
            overflow-y: auto;
            box-shadow: 0 25px 60px rgba(0,0,0,0.5);
            animation: modalSlide 0.3s ease;
        }
        @keyframes modalSlide {
            from { opacity: 0; transform: scale(0.95) translateY(10px); }
            to   { opacity: 1; transform: scale(1) translateY(0); }
        }
        .modal-header {
            padding: 22px 24px 14px;
            border-bottom: 1px solid var(--border);
            display: flex;
            align-items: center;
            justify-content: space-between;
        }
        .modal-header h2 {
            font-size: 17px;
            font-weight: 700;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .modal-close {
            background: var(--glass);
            border: 1px solid var(--border);
            color: var(--text-secondary);
            width: 32px;
            height: 32px;
            border-radius: 8px;
            font-size: 16px;
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
            transition: all 0.2s;
        }
        .modal-close:hover {
            background: rgba(233,69,96,0.15);
            border-color: var(--accent);
            color: var(--accent);
        }
        .modal-body {
            padding: 20px 24px;
        }
        .modal-info {
            background: var(--glass);
            border: 1px solid var(--border);
            border-radius: 10px;
            padding: 12px 16px;
            margin-bottom: 18px;
            font-size: 13px;
            display: grid;
            grid-template-columns: auto 1fr;
            gap: 6px 14px;
        }
        .modal-info-label {
            color: var(--text-muted);
            font-size: 12px;
        }
        .modal-info-value {
            font-weight: 600;
        }
        .split-items-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 10px;
        }
        .split-items-header h4 {
            font-size: 13px;
            font-weight: 600;
            color: var(--text-secondary);
        }
        .btn-add-item {
            background: var(--glass);
            border: 1px dashed var(--border);
            color: var(--blue);
            padding: 5px 12px;
            border-radius: 8px;
            font-size: 12px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
        }
        .btn-add-item:hover {
            border-color: var(--blue);
            background: rgba(92,124,250,0.1);
        }
        .split-item-row {
            display: flex;
            gap: 10px;
            align-items: center;
            margin-bottom: 10px;
            animation: slideDown 0.2s ease;
        }
        .split-item-row input {
            background: var(--bg-secondary);
            border: 1px solid var(--border);
            color: var(--text-primary);
            padding: 8px 12px;
            border-radius: 8px;
            font-family: 'Inter', sans-serif;
            font-size: 13px;
            outline: none;
            transition: all 0.2s;
        }
        .split-item-row input:focus {
            border-color: var(--blue);
            box-shadow: 0 0 0 2px var(--blue-glow);
        }
        .split-item-row input.item-name { flex: 1; }
        .split-item-row input.item-price { width: 100px; text-align: right; }
        .split-item-remove {
            background: none;
            border: none;
            color: var(--text-muted);
            font-size: 16px;
            cursor: pointer;
            padding: 4px;
            border-radius: 6px;
            transition: all 0.2s;
        }
        .split-item-remove:hover {
            color: var(--accent);
            background: rgba(233,69,96,0.1);
        }
        .split-total {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 10px 0;
            margin-top: 6px;
            border-top: 1px solid var(--border);
            font-size: 13px;
        }
        .split-total-label { color: var(--text-secondary); }
        .split-total-value {
            font-weight: 700;
            font-size: 16px;
            color: var(--green);
        }
        .split-total-mismatch { color: var(--accent); }
        .modal-footer {
            padding: 14px 24px 22px;
            display: flex;
            align-items: center;
            justify-content: flex-end;
            gap: 10px;
        }
        
        /* Row highlight for needs_split */
        tr.needs-split td {
            background: rgba(255,146,43,0.04);
        }
        tr.needs-split:hover td {
            background: rgba(255,146,43,0.08);
        }
        tr.was-split td {
            background: rgba(0,212,170,0.04);
        }
        
        /* Responsive */
        @media (max-width: 768px) {
            .header { padding: 16px; }
            .main { padding: 16px; }
            .actions { grid-template-columns: 1fr; }
            .modal { width: 100%; border-radius: 16px 16px 0 0; }
        }
    </style>
</head>
<body>
    <div class="header">
        <div>
            <h1>📦 HDG Orders Dashboard</h1>
            <div class="header-sub">Facebook Orders → Google Sheet Automation</div>
        </div>
        <div class="header-right">
            <button class="btn-refresh" onclick="loadOrders()" title="รีเฟรช orders">&#x1f504;</button>
            <div class="month-selector">
                <label>เดือน:</label>
                <select id="monthSelect" onchange="loadOrders()">
                    <option value="1">ม.ค. 69</option>
                    <option value="2">ก.พ. 69</option>
                    <option value="3" selected>มี.ค. 69</option>
                    <option value="4">เม.ย. 69</option>
                    <option value="5">พ.ค. 69</option>
                    <option value="6">มิ.ย. 69</option>
                    <option value="7">ก.ค. 69</option>
                    <option value="8">ส.ค. 69</option>
                    <option value="9">ก.ย. 69</option>
                    <option value="10">ต.ค. 69</option>
                    <option value="11">พ.ย. 69</option>
                    <option value="12">ธ.ค. 69</option>
                </select>
            </div>
        </div>
    </div>
    
    <div class="main">
        <!-- Action Cards -->
        <div class="actions">
            <div class="action-card fetch">
                <div class="action-icon">📥</div>
                <div class="action-title">ดึง Orders จาก Facebook</div>
                <div class="action-desc">เปิด Facebook Business Suite → อ่านรายการสั่งซื้อ → Export JSON + CSV</div>
                <div class="action-btns">
                    <button class="btn btn-primary" onclick="fetchOrders()" id="btnFetch">
                        ▶ ดึง Orders
                    </button>
                    <button class="btn btn-outline" onclick="document.getElementById('uploadJsonInput').click()">
                        📤 อัปโหลด JSON
                    </button>
                    <input type="file" id="uploadJsonInput" accept=".json" style="display:none" onchange="uploadOrders(this)">
                </div>
            </div>
            
            <div class="action-card sync">
                <div class="action-icon">📊</div>
                <div class="action-title">Sync ลง Google Sheet</div>
                <div class="action-desc">แทรกข้อมูล orders ลง Sheet "บัญชี HDG 69" เรียงตามวันที่</div>
                <div class="action-btns">
                    <button class="btn btn-outline" onclick="syncSheet(true)">👁 Dry Run</button>
                    <button class="btn btn-green" onclick="syncSheet(false)" id="btnSync">
                        ▶ Sync จริง
                    </button>
                </div>
            </div>
            
            <div class="action-card bills">
                <div class="action-icon">🧾</div>
                <div class="action-title">ดึงบิลขนส่ง</div>
                <div class="action-desc">เปิดแชทลูกค้า → ดาวน์โหลดรูปบิลขนส่งจาก Messenger</div>
                <div class="action-btns">
                    <button class="btn btn-blue" onclick="fetchBills()" id="btnBills">
                        ▶ ดึงบิล
                    </button>
                </div>
            </div>
            
            <div class="action-card" style="border-top: 3px solid var(--blue);">
                <div class="action-icon">🗂️</div>
                <div class="action-title">Sort Google Sheet</div>
                <div class="action-desc">เรียงข้อมูลใน Google Sheet ตามวันที่ (เก่า→ใหม่) เลือก column ที่ต้องการ sort</div>
                <div class="action-btns" style="align-items:center">
                    <select id="sortColSelect" style="background:var(--bg-card);border:1px solid var(--border);color:var(--text-primary);padding:6px 12px;border-radius:8px;font-family:'Inter',sans-serif;font-size:13px;outline:none;">
                        <option value="B" selected>B — วันที่สั่ง</option>
                        <option value="N">N — วันที่ส่ง</option>
                    </select>
                    <button class="btn btn-outline" onclick="sortSheet(true)">👁 Dry Run</button>
                    <button class="btn btn-blue" onclick="sortSheet(false)" id="btnSortSheet">
                        ▶ Sort Sheet
                    </button>
                </div>
            </div>
            
            <div class="action-card" style="border-top: 3px solid var(--text-muted);">
                <div class="action-icon">🧹</div>
                <div class="action-title">Cleanup JSON Files</div>
                <div class="action-desc">ลบ JSON files เก่าที่ซ้ำซ้อน เหลือแค่ latest per month + named files</div>
                <div class="action-btns">
                    <button class="btn btn-outline" onclick="cleanupJson()" id="btnCleanup">
                        🧹 Cleanup
                    </button>
                    <span id="cleanupMsg" style="font-size:12px;color:var(--text-secondary)"></span>
                </div>
            </div>
        </div>
        
        <!-- Status Panel -->
        <div class="status-panel" id="statusPanel">
            <div class="status-header">
                <h3>
                    <span class="status-dot" id="statusDot"></span>
                    <span id="statusTitle">กำลังทำงาน...</span>
                </h3>
                <button class="btn btn-outline" onclick="hideStatus()" style="padding:4px 12px;font-size:12px;">✕ ปิด</button>
            </div>
            <div class="status-log" id="statusLog"></div>
        </div>
        
        <!-- Split Warning Banner -->
        <div class="split-warning" id="splitWarning">
            <div class="split-warning-icon">⚠️</div>
            <div class="split-warning-body">
                <div class="split-warning-title">พบ orders ที่อาจต้องแยกรายการ (ราคา > ฿890)</div>
                <div class="split-warning-desc" id="splitWarningDesc">กรุณาตรวจสอบและแยกสินค้าก่อน Sync ลง Sheet</div>
                <div class="split-warning-list" id="splitWarningList"></div>
            </div>
        </div>
        
        <!-- Split Modal -->
        <div class="modal-overlay" id="splitModal">
            <div class="modal">
                <div class="modal-header">
                    <h2>✂️ แยกรายการสินค้า</h2>
                    <button class="modal-close" onclick="closeSplitModal()">✕</button>
                </div>
                <div class="modal-body">
                    <div class="modal-info" id="splitModalInfo"></div>
                    <div class="split-items-header">
                        <h4>รายการสินค้าย่อย</h4>
                        <button class="btn-add-item" onclick="addSplitItem()">+ เพิ่มรายการ</button>
                    </div>
                    <div id="splitItemsList"></div>
                    <div class="split-total">
                        <span class="split-total-label">รวมทั้งหมด</span>
                        <span class="split-total-value" id="splitTotalValue">฿0</span>
                    </div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-outline" onclick="closeSplitModal()">ยกเลิก</button>
                    <button class="btn btn-green" onclick="submitSplit()" id="btnSubmitSplit">✂️ บันทึกการแยก</button>
                </div>
            </div>
        </div>
        
        <!-- Orders Table -->
        <div class="table-container">
            <div class="table-header">
                <h3>📋 รายการ Orders</h3>
                <span class="table-count" id="orderCount">0 รายการ</span>
            </div>
            <div class="scroll-table">
                <table>
                    <thead>
                        <tr>
                            <th>#</th>
                            <th class="sortable" data-sort="date" onclick="toggleSort('date')">วันที่สั่ง</th>
                            <th class="sortable" data-sort="customer" onclick="toggleSort('customer')">ลูกค้า</th>
                            <th>สินค้า</th>
                            <th class="sortable" data-sort="price" onclick="toggleSort('price')">ราคา</th>
                            <th>📅 วันที่ส่ง</th>
                            <th>💰 ค่าส่ง</th>
                            <th class="sortable" data-sort="status" onclick="toggleSort('status')">สถานะ</th>
                            <th></th>
                        </tr>
                    </thead>
                    <tbody id="ordersBody">
                        <tr>
                            <td colspan="9">
                                <div class="empty-state">
                                    <div class="icon">📦</div>
                                    <p>เลือกเดือนแล้วกด "ดึง Orders" หรือรอโหลดข้อมูล...</p>
                                </div>
                            </td>
                        </tr>
                    </tbody>
                </table>
            </div>
            <div class="save-all-bar">
                <span class="save-msg" id="saveMsg">✅ บันทึกเรียบร้อย!</span>
                <button class="btn btn-green" onclick="saveAllShipping()" id="btnSaveAll">
                    💾 บันทึกทั้งหมดลง Sheet
                </button>
            </div>
        </div>
    </div>
    
    <script>
        let currentTaskId = null;
        let pollInterval = null;
        // Store current orders globally for split modal
        let currentOrders = [];
        let currentSplitOrder = null;
        // Sort state
        let currentSort = { column: 'date', direction: 'asc' };
        
        function getMonth() {
            return document.getElementById('monthSelect').value;
        }
        
        // ============ Sort Logic ============
        function parseDateForSort(dateStr) {
            // Parse "d/m/yyyy" → comparable number (yyyymmdd)
            try {
                const parts = dateStr.split('/');
                if (parts.length === 3) {
                    const d = parseInt(parts[0]), m = parseInt(parts[1]), y = parseInt(parts[2]);
                    return y * 10000 + m * 100 + d;
                }
            } catch(e) {}
            return 0;
        }
        
        function toggleSort(column) {
            if (currentSort.column === column) {
                // สลับทิศทาง
                currentSort.direction = currentSort.direction === 'asc' ? 'desc' : 'asc';
            } else {
                currentSort.column = column;
                currentSort.direction = 'asc';
            }
            // Re-render ด้วย sort ใหม่
            renderOrders(currentOrders);
        }
        
        function sortOrders(orders) {
            const col = currentSort.column;
            const dir = currentSort.direction === 'asc' ? 1 : -1;
            
            return [...orders].sort((a, b) => {
                let va, vb;
                switch (col) {
                    case 'date':
                        va = parseDateForSort(a.date || '');
                        vb = parseDateForSort(b.date || '');
                        return (va - vb) * dir;
                    case 'customer':
                        va = (a.customer || '').toLowerCase();
                        vb = (b.customer || '').toLowerCase();
                        return va.localeCompare(vb) * dir;
                    case 'price':
                        va = parseFloat((a.price || '0').replace(/,/g, ''));
                        vb = parseFloat((b.price || '0').replace(/,/g, ''));
                        return (va - vb) * dir;
                    case 'status':
                        va = a.status || '';
                        vb = b.status || '';
                        return va.localeCompare(vb) * dir;
                    default:
                        return 0;
                }
            });
        }
        
        function updateSortHeaders() {
            document.querySelectorAll('th.sortable').forEach(th => {
                th.classList.remove('active', 'asc', 'desc');
                if (th.dataset.sort === currentSort.column) {
                    th.classList.add('active', currentSort.direction);
                }
            });
        }
        
        async function loadOrders() {
            const month = getMonth();
            try {
                const resp = await fetch(`/api/orders?month=${month}`);
                const data = await resp.json();
                renderOrders(data.orders);
            } catch(e) {
                console.error('Load error:', e);
            }
        }
        
        function renderOrders(orders) {
            currentOrders = orders;
            const tbody = document.getElementById('ordersBody');
            // Filter out cancelled orders for display
            let activeOrders = orders.filter(o => o.status !== 'ยกเลิกแล้ว');
            
            // Apply sort
            activeOrders = sortOrders(activeOrders);
            updateSortHeaders();
            
            document.getElementById('orderCount').textContent = `${activeOrders.length} รายการ (${orders.length} ทั้งหมด)`;
            
            // Show/hide split warning banner
            const splitOrders = activeOrders.filter(o => o.needs_split && !o.split_manually);
            const warningEl = document.getElementById('splitWarning');
            if (splitOrders.length > 0) {
                warningEl.classList.add('active');
                document.getElementById('splitWarningDesc').textContent = 
                    `พบ ${splitOrders.length} รายการที่ราคารวม > ฿890 — กรุณาตรวจสอบและแยกสินค้าก่อน Sync`;
                document.getElementById('splitWarningList').innerHTML = splitOrders.map(o => 
                    `<span class="split-warning-chip" onclick="openSplitModal('${(o.order_number||'').replace(/'/g,"\\\\'")}')">`+
                    `${o.customer || 'N/A'} — ฿${parseFloat(o.price||0).toLocaleString()}</span>`
                ).join('');
            } else {
                warningEl.classList.remove('active');
            }
            
            if (orders.length === 0) {
                tbody.innerHTML = `<tr><td colspan="9">
                    <div class="empty-state">
                        <div class="icon">📭</div>
                        <p>ไม่พบ orders เดือนนี้ — กด "ดึง Orders" เพื่อดึงข้อมูลใหม่</p>
                    </div>
                </td></tr>`;
                return;
            }
            
            tbody.innerHTML = activeOrders.map((o, i) => {
                const prodList = o.products || [];
                const productNames = prodList.map(p => p.name).join(', ') || '(ไม่ทราบ)';
                let productCell = '';
                if (prodList.length > 1) {
                    productCell = `<div class="product-lines">
                        <span class="item-count">${prodList.length} items</span>
                        ${prodList.map(p => `<div class="product-line"><span class="pl-name">${p.name}</span><span class="pl-price">${p.price ? '\u0e3f'+parseFloat(p.price).toLocaleString() : ''}</span></div>`).join('')}
                    </div>`;
                } else {
                    productCell = `<div class="product-name" title="${productNames}">${productNames}</div>`;
                }
                const price = o.price ? `฿${parseFloat(o.price).toLocaleString()}` : '-';
                const orderDate = o.date || o.date_raw || '-';
                const custSafe = (o.customer || '').replace(/'/g, "\\'");
                
                let statusBadge = '';
                if (o.status === 'แนบสลิปแล้ว') statusBadge = '<span class="badge badge-green">แนบสลิปแล้ว</span>';
                else if (o.status === 'ยกเลิกแล้ว') statusBadge = '<span class="badge badge-red">ยกเลิกแล้ว</span>';
                else statusBadge = `<span class="badge badge-yellow">${o.status || '-'}</span>`;
                
                const orderNum = (o.order_number || '').replace(/'/g, "\\'");
                
                // Split badge
                let splitBadge = '';
                let splitBtn = '';
                let rowClass = '';
                if (o.split_manually) {
                    splitBadge = '<span class="badge-split badge-split-done">✅ แยกแล้ว</span>';
                    rowClass = 'was-split';
                } else if (o.needs_split) {
                    splitBadge = '<span class="badge-split">⚠️ ต้องแยก</span>';
                    splitBtn = `<button class="btn-split" onclick="openSplitModal('${orderNum}')">✂️ แยก</button><button class="btn-dismiss" onclick="dismissSplit(this, '${orderNum}')">✓ OK</button>`;
                    rowClass = 'needs-split';
                }
                
                return `<tr class="${rowClass}" data-customer="${custSafe}" data-order-date="${orderDate}" data-order-num="${orderNum}">
                    <td style="color:var(--text-muted)">${i+1}</td>
                    <td>${orderDate}</td>
                    <td><strong>${o.customer || '-'}</strong>${splitBadge}</td>
                    <td>${productCell}</td>
                    <td><span class="price">${price}</span></td>
                    <td><input type="text" class="ship-input ship-date" placeholder="เช่น 8/3/2026" data-idx="${i}"></td>
                    <td><input type="text" class="ship-input cost ship-cost" placeholder="30" data-idx="${i}"></td>
                    <td>${statusBadge}</td>
                    <td style="display:flex;gap:4px;align-items:center">
                        ${splitBtn}
                        <button class="btn-save-row" onclick="saveOneShipping(this, '${custSafe}', '${orderDate}', ${i})">💾</button>
                        <button class="btn-mark-ship" onclick="markShipped(this, '${orderNum}', '${custSafe}')">🚚 จัดส่ง</button>
                    </td>
                </tr>`;
            }).join('');
        }
        
        async function fetchOrders() {
            const month = getMonth();
            const btn = document.getElementById('btnFetch');
            btn.disabled = true;
            btn.textContent = '⏳ กำลังดึง...';
            
            try {
                const resp = await fetch('/api/fetch-orders', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({month: parseInt(month)})
                });
                const data = await resp.json();
                if (data.error) {
                    alert(data.error);
                    btn.disabled = false;
                    btn.textContent = '▶ ดึง Orders';
                    return;
                }
                showStatus(data.message, data.task_id);
            } catch(e) {
                alert('Error: ' + e.message);
                btn.disabled = false;
                btn.textContent = '▶ ดึง Orders';
            }
        }
        
        async function uploadOrders(input) {
            if (!input.files || !input.files[0]) return;
            const file = input.files[0];
            
            const formData = new FormData();
            formData.append('file', file);
            
            try {
                const resp = await fetch('/api/upload-orders', {
                    method: 'POST',
                    body: formData
                });
                const data = await resp.json();
                if (data.error) {
                    alert('❌ ' + data.error);
                } else {
                    alert(`✅ อัปโหลดสำเร็จ!\n\nไฟล์: ${data.filename}\nจำนวน orders: ${data.count}`);
                    loadOrders();
                }
            } catch(e) {
                alert('Error: ' + e.message);
            }
            input.value = '';
        }
        
        async function dismissSplit(btn, orderNumber) {
            try {
                btn.disabled = true;
                btn.textContent = '⏳';
                const resp = await fetch('/api/dismiss-split', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({order_number: orderNumber})
                });
                const data = await resp.json();
                if (data.error) {
                    alert('❌ ' + data.error);
                    btn.disabled = false;
                    btn.textContent = '✓ OK';
                } else {
                    btn.textContent = '✅';
                    btn.style.background = 'rgba(0,212,170,0.2)';
                    btn.style.color = 'var(--green)';
                    btn.style.borderColor = 'var(--green)';
                    // Reload orders after brief delay
                    setTimeout(() => loadOrders(), 500);
                }
            } catch(e) {
                alert('Error: ' + e.message);
                btn.disabled = false;
                btn.textContent = '✓ OK';
            }
        }
        
        async function syncSheet(dryRun) {
            // ตรวจ unsplit orders ก่อน Sync จริง
            const unsplit = currentOrders.filter(o => o.needs_split && !o.split_manually && o.status !== 'ยกเลิกแล้ว');
            if (!dryRun && unsplit.length > 0) {
                const names = unsplit.map(o => `\u2022 ${o.customer} (\u0e3f${parseFloat(o.price||0).toLocaleString()})`).join('\n');
                if (!confirm(`\u26a0\ufe0f มี ${unsplit.length} orders ที่ยังไม่ได้แยกรายการ (ราคา > \u0e3f890):\n\n${names}\n\nต้องการ Sync ต่อหรือไม่?`)) {
                    return;
                }
            }
            
            const month = getMonth();
            const btn = document.getElementById('btnSync');
            btn.disabled = true;
            
            try {
                const resp = await fetch('/api/sync-sheet', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({month: parseInt(month), dry_run: dryRun})
                });
                const data = await resp.json();
                if (data.error) {
                    alert(data.error);
                    btn.disabled = false;
                    return;
                }
                showStatus(data.message, data.task_id);
            } catch(e) {
                alert('Error: ' + e.message);
                btn.disabled = false;
            }
        }
        
        async function fetchBills() {
            const btn = document.getElementById('btnBills');
            btn.disabled = true;
            btn.textContent = '⏳ กำลังดึง...';
            
            try {
                const resp = await fetch('/api/fetch-bills', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({})
                });
                const data = await resp.json();
                if (data.error) {
                    alert(data.error);
                    btn.disabled = false;
                    btn.textContent = '▶ ดึงบิล';
                    return;
                }
                showStatus(data.message, data.task_id);
            } catch(e) {
                alert('Error: ' + e.message);
                btn.disabled = false;
                btn.textContent = '▶ ดึงบิล';
            }
        }
        
        async function saveOneShipping(btn, customer, orderDate, idx) {
            const row = btn.closest('tr');
            const shipDate = row.querySelector('.ship-date').value.trim();
            const shipCost = row.querySelector('.ship-cost').value.trim();
            
            if (!shipDate && !shipCost) {
                alert('กรุณากรอกวันที่ส่ง หรือ ค่าส่ง อย่างน้อย 1 อย่าง');
                return;
            }
            
            btn.textContent = '⏳';
            btn.disabled = true;
            
            try {
                const resp = await fetch('/api/update-shipping', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({
                        month: parseInt(getMonth()),
                        updates: [{
                            customer: customer,
                            order_date: orderDate,
                            ship_date: shipDate,
                            ship_cost: shipCost
                        }]
                    })
                });
                const data = await resp.json();
                if (data.error) {
                    alert('Error: ' + data.error);
                } else {
                    btn.textContent = '✅';
                    btn.classList.add('saved');
                    setTimeout(() => { btn.textContent = '💾'; btn.classList.remove('saved'); }, 3000);
                }
            } catch(e) {
                alert('Error: ' + e.message);
            }
            btn.disabled = false;
        }
        
        async function saveAllShipping() {
            const rows = document.querySelectorAll('#ordersBody tr[data-customer]');
            const updates = [];
            
            rows.forEach(row => {
                const customer = row.dataset.customer;
                const orderDate = row.dataset.orderDate;
                const shipDate = row.querySelector('.ship-date')?.value.trim() || '';
                const shipCost = row.querySelector('.ship-cost')?.value.trim() || '';
                
                if (shipDate || shipCost) {
                    updates.push({ customer, order_date: orderDate, ship_date: shipDate, ship_cost: shipCost });
                }
            });
            
            if (updates.length === 0) {
                alert('ไม่มีข้อมูลที่กรอก — กรุณากรอกวันที่ส่ง หรือ ค่าส่ง ก่อน');
                return;
            }
            
            const btn = document.getElementById('btnSaveAll');
            btn.disabled = true;
            btn.textContent = '⏳ กำลังบันทึก...';
            
            try {
                const resp = await fetch('/api/update-shipping', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({ month: parseInt(getMonth()), updates })
                });
                const data = await resp.json();
                if (data.error) {
                    alert('Error: ' + data.error);
                } else {
                    const msg = document.getElementById('saveMsg');
                    msg.textContent = `✅ บันทึก ${data.count} รายการเรียบร้อย!`;
                    msg.classList.add('active');
                    setTimeout(() => msg.classList.remove('active'), 5000);
                }
            } catch(e) {
                alert('Error: ' + e.message);
            }
            btn.disabled = false;
            btn.textContent = '💾 บันทึกทั้งหมดลง Sheet';
        }
        
        async function markShipped(btn, orderNumber, customer) {
            if (!orderNumber) {
                alert('ไม่มี Order Number');
                return;
            }
            if (!confirm(`ทำเครื่องหมายว่าจัดส่งแล้ว\nOrder: ${orderNumber}\nCustomer: ${customer}\n\nจะเปิด browser อัตโนมัติ`)) {
                return;
            }
            
            btn.disabled = true;
            btn.textContent = '⏳ ...';
            
            try {
                const resp = await fetch('/api/mark-shipped', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({ order_number: orderNumber, customer: customer })
                });
                const data = await resp.json();
                if (data.error) {
                    alert('Error: ' + data.error);
                    btn.textContent = '🚚 จัดส่ง';
                } else {
                    showStatus(data.message, data.task_id);
                    btn.textContent = '🚚 จัดส่ง';
                }
            } catch(e) {
                alert('Error: ' + e.message);
                btn.textContent = '🚚 จัดส่ง';
            }
            btn.disabled = false;
        }
        
        function showStatus(title, taskId) {
            const panel = document.getElementById('statusPanel');
            panel.classList.add('active');
            document.getElementById('statusTitle').textContent = title;
            document.getElementById('statusLog').textContent = 'เริ่มต้น...\n';
            document.getElementById('statusDot').className = 'status-dot';
            
            currentTaskId = taskId;
            if (pollInterval) clearInterval(pollInterval);
            pollInterval = setInterval(() => pollTask(taskId), 2000);
        }
        
        async function pollTask(taskId) {
            try {
                const resp = await fetch(`/api/task/${taskId}`);
                const data = await resp.json();
                
                document.getElementById('statusLog').textContent = data.output || '';
                // Auto scroll
                const log = document.getElementById('statusLog');
                log.scrollTop = log.scrollHeight;
                
                if (data.status === 'done' || data.status === 'error') {
                    clearInterval(pollInterval);
                    pollInterval = null;
                    
                    const dot = document.getElementById('statusDot');
                    dot.className = `status-dot ${data.status}`;
                    
                    const title = document.getElementById('statusTitle');
                    title.textContent = data.status === 'done' ? '✅ เสร็จแล้ว!' : '❌ เกิดข้อผิดพลาด';
                    
                    // Re-enable buttons
                    document.querySelectorAll('.btn').forEach(b => b.disabled = false);
                    document.getElementById('btnFetch').textContent = '▶ ดึง Orders';
                    document.getElementById('btnBills').textContent = '▶ ดึงบิล';
                    
                    // Reload orders
                    if (data.status === 'done') {
                        setTimeout(loadOrders, 1000);
                    }
                }
            } catch(e) {
                console.error('Poll error:', e);
            }
        }
        
        function hideStatus() {
            document.getElementById('statusPanel').classList.remove('active');
            if (pollInterval) { clearInterval(pollInterval); pollInterval = null; }
        }
        
        // ============ Split Modal Logic ============
        function openSplitModal(orderNumber) {
            const order = currentOrders.find(o => 
                (o.order_number || '').replace('#','') === orderNumber.replace('#','')
            );
            if (!order) {
                alert('ไม่พบ order: ' + orderNumber);
                return;
            }
            currentSplitOrder = order;
            
            // Fill info
            const infoEl = document.getElementById('splitModalInfo');
            infoEl.innerHTML = `
                <span class="modal-info-label">Order</span>
                <span class="modal-info-value">#${order.order_number || '-'}</span>
                <span class="modal-info-label">ลูกค้า</span>
                <span class="modal-info-value">${order.customer || '-'}</span>
                <span class="modal-info-label">ราคารวม</span>
                <span class="modal-info-value" style="color:var(--orange)">฿${parseFloat(order.price || 0).toLocaleString()}</span>
                <span class="modal-info-label">สินค้าเดิม</span>
                <span class="modal-info-value">${(order.products || []).map(p=>p.name).join(', ') || '-'}</span>
            `;
            
            // Pre-fill items from existing products or create 2 empty rows
            const itemsList = document.getElementById('splitItemsList');
            itemsList.innerHTML = '';
            
            const existingProducts = order.products || [];
            if (existingProducts.length > 1) {
                existingProducts.forEach(p => addSplitItem(p.name, p.price || ''));
            } else {
                // Start with 2 empty rows
                addSplitItem(existingProducts[0]?.name || '', '');
                addSplitItem('', '');
            }
            
            updateSplitTotal();
            document.getElementById('splitModal').classList.add('active');
        }
        
        function closeSplitModal() {
            document.getElementById('splitModal').classList.remove('active');
            currentSplitOrder = null;
        }
        
        function addSplitItem(name, price) {
            const list = document.getElementById('splitItemsList');
            const row = document.createElement('div');
            row.className = 'split-item-row';
            row.innerHTML = `
                <input type="text" class="item-name" placeholder="ชื่อสินค้า เช่น สบู่ HDG" value="${name || ''}"  oninput="updateSplitTotal()">
                <input type="text" class="item-price" placeholder="ราคา" value="${price || ''}" oninput="updateSplitTotal()">
                <button class="split-item-remove" onclick="removeSplitItem(this)">✕</button>
            `;
            list.appendChild(row);
            updateSplitTotal();
        }
        
        function removeSplitItem(btn) {
            const row = btn.closest('.split-item-row');
            row.style.opacity = '0';
            row.style.transform = 'translateX(20px)';
            row.style.transition = 'all 0.2s ease';
            setTimeout(() => { row.remove(); updateSplitTotal(); }, 200);
        }
        
        function updateSplitTotal() {
            const rows = document.querySelectorAll('#splitItemsList .split-item-row');
            let total = 0;
            rows.forEach(r => {
                const val = parseFloat(r.querySelector('.item-price')?.value || '0');
                if (!isNaN(val)) total += val;
            });
            
            const totalEl = document.getElementById('splitTotalValue');
            totalEl.textContent = `฿${total.toLocaleString()}`;
            
            // Check mismatch with original price
            if (currentSplitOrder) {
                const origPrice = parseFloat(currentSplitOrder.price || '0');
                if (total > 0 && Math.abs(total - origPrice) > 1) {
                    totalEl.classList.add('split-total-mismatch');
                    totalEl.classList.remove('split-total-value');
                    totalEl.textContent += ` (เดิม ฿${origPrice.toLocaleString()})`;
                } else {
                    totalEl.classList.remove('split-total-mismatch');
                }
            }
        }
        
        async function submitSplit() {
            if (!currentSplitOrder) return;
            
            const rows = document.querySelectorAll('#splitItemsList .split-item-row');
            const products = [];
            rows.forEach(r => {
                const name = r.querySelector('.item-name')?.value.trim();
                const price = r.querySelector('.item-price')?.value.trim();
                if (name) {
                    products.push({ name, price: price || '0' });
                }
            });
            
            if (products.length < 2) {
                alert('กรุณาเพิ่มอย่างน้อย 2 รายการสินค้า');
                return;
            }
            
            const btn = document.getElementById('btnSubmitSplit');
            btn.disabled = true;
            btn.textContent = '⏳ กำลังบันทึก...';
            
            try {
                const resp = await fetch('/api/split-order', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({
                        order_number: currentSplitOrder.order_number,
                        products: products
                    })
                });
                const data = await resp.json();
                if (data.error) {
                    alert('Error: ' + data.error);
                } else {
                    closeSplitModal();
                    // Reload to reflect changes
                    loadOrders();
                }
            } catch(e) {
                alert('Error: ' + e.message);
            }
            btn.disabled = false;
            btn.textContent = '✂️ บันทึกการแยก';
        }
        
        // Close modal on overlay click
        document.getElementById('splitModal').addEventListener('click', function(e) {
            if (e.target === this) closeSplitModal();
        });
        
        // Close modal on Escape key
        document.addEventListener('keydown', function(e) {
            if (e.key === 'Escape') closeSplitModal();
        });
        
        // Initial load
        loadOrders();
    </script>
</body>
</html>
"""


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print("\n" + "=" * 50)
    print("📦 HDG Orders Dashboard")
    print("=" * 50)
    print(f"\n🌐 เปิด browser ไปที่: http://localhost:{port}")
    print(f"📂 โฟลเดอร์: {BASE_DIR}")
    print(f"\nกด Ctrl+C เพื่อหยุด\n")
    
    app.run(host='0.0.0.0', port=port, debug=False)
