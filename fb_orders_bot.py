"""
Facebook Orders Bot - ดึงข้อมูล orders จาก Facebook Business Suite
ใช้ Playwright เปิด browser → อ่าน orders → export เป็น CSV

วิธีใช้:
1. รันครั้งแรก: python fb_orders_bot.py --login
   (เปิด browser ให้ login Facebook → บันทึก session)
2. รันดึงข้อมูล: python fb_orders_bot.py
   (ดึง orders ทั้งหมด → export CSV)
3. กรองเดือน: python fb_orders_bot.py --month 2
   (ดึงเฉพาะเดือน ก.พ.)
4. กรองสถานะ: python fb_orders_bot.py --status slip
   (เฉพาะ "แนบสลิปแล้ว")
"""

import argparse
import calendar
import csv
import re
import json
import os
import sys
import time
from datetime import datetime
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

# === Configuration ===
ORDERS_URL_BASE = "https://business.facebook.com/latest/orders/orders_list/?asset_id=114336388182180"
SESSION_DIR = Path(__file__).parent / "fb_session"
OUTPUT_DIR = Path(__file__).parent / "output"


def get_month_timestamps(month, year=2026):
    """คำนวณ start_time/end_time สำหรับเดือนที่ระบุ (UTC+7 → Unix timestamp)"""
    import calendar as cal
    from datetime import datetime, timezone, timedelta
    
    tz_bkk = timezone(timedelta(hours=7))
    # เริ่มต้นเดือน 00:00:00 Bangkok time
    start_dt = datetime(year, month, 1, 0, 0, 0, tzinfo=tz_bkk)
    # สิ้นสุดเดือน: วันที่ 1 ของเดือนถัดไป 00:00:00 Bangkok time
    if month == 12:
        end_dt = datetime(year + 1, 1, 1, 0, 0, 0, tzinfo=tz_bkk)
    else:
        end_dt = datetime(year, month + 1, 1, 0, 0, 0, tzinfo=tz_bkk)
    
    start_ts = int(start_dt.timestamp())
    end_ts = int(end_dt.timestamp())
    return start_ts, end_ts


def build_orders_url(month=None, year=2026):
    """สร้าง URL พร้อม date filter สำหรับเดือนที่ระบุ"""
    url = ORDERS_URL_BASE
    if month:
        start_ts, end_ts = get_month_timestamps(month, year)
        url += f"&start_time={start_ts}&end_time={end_ts}"
    return url
THAI_MONTHS = {
    'ม.ค.': 1, 'ก.พ.': 2, 'มี.ค.': 3, 'เม.ย.': 4,
    'พ.ค.': 5, 'มิ.ย.': 6, 'ก.ค.': 7, 'ส.ค.': 8,
    'ก.ย.': 9, 'ต.ค.': 10, 'พ.ย.': 11, 'ธ.ค.': 12
}


def clean_product_name(raw_name):
    """ทำความสะอาดชื่อสินค้า ลบ suffix ที่ไม่จำเป็น"""
    name = raw_name.strip()
    
    # ลบ suffix ที่ไม่จำเป็น (เฉพาะคำลงท้าย)
    name = re.sub(r'ครับ\s*$', '', name)
    name = re.sub(r'ค่ะ\s*$', '', name)
    name = name.strip()
    # ลบ . ท้ายชื่อ
    name = re.sub(r'\.\s*$', '', name)
    
    # ลบ "ลด xxx บาท" suffix
    name = re.sub(r'\s*ลด\s*\d+\s*บาท\s*$', '', name)
    
    # กรณีพิเศษ: ชื่อสินค้ายุ่งๆ เช่น "Ptbk. และในส่วนที่เป็นกริบจับใช้เป็นLtbk.ครับ Nikon Z6iii"
    # → ต้องหารุ่นกล้อง/เลนส์ + pattern ลาย
    camera_brands = r'(?:Canon|Nikon|Sony|Fuji|Fujifilm|Panasonic|Olympus|Sigma|Tamron|Viltrox|Leica)'
    camera_models = r'(?:R\d+|Z\d+\w*|A\d+\w*|X-?\w+|GH\d+|OM-?\d+|RF\d+|EOS\s*R?\d*|a\d+\w*|Legion\s*go\d*)'
    
    # ถ้าชื่อมีคำอธิบายยาวๆ (> 60 ตัวอักษร) + มีชื่อกล้อง → ลองจัดใหม่
    if len(name) > 60:
        # หารุ่นกล้อง/อุปกรณ์
        model_match = re.search(rf'({camera_brands}\s*{camera_models}|{camera_models})', name, re.IGNORECASE)
        # หา pattern ลาย (คำย่อ 2-6 ตัวอักษร เช่น ptbk, mbbk, cmd, slpg)
        pattern_matches = re.findall(r'\b([a-zA-Z]{2,6})\b', name)
        # กรอง pattern ที่เป็นชื่อกล้อง/คำทั่วไปออก
        skip_words = {'and', 'the', 'for', 'pro', 'with', 'usb', 'hub', 'iii', 'mark'}
        patterns = [p.lower() for p in pattern_matches if p.lower() not in skip_words and len(p) >= 3]
        
        if model_match and patterns:
            model_name = model_match.group(1).strip()
            # ใช้ pattern ที่ไม่ใช่ชื่อรุ่น
            model_lower = model_name.lower()
            skin_patterns = [p for p in patterns if p not in model_lower]
            if skin_patterns:
                name = f"{model_name} ลาย {'/'.join(skin_patterns[:2])}"
    
    return name


def split_multi_items(product_name, total_price=''):
    """
    แยกรายการที่มี 'กับ' หรือ 'และ' เป็นหลายชิ้น
    ตัวอย่าง:
      "กล้อง Sony a7iii กับ เลนส์ Sony24-70 2.8 II ลาย mbbk"
      → [("Sony a7iii ลาย mbbk", ""), ("Sony24-70 2.8 II ลาย mbbk", "")]
    """
    # หา pattern "X กับ Y" หรือ "X และ Y"
    split_match = re.match(
        r'(?:กล้อง\s*)?(.+?)\s+(?:กับ|และ)\s+(?:เลนส์\s*)?(.+)',
        product_name, re.IGNORECASE
    )
    
    if split_match:
        part1 = split_match.group(1).strip()
        part2 = split_match.group(2).strip()
        
        # หา "ลาย xxx" ที่อยู่ท้ายสุด → ใช้กับทั้ง 2 ชิ้น
        pattern_match = re.search(r'ลาย\s+(\S+(?:\s+\S+)?)', part2)
        pattern_name = pattern_match.group(1) if pattern_match else ''
        
        # ถ้า part1 ไม่มี "ลาย" → เพิ่ม pattern จาก part2
        if 'ลาย' not in part1 and pattern_name:
            part1 = f"{part1} ลาย {pattern_name}"
        
        return [(part1, ''), (part2, '')]
    
    return [(product_name, total_price)]


def process_products(products, order_price=''):
    """Clean + Split products → return list of processed products"""
    result = []
    for p in products:
        raw_name = p.get('name', '')
        price = p.get('price', '') or order_price
        qty = p.get('qty', 1)
        
        # Clean ชื่อ
        cleaned = clean_product_name(raw_name)
        
        # แยกรายการ
        split_items = split_multi_items(cleaned, price)
        
        for name, item_price in split_items:
            result.append({
                'name': name,
                'price': item_price,
                'qty': qty
            })
    
    return result


class FacebookOrdersBot:
    def __init__(self, headless=False):
        self.headless = headless
        self.browser = None
        self.context = None
        self.page = None
        self.orders = []
        
    def start(self):
        """เริ่ม browser พร้อม session ที่บันทึกไว้"""
        self.playwright = sync_playwright().start()
        
        # ใช้ persistent context เพื่อเก็บ login session
        SESSION_DIR.mkdir(parents=True, exist_ok=True)
        
        self.context = self.playwright.chromium.launch_persistent_context(
            user_data_dir=str(SESSION_DIR),
            headless=self.headless,
            viewport={"width": 1920, "height": 1080},
            locale="th-TH",
            timezone_id="Asia/Bangkok",
            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox",
            ]
        )
        self.page = self.context.pages[0] if self.context.pages else self.context.new_page()
        print("✅ Browser เปิดแล้ว")
        
    def stop(self):
        """ปิด browser"""
        if self.context:
            self.context.close()
        if hasattr(self, 'playwright'):
            self.playwright.stop()
        print("🛑 Browser ปิดแล้ว")

    def login_flow(self):
        """เปิด browser ให้ user login Facebook ด้วยตัวเอง"""
        print("\n🔐 กรุณา Login Facebook ในหน้าต่าง browser ที่เปิดขึ้น")
        print("   หลัง login เสร็จ → กด Enter ที่นี่เพื่อบันทึก session\n")
        
        self.page.goto("https://www.facebook.com/", wait_until="domcontentloaded", timeout=60000)
        input("✋ กด Enter หลัง login เสร็จแล้ว... ")
        
        # ทดสอบว่า login สำเร็จ
        try:
            self.page.goto(ORDERS_URL, wait_until="domcontentloaded", timeout=60000)
        except PlaywrightTimeout:
            print("⏳ Facebook โหลดช้า แต่ไม่เป็นไร...")
        time.sleep(5)
        
        if "login" in self.page.url.lower():
            print("❌ ยังไม่ได้ login! ลองใหม่อีกครั้ง")
            return False
        
        print("✅ Login สำเร็จ! Session ถูกบันทึกแล้ว")
        print("   ครั้งต่อไปรัน: python fb_orders_bot.py (ไม่ต้อง login ใหม่)")
        return True
    
    def navigate_to_orders(self, month=None, year=2026):
        """ไปที่หน้า orders พร้อม date filter"""
        url = build_orders_url(month, year)
        month_label = f" เดือน {THAI_MONTHS.get(month, month)}/{year}" if month else " (ทั้งหมด)"
        print(f"📂 กำลังไปหน้า Orders{month_label}...")
        print(f"   URL: {url[:120]}...")
        try:
            self.page.goto(url, wait_until="domcontentloaded", timeout=60000)
        except PlaywrightTimeout:
            print("⏳ Facebook โหลดช้า รอสักครู่...")
        time.sleep(5)
        
        # ตรวจว่าต้อง login ไหม
        if "login" in self.page.url.lower():
            print("❌ Session หมดอายุ! กรุณารัน: python fb_orders_bot.py --login")
            return False
        
        print("✅ เปิดหน้า Orders แล้ว")
        return True
    
    def set_date_filter(self, month=None, year=2026):
        """ตั้งค่า filter วันที่ (ถ้าระบุเดือน)"""
        if month is None:
            print("📅 ไม่ได้ระบุเดือน — ดึงทุก order ที่แสดง")
            return True
            
        print(f"📅 กำลังตั้งค่า filter เดือน {month}/{year}...")
        
        # คลิก date range picker
        try:
            date_button = self.page.locator('[aria-label*="date"], [aria-label*="วัน"]').first
            if date_button.is_visible(timeout=3000):
                date_button.click()
                time.sleep(1)
            else:
                # ลองหา button ที่มีข้อความวันที่
                date_buttons = self.page.locator('div[role="button"]').all()
                for btn in date_buttons:
                    text = btn.inner_text()
                    if any(m in text for m in THAI_MONTHS.keys()) or '/' in text or '-' in text:
                        btn.click()
                        time.sleep(1)
                        break
        except Exception as e:
            print(f"⚠️ ไม่สามารถตั้งค่า filter วันที่: {e}")
            print("   จะดึงข้อมูลทั้งหมดที่แสดงแทน")
        
        return True
    
    def extract_orders_from_page(self):
        """ดึงข้อมูล orders จากหน้า (DOM parsing)"""
        print("📊 กำลังอ่าน orders จากหน้าจอ...")
        
        orders_data = self.page.evaluate("""
        () => {
            const orders = [];
            
            // หา order rows ในตาราง
            const rows = document.querySelectorAll('table tbody tr, [role="row"]');
            
            if (rows.length > 0) {
                rows.forEach(row => {
                    const cells = row.querySelectorAll('td, [role="cell"], [role="gridcell"]');
                    if (cells.length >= 3) {
                        const texts = Array.from(cells).map(c => c.innerText.trim());
                        orders.push({
                            raw: texts,
                            html: row.innerHTML.substring(0, 500)
                        });
                    }
                });
            }
            
            // ถ้าไม่เจอตาราง ลองหาแบบ list
            if (orders.length === 0) {
                // Facebook ใช้ list-based layout
                const listItems = document.querySelectorAll('[data-testid*="order"], [class*="order"]');
                listItems.forEach(item => {
                    orders.push({
                        raw: [item.innerText.trim()],
                        html: item.innerHTML.substring(0, 500)
                    });
                });
            }
            
            // ลองดึงจาก DOM structure ทั่วไป
            if (orders.length === 0) {
                // หา elements ที่มีข้อความ order-like
                const allDivs = document.querySelectorAll('div');
                const orderDivs = [];
                
                allDivs.forEach(div => {
                    const text = div.innerText || '';
                    // มองหา pattern: ชื่อ + # + วันที่ + สถานะ
                    if (text.includes('#') && text.includes('2026') && 
                        (text.includes('แนบสลิป') || text.includes('ยกเลิก') || 
                         text.includes('รอดำเนินการ') || text.includes('ชำระ'))) {
                        // ตรวจว่าไม่ใช่ parent element ใหญ่เกินไป
                        if (text.length < 500 && div.children.length < 20) {
                            orderDivs.push({
                                raw: [text],
                                html: div.innerHTML.substring(0, 500)
                            });
                        }
                    }
                });
                
                orders.push(...orderDivs);
            }
            
            return {
                ordersFound: orders.length,
                orders: orders.slice(0, 50), // limit to 50
                pageTitle: document.title,
                url: window.location.href,
                bodyText: document.body.innerText.substring(0, 2000)
            };
        }
        """)
        
        return orders_data
    
    def extract_order_details(self):
        """ดึงรายละเอียดจากทุก order โดยคลิกทีละรายการ"""
        print("🔍 กำลังดึงรายละเอียดจากแต่ละ order...")
        
        # หา order rows
        order_rows = self.page.evaluate("""
        () => {
            // Facebook Business Suite ใช้ div-based rows
            const results = [];
            const allElements = document.querySelectorAll('a[href*="order"], [role="link"][href*="order"], [role="row"]');
            
            // ลองหาจาก text patterns
            if (allElements.length === 0) {
                const spans = document.querySelectorAll('span');
                const seen = new Set();
                spans.forEach(span => {
                    const text = span.innerText || '';
                    if (text.startsWith('#') && text.length > 10 && !seen.has(text)) {
                        seen.add(text);
                        const parent = span.closest('[role="row"], tr, [class*="row"]') || span.parentElement?.parentElement;
                        if (parent) {
                            results.push({
                                orderNumber: text,
                                fullText: parent.innerText,
                                boundingBox: parent.getBoundingClientRect()
                            });
                        }
                    }
                });
            }
            
            return results;
        }
        """)
        
        print(f"   พบ {len(order_rows)} order rows")
        
        detailed_orders = []
        for i, row in enumerate(order_rows):
            try:
                # คลิกที่ order row
                bb = row.get('boundingBox', {})
                if bb.get('x') and bb.get('y'):
                    self.page.mouse.click(
                        bb['x'] + bb.get('width', 100) / 2,
                        bb['y'] + bb.get('height', 30) / 2
                    )
                    time.sleep(1.5)
                    
                    # ดึงรายละเอียดจาก detail panel
                    detail = self.page.evaluate("""
                    () => {
                        const detailPanel = document.querySelector('[class*="detail"], [class*="sidebar"], [role="dialog"]');
                        if (detailPanel) {
                            return {
                                text: detailPanel.innerText,
                                found: true
                            };
                        }
                        return { text: document.body.innerText.substring(0, 3000), found: false };
                    }
                    """)
                    
                    detailed_orders.append({
                        'index': i + 1,
                        'orderNumber': row.get('orderNumber', ''),
                        'rowText': row.get('fullText', ''),
                        'detailText': detail.get('text', ''),
                        'detailFound': detail.get('found', False)
                    })
                    
                    print(f"   [{i+1}/{len(order_rows)}] {row.get('orderNumber', 'N/A')}")
                    
            except Exception as e:
                print(f"   ⚠️ Error on order {i+1}: {e}")
        
        return detailed_orders
    
    def scroll_and_collect(self):
        """Scroll ลงเพื่อโหลด orders เพิ่ม แล้วเก็บข้อมูลทั้งหมด"""
        print("📜 กำลัง scroll เพื่อโหลด orders ทั้งหมด...")
        
        # นับจำนวน order ก่อน scroll
        prev_order_count = 0
        no_new_orders = 0
        scroll_attempts = 0
        max_scrolls = 50  # เพิ่มจาก 20 เป็น 50
        
        while scroll_attempts < max_scrolls:
            # นับ order numbers ปัจจุบัน
            current_count = self.page.evaluate("""
            () => {
                const spans = document.querySelectorAll('span');
                let count = 0;
                spans.forEach(s => {
                    if (/^#\d{10,}$/.test((s.innerText || '').trim())) count++;
                });
                return count;
            }
            """)
            
            if current_count == prev_order_count:
                no_new_orders += 1
                if no_new_orders >= 3:  # 3 ครั้งติดไม่มี order ใหม่ = หยุด
                    break
            else:
                no_new_orders = 0
                prev_order_count = current_count
            
            # Scroll ลง (ลอง scroll ทั้ง body และ scrollable container)
            self.page.evaluate("""
            () => {
                // ลอง scroll container ที่มี orders
                const containers = document.querySelectorAll('[role="main"], [class*="scroll"], [class*="list"]');
                let scrolled = false;
                containers.forEach(c => {
                    if (c.scrollHeight > c.clientHeight + 100) {
                        c.scrollBy(0, 600);
                        scrolled = true;
                    }
                });
                if (!scrolled) {
                    window.scrollBy(0, 800);
                }
            }
            """)
            time.sleep(1.5)
            scroll_attempts += 1
        
        print(f"   Scroll {scroll_attempts} ครั้ง, พบ {prev_order_count} orders")
        
        # ดึงข้อมูลจาก DOM elements โดยตรง
        print("📊 กำลังดึงข้อมูลจาก DOM...")
        orders = self.page.evaluate("""
        () => {
            const results = [];
            
            // Facebook Business Suite order list มี structure:
            // แต่ละ order row มี: ชื่อลูกค้า, #order_number, วันที่, สถานะ
            // ลองหาจาก column headers
            const allText = document.body.innerText;
            const lines = allText.split('\\n').map(l => l.trim()).filter(l => l);
            
            // Parse line by line - order records follow pattern:
            // [customer_name]
            // [#order_number]
            // [date]
            // [status]
            // [status] (duplicate)
            // [฿price]
            let i = 0;
            while (i < lines.length) {
                const line = lines[i];
                
                // ตรวจหา order number pattern
                if (/^#\\d{10,}$/.test(line)) {
                    const orderNum = line;
                    
                    // ชื่อลูกค้าอยู่บรรทัดก่อน
                    let customer = '';
                    if (i > 0 && !lines[i-1].match(/^#/) && !lines[i-1].match(/^\\d/) && 
                        !lines[i-1].match(/^฿/) &&
                        !['แนบสลิปแล้ว','ยกเลิกแล้ว','หมดเขต','รอดำเนินการ','คำสั่งซื้อ','ทั้งหมด','ลูกค้า','สถานะ','เวลาที่สร้าง','จำนวน'].includes(lines[i-1])) {
                        customer = lines[i-1];
                    }
                    
                    // วันที่ + สถานะ + ราคา อยู่บรรทัดถัดไป
                    let date = '';
                    let status = '';
                    let price = '';
                    
                    for (let j = i + 1; j < Math.min(i + 8, lines.length); j++) {
                        // วันที่ pattern: "7 มี.ค. 2026 20:57"
                        const dateMatch = lines[j].match(/(\\d{1,2})\\s+(ม\\.ค\\.|ก\\.พ\\.|มี\\.ค\\.|เม\\.ย\\.|พ\\.ค\\.|มิ\\.ย\\.|ก\\.ค\\.|ส\\.ค\\.|ก\\.ย\\.|ต\\.ค\\.|พ\\.ย\\.|ธ\\.ค\\.)\\s+(\\d{4})/);
                        if (dateMatch && !date) {
                            date = lines[j];
                        }
                        
                        // สถานะ
                        if (!status && (lines[j] === 'แนบสลิปแล้ว' || lines[j] === 'ยกเลิกแล้ว' || 
                            lines[j] === 'รอดำเนินการ' || lines[j] === 'หมดเขต')) {
                            status = lines[j];
                        }
                        
                        // ราคา pattern: ฿890.00 or ฿1,512.00
                        const priceMatch = lines[j].match(/^฿([\\d,]+\\.\\d{2})$/);
                        if (priceMatch && !price) {
                            price = priceMatch[1].replace(/,/g, '');
                        }
                        
                        // ถ้าเจอ # ถัดไป = order ใหม่ หยุด
                        if (j > i && /^#\\d{10,}$/.test(lines[j])) break;
                    }
                    
                    results.push({
                        order_number: orderNum,
                        customer: customer,
                        date_raw: date,
                        status: status,
                        price: price
                    });
                }
                i++;
            }
            
            return results;
        }
        """)
        
        # Save raw text for debugging
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        raw_text = self.page.evaluate("document.body.innerText")
        debug_path = OUTPUT_DIR / "debug_raw_text.txt"
        with open(debug_path, 'w', encoding='utf-8') as f:
            f.write(raw_text)
        
        print(f"   พบ {len(orders)} orders จาก DOM")
        return orders
    
    def fetch_order_details(self, orders):
        """คลิกเข้าแต่ละ order เพื่อดึงชื่อสินค้าจริง จาก detail panel ด้านขวา"""
        print(f"\n🔍 กำลังดึงรายละเอียดสินค้าจาก {len(orders)} orders...")
        
        all_details = []
        
        for idx, order in enumerate(orders):
            order_num = order.get('order_number', '').replace('#', '')
            print(f"   [{idx+1}/{len(orders)}] #{order_num} - {order.get('customer', '')}...", end=" ")
            
            try:
                # Navigate ตรงไปที่ order page โดยใช้ order_id ใน URL
                order_url = f"{ORDERS_URL_BASE}&order_id={order_num}"
                try:
                    self.page.goto(order_url, wait_until="domcontentloaded", timeout=30000)
                except PlaywrightTimeout:
                    pass
                time.sleep(4)  # รอ detail panel โหลด
                
                # ตรวจว่า detail panel แสดง order ที่ถูกต้อง
                page_text = self.page.evaluate("document.body.innerText")
                if order_num not in page_text:
                    print(f"⚠️ Panel ไม่แสดง #{order_num}, รอเพิ่ม...")
                    time.sleep(3)
                
                # ดึงข้อมูลจาก DETAIL PANEL ด้านขวา (x > 900)
                # Panel มี structure: ชื่อลูกค้า → หมายเลขคำสั่งซื้อ → รายละเอียดคำสั่งซื้อ → สินค้า
                detail = self.page.evaluate("""
                () => {
                    // หา elements ที่อยู่ด้านขวา (right panel)
                    const allDivs = document.querySelectorAll('div');
                    let panelText = '';
                    let bestPanel = null;
                    let bestLen = 0;
                    
                    allDivs.forEach(div => {
                        const rect = div.getBoundingClientRect();
                        const text = div.innerText || '';
                        
                        // Detail panel อยู่ด้านขวา (x > 900) และมี "รายละเอียดคำสั่งซื้อ"
                        if (rect.x > 900 && text.includes('รายละเอียดคำสั่งซื้อ') && 
                            text.length > 50 && text.length < 3000 &&
                            div.children.length < 30) {
                            // เอาตัวที่สั้นที่สุด (ใกล้ตัว detail section มากที่สุด)
                            if (!bestPanel || text.length < bestLen) {
                                bestPanel = div;
                                bestLen = text.length;
                                panelText = text;
                            }
                        }
                    });
                    
                    if (!panelText) {
                        // Fallback: ลองหาจาก body แต่เฉพาะส่วน detail
                        const body = document.body.innerText;
                        panelText = body;
                    }
                    
                    const lines = panelText.split('\\n').map(l => l.trim()).filter(l => l);
                    
                    const products = [];
                    let detailLines = [];
                    let inDetail = false;
                    
                    // หาส่วน "รายละเอียดคำสั่งซื้อ"
                    for (let i = 0; i < lines.length; i++) {
                        if (lines[i].includes('รายละเอียดคำสั่งซื้อ')) {
                            inDetail = true;
                            continue;
                        }
                        if (inDetail) {
                            detailLines.push(lines[i]);
                            // หยุดเมื่อเจอ section ถัดไป
                            if (lines[i].includes('สถานะคำสั่งซื้อ') || 
                                lines[i].includes('ข้อมูลการจัดส่ง') ||
                                lines[i].includes('การชำระเงิน') ||
                                detailLines.length > 30) {
                                detailLines.pop(); // เอาบรรทัด header ออก
                                break;
                            }
                        }
                    }
                    
                    // Parse products: ชื่อสินค้า → จำนวน: N → THBxxx.xx
                    if (detailLines.length > 0) {
                        let currentProduct = null;
                        
                        for (let i = 0; i < detailLines.length; i++) {
                            const line = detailLines[i];
                            
                            // ข้ามบรรทัดที่ไม่ใช่ product info
                            if (line === 'ยอดรวม' || line.includes('สถานะคำสั่งซื้อ')) break;
                            
                            // จำนวน pattern: "จำนวน: 1"
                            const qtyMatch = line.match(/จำนวน\\s*[:：]\\s*(\\d+)/);
                            if (qtyMatch) {
                                if (currentProduct) {
                                    currentProduct.qty = parseInt(qtyMatch[1]);
                                }
                                continue;
                            }
                            
                            // ราคา pattern: "THB890.00" or "THB1,512.00"
                            const priceMatch = line.match(/^THB\\s*([\\d,]+\\.\\d{2})$/);
                            if (priceMatch) {
                                if (currentProduct) {
                                    currentProduct.price = priceMatch[1].replace(/,/g, '');
                                    products.push({...currentProduct});
                                    currentProduct = null;
                                }
                                continue;
                            }
                            
                            // ข้ามบรรทัด "ยอดรวม" หรือราคารวม ที่ซ้ำ
                            if (line === 'ยอดรวม' || line.match(/^฿/) || line.match(/^THB/)) {
                                continue;
                            }
                            
                            // ชื่อสินค้า (ไม่ใช่จำนวน, ไม่ใช่ราคา, ยาว > 3 ตัวอักษร)
                            if (line.length > 3) {
                                // ถ้ามี product ก่อนหน้าที่ยังไม่มีราคา → บันทึกก่อน
                                if (currentProduct) {
                                    products.push({...currentProduct});
                                }
                                currentProduct = {
                                    name: line,
                                    price: '',
                                    qty: 1
                                };
                            }
                        }
                        
                        // ถ้ายังมี product ค้างอยู่
                        if (currentProduct) {
                            products.push(currentProduct);
                        }
                    }
                    
                    return {
                        products: products,
                        detailLines: detailLines,
                        panelFound: !!bestPanel,
                        panelTextLength: panelText.length,
                        url: window.location.href
                    };
                }
                """)
                
                raw_products = detail.get('products', [])
                # Clean + Split products (แยกรายการที่มี กับ/และ)
                order['products'] = process_products(raw_products, order.get('price', ''))
                product_names = [p.get('name', '') for p in order['products']]
                
                all_details.append({
                    'order_number': order.get('order_number', ''),
                    'customer': order.get('customer', ''),
                    'products': order['products'],
                    'detail_lines': detail.get('detailLines', [])[:20],
                    'panel_found': detail.get('panelFound', False)
                })
                
                panel_icon = "📋" if detail.get('panelFound') else "⚠️"
                if product_names:
                    prices = [p.get('price', '') for p in order['products']]
                    print(f"{panel_icon} ✅ {', '.join(product_names[:3])} | THB{', '.join(prices[:3])}")
                else:
                    print(f"{panel_icon} ⚠️ ไม่เจอสินค้า (detailLines: {len(detail.get('detailLines', []))})")
                
            except Exception as e:
                print(f"❌ Error: {e}")
                order['products'] = []
        
        # Save debug details
        debug_path = OUTPUT_DIR / "debug_order_details.json"
        with open(debug_path, 'w', encoding='utf-8') as f:
            json.dump(all_details, f, ensure_ascii=False, indent=2)
        print(f"\n   💾 Debug details: {debug_path}")
        
        return orders
    
    def parse_orders_text(self, text, status_filter=None, month_filter=None):
        """แยก orders จาก text ที่ดึงมา"""
        import re
        
        lines = text.split('\n')
        orders = []
        current_order = {}
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # หา order number (#xxxx)
            order_match = re.search(r'#(\d{10,})', line)
            if order_match:
                if current_order:
                    orders.append(current_order)
                current_order = {
                    'order_number': '#' + order_match.group(1),
                    'raw_text': line
                }
                continue
            
            # หาชื่อลูกค้า (ชื่อที่อยู่ก่อน #)
            if current_order and 'customer' not in current_order:
                name_match = re.search(r'^([^\d#].{2,30})$', line)
                if name_match:
                    current_order['customer'] = name_match.group(1).strip()
            
            # หาวันที่
            date_match = re.search(r'(\d{1,2})\s*(ม\.ค\.|ก\.พ\.|มี\.ค\.|เม\.ย\.|พ\.ค\.|มิ\.ย\.|ก\.ค\.|ส\.ค\.|ก\.ย\.|ต\.ค\.|พ\.ย\.|ธ\.ค\.)\s*(\d{4})', line)
            if date_match and current_order:
                day = int(date_match.group(1))
                month = THAI_MONTHS.get(date_match.group(2), 0)
                year = int(date_match.group(3))
                current_order['date'] = f"{day}/{month}/{year}"
                current_order['month'] = month
                current_order['year'] = year
            
            # หาสถานะ
            if current_order:
                if 'แนบสลิปแล้ว' in line:
                    current_order['status'] = 'แนบสลิปแล้ว'
                elif 'ยกเลิกแล้ว' in line:
                    current_order['status'] = 'ยกเลิกแล้ว'
                elif 'รอดำเนินการ' in line:
                    current_order['status'] = 'รอดำเนินการ'
                elif 'ชำระเงินแล้ว' in line:
                    current_order['status'] = 'ชำระเงินแล้ว'
            
            # หาราคา
            price_match = re.search(r'THB\s*([\d,]+\.?\d*)', line)
            if price_match and current_order:
                current_order['price'] = price_match.group(1).replace(',', '')
        
        # เพิ่ม order สุดท้าย
        if current_order:
            orders.append(current_order)
        
        # กรอง
        filtered = orders
        if status_filter:
            if status_filter == 'slip':
                filtered = [o for o in filtered if o.get('status') == 'แนบสลิปแล้ว']
            elif status_filter == 'cancelled':
                filtered = [o for o in filtered if o.get('status') == 'ยกเลิกแล้ว']
        
        if month_filter:
            filtered = [o for o in filtered if o.get('month') == month_filter]
        
        return filtered
    
    def export_csv(self, orders, filename=None):
        """Export orders เป็น CSV"""
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        
        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"orders_{timestamp}.csv"
        
        filepath = OUTPUT_DIR / filename
        
        # ตรวจว่ามี product details หรือไม่
        has_products = any(order.get('products') for order in orders)
        
        with open(filepath, 'w', newline='', encoding='utf-8-sig') as f:
            if has_products:
                # Format แบบมีรายละเอียดสินค้า (ตาม format sheet)
                fieldnames = ['date', 'product_name', 'price', 'customer', 'order_number', 'status']
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                for order in orders:
                    products = order.get('products', [])
                    if products:
                        for p in products:
                            writer.writerow({
                                'date': order.get('date', ''),
                                'product_name': p.get('name', ''),
                                'price': p.get('price', order.get('price', '')),
                                'customer': f"fb.{order.get('customer', '')}",
                                'order_number': order.get('order_number', ''),
                                'status': order.get('status', '')
                            })
                    else:
                        writer.writerow({
                            'date': order.get('date', ''),
                            'product_name': '(ไม่ทราบชื่อสินค้า)',
                            'price': order.get('price', ''),
                            'customer': f"fb.{order.get('customer', '')}",
                            'order_number': order.get('order_number', ''),
                            'status': order.get('status', '')
                        })
            else:
                # Format แบบไม่มีรายละเอียดสินค้า
                fieldnames = ['order_number', 'customer', 'date', 'status', 'price']
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                for order in orders:
                    writer.writerow({
                        'order_number': order.get('order_number', ''),
                        'customer': order.get('customer', ''),
                        'date': order.get('date', ''),
                        'status': order.get('status', ''),
                        'price': order.get('price', '')
                    })
        
        print(f"\n📁 Export เสร็จ: {filepath}")
        print(f"   จำนวน: {len(orders)} รายการ")
        return filepath
    
    def export_json(self, orders, filename=None):
        """Export orders เป็น JSON"""
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        
        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"orders_{timestamp}.json"
        
        filepath = OUTPUT_DIR / filename
        
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(orders, f, ensure_ascii=False, indent=2)
        
        print(f"📁 Export JSON: {filepath}")
        return filepath
    
    def run(self, month=None, status=None, export_format='csv', fetch_details=False):
        """รัน bot ดึง orders"""
        self._fetch_details = fetch_details
        print("\n" + "=" * 60)
        print("🤖 Facebook Orders Bot - HDG Wrap Skin")
        print("=" * 60)
        
        self.start()
        
        try:
            if not self.navigate_to_orders(month=month):
                return None
            
            # รอให้ orders โหลด
            time.sleep(3)
            
            # ดึง DOM data ก่อน
            print("\n📊 กำลังวิเคราะห์หน้า Orders...")
            dom_data = self.extract_orders_from_page()
            
            print(f"   📄 Page: {dom_data.get('pageTitle', 'N/A')}")
            print(f"   📦 พบ elements: {dom_data.get('ordersFound', 0)}")
            
            # Scroll เพื่อโหลดทั้งหมด + ดึง DOM orders
            dom_orders = self.scroll_and_collect()
            
            # Parse dates and apply filters
            import re
            orders = []
            for o in dom_orders:
                # Parse date from date_raw
                date_raw = o.get('date_raw', '')
                date_match = re.search(r'(\d{1,2})\s*(ม\.ค\.|ก\.พ\.|มี\.ค\.|เม\.ย\.|พ\.ค\.|มิ\.ย\.|ก\.ค\.|ส\.ค\.|ก\.ย\.|ต\.ค\.|พ\.ย\.|ธ\.ค\.)\s*(\d{4})', date_raw)
                if date_match:
                    day = int(date_match.group(1))
                    m = THAI_MONTHS.get(date_match.group(2), 0)
                    year = int(date_match.group(3))
                    o['date'] = f"{day}/{m}/{year}"
                    o['month'] = m
                    o['year'] = year
                
                # กรองสถานะ
                if status == 'slip' and o.get('status') != 'แนบสลิปแล้ว':
                    continue
                if status == 'cancelled' and o.get('status') != 'ยกเลิกแล้ว':
                    continue
                
                # กรองเดือน
                if month and o.get('month') != month:
                    continue
                
                orders.append(o)
            
            print(f"\n✅ ดึง orders ได้ {len(orders)} รายการ (จากทั้งหมด {len(dom_orders)})")
            
            if orders:
                # แสดงตัวอย่าง
                print("\n📋 ตัวอย่าง orders:")
                for i, o in enumerate(orders[:10]):
                    price_str = f" | ฿{o.get('price', 'N/A')}" if o.get('price') else ''
                    print(f"   {i+1}. {o.get('date', 'N/A')} | {o.get('customer', 'N/A')} | {o.get('status', 'N/A')}{price_str}")
                
                if len(orders) > 10:
                    print(f"   ... (อีก {len(orders)-10} รายการ)")
                
                # ดึงรายละเอียดสินค้า (ถ้า flag --details)
                if self._fetch_details:
                    orders = self.fetch_order_details(orders)
                
                # Export
                if export_format == 'csv':
                    self.export_csv(orders)
                elif export_format == 'json':
                    self.export_json(orders)
                elif export_format == 'both':
                    self.export_csv(orders)
                    self.export_json(orders)
            else:
                print("\n⚠️ ไม่พบ orders ที่ตรงเงื่อนไข")
                print(f"   บันทึก raw text: output/debug_raw_text.txt")
            
            return orders
            
        except Exception as e:
            print(f"\n❌ Error: {e}")
            import traceback
            traceback.print_exc()
            return None
        finally:
            self.stop()


def main():
    parser = argparse.ArgumentParser(description='Facebook Orders Bot - HDG Wrap Skin')
    parser.add_argument('--login', action='store_true', help='เปิด browser เพื่อ login Facebook')
    parser.add_argument('--month', type=int, help='กรองเดือน (1-12)')
    parser.add_argument('--status', choices=['slip', 'cancelled', 'all'], default='all',
                       help='กรองสถานะ: slip=แนบสลิปแล้ว, cancelled=ยกเลิก, all=ทั้งหมด')
    parser.add_argument('--format', choices=['csv', 'json', 'both'], default='csv',
                       help='รูปแบบ export')
    parser.add_argument('--headless', action='store_true', help='รันแบบไม่แสดง browser')
    parser.add_argument('--details', action='store_true', help='ดึงรายละเอียดสินค้าจากแต่ละ order')
    
    args = parser.parse_args()
    
    bot = FacebookOrdersBot(headless=args.headless)
    
    if args.login:
        bot.start()
        try:
            bot.login_flow()
        finally:
            bot.stop()
    else:
        status_filter = None if args.status == 'all' else args.status
        bot.run(
            month=args.month,
            status=status_filter,
            export_format=args.format,
            fetch_details=args.details
        )


if __name__ == '__main__':
    main()
