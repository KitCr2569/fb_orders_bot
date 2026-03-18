#!/usr/bin/env python3
"""
Fetch Shipping Info v6 - ดึงข้อมูลค่าส่ง + วันที่ส่ง จากแชท
1. เปิดแชทลูกค้า
2. Scroll หาข้อความเกี่ยวกับการส่ง (เลขพัสดุ, ค่าส่ง, วันที่ส่ง)
3. ดึงจากรูปบิลขนส่งและข้อความใกล้เคียง
"""
import json
import time
import re
import argparse
from pathlib import Path
from playwright.sync_api import sync_playwright

SESSION_DIR = Path(__file__).parent / "fb_session"
OUTPUT_DIR = Path(__file__).parent / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

ASSET_ID = "114336388182180"
INBOX_URL = f"https://business.facebook.com/latest/inbox/all?asset_id={ASSET_ID}"


def open_customer_chat(page, customer, max_retries=2):
    """search ชื่อ → คลิก conversation item ใน sidebar"""
    print(f"   🔍 หา '{customer}'...")
    customer_js = customer.replace("'", "\\'")

    for attempt in range(max_retries):
        if attempt > 0:
            print(f"      🔄 retry {attempt+1}...")
            time.sleep(2)
        
        try:
            search_box = page.locator('input[placeholder*="ค้นหา"], input[placeholder*="Search"]').first
            if search_box.is_visible(timeout=5000):
                search_box.click()
                time.sleep(0.5)
                search_box.click(click_count=3)
                time.sleep(0.3)
                search_box.fill('')
                time.sleep(0.5)
                search_box.fill(customer)
                time.sleep(5)
        except Exception as e:
            print(f"      ⚠️ Search error: {e}")
            continue
        
        click_target = page.evaluate(f"""
        () => {{
            const spans = document.querySelectorAll('span');
            const results = [];
            let messengerSectionY = 0;
            
            for (const span of spans) {{
                const text = (span.innerText || '').trim();
                if (text.includes('การสนทนาใน') || text === 'MESSENGER') {{
                    const rect = span.getBoundingClientRect();
                    if (rect.x < 420) messengerSectionY = Math.max(messengerSectionY, rect.y);
                }}
            }}
            
            for (const span of spans) {{
                const text = (span.innerText || '').trim();
                if (text === '{customer_js}') {{
                    const rect = span.getBoundingClientRect();
                    if (rect.x < 420 && rect.y > 150 && rect.width > 20) {{
                        results.push({{
                            spanY: rect.y + rect.height / 2,
                            belowMessenger: rect.y > messengerSectionY
                        }});
                    }}
                }}
            }}
            
            if (results.length === 0) return null;
            const below = results.filter(r => r.belowMessenger);
            return below.length > 0 ? below[0] : results[results.length - 1];
        }}
        """)
        
        if not click_target:
            first_name = customer.split()[0]
            click_target = page.evaluate(f"""
            () => {{
                const spans = document.querySelectorAll('span');
                for (const span of spans) {{
                    const text = (span.innerText || '').trim();
                    if (text.includes('{first_name}') && text.length < 50) {{
                        const rect = span.getBoundingClientRect();
                        if (rect.x < 420 && rect.y > 200 && rect.width > 20) {{
                            return {{ spanY: rect.y + rect.height / 2 }};
                        }}
                    }}
                }}
                return null;
            }}
            """)
            if not click_target:
                if attempt < max_retries - 1: continue
                return False
        
        cx, cy = 200, click_target['spanY']
        page.mouse.click(cx, cy)
        time.sleep(3)
        page.keyboard.press('Escape')
        time.sleep(1)
        page.mouse.click(cx, cy)
        time.sleep(4)
        return True
    return False


def get_chat_text_messages(page):
    """ดึงข้อความทั้งหมดจากแชท (โดยเฉพาะข้อความเกี่ยวกับการส่ง)"""
    return page.evaluate("""
    () => {
        const messages = [];
        // หา message bubbles ในพื้นที่แชท (x > 380)
        const allDivs = document.querySelectorAll('div[dir="auto"], span[dir="auto"]');
        
        allDivs.forEach(el => {
            const rect = el.getBoundingClientRect();
            const text = (el.innerText || '').trim();
            
            // เฉพาะพื้นที่แชท (ไม่ใช่ sidebar หรือ header)
            if (rect.x > 380 && text.length > 2 && text.length < 1000 && rect.width > 30) {
                // ข้ามชื่อ header, timestamps, etc.
                if (text.includes('ข้อความตอบกลับ') || text.includes('ตอบกลับใน') || 
                    text.includes('คลิกเพื่อแทรก')) return;
                
                messages.push({
                    text: text,
                    x: Math.round(rect.x),
                    y: Math.round(rect.y),
                    width: Math.round(rect.width),
                    height: Math.round(rect.height)
                });
            }
        });
        
        // unique by text + y position
        const seen = new Set();
        return messages.filter(m => {
            const key = m.text.substring(0, 50) + '_' + Math.round(m.y / 20);
            if (seen.has(key)) return false;
            seen.add(key);
            return true;
        }).sort((a, b) => a.y - b.y);
    }
    """)


def extract_shipping_info(messages):
    """วิเคราะห์ข้อความหาข้อมูลการส่ง"""
    info = {
        'tracking_numbers': [],
        'shipping_cost': '',
        'shipping_date': '',
        'carrier': '',
        'raw_shipping_messages': []
    }
    
    shipping_keywords = ['ส่งให้แล้ว', 'ส่งแล้ว', 'เลขพัสดุ', 'tracking', 'EMS', 'Kerry', 'KEX', 
                         'Flash', 'J&T', 'SPX', 'ค่าส่ง', 'PBSK', 'TH', 'แจ้งเลข', 'จัดส่ง',
                         'นำส่ง', 'ส่งของ', 'พัสดุ']
    
    for msg in messages:
        text = msg['text']
        text_lower = text.lower()
        
        is_shipping = any(kw.lower() in text_lower for kw in shipping_keywords)
        
        if is_shipping:
            info['raw_shipping_messages'].append(text[:200])
        
        # หาเลขพัสดุ
        # KEX pattern: PBSK + digits
        for match in re.finditer(r'(PBSK\w+)', text):
            if match.group(1) not in info['tracking_numbers']:
                info['tracking_numbers'].append(match.group(1))
        
        # Flash/Kerry/etc
        for match in re.finditer(r'(TH\d{13,20}|SPXTH\w+|JT\d+|KERTH\w+)', text):
            if match.group(1) not in info['tracking_numbers']:
                info['tracking_numbers'].append(match.group(1))
        
        # หาค่าส่ง
        cost_match = re.search(r'ค่าส่ง\s*[:=]?\s*(\d+)', text)
        if cost_match:
            info['shipping_cost'] = cost_match.group(1)
        
        # ราคารวมค่าส่ง
        cost_match2 = re.search(r'(\d+)\s*(?:บาท|฿)\s*(?:ค่าส่ง|shipping)', text, re.IGNORECASE)
        if cost_match2 and not info['shipping_cost']:
            info['shipping_cost'] = cost_match2.group(1)
        
        # หา carrier
        if 'KEX' in text or 'Kerry' in text:
            info['carrier'] = 'KEX'
        elif 'Flash' in text:
            info['carrier'] = 'Flash'
        elif 'J&T' in text:
            info['carrier'] = 'J&T'
        elif 'EMS' in text or 'ไปรษณีย์' in text:
            info['carrier'] = 'EMS'
        elif 'SPX' in text:
            info['carrier'] = 'Shopee Express'
        
        # หาวันที่ส่ง (Thai format: 17/3/69, 17 มี.ค. 69)
        date_match = re.search(r'(\d{1,2})[/.](\d{1,2})[/.](\d{2,4})', text)
        if date_match and is_shipping:
            d, m, y = date_match.group(1), date_match.group(2), date_match.group(3)
            if len(y) == 2:
                y = str(int(y) + 2500 - 543)  # BE to CE
            info['shipping_date'] = f"{d}/{m}/{y}"
    
    return info


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--json', help='Path to orders JSON file')
    args = parser.parse_args()
    
    print("\n" + "=" * 60)
    print("📦 Fetch Shipping Info v6 - ค่าส่ง + วันที่ส่ง")
    print("=" * 60)
    
    if args.json:
        latest_json = Path(args.json)
    else:
        json_files = sorted(OUTPUT_DIR.glob("orders_march2026_active.json"))
        if json_files:
            latest_json = json_files[-1]
        else:
            json_files = sorted(OUTPUT_DIR.glob("orders_2026*.json"))
            latest_json = json_files[-1]
    
    print(f"📂 ใช้: {latest_json.name}")
    
    with open(latest_json, 'r', encoding='utf-8') as f:
        orders = json.load(f)
    
    print(f"📦 พบ {len(orders)} orders")
    
    # เปิด browser
    pw = sync_playwright().start()
    context = pw.chromium.launch_persistent_context(
        user_data_dir=str(SESSION_DIR),
        headless=False,
        viewport={"width": 1920, "height": 1080},
        locale="th-TH",
        timezone_id="Asia/Bangkok",
        args=["--disable-blink-features=AutomationControlled", "--no-sandbox"]
    )
    page = context.pages[0] if context.pages else context.new_page()
    
    print("📂 กำลังเปิด Inbox...")
    try:
        page.goto(INBOX_URL, wait_until="domcontentloaded", timeout=30000)
    except Exception:
        pass
    time.sleep(5)
    print("✅ เปิด Inbox แล้ว")
    
    results = []
    
    for idx, order in enumerate(orders):
        customer = order.get('customer', '')
        order_date = order.get('date', '')
        print(f"\n[{idx+1}/{len(orders)}] {customer} | สั่ง {order_date}")
        
        found = open_customer_chat(page, customer)
        
        if not found:
            print(f"   ❌ ไม่เจอแชท")
            results.append({
                'customer': customer,
                'order_date': order_date,
                'shipping_info': {},
                'error': 'chat not found'
            })
            try:
                page.goto(INBOX_URL, wait_until="domcontentloaded", timeout=15000)
            except Exception:
                pass
            time.sleep(3)
            continue
        
        # Scroll ลงล่าง (ไปข้อความล่าสุด) แล้ว scroll ขึ้นหาข้อความส่ง
        # กด End เพื่อไปล่างสุด
        print(f"   📜 กำลังอ่านข้อความ...")
        
        # Scroll ขึ้น 3 ครั้ง เพื่อหาข้อความเก่ากว่า
        page.mouse.click(700, 400)
        time.sleep(0.5)
        for _ in range(5):
            page.mouse.wheel(0, -2000)
            time.sleep(1)
        time.sleep(2)
        
        # ดึงข้อความ
        messages = get_chat_text_messages(page)
        print(f"   💬 ข้อความ: {len(messages)}")
        
        # วิเคราะห์
        shipping_info = extract_shipping_info(messages)
        
        # แสดงผล
        if shipping_info['tracking_numbers']:
            print(f"   📦 Tracking: {', '.join(shipping_info['tracking_numbers'])}")
        if shipping_info['shipping_cost']:
            print(f"   💰 ค่าส่ง: {shipping_info['shipping_cost']} บาท")
        if shipping_info['shipping_date']:
            print(f"   📅 วันที่ส่ง: {shipping_info['shipping_date']}")
        if shipping_info['carrier']:
            print(f"   🚚 ขนส่ง: {shipping_info['carrier']}")
        if shipping_info['raw_shipping_messages']:
            for msg in shipping_info['raw_shipping_messages'][:3]:
                print(f"   📝 \"{msg[:80]}\"")
        
        if not shipping_info['tracking_numbers'] and not shipping_info['raw_shipping_messages']:
            print(f"   ⚠️ ไม่เจอข้อมูลการส่ง — อาจต้อง scroll มากกว่านี้")
            # ลอง scroll ขึ้นอีก
            for _ in range(5):
                page.mouse.wheel(0, -2000)
                time.sleep(1)
            time.sleep(2)
            
            messages2 = get_chat_text_messages(page)
            shipping_info2 = extract_shipping_info(messages2)
            if shipping_info2['tracking_numbers'] or shipping_info2['raw_shipping_messages']:
                shipping_info = shipping_info2
                if shipping_info['tracking_numbers']:
                    print(f"   📦 Tracking (retry): {', '.join(shipping_info['tracking_numbers'])}")
                if shipping_info['raw_shipping_messages']:
                    for msg in shipping_info['raw_shipping_messages'][:3]:
                        print(f"   📝 (retry) \"{msg[:80]}\"")
        
        results.append({
            'customer': customer,
            'order_date': order_date,
            'order_number': order.get('order_number', ''),
            'shipping_info': shipping_info
        })
        
        # กลับ Inbox
        try:
            page.goto(INBOX_URL, wait_until="domcontentloaded", timeout=15000)
        except Exception:
            pass
        time.sleep(3)
    
    # Save results
    results_path = OUTPUT_DIR / "shipping_info_v6.json"
    with open(results_path, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    
    print(f"\n{'=' * 60}")
    print(f"📊 สรุป:")
    for r in results:
        si = r.get('shipping_info', {})
        tracking = ', '.join(si.get('tracking_numbers', []))[:30] or '-'
        cost = si.get('shipping_cost', '-') or '-'
        date = si.get('shipping_date', '-') or '-'
        carrier = si.get('carrier', '-') or '-'
        print(f"   {r['customer']:<25} | ส่ง {date:<12} | ค่าส่ง {cost:<6} | {carrier:<6} | {tracking}")
    
    print(f"\n💾 {results_path}")
    
    print("\n🛑 ปิด browser...")
    context.close()
    pw.stop()
    print("✅ เสร็จ!")


if __name__ == '__main__':
    main()
