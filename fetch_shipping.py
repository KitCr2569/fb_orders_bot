"""
Fetch Shipping Bills v5 - ตาม flow จริงที่ user ทำ:
1. ไป Inbox → search ชื่อลูกค้า
2. คลิก conversation item ใน sidebar (ไม่ใช่ ผู้คน) 
3. แชทเปิดที่ข้อความล่าสุด → scroll ขึ้น เพื่อหาบิล
4. ดึงรูปบิลขนส่ง
"""
import json
import time
import re
import requests
import argparse
from pathlib import Path
from playwright.sync_api import sync_playwright

SESSION_DIR = Path(__file__).parent / "fb_session"
OUTPUT_DIR = Path(__file__).parent / "output"
BILLS_DIR = OUTPUT_DIR / "shipping_bills"
BILLS_DIR.mkdir(parents=True, exist_ok=True)

ASSET_ID = "114336388182180"
INBOX_URL = f"https://business.facebook.com/latest/inbox/all?asset_id={ASSET_ID}"


def open_customer_chat(page, customer, max_retries=2):
    """search ชื่อ → คลิก conversation item ใน sidebar (with retry)"""
    print(f"   🔍 หา '{customer}'...")
    
    # escape single quote สำหรับ JS
    customer_js = customer.replace("'", "\\'")

    for attempt in range(max_retries):
        if attempt > 0:
            print(f"      🔄 retry {attempt+1}/{max_retries}...")
            time.sleep(2)
        
        # 1. Search - clear แล้วพิมพ์ใหม่
        try:
            search_box = page.locator('input[placeholder*="ค้นหา"], input[placeholder*="Search"]').first
            if search_box.is_visible(timeout=5000):
                search_box.click()
                time.sleep(0.5)
                # triple click เพื่อ select all แล้ว clear
                search_box.click(click_count=3)
                time.sleep(0.3)
                search_box.fill('')
                time.sleep(0.5)
                search_box.fill(customer)
                time.sleep(5)  # รอผลค้นหานานขึ้น
            else:
                print(f"      ⚠️ ไม่เจอ search box")
                continue
        except Exception as e:
            print(f"      ⚠️ Search error: {e}")
            continue
        
        # 2. คลิก conversation item ใน sidebar  
        # ใช้ JS เพื่อหา span ที่มีชื่อลูกค้า อยู่ใน sidebar (x < 420)
        # หาตัวที่อยู่ใต้ "การสนทนาใน MESSENGER" (ไม่ใช่ section ผู้คน)
        
        click_target = page.evaluate(f"""
        () => {{
            const spans = document.querySelectorAll('span');
            const results = [];
            let messengerSectionY = 0;
            
            // หา "การสนทนาใน" หรือ "MESSENGER" section header
            for (const span of spans) {{
                const text = (span.innerText || '').trim();
                if (text.includes('การสนทนาใน') || text === 'MESSENGER') {{
                    const rect = span.getBoundingClientRect();
                    if (rect.x < 420) {{
                        messengerSectionY = Math.max(messengerSectionY, rect.y);
                    }}
                }}
            }}
            
            for (const span of spans) {{
                const text = (span.innerText || '').trim();
                if (text === '{customer_js}') {{
                    const rect = span.getBoundingClientRect();
                    if (rect.x < 420 && rect.y > 150 && rect.width > 20) {{
                        results.push({{
                            spanX: rect.x,
                            spanY: rect.y + rect.height / 2,
                            spanW: rect.width,
                            spanH: rect.height,
                            belowMessenger: rect.y > messengerSectionY
                        }});
                    }}
                }}
            }}
            
            if (results.length === 0) return null;
            
            // ให้ความสำคัญกับตัวที่อยู่ใต้ "การสนทนาใน MESSENGER"
            const belowMsgr = results.filter(r => r.belowMessenger);
            if (belowMsgr.length > 0) {{
                return belowMsgr[0];
            }}
            
            // fallback: เอาตัวสุดท้าย (y มากสุด)
            results.sort((a, b) => b.spanY - a.spanY);
            return results[0];
        }}
        """)
        
        if not click_target:
            # ลอง partial match ด้วยชื่อแรก
            first_name = customer.split()[0]
            click_target = page.evaluate(f"""
            () => {{
                const spans = document.querySelectorAll('span');
                for (const span of spans) {{
                    const text = (span.innerText || '').trim();
                    if (text.includes('{first_name}') && text.length < 50) {{
                        const rect = span.getBoundingClientRect();
                        if (rect.x < 420 && rect.y > 200 && rect.width > 20) {{
                            return {{
                                spanX: rect.x,
                                spanY: rect.y + rect.height / 2,
                                spanW: rect.width,
                                spanH: rect.height
                            }};
                        }}
                    }}
                }}
                return null;
            }}
            """)
            
            if not click_target:
                if attempt < max_retries - 1:
                    continue
                print(f"      ❌ ไม่เจอ '{customer}' ใน sidebar")
                return False
        
        # คลิก 2 ครั้ง ตาม flow จริงที่ user ทำ
        cx = 200  # กลาง sidebar
        cy = click_target['spanY']
        
        print(f"      📍 คลิกครั้งที่ 1: x={cx}, y={cy:.0f}")
        page.mouse.click(cx, cy)
        time.sleep(3)
        
        # หลังคลิกครั้งแรก sidebar อาจเปลี่ยน → หา position ใหม่
        page.keyboard.press('Escape')  # ปิด image viewer ถ้าเปิดมา
        time.sleep(1)
        
        click_target2 = page.evaluate(f"""
        () => {{
            const spans = document.querySelectorAll('span');
            for (const span of spans) {{
                const text = (span.innerText || '').trim();
                if (text === '{customer_js}' || text.includes('{customer_js}'.split(' ')[0])) {{
                    const rect = span.getBoundingClientRect();
                    if (rect.x < 420 && rect.y > 150) {{
                        return {{
                            x: rect.x + rect.width / 2,
                            y: rect.y + rect.height / 2
                        }};
                    }}
                }}
            }}
            return null;
        }}
        """)
        
        if click_target2:
            cx2, cy2 = click_target2['x'], click_target2['y']
            print(f"      📍 คลิกครั้งที่ 2: x={cx2:.0f}, y={cy2:.0f}")
            page.mouse.click(cx2, cy2)
            time.sleep(4)
        else:
            # fallback: คลิก x=200 ใน sidebar ตำแหน่งเดิม
            print(f"      📍 คลิกครั้งที่ 2 (fallback): x={cx}, y={cy:.0f}")
            page.mouse.click(cx, cy)
            time.sleep(4)
        
        # ตรวจสอบ header ว่าเปิดแชทถูกคนไหม
        header = page.evaluate("""
        () => {
            const main = document.querySelector('[role="main"]');
            if (main) {
                const h = main.querySelector('h2');
                if (h) return h.innerText || '';
                const spans = main.querySelectorAll('span[dir="auto"]');
                for (const span of spans) {
                    const text = span.innerText || '';
                    if (text.length > 2 && text.length < 60) return text;
                }
            }
            return '';
        }
        """)
        first = customer.split()[0]
        if first in header:
            print(f"      ✅ เปิดแชท '{header.strip()[:30]}'")
            return True
        else:
            print(f"      ⚠️ Header = '{header.strip()[:30]}' — อาจยังไม่ตรง")
            return True  # ลองต่อดูก่อน
    
    return False


def scroll_chat_up(page, scroll_count=5):
    """scroll ขึ้น เพื่อหาข้อความเก่ากว่า (บิลขนส่ง)"""
    
    # คลิกในพื้นที่แชทก่อน เพื่อให้ focus
    try:
        page.mouse.click(700, 400)
        time.sleep(0.5)
    except Exception:
        pass
    
    for i in range(scroll_count):
        # Mouse wheel scroll ขึ้น
        page.mouse.wheel(0, -2000)
        time.sleep(1)
    
    time.sleep(2)


def get_chat_images(page):
    """ดึงรูปจากแชท"""
    return page.evaluate("""
    () => {
        const images = [];
        const allImgs = document.querySelectorAll('img');
        
        allImgs.forEach(img => {
            const src = img.src || '';
            const rect = img.getBoundingClientRect();
            
            if (rect.width > 60 && rect.height > 60 && 
                src.includes('scontent') && !src.includes('profile') &&
                !src.includes('emoji') && !src.includes('avatar') &&
                rect.x > 380) {  // เฉพาะพื้นที่แชท
                
                const ratio = rect.height / rect.width;
                
                images.push({
                    src: src,
                    width: Math.round(rect.width),
                    height: Math.round(rect.height),
                    x: Math.round(rect.x),
                    y: Math.round(rect.y),
                    ratio: Math.round(ratio * 100) / 100,
                    isVertical: ratio > 1.2
                });
            }
        });
        
        // เรียงจาก y น้อย → มาก (บน → ล่าง, เก่า → ใหม่)
        images.sort((a, b) => a.y - b.y);
        return images;
    }
    """)


def download_image(url, filepath):
    try:
        resp = requests.get(url, timeout=15)
        if resp.status_code == 200:
            with open(filepath, 'wb') as f:
                f.write(resp.content)
            return True
    except Exception as e:
        print(f"      ❌ Download error: {e}")
    return False


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--json', help='Path to orders JSON file')
    args = parser.parse_args()
    
    print("\n" + "=" * 60)
    print("📦 Fetch Shipping Bills v5")
    print("=" * 60)
    
    if args.json:
        latest_json = Path(args.json)
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
    
    # ไป Inbox
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
        order_num = order.get('order_number', '').replace('#', '')
        order_date = order.get('date', '')
        print(f"\n[{idx+1}/{len(orders)}] {customer} | สั่ง {order_date}")
        
        # เปิดแชทลูกค้า
        found = open_customer_chat(page, customer)
        
        if not found:
            print(f"   ❌ ไม่เจอแชท")
            results.append({
                'customer': customer,
                'order_number': order_num,
                'order_date': order_date,
                'bill_images': [],
                'error': 'chat not found'
            })
            try:
                page.goto(INBOX_URL, wait_until="domcontentloaded", timeout=15000)
            except Exception:
                pass
            time.sleep(3)
            continue
        
        # Screenshot ก่อน scroll
        safe_name = re.sub(r'[^\w]', '_', customer)[:25]
        ss0 = BILLS_DIR / f"chat5_before_{safe_name}.png"
        page.screenshot(path=str(ss0), full_page=False)
        
        # Scroll ขึ้น เพื่อหาบิล (บิลส่งก่อนข้อความล่าสุด)
        print(f"   ⬆️ Scroll ขึ้นหาบิล...")
        scroll_chat_up(page, scroll_count=8)
        
        # Screenshot หลัง scroll
        ss1 = BILLS_DIR / f"chat5_after_{safe_name}.png"
        page.screenshot(path=str(ss1), full_page=False)
        
        # ดึงรูป
        images = get_chat_images(page)
        print(f"   📷 รูปทั้งหมด: {len(images)}")
        
        # แสดง info
        for i, img in enumerate(images[:10]):
            tag = "📄" if img.get('isVertical') else "🖼️"
            print(f"      {tag} [{i+1}] {img['width']}x{img['height']} (ratio={img['ratio']}) y={img['y']}")
        
        # Download รูปแนวตั้ง (บิล) ก่อน, แล้วรูปทั่วไป
        vertical = [img for img in images if img.get('isVertical')]
        horizontal = [img for img in images if not img.get('isVertical')]
        
        saved = []
        # ดาวน์โหลดบิล (แนวตั้ง) ก่อน
        for i, img in enumerate(vertical[:3]):
            filename = f"bill5_{safe_name}_v{i+1}.jpg"
            filepath = BILLS_DIR / filename
            if download_image(img['src'], str(filepath)):
                saved.append(str(filepath))
                print(f"   💾 📄บิล: {filename} ({img['width']}x{img['height']})")
        
        # ถ้าไม่มีแนวตั้ง ก็ดาวน์โหลดแนวนอนด้วย
        if not vertical:
            for i, img in enumerate(horizontal[:3]):
                filename = f"bill5_{safe_name}_h{i+1}.jpg"
                filepath = BILLS_DIR / filename
                if download_image(img['src'], str(filepath)):
                    saved.append(str(filepath))
                    print(f"   💾 🖼️รูป: {filename} ({img['width']}x{img['height']})")
        
        results.append({
            'customer': customer,
            'order_number': order_num,
            'order_date': order_date,
            'bill_images': saved,
            'screenshots': [str(ss0), str(ss1)],
            'total_images': len(images),
            'vertical_images': len(vertical)
        })
        
        # กลับ Inbox
        try:
            page.goto(INBOX_URL, wait_until="domcontentloaded", timeout=15000)
        except Exception:
            pass
        time.sleep(3)
    
    # Save results
    results_path = OUTPUT_DIR / "shipping_bills_v5.json"
    with open(results_path, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    
    print(f"\n{'=' * 60}")
    print(f"📊 สรุป:")
    for r in results:
        bills = len(r.get('bill_images', []))
        vert = r.get('vertical_images', 0)
        icon = "✅" if bills else "❌"
        err = f" [{r.get('error','')}]" if r.get('error') else ""
        print(f"   {icon} {r['customer']}: {bills} รูป ({vert} บิล) สั่ง {r.get('order_date', 'N/A')}{err}")
    
    print(f"\n   📁 {BILLS_DIR}")
    
    print("\n🛑 ปิด browser...")
    context.close()
    pw.stop()
    print("✅ เสร็จ!")


if __name__ == '__main__':
    main()
