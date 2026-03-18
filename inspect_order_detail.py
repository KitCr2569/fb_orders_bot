"""
Inspect Order Detail - เปิด order detail page เพื่อดู DOM structure จริงๆ
ดูว่าชื่อสินค้าจริง (ที่คุณพิมพ์ตอนสร้างคำสั่งซื้อ) อยู่ตรงไหนใน DOM
"""
import json
import time
from pathlib import Path
from playwright.sync_api import sync_playwright

SESSION_DIR = Path(__file__).parent / "fb_session"
OUTPUT_DIR = Path(__file__).parent / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Orders จากข้อมูลที่ดึงมาเมื่อ 17 มี.ค. 2026
SAMPLE_ORDERS = [
    {"order_number": "1575080583722849", "customer": "Jeerasak Janson", "price": "890"},
    {"order_number": "3841196412855574", "customer": "Watchara Samsuvan", "price": "990"},
    {"order_number": "1418218653124538", "customer": "Nasri E'eso", "price": "1512"},
    {"order_number": "773042985873899", "customer": "Akkasart Samrandee", "price": "890"},
]

ASSET_ID = "114336388182180"


def inspect_order(page, order_num, customer):
    """ดึงข้อมูลจาก order detail panel อย่างละเอียด"""
    
    # URL สำหรับเปิด order detail
    detail_url = f"https://business.facebook.com/latest/orders/orders_list?asset_id={ASSET_ID}&order_id={order_num}"
    
    print(f"\n📂 กำลังเปิด order #{order_num} ({customer})...")
    try:
        page.goto(detail_url, wait_until="domcontentloaded", timeout=30000)
    except Exception:
        print("⏳ โหลดช้า...")
    
    time.sleep(5)
    
    # Screenshot
    ss_path = OUTPUT_DIR / f"order_detail_{order_num}.png"
    page.screenshot(path=str(ss_path), full_page=False)
    print(f"📸 Screenshot: {ss_path}")
    
    # ดึง DOM อย่างละเอียด
    detail_data = page.evaluate("""
    () => {
        const result = {
            url: window.location.href,
            title: document.title,
        };
        
        // ดึง text ทั้งหมดจาก body
        result.bodyText = document.body.innerText;
        
        // หา "รายละเอียดคำสั่งซื้อ" section
        const allElements = document.querySelectorAll('*');
        const orderDetailSections = [];
        
        allElements.forEach(el => {
            const text = el.innerText || '';
            if (text.includes('รายละเอียดคำสั่งซื้อ') && text.length < 3000 && el.children.length < 30) {
                orderDetailSections.push({
                    text: text,
                    tagName: el.tagName,
                    className: (el.className || '').substring(0, 200),
                    role: el.getAttribute('role'),
                    dataTestId: el.getAttribute('data-testid'),
                    innerHTML: el.innerHTML.substring(0, 2000),
                    childCount: el.children.length,
                    rect: el.getBoundingClientRect()
                });
            }
        });
        
        result.orderDetailSections = orderDetailSections.slice(0, 10);
        
        // หา elements ที่มีข้อความ "สินค้า", "รายการ", "item"
        const itemSections = [];
        allElements.forEach(el => {
            const text = el.innerText || '';
            if ((text.includes('สินค้า') || text.includes('รายการ') || text.includes('item') ||
                 text.includes('Product') || text.includes('product')) &&
                text.length < 1000 && text.length > 10 && el.children.length < 15) {
                itemSections.push({
                    text: text.substring(0, 500),
                    tagName: el.tagName,
                    className: (el.className || '').substring(0, 200),
                });
            }
        });
        result.itemSections = itemSections.slice(0, 15);
        
        // หา elements ที่มี "THB" หรือ "฿" (ราคาสินค้า)
        const priceSections = [];
        allElements.forEach(el => {
            const text = (el.innerText || '').trim();
            if ((text.match(/THB\s*[\d,]+/) || text.match(/฿[\d,]+/)) &&
                text.length < 300 && text.length > 3 && el.children.length < 10) {
                priceSections.push({
                    text: text,
                    tagName: el.tagName,
                    className: (el.className || '').substring(0, 100),
                    parentText: (el.parentElement?.innerText || '').substring(0, 300)
                });
            }
        });
        result.priceSections = priceSections.slice(0, 20);
        
        // หา slide-out panel / sidebar / dialog
        const panels = [];
        const panelSelectors = [
            '[role="dialog"]',
            '[role="complementary"]', 
            '[class*="sidebar"]',
            '[class*="panel"]',
            '[class*="detail"]',
            '[class*="drawer"]',
            '[data-pagelet*="order"]',
            '[data-pagelet*="Order"]',
        ];
        
        panelSelectors.forEach(selector => {
            document.querySelectorAll(selector).forEach(el => {
                const text = el.innerText || '';
                if (text.length > 50 && text.length < 5000) {
                    panels.push({
                        selector: selector,
                        text: text.substring(0, 2000),
                        tagName: el.tagName,
                        className: (el.className || '').substring(0, 200),
                        role: el.getAttribute('role'),
                        dataPagelet: el.getAttribute('data-pagelet'),
                        rect: el.getBoundingClientRect()
                    });
                }
            });
        });
        result.panels = panels.slice(0, 10);
        
        // ดึง data-pagelet ทั้งหมดที่เกี่ยวกับ order
        const pagelets = [];
        document.querySelectorAll('[data-pagelet]').forEach(el => {
            const pagelet = el.getAttribute('data-pagelet');
            if (pagelet) {
                pagelets.push({
                    pagelet: pagelet,
                    text: (el.innerText || '').substring(0, 500),
                    childCount: el.children.length
                });
            }
        });
        result.pagelets = pagelets;
        
        return result;
    }
    """)
    
    return detail_data


def main():
    print("\n" + "=" * 60)
    print("🔍 Inspect Order Detail - ดู DOM ของหน้า order detail")
    print("=" * 60)
    
    pw = sync_playwright().start()
    
    context = pw.chromium.launch_persistent_context(
        user_data_dir=str(SESSION_DIR),
        headless=False,
        viewport={"width": 1920, "height": 1080},
        locale="th-TH",
        timezone_id="Asia/Bangkok",
        args=[
            "--disable-blink-features=AutomationControlled",
            "--no-sandbox",
        ]
    )
    page = context.pages[0] if context.pages else context.new_page()
    
    print("✅ Browser เปิดแล้ว")
    
    # ทดสอบกับ order แรก
    order = SAMPLE_ORDERS[0]
    detail = inspect_order(page, order['order_number'], order['customer'])
    
    # Save
    output_path = OUTPUT_DIR / f"order_detail_inspection_{order['order_number']}.json"
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(detail, f, ensure_ascii=False, indent=2)
    
    print(f"\n💾 บันทึก: {output_path}")
    
    # แสดงผล
    print(f"\n📄 Pagelets ที่เจอ:")
    for p in detail.get('pagelets', []):
        print(f"   [{p['pagelet']}] ({p['childCount']} children) → {p['text'][:80]}...")
    
    print(f"\n📦 Order Detail Sections: {len(detail.get('orderDetailSections', []))}")
    for i, s in enumerate(detail.get('orderDetailSections', [])[:5]):
        print(f"\n--- Section #{i+1} ---")
        print(f"Tag: {s['tagName']} | Children: {s['childCount']}")
        print(f"Text: {s['text'][:300]}")
    
    print(f"\n🏷️ Item Sections: {len(detail.get('itemSections', []))}")
    for s in detail.get('itemSections', [])[:5]:
        print(f"   > {s['text'][:150]}")
    
    print(f"\n💰 Price Sections: {len(detail.get('priceSections', []))}")
    for s in detail.get('priceSections', [])[:10]:
        print(f"   > {s['text'][:100]}")
        if s.get('parentText'):
            print(f"     Parent: {s['parentText'][:150]}")
    
    print(f"\n🖼️ Panels: {len(detail.get('panels', []))}")
    for i, p in enumerate(detail.get('panels', [])[:5]):
        print(f"\n--- Panel #{i+1} ({p['selector']}) ---")
        print(f"Pagelet: {p.get('dataPagelet', 'N/A')}")
        print(f"Text: {p['text'][:400]}")
    
    # รอให้ user ดูหน้าจอ
    print("\n" + "=" * 60)
    print("👀 ดูหน้าจอ browser เพื่อเปรียบเทียบ")
    print("   ถ้าต้องการ scroll ดู order detail → ทำได้เลย")
    print("   แล้วกด Enter เพื่อ capture DOM อีกรอบ")
    
    while True:
        action = input("\n👉 กด Enter เพื่อ re-capture (หรือ 'n' → order ถัดไป, 'q' → จบ): ").strip()
        
        if action.lower() == 'q':
            break
        elif action.lower() == 'n':
            # ลอง order ถัดไป
            for next_order in SAMPLE_ORDERS[1:]:
                detail = inspect_order(page, next_order['order_number'], next_order['customer'])
                out = OUTPUT_DIR / f"order_detail_inspection_{next_order['order_number']}.json"
                with open(out, 'w', encoding='utf-8') as f:
                    json.dump(detail, f, ensure_ascii=False, indent=2)
                print(f"💾 บันทึก: {out}")
                
                # แสดง panels
                for i, p in enumerate(detail.get('panels', [])[:3]):
                    print(f"\n--- Panel #{i+1} ({p['selector']}) ---")
                    print(f"Text: {p['text'][:400]}")
                
                time.sleep(2)
            break
        else:
            # Re-capture current page
            detail2 = page.evaluate("""
            () => {
                const body = document.body.innerText;
                const sections = [];
                
                // หา div ที่อยู่ในส่วน order detail (ขวามือ)
                const allDivs = document.querySelectorAll('div');
                allDivs.forEach(div => {
                    const text = div.innerText || '';
                    const rect = div.getBoundingClientRect();
                    
                    // เอาเฉพาะ elements ที่อยู่ด้านขวา (x > 800) = detail panel
                    if (rect.x > 800 && text.length > 20 && text.length < 2000 && 
                        div.children.length < 20) {
                        sections.push({
                            text: text,
                            x: rect.x,
                            y: rect.y,
                            w: rect.width,
                            h: rect.height,
                            tagName: div.tagName,
                            className: (div.className || '').substring(0, 100)
                        });
                    }
                });
                
                // Sort by position
                sections.sort((a, b) => a.y - b.y);
                
                return {
                    url: window.location.href,
                    bodyTextLength: body.length,
                    rightPanelSections: sections.slice(0, 20)
                };
            }
            """)
            
            page.screenshot(path=str(OUTPUT_DIR / "order_detail_recapture.png"), full_page=False)
            
            print(f"\n📊 Right panel sections: {len(detail2.get('rightPanelSections', []))}")
            for i, s in enumerate(detail2.get('rightPanelSections', [])[:15]):
                print(f"\n--- [{i+1}] x={s['x']:.0f} y={s['y']:.0f} ---")
                print(f"   {s['text'][:200]}")
            
            # Save
            with open(OUTPUT_DIR / "order_detail_recapture.json", 'w', encoding='utf-8') as f:
                json.dump(detail2, f, ensure_ascii=False, indent=2)
    
    print("\n🛑 ปิด browser...")
    context.close()
    pw.stop()
    print("✅ เสร็จ!")


if __name__ == '__main__':
    main()
