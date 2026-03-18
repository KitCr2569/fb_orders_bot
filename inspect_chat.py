"""
Inspect Chat - เปิด Business Suite Inbox เพื่อดู DOM structure ของ order card ในแชท
"""
import json
import time
from pathlib import Path
from playwright.sync_api import sync_playwright

SESSION_DIR = Path(__file__).parent / "fb_session"
OUTPUT_DIR = Path(__file__).parent / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# ลูกค้าจาก orders ที่ดึงมาแล้ว (ใช้เป็นตัวอย่าง)
SAMPLE_CUSTOMERS = [
    "Jeerasak Janson",
    "Watchara Samsuvan",
    "Nasri E'eso",
    "Akkasart Samrandee",
]

INBOX_URL = "https://business.facebook.com/latest/inbox/all?asset_id=114336388182180"


def main():
    print("\n" + "=" * 60)
    print("🔍 Inspect Chat - ดู order card ใน Business Suite Inbox")
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
    print("\n📂 กำลังไปหน้า Inbox...")
    
    try:
        page.goto(INBOX_URL, wait_until="domcontentloaded", timeout=60000)
    except Exception:
        print("⏳ โหลดช้า...")
    
    time.sleep(5)
    
    print(f"📄 URL: {page.url}")
    
    # ถ้าต้อง login ใหม่
    if "login" in page.url.lower():
        print("\n🔐 Session หมดอายุ! กรุณา login ใน browser ที่เปิด")
        print("   หลัง login เสร็จ → กด Enter")
        input("✋ กด Enter หลัง login... ")
        try:
            page.goto(INBOX_URL, wait_until="domcontentloaded", timeout=60000)
        except Exception:
            pass
        time.sleep(5)
    
    print("\n" + "=" * 60)
    print("📨 หน้า Inbox เปิดแล้ว!")
    print("=" * 60)
    print("\n👉 กรุณาทำตามนี้:")
    print("   1. คลิกเปิดแชทของลูกค้าที่มี order (เช่น Watchara, Nasri)")
    print("   2. Scroll ไปหา order card ในแชท (จะเห็นคำสั่งซื้อ)")
    print("   3. เมื่อเห็น order card แล้ว → กลับมากด Enter ที่นี่")
    print()
    input("✋ กด Enter เมื่อเปิดแชทและเห็น order card แล้ว... ")
    
    # ดึง DOM ทั้งหมดจากหน้า chat
    print("\n📊 กำลังดึง DOM structure...")
    
    chat_data = page.evaluate("""
    () => {
        const result = {
            url: window.location.href,
            title: document.title,
        };
        
        // ดึง body text ทั้งหมด (เฉพาะส่วน chat)
        const chatArea = document.querySelector('[role="main"]') || document.body;
        result.chatText = chatArea.innerText;
        
        // หา elements ที่อาจเป็น order card
        const allDivs = chatArea.querySelectorAll('div');
        const orderCards = [];
        
        allDivs.forEach(div => {
            const text = div.innerText || '';
            // หา order card: มี "คำสั่งซื้อ" หรือ "order" หรือ "THB" หรือ "#"
            if ((text.includes('คำสั่งซื้อ') || text.includes('สั่งซื้อ') || 
                 text.includes('Order') || text.includes('order')) &&
                (text.includes('THB') || text.includes('฿') || text.includes('บาท'))) {
                
                // ไม่เอา element ที่ใหญ่เกินไป
                if (text.length < 2000 && text.length > 20) {
                    const rect = div.getBoundingClientRect();
                    orderCards.push({
                        text: text,
                        tagName: div.tagName,
                        className: (div.className || '').substring(0, 200),
                        role: div.getAttribute('role'),
                        dataTestId: div.getAttribute('data-testid'),
                        ariaLabel: div.getAttribute('aria-label'),
                        rect: { x: rect.x, y: rect.y, w: rect.width, h: rect.height },
                        childCount: div.children.length,
                        innerHTML: div.innerHTML.substring(0, 1000)
                    });
                }
            }
        });
        
        result.orderCards = orderCards;
        
        // หา spans/divs ที่มี product names
        const productElements = [];
        allDivs.forEach(div => {
            const text = (div.innerText || '').trim();
            // หาชื่อที่ดูเหมือนสินค้า (มีคำว่า ลาย, body, lens, skin ฯลฯ)
            if (text.length > 5 && text.length < 200 &&
                (text.includes('ลาย') || text.includes('body') || text.includes('lens') ||
                 text.includes('Skin') || text.includes('กล้อง') || text.includes('เลนส์') ||
                 text.match(/[A-Z][a-z]*\s*(a|A|Z|R|X)\d/))) {
                
                productElements.push({
                    text: text,
                    tagName: div.tagName,
                    className: (div.className || '').substring(0, 100)
                });
            }
        });
        
        result.productElements = productElements.slice(0, 30);
        
        // ดึง links ที่มี "order" ใน href
        const links = document.querySelectorAll('a[href*="order"]');
        result.orderLinks = Array.from(links).map(a => ({
            href: a.href,
            text: a.innerText.substring(0, 200)
        }));
        
        return result;
    }
    """)
    
    # Save to file
    output_path = OUTPUT_DIR / "chat_inspection.json"
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(chat_data, f, ensure_ascii=False, indent=2)
    
    print(f"\n💾 บันทึกข้อมูลที่: {output_path}")
    print(f"   📦 พบ order cards: {len(chat_data.get('orderCards', []))}")
    print(f"   🏷️ พบ product elements: {len(chat_data.get('productElements', []))}")
    print(f"   🔗 พบ order links: {len(chat_data.get('orderLinks', []))}")
    
    # แสดงตัวอย่าง order cards
    for i, card in enumerate(chat_data.get('orderCards', [])[:5]):
        print(f"\n--- Order Card #{i+1} ---")
        print(f"Text: {card['text'][:300]}")
        print(f"DataTestId: {card.get('dataTestId', 'N/A')}")
        print(f"Size: {card['rect']['w']:.0f}x{card['rect']['h']:.0f}")
    
    # แสดงตัวอย่าง product elements
    if chat_data.get('productElements'):
        print(f"\n--- Product Elements ---")
        for pe in chat_data['productElements'][:10]:
            print(f"  > {pe['text'][:100]}")
    
    # ดึง screenshots
    print("\n📸 กำลัง screenshot...")
    page.screenshot(path=str(OUTPUT_DIR / "chat_screenshot.png"), full_page=False)
    print(f"   💾 {OUTPUT_DIR / 'chat_screenshot.png'}")
    
    # ถามว่าจะดู chat อื่นอีกไหม
    print("\n" + "=" * 60)
    while True:
        action = input("\n👉 กด Enter เพื่อ inspect อีกรอบ (หรือพิมพ์ 'q' เพื่อจบ): ").strip()
        if action.lower() == 'q':
            break
        
        print("📊 กำลังดึงข้อมูลอีกรอบ...")
        
        # ดึง DOM อีกรอบ
        chat_data2 = page.evaluate("""
        () => {
            const chatArea = document.querySelector('[role="main"]') || document.body;
            const text = chatArea.innerText;
            
            // หา order-related text blocks
            const blocks = [];
            const allDivs = chatArea.querySelectorAll('div');
            
            allDivs.forEach(div => {
                const t = div.innerText || '';
                if ((t.includes('คำสั่งซื้อ') || t.includes('THB') || t.includes('฿')) &&
                    t.length > 20 && t.length < 1500 && div.children.length < 15) {
                    blocks.push({
                        text: t,
                        html: div.innerHTML.substring(0, 800),
                        className: (div.className || '').substring(0, 100),
                        dataTestId: div.getAttribute('data-testid'),
                    });
                }
            });
            
            return {
                url: window.location.href,
                fullText: text.substring(0, 5000),
                blocks: blocks.slice(0, 10)
            };
        }
        """)
        
        # Save
        output_path2 = OUTPUT_DIR / "chat_inspection_2.json"
        with open(output_path2, 'w', encoding='utf-8') as f:
            json.dump(chat_data2, f, ensure_ascii=False, indent=2)
        
        page.screenshot(path=str(OUTPUT_DIR / "chat_screenshot_2.png"), full_page=False)
        
        print(f"💾 บันทึกแล้ว: {output_path2}")
        print(f"📦 พบ blocks: {len(chat_data2.get('blocks', []))}")
        
        for i, block in enumerate(chat_data2.get('blocks', [])[:5]):
            print(f"\n--- Block #{i+1} ---")
            print(f"{block['text'][:300]}")
    
    print("\n🛑 ปิด browser...")
    context.close()
    pw.stop()
    print("✅ เสร็จ!")


if __name__ == '__main__':
    main()
