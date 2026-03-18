#!/usr/bin/env python3
"""
Mark order as shipped on Facebook Business Suite
ใช้ Playwright เปิดหน้า order แล้วกด "ทำเครื่องหมายว่าจัดส่งแล้ว"
"""
import sys
import json
import time
import argparse
from pathlib import Path
from playwright.sync_api import sync_playwright

BASE_DIR = Path(__file__).parent
SESSION_DIR = BASE_DIR / "fb_session"
ASSET_ID = "114336388182180"

# URL ของ order detail page
ORDER_LIST_URL = f"https://business.facebook.com/latest/orders/orders_list?asset_id={ASSET_ID}"


def mark_as_shipped(order_number: str, customer: str = ""):
    """เปิด FB Business Suite → หา order → กด 'ทำเครื่องหมายว่าจัดส่งแล้ว'"""
    
    clean_number = order_number.replace('#', '').strip()
    
    with sync_playwright() as p:
        browser = p.chromium.launch_persistent_context(
            str(SESSION_DIR),
            headless=False,
            viewport={"width": 1400, "height": 900},
            locale="th-TH",
        )
        
        page = browser.pages[0] if browser.pages else browser.new_page()
        
        try:
            print(f"📦 กำลังเปิดหน้า Orders...")
            page.goto(ORDER_LIST_URL, wait_until="domcontentloaded", timeout=30000)
            time.sleep(3)
            
            # หา order ด้วย order number
            print(f"🔍 หา order #{clean_number} ({customer})...")
            
            # ลองคลิกที่ row ที่มี order number
            order_row = None
            
            # วิธี 1: หาจาก text ที่มี order number
            try:
                order_row = page.locator(f"text=#{clean_number}").first
                if order_row:
                    order_row.click(timeout=5000)
                    print(f"   ✅ คลิก order #{clean_number}")
                    time.sleep(3)
            except:
                pass
            
            if not order_row:
                # วิธี 2: หาจากชื่อลูกค้า
                try:
                    order_row = page.locator(f"text={customer}").first
                    if order_row:
                        order_row.click(timeout=5000)
                        print(f"   ✅ คลิก order ของ {customer}")
                        time.sleep(3)
                except:
                    pass
            
            # หาปุ่ม "ทำเครื่องหมายว่าจัดส่งแล้ว"
            print(f"🔍 หาปุ่ม 'ทำเครื่องหมายว่าจัดส่งแล้ว'...")
            
            shipped_btn = None
            
            # ลองหลายวิธี
            selectors = [
                "text=ทำเครื่องหมายว่าจัดส่งแล้ว",
                "button:has-text('ทำเครื่องหมายว่าจัดส่งแล้ว')",
                "[aria-label='ทำเครื่องหมายว่าจัดส่งแล้ว']",
                "div[role='button']:has-text('ทำเครื่องหมายว่าจัดส่งแล้ว')",
            ]
            
            for sel in selectors:
                try:
                    btn = page.locator(sel).first
                    if btn and btn.is_visible(timeout=3000):
                        shipped_btn = btn
                        print(f"   ✅ เจอปุ่ม: {sel}")
                        break
                except:
                    continue
            
            if shipped_btn:
                # สกรีนช็อตก่อนกด
                page.screenshot(path=str(BASE_DIR / "output" / f"before_mark_{clean_number}.png"))
                
                shipped_btn.click()
                print(f"   ✅ กดปุ่ม 'ทำเครื่องหมายว่าจัดส่งแล้ว'")
                time.sleep(3)
                
                # สกรีนช็อตหลังกด
                page.screenshot(path=str(BASE_DIR / "output" / f"after_mark_{clean_number}.png"))
                
                # ดูว่ามี confirmation dialog ไหม
                try:
                    confirm = page.locator("text=ยืนยัน").first
                    if confirm and confirm.is_visible(timeout=2000):
                        confirm.click()
                        print(f"   ✅ กดยืนยัน")
                        time.sleep(2)
                except:
                    pass
                
                print(f"✅ เสร็จ! order #{clean_number} ถูกทำเครื่องหมายว่าจัดส่งแล้ว")
                return {"status": "ok", "order": clean_number, "customer": customer}
            else:
                page.screenshot(path=str(BASE_DIR / "output" / f"notfound_mark_{clean_number}.png"))
                print(f"❌ ไม่เจอปุ่ม 'ทำเครื่องหมายว่าจัดส่งแล้ว' — อาจจัดส่งแล้วหรือหน้าเปลี่ยน")
                return {"status": "not_found", "order": clean_number, "message": "ไม่เจอปุ่ม"}
                
        except Exception as e:
            print(f"❌ Error: {e}")
            page.screenshot(path=str(BASE_DIR / "output" / f"error_mark_{clean_number}.png"))
            return {"status": "error", "order": clean_number, "message": str(e)}
        finally:
            browser.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--order", required=True, help="Order number เช่น #1575080583722849")
    parser.add_argument("--customer", default="", help="ชื่อลูกค้า")
    args = parser.parse_args()
    
    result = mark_as_shipped(args.order, args.customer)
    print(json.dumps(result, ensure_ascii=False, indent=2))
