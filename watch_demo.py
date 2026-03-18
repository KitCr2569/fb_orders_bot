"""
เปิด browser → ไป Inbox → Search Watchara → รอให้ user คลิก
Screenshot ทุก 2 วินาที (90 วิ = 45 screenshots)
"""
import time
from pathlib import Path
from playwright.sync_api import sync_playwright

SESSION_DIR = Path(__file__).parent / "fb_session"
DEMO_DIR = Path(__file__).parent / "output" / "demo2"
DEMO_DIR.mkdir(parents=True, exist_ok=True)

ASSET_ID = "114336388182180"
INBOX_URL = f"https://business.facebook.com/latest/inbox/all?asset_id={ASSET_ID}"

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

print("📂 เปิด Inbox...")
try:
    page.goto(INBOX_URL, wait_until="domcontentloaded", timeout=30000)
except Exception:
    pass
time.sleep(5)

# Search Watchara
print("🔍 Search Watchara Samsuvan...")
try:
    sb = page.locator('input[placeholder*="ค้นหา"], input[placeholder*="Search"]').first
    if sb.is_visible(timeout=3000):
        sb.click()
        time.sleep(0.3)
        sb.fill("Watchara Samsuvan")
        time.sleep(3)
except Exception:
    pass

print("✅ พร้อมแล้ว! คลิกเปิดแชท Watchara ได้เลยครับ!")
print("📸 จะ screenshot ทุก 2 วินาที (90 วิ)")

for i in range(45):
    ss = DEMO_DIR / f"demo_{i+1:02d}.png"
    page.screenshot(path=str(ss), full_page=False)
    print(f"  📸 [{i+1}/45] saved")
    time.sleep(2)

print("\n🛑 ปิด browser...")
context.close()
pw.stop()
print("✅ เสร็จ!")
