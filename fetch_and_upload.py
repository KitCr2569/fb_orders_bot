#!/usr/bin/env python3
"""
🔄 Fetch & Upload — ดึง orders จาก Facebook แล้วอัปโหลดขึ้น Dashboard ออนไลน์
ใช้: python fetch_and_upload.py [--month 3]
"""
import argparse
import json
import subprocess
import sys
import requests
from pathlib import Path
from datetime import datetime

CLOUD_URL = "https://fb-orders-bot.onrender.com"
OUTPUT_DIR = Path(__file__).parent / "output"


def main():
    parser = argparse.ArgumentParser(description='ดึง orders จาก FB แล้วอัปโหลดขึ้น cloud')
    parser.add_argument('--month', type=int, default=datetime.now().month, help='เดือนที่ต้องการ (1-12)')
    parser.add_argument('--upload-only', action='store_true', help='อัปโหลด JSON ล่าสุดโดยไม่ดึงใหม่')
    args = parser.parse_args()

    if not args.upload_only:
        # Step 1: ดึง orders จาก Facebook
        print(f"\n📥 กำลังดึง orders เดือน {args.month}...")
        cmd = [sys.executable, 'fb_orders_bot.py', '--month', str(args.month), '--format', 'both', '--details']
        result = subprocess.run(cmd, cwd=str(Path(__file__).parent))
        
        if result.returncode != 0:
            print("❌ ดึง orders ไม่สำเร็จ")
            return

    # Step 2: หาไฟล์ JSON ล่าสุด (เฉพาะไฟล์ที่มี timestamp)
    json_files = sorted(
        [f for f in OUTPUT_DIR.glob("orders_2*.json")],  # Only timestamped files
        reverse=True
    )
    if not json_files:
        print("❌ ไม่พบไฟล์ JSON ในโฟลเดอร์ output/")
        return
    
    latest = json_files[0]
    with open(latest, 'r', encoding='utf-8') as f:
        orders = json.load(f)
    
    print(f"📄 ไฟล์: {latest.name} ({len(orders)} orders)")

    # Step 3: อัปโหลดขึ้น cloud
    print(f"\n🌐 กำลังอัปโหลดขึ้น {CLOUD_URL}...")
    try:
        with open(latest, 'rb') as f:
            resp = requests.post(
                f"{CLOUD_URL}/api/upload-orders",
                files={'file': (latest.name, f, 'application/json')}
            )
        
        data = resp.json()
        if data.get('success'):
            print(f"✅ อัปโหลดสำเร็จ! {data.get('count', 0)} orders → {CLOUD_URL}")
        else:
            print(f"❌ Error: {data.get('error', 'Unknown error')}")
    except Exception as e:
        print(f"❌ อัปโหลดไม่สำเร็จ: {e}")


if __name__ == '__main__':
    main()
