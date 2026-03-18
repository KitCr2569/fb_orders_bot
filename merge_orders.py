#!/usr/bin/env python3
"""
Merge orders: ใช้ products จาก JSON เก่า (ที่มี products ครบ) มาเสริม JSON ใหม่ (ที่ products ขาด)
"""
import json
from pathlib import Path

OUTPUT_DIR = Path(__file__).parent / "output"

# อ่าน JSON ใหม่
new_file = OUTPUT_DIR / "orders_20260318_103802.json"
old_file = OUTPUT_DIR / "orders_20260318_082840.json"

with open(new_file, 'r', encoding='utf-8') as f:
    new_orders = json.load(f)

with open(old_file, 'r', encoding='utf-8') as f:
    old_orders = json.load(f)

# สร้าง lookup จาก old orders (by order_number)
old_lookup = {}
for o in old_orders:
    key = o.get('order_number', '').replace('#', '')
    old_lookup[key] = o

# Merge: ถ้า new order ไม่มี products → ใช้จาก old
merged = 0
for o in new_orders:
    key = o.get('order_number', '').replace('#', '')
    if not o.get('products') and key in old_lookup:
        old = old_lookup[key]
        if old.get('products'):
            o['products'] = old['products']
            merged += 1
            print(f"  ✅ Merged products for {o.get('customer', '')}: {[p.get('name', '') for p in o['products']]}")

# ลบ order ที่ยกเลิก (สำหรับ sheet sync)
active_orders = [o for o in new_orders if o.get('status') != 'ยกเลิกแล้ว']

print(f"\n📊 สรุป:")
print(f"   Orders ทั้งหมด: {len(new_orders)}")
print(f"   Active (ไม่รวมยกเลิก): {len(active_orders)}")
print(f"   Merged products: {merged}")

# Save merged JSON (orders ทั้งหมด)
merged_all_path = OUTPUT_DIR / "orders_march2026_all.json"
with open(merged_all_path, 'w', encoding='utf-8') as f:
    json.dump(new_orders, f, ensure_ascii=False, indent=2)
print(f"\n💾 All orders: {merged_all_path}")

# Save active-only JSON (สำหรับ sheet sync)
merged_active_path = OUTPUT_DIR / "orders_march2026_active.json"
with open(merged_active_path, 'w', encoding='utf-8') as f:
    json.dump(active_orders, f, ensure_ascii=False, indent=2)
print(f"💾 Active orders: {merged_active_path}")

# แสดงข้อมูลที่จะ sync
print(f"\n📋 Orders ที่จะ sync ลง Sheet:")
for i, o in enumerate(active_orders):
    products = o.get('products', [])
    names = [p.get('name', '?') for p in products] if products else ['(ไม่ทราบ)']
    price = o.get('price', '')
    print(f"   {i+1}. {o.get('date', '')} | {', '.join(names)} | ฿{price} | {o.get('customer', '')}")
