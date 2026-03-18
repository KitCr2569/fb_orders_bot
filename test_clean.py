#!/usr/bin/env python3
"""Test clean + split functions"""
import sys
sys.path.insert(0, '.')
from sheet_sync import clean_product_name, split_multi_items

tests = [
    ('rf24-105 f2.8 ลาย mamba red', '990'),
    ('Sony A6600 ลาย cmd ลด 700 บาท', '700'),
    ('กล้อง Sony a7iii กับ เลนส์ Sony24-70 2.8 II ลาย mbbk', '1512'),
    ('legion go2 ลาย ctwt', '890'),
    ('RF50 F1.2 ลาย mabk+mtbk', '790'),
    ('Ptbk. และในส่วนที่เป็นกริบจับใช้เป็นLtbk.ครับ Nikon Z6iiiครับ', '890'),
    ('Canon R6III ลาย slpg', '890'),
]

for name, price in tests:
    cleaned = clean_product_name(name)
    items = split_multi_items(cleaned, price)
    if len(items) > 1:
        print(f'SPLIT: "{name}"')
        for n, p in items:
            print(f'  → Skin {n} | {p}')
    else:
        n, p = items[0]
        print(f'Skin {n} | {p}')
