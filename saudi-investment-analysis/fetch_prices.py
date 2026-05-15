#!/usr/bin/env python3
"""
جالب أسعار أسهم تداول من Google Sheets
======================================
الإعداد (مرة واحدة فقط):
  1. افتح Google Sheets جديد على: https://sheets.google.com
  2. الصق المحتوى أدناه في الخلية A1 (ثم انسخ للأسفل)
  3. File → Share → Publish to web → اختر "الورقة الأولى" + CSV → انشر
  4. انسخ الرابط الناتج وضعه في متغير SHEET_CSV_URL أدناه

نموذج الصيغ في Google Sheets (انسخها في العمودين A و B):
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  A1: الراجحي      B1: =GOOGLEFINANCE("TADAWUL:1120","closeyest")
  A2: الإنماء     B2: =GOOGLEFINANCE("TADAWUL:1150","closeyest")
  A3: STC          B3: =GOOGLEFINANCE("TADAWUL:7010","closeyest")
  A4: سال          B4: =GOOGLEFINANCE("TADAWUL:4263","closeyest")
  A5: المواساة     B5: =GOOGLEFINANCE("TADAWUL:4002","closeyest")
  A6: بدجت         B6: =GOOGLEFINANCE("TADAWUL:4260","closeyest")
  A7: اكسترا       B7: =GOOGLEFINANCE("TADAWUL:4003","closeyest")
  A8: المتقدمة     B8: =GOOGLEFINANCE("TADAWUL:2330","closeyest")
  A9: بنيان ريت    B9: =GOOGLEFINANCE("TADAWUL:4347","closeyest")
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""

import json
import csv
import sys
import io
import os
from datetime import date

# ─── ضع رابط CSV من Google Sheets هنا ────────────────────────────────────────
SHEET_CSV_URL = ""   # <-- مثال: "https://docs.google.com/spreadsheets/d/XXXX/pub?output=csv"
# ─────────────────────────────────────────────────────────────────────────────

PRICES_FILE = os.path.join(os.path.dirname(__file__), "prices.json")

COMPANY_ORDER = [
    "الراجحي", "الإنماء", "STC", "سال",
    "المواساة", "بدجت", "اكسترا", "المتقدمة", "بنيان ريت"
]


def fetch_from_sheet(url: str) -> dict:
    """يقرأ أسعار الإغلاق من Google Sheet منشور كـ CSV."""
    import urllib.request
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=15) as resp:
            content = resp.read().decode("utf-8-sig")
        reader = csv.reader(io.StringIO(content))
        prices = {}
        for row in reader:
            if len(row) >= 2:
                name = row[0].strip()
                try:
                    price = float(row[1].replace(",", ""))
                    if name in COMPANY_ORDER and price > 0:
                        prices[name] = price
                except ValueError:
                    pass
        return prices
    except Exception as e:
        print(f"  ⚠ تعذّر قراءة Google Sheet: {e}")
        return {}


def fetch_via_yfinance() -> dict:
    """يحاول جلب الأسعار عبر yfinance (يعمل على جهازك المحلي)."""
    try:
        import yfinance as yf
        tickers_map = {
            "الراجحي":   "1120.SR",
            "الإنماء":   "1150.SR",
            "STC":       "7010.SR",
            "سال":       "4263.SR",
            "المواساة":  "4002.SR",
            "بدجت":      "4260.SR",
            "اكسترا":    "4003.SR",
            "المتقدمة":  "2330.SR",
            "بنيان ريت": "4347.SR",
        }
        prices = {}
        for company, ticker in tickers_map.items():
            try:
                info = yf.Ticker(ticker).fast_info
                price = getattr(info, "last_price", None) or getattr(info, "regular_market_price", None)
                if price and price > 0:
                    prices[company] = round(float(price), 2)
                    print(f"  ✅ {company}: {prices[company]}")
                else:
                    print(f"  ⚠ {company}: لا يوجد سعر")
            except Exception as e:
                print(f"  ✗ {company}: {e}")
        return prices
    except ImportError:
        print("  ⚠ yfinance غير مثبت — شغّل: pip install yfinance")
        return {}


def load_existing_prices() -> dict:
    """يحمّل الأسعار الحالية من prices.json."""
    try:
        with open(PRICES_FILE, encoding="utf-8") as f:
            data = json.load(f)
        return {k: v for k, v in data.items() if not k.startswith("_")}
    except Exception:
        return {}


def save_prices(prices: dict):
    """يحفظ الأسعار في prices.json."""
    old = load_existing_prices()
    merged = {**old, **prices}
    output = {
        "_note": "قيم الإغلاق — تُحدَّث تلقائياً بواسطة fetch_prices.py أو يدوياً",
        "_last_updated": str(date.today()),
        "_source": "Google Sheets / GOOGLEFINANCE",
    }
    for c in COMPANY_ORDER:
        if c in merged:
            output[c] = merged[c]
    with open(PRICES_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n✅ prices.json محدَّث ({PRICES_FILE})")


def print_prices(prices: dict):
    print("\n" + "─" * 45)
    print(f"{'الشركة':<15} {'السعر':>10} {'التغيير':>12}")
    print("─" * 45)
    old = load_existing_prices()
    for c in COMPANY_ORDER:
        p = prices.get(c)
        if p:
            change = f"{((p - old[c]) / old[c] * 100):+.1f}%" if c in old and old[c] else ""
            print(f"{c:<15} {p:>10.2f} {change:>12}")
    print("─" * 45)


def main():
    print("🔄 جلب أسعار الإغلاق لأسهم تداول...\n")

    prices = {}

    # 1. حاول من Google Sheet أولاً
    if SHEET_CSV_URL:
        print("1. قراءة من Google Sheets...")
        prices = fetch_from_sheet(SHEET_CSV_URL)
        if len(prices) == 9:
            print(f"  ✅ تم جلب {len(prices)} أسهم من Google Sheets")

    # 2. احتياطي: yfinance (يعمل على الجهاز المحلي)
    if len(prices) < 9:
        print("2. محاولة yfinance...")
        yf_prices = fetch_via_yfinance()
        for k, v in yf_prices.items():
            if k not in prices:
                prices[k] = v

    if not prices:
        print("\n❌ تعذّر جلب الأسعار تلقائياً.")
        print("   الحل: ضع رابط CSV من Google Sheets في متغير SHEET_CSV_URL\n")
        print("   أو أدخل الأسعار يدوياً الآن (اضغط Enter للتخطي):")
        for c in COMPANY_ORDER:
            val = input(f"  {c} (الحالي: {load_existing_prices().get(c, '?')}): ").strip()
            if val:
                try:
                    prices[c] = float(val)
                except ValueError:
                    pass

    if prices:
        print_prices(prices)
        save_prices(prices)

        # إعادة توليد ملف الإكسل تلقائياً
        regen = input("\nهل تريد إعادة توليد ملف الإكسل الآن؟ (y/n): ").strip().lower()
        if regen == "y":
            import subprocess
            script = os.path.join(os.path.dirname(__file__), "generate_excel.py")
            subprocess.run([sys.executable, script])


if __name__ == "__main__":
    main()
