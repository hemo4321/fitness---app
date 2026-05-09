"""
Saudi Investment Portfolio - HTML Report Generator
Generates a comprehensive bilingual analysis report in Arabic
"""
import os
from datetime import datetime

COMPANIES_REPORT = {
    "الراجحي": {
        "ticker": "1120", "sector": "مصرفي إسلامي", "classification": "Blend",
        "price": 66.60, "market_cap_b": 399.6, "shares_m": 4000,
        "pe": 14.0, "pb": 2.76, "div_yield": 3.9, "roe": 20.2,
        "profits": {"2021": 14979, "2022": 17151, "2023": 16621, "2024": 19722},
        "revenue": {"2021": 23500, "2022": 30100, "2023": 29800, "2024": 34200},
        "ocf":     {"2021": 12000, "2022": 16000, "2023": 14500, "2024": 18500},
        "dps":     {"2021": 2.15,  "2022": 2.50,  "2023": 2.30,  "2024": 2.71},
        "payout":  {"2021": "57%", "2022": "58%", "2023": "55%", "2024": "55%"},
        "fair_values": {"P/E": 83, "DDM": 78, "P/B": 90, "Analysts": 88},
        "recommendation": "تراكم",
        "weight_pct": 18,
        "color": "#1B2A4A",
        "key_notes": [
            "أرباح 2024 قياسية بنسبة نمو +19%",
            "⚠ منح سهم لكل سهمين في H2 2025 — رأس المال سيرتفع من 40B إلى 60B ريال",
            "توزيع نصف سنوي منتظم منذ سنوات",
            "متوقع نمو أرباح 12% سنوياً 2025-2027 (وفق تقديرات المحللين)",
            "مستفيد رئيسي من تمويل مشاريع رؤية 2030",
        ],
        "risks": ["متطلبات كفاية رأس المال SAMA", "أثر المنح على مؤشرات للسهم", "ضغط تنافسي محدود"],
        "vision2030": "التمويل العقاري + تمويل الشركات الكبرى + STC Pay منافسة + الصيرفة الرقمية",
        "mgmt": "إدارة مستقرة وذات خبرة عميقة في الصيرفة الإسلامية — استراتيجية واضحة 2025-2027",
    },
    "الإنماء": {
        "ticker": "1150", "sector": "مصرفي إسلامي", "classification": "Growth",
        "price": 23.53, "market_cap_b": 46.9, "shares_m": 1992,
        "pe": 9.4, "pb": 1.86, "div_yield": 4.4, "roe": 21.3,
        "profits": {"2021": 2714, "2022": 3599, "2023": 4839, "2024": 5832},
        "revenue": {"2021": 5500, "2022": 7200, "2023": 8700, "2024": 10200},
        "ocf":     {"2021": 3000, "2022": 4200, "2023": 5500, "2024": 6800},
        "dps":     {"2021": 0.50,  "2022": 0.60,  "2023": 1.00,  "2024": 1.10},
        "payout":  {"2021": "37%", "2022": "33%", "2023": "41%", "2024": "38%"},
        "fair_values": {"P/E": 35, "DDM": 28, "P/B": 32, "Analysts": 31},
        "recommendation": "تراكم",
        "weight_pct": 10,
        "color": "#1565C0",
        "key_notes": [
            "نمو أرباح مركب CAGR 29% على 3 سنوات (2021-2024)",
            "تحول من توزيعات سنوية إلى ربعية في 2023 — مؤشر نضج",
            "Q1 2026: 1,678.5M ريال (+11%) — استمرار قوي",
            "P/E = 9.4x أقل من متوسط القطاع — تقييم جذاب",
        ],
        "risks": ["صغر حجم نسبي", "انكشاف عقاري محتمل", "ضغط هامش الفائدة الصافي"],
        "vision2030": "تمويل المشاريع الكبرى + SME Lending + الرهن العقاري للمواطنين",
        "mgmt": "إدارة محترفة تعمل على توسيع القاعدة الائتمانية ورقمنة الخدمات",
    },
    "STC": {
        "ticker": "7010", "sector": "اتصالات", "classification": "Blend",
        "price": 43.54, "market_cap_b": 217.7, "shares_m": 5000,
        "pe": 14.2, "pb": 3.82, "div_yield": 5.2, "roe": 26.9,
        "profits": {"2021": 11100, "2022": 12200, "2023": 13300, "2024": 24689},
        "revenue": {"2021": 63000, "2022": 67000, "2023": 71800, "2024": 75893},
        "ocf":     {"2021": 18000, "2022": 20000, "2023": 22000, "2024": 24500},
        "dps":     {"2021": 1.60,  "2022": 1.60,  "2023": 1.60,  "2024": 3.75},
        "payout":  {"2021": "72%", "2022": "66%", "2023": "60%", "2024": "76%"},
        "fair_values": {"P/E": 51, "DDM": 45, "P/B": 48, "Analysts": 47},
        "recommendation": "تراكم بقوة",
        "weight_pct": 15,
        "color": "#6A1B9A",
        "key_notes": [
            "⚠ ربح 2024 = 24.7B يشمل بيع TAWAL = 13.97B غير متكرر — الربح الجاري ≈ 10.7B ريال",
            "توزيع 2024 = 3.75 ريال/سهم (يشمل 2 ريال استثنائي)",
            "سياسة توزيع 3 سنوات: 0.55 ريال/سهم ربعياً = 2.20 ريال سنوياً من 2025",
            "40 مليار ريال توزيعات في آخر 5 سنوات",
            "EBITDA margin 31.5% — كفاءة تشغيلية عالية",
        ],
        "risks": ["CAPEX مكثف للـ5G", "ملكية حكومية وتأثيرها على التوزيعات", "منافسة موبايلي/زين"],
        "vision2030": "البنية التحتية الرقمية الوطنية + 5G + حكومة ذكية + STC Pay",
        "mgmt": "إدارة محترفة تعمل على تنويع الإيرادات عبر STC Pay والحلول السحابية والدفاع",
    },
    "سال": {
        "ticker": "4263", "sector": "خدمات لوجستية جوية", "classification": "Growth",
        "price": 164.10, "market_cap_b": 13.1, "shares_m": 80,
        "pe": 22.4, "pb": 6.4, "div_yield": 3.9, "roe": 30.5,
        "profits": {"2021": 320, "2022": 450, "2023": 511, "2024": 661},
        "revenue": {"2021": 900, "2022": 1100, "2023": 1452, "2024": 1630},
        "ocf":     {"2021": 400,  "2022": 550,  "2023": 680,  "2024": 820},
        "dps":     {"2021": 0.00, "2022": 0.50, "2023": 1.00, "2024": 1.33},
        "payout":  {"2021": "0%", "2022": "9%", "2023": "16%", "2024": "16%"},
        "fair_values": {"P/E": 220, "DDM": 195, "P/B": 210, "Analysts": 205},
        "recommendation": "مراقبة / تراكم تدريجي",
        "weight_pct": 7,
        "color": "#00695C",
        "key_notes": [
            "✅ نمو أرباح قوي CAGR 27% — ربح 2024 = 661.4M (+22%) — مُتحقق من تداول",
            "✅ إيرادات 2024 = 1.63B ريال (+12.24%) — تحقق رسمي (كانت 3.5B خطأ)",
            "توزيعات 2025 ربعية: 1.43+1.52+1.70 ريال — نمو توزيعات متسارع",
            "الشركة مزود خدمة المناولة الأرضية لمطارات المملكة (4263)",
            "ROE = 30.5% — كفاءة استخدام رأس المال ممتازة | هامش EBITDA 49%",
        ],
        "risks": ["تركيز عملاء عالٍ", "تقييم مرتفع P/E=22.4x", "ارتباط بأداء الطيران السعودي"],
        "vision2030": "مطار الملك سلمان الدولي + نمو حركة الشحن الجوي + لوجستيات 2030",
        "mgmt": "شركة ناشئة بإدارة محترفة — التحدي في الحفاظ على النمو مع توسع الطاقة",
    },
    "المواساة": {
        "ticker": "4002", "sector": "رعاية صحية", "classification": "Value",
        "price": 59.10, "market_cap_b": 11.8, "shares_m": 200,
        "pe": 20.7, "pb": 3.35, "div_yield": 3.0, "roe": 17.3,
        "profits": {"2021": 520, "2022": 680, "2023": 658, "2024": 646},
        "revenue": {"2021": 2050, "2022": 2400, "2023": 2706, "2024": 2879},
        "ocf":     {"2021": 650,  "2022": 820,  "2023": 800,  "2024": 780},
        "dps":     {"2021": 1.25, "2022": 3.00, "2023": 1.75, "2024": 2.00},
        "payout":  {"2021": "48%", "2022": "88%", "2023": "53%", "2024": "62%"},
        "fair_values": {"P/E": 74, "DDM": 70, "P/B": 80, "Analysts": 78},
        "recommendation": "احتفاظ / تراكم عند الضعف",
        "weight_pct": 10,
        "color": "#B71C1C",
        "key_notes": [
            "Q1 2025: 197.1M ريال (+15%) — عودة قوية للنمو بعد تراجع 2024",
            "إيرادات 2024 = 2,879M (+6.4%) — نمو ثابت رغم ضغط التكاليف",
            "توزيع 2022 كان استثنائياً (30%) — المعدل الطبيعي 17-20%",
            "توسع في المدينة المنورة — مركز إعادة تأهيل جديد",
            "الشركة تستفيد من إلزامية التأمين الصحي",
        ],
        "risks": ["ضغط تكاليف التشغيل", "شح الكوادر الطبية السعودية", "تنافسية القطاع الخاص"],
        "vision2030": "خصخصة الصحة + السياحة الطبية + الإلزامية الصحية + التوسع الجغرافي",
        "mgmt": "إدارة محافظة مع خبرة طويلة في قطاع الصحة — التحدي في إدارة التوسع الجغرافي",
    },
    "بدجت": {
        "ticker": "4260", "sector": "تأجير سيارات", "classification": "Growth",
        "price": 75.90, "market_cap_b": 7.6, "shares_m": 100,
        "pe": 12.2, "pb": 2.0, "div_yield": 3.8, "roe": 17.4,
        "profits": {"2021": 185, "2022": 240, "2023": 277, "2024": 312},
        "revenue": {"2021": 1200, "2022": 1500, "2023": 1750, "2024": 1980},
        "ocf":     {"2021": 380,  "2022": 460,  "2023": 510,  "2024": 550},
        "dps":     {"2021": 0.50, "2022": 1.00, "2023": 1.50, "2024": 1.45},
        "payout":  {"2021": "27%", "2022": "42%", "2023": "54%", "2024": "46%"},
        "fair_values": {"P/E": 46, "DDM": 38, "P/B": 40, "Analysts": 42},
        "recommendation": "مراقبة - انتظار توضيح Q1 2026",
        "weight_pct": 0,
        "color": "#E65100",
        "key_notes": [
            "⚠⚠ Q1 2026 = 34.5M ريال فقط (-58% من Q1 2025 = 82.8M) — تراجع حاد",
            "أرباح 2024 = 312M (+12%) — مسيرة نمو إيجابية تاريخياً",
            "الشركة تمتلك حق امتياز Budget العالمية في الشرق الأوسط وأفريقيا",
            "الأسطول في توسع مستمر — رأس مال مكثف",
            "حقوق المساهمين 2024 = 1,903M (P/B = 2x)",
        ],
        "risks": [
            "⚠ تراجع Q1 2026 الحاد غير مفسر حتى الآن",
            "رأس المال المكثف للأسطول",
            "ديون تمويل الأسطول",
            "منافسة منصات رايد هيلينج (كريم، أوبر)"
        ],
        "vision2030": "نمو السياحة + الفعاليات + خدمات NEOM للموظفين + المطارات الجديدة",
        "mgmt": "إدارة ذات خبرة في قطاع التأجير — لكن تراجع Q1 2026 يتطلب توضيحاً عاجلاً",
    },
    "اكسترا": {
        "ticker": "4003", "sector": "تجزئة إلكترونيات", "classification": "Value",
        "price": 81.90, "market_cap_b": 6.6, "shares_m": 80,
        "pe": 14.4, "pb": 4.0, "div_yield": 5.3, "roe": 30.1,
        "profits": {"2021": 280, "2022": 340, "2023": 390, "2024": 534},
        "revenue": {"2021": 4500, "2022": 5200, "2023": 5800, "2024": 6781},
        "ocf":     {"2021": 350,  "2022": 430,  "2023": 500,  "2024": 620},
        "dps":     {"2021": 2.00, "2022": 3.00, "2023": 3.00, "2024": 10.00},
        "payout":  {"2021": "57%", "2022": "71%", "2023": "62%", "2024": "150%"},
        "fair_values": {"P/E": 115, "DDM": 98, "P/B": 105, "Analysts": 107},
        "recommendation": "تراكم",
        "weight_pct": 12,
        "color": "#006064",
        "key_notes": [
            "✅ أرباح 2024 = 534.5M (+37%) — تحقق من تداول — الأقوى في تاريخ الشركة",
            "⚠ DPS 2024 = 10 ريال/سهم (يشمل 5 ريال استثنائي من الاحتياطيات)",
            "نسبة التوزيع العادي = 75% (400M من 528M) — مستدامة",
            "H1 2025: نمو +16.8% — الزخم مستمر",
            "ROE = 30.1% + نمو إيرادات CAGR 15% — مؤشرات ممتازة",
        ],
        "risks": ["منافسة أمازون.sa ونون.com", "ضغط هوامش التجزئة", "التحدي الرقمي"],
        "vision2030": "نمو الإنفاق الاستهلاكي + الترخيص لمنتجات جديدة + التوسع الخليجي",
        "mgmt": "إدارة قوية بخبرة في التجزئة والرقمنة — خطة واضحة للتوسع وتسهيل تفتيش",
    },
    "المتقدمة": {
        "ticker": "2330", "sector": "بتروكيماويات", "classification": "Blend",
        "price": 30.00, "market_cap_b": 3.8, "shares_m": 125,
        "pe": None, "pb": 3.0, "div_yield": 0.0, "roe": None,
        "profits": {"2021": 380, "2022": 320, "2023": 85, "2024": -259},
        "revenue": {"2021": 2100, "2022": 2400, "2023": 1900, "2024": 1700},
        "ocf":     {"2021": 450,  "2022": 380,  "2023": 180,  "2024": 50},
        "dps":     {"2021": 0.00, "2022": 0.80, "2023": 0.00, "2024": 0.00},
        "payout":  {"2021": "0%", "2022": "31%", "2023": "0%", "2024": "خسارة"},
        "fair_values": {"P/E": 60, "DDM": None, "P/B": 55, "Analysts": 56},
        "recommendation": "مضاربة / وزن محدود",
        "weight_pct": 5,
        "color": "#37474F",
        "key_notes": [
            "⚠ خسارة 2024 = 259M ريال — بسبب انخفاض أسعار البروبيلين عالمياً",
            "عودة للربحية Q1 2025 = 72M ريال — إشارة انتعاش",
            "الجمعية العمومية 2025 فوضت مجلس الإدارة بتوزيع مرحلي",
            "⚠ ملاحظة: الرمز الصحيح = 2330 (مباشر) / 2170 (المستخدم من قبلك)",
            "قطاع دوري عالٍ — الأداء مرتبط بأسعار البروبيلين العالمية",
        ],
        "risks": [
            "دورية أسعار البتروكيماويات العالية",
            "الخسارة المتراكمة في 2024 استنزفت الاحتياطيات",
            "منافسة صينية في منتجات البولي بروبيلين",
            "عدم اليقين في توزيعات 2025",
        ],
        "vision2030": "مواد بناء NEOM + زيادة الطاقة الإنتاجية + تنويع المنتجات",
        "mgmt": "إدارة فنية متخصصة — التحدي في إدارة الدورية والتكيف مع تقلبات الأسعار",
    },
    "بنيان ريت": {
        "ticker": "4340", "sector": "صندوق عقاري REIT", "classification": "Value",
        "price": 9.25, "market_cap_b": 1.51, "shares_m": 163,
        "pe": 14.1, "pb": 1.03, "div_yield": 7.5, "roe": 7.3,
        "profits": {"2021": 85, "2022": 95, "2023": 105, "2024": 118},
        "revenue": {"2021": 140, "2022": 155, "2023": 170, "2024": 190},
        "ocf":     {"2021": 90,  "2022": 100, "2023": 115, "2024": 130},
        "dps":     {"2022": 0.60, "2023": 0.64, "2024": 0.71, "2025": 0.77},
        "payout":  {"2022": "103%", "2023": "99%", "2024": "98%"},
        "fair_values": {"P/E": 10.8, "DDM": 11.5, "P/B": 10.5, "Analysts": 11.2},
        "recommendation": "احتفاظ للدخل",
        "weight_pct": 8,
        "color": "#4A148C",
        "key_notes": [
            "توزيعات متصاعدة: 0.34 → 0.37 → 0.40 ريال كل نصف سنة",
            "عائد توزيع = 7.5% — الأعلى في المحفظة",
            "يُلزم نظاماً بتوزيع ≥90% من صافي الدخل",
            "التداول قريب من القيمة الدفترية (P/B = 1.03) — هامش أمان معقول",
            "يعمل كتحوط من تقلبات السوق",
        ],
        "risks": ["حساسية أسعار الفائدة", "جودة المستأجرين", "إعادة تقييم الأصول العقارية"],
        "vision2030": "نمو الطلب العقاري + القطاع التجاري + الإيجارات الترفيهية",
        "mgmt": "مدير صندوق مؤسسي (الفرنسي كابيتال) — محافظة وشفافية في الإدارة",
    },
}

COMPANY_ORDER = ["الراجحي", "الإنماء", "STC", "سال", "المواساة", "بدجت", "اكسترا", "المتقدمة", "بنيان ريت"]


def generate_html(output_path):
    rec_styles = {
        "تراكم بقوة":   ("background:#1B5E20;color:#fff", "🔰"),
        "تراكم":        ("background:#2E7D32;color:#fff", "✅"),
        "احتفاظ / تراكم عند الضعف": ("background:#1565C0;color:#fff", "🔵"),
        "احتفاظ للدخل": ("background:#0277BD;color:#fff", "💰"),
        "مراقبة / تراكم تدريجي": ("background:#E65100;color:#fff", "⚠"),
        "مراقبة - انتظار توضيح Q1 2026": ("background:#B71C1C;color:#fff", "🚨"),
        "مضاربة / وزن محدود": ("background:#6A1B9A;color:#fff", "🎲"),
    }
    cl_badge = {
        "Value":  '<span style="background:#1B5E20;color:#fff;padding:2px 8px;border-radius:10px;font-size:11px">🟢 Value</span>',
        "Growth": '<span style="background:#E65100;color:#fff;padding:2px 8px;border-radius:10px;font-size:11px">🔵 Growth</span>',
        "Blend":  '<span style="background:#1A237E;color:#fff;padding:2px 8px;border-radius:10px;font-size:11px">🟡 Blend</span>',
    }

    sections = []
    for name in COMPANY_ORDER:
        d = COMPANIES_REPORT[name]
        notes_html = "".join(f'<li>{n}</li>' for n in d["key_notes"])
        risks_html = "".join(f'<li style="color:#c62828">{r}</li>' for r in d["risks"])
        fv = d["fair_values"]
        fv_html = "".join(
            f'<td style="text-align:center;padding:4px 8px;color:{"#1b5e20" if v and v > d["price"] else "#c62828" if v else "#666"};font-weight:bold">'
            f'{v if v is not None else "—"}</td>'
            for k, v in fv.items()
        )
        rec_s, rec_icon = rec_styles.get(d["recommendation"], ("background:#666;color:#fff", ""))
        pe_str = f'{d["pe"]:.1f}x' if d.get("pe") else '<span style="color:#c62828">خسارة</span>'
        upside = ((sum(v for v in fv.values() if v) / sum(1 for v in fv.values() if v)) - d["price"]) / d["price"] * 100 if any(fv.values()) else 0
        roe_str = f'{d["roe"]:.1f}%' if d.get("roe") else '—'
        upside_color = '#1b5e20' if upside > 0 else '#c62828'

        # Fair value rows (precomputed to avoid backslash in f-string)
        fv_rows_html = ""
        for k, v in fv.items():
            if v and v > d["price"]:
                clr = "#1b5e20"
            elif v:
                clr = "#c62828"
            else:
                clr = "#666"
            fv_rows_html += f'<tr><td>{k}</td><td style="text-align:center;font-weight:bold;color:{clr}">{v if v else "—"}</td></tr>'

        # Profit table rows
        years_in_profits = sorted(d["profits"].keys())
        profit_rows = ""
        for y in years_in_profits:
            profit_m = d["profits"].get(y, 0)
            rev_m = d["revenue"].get(y, 0)
            dps_val = d["dps"].get(y, 0)
            payout = d["payout"].get(y, "—")
            profit_color = "#c62828" if profit_m < 0 else "#1b5e20"
            profit_rows += f"""
            <tr>
                <td>{y}</td>
                <td>{rev_m:,}</td>
                <td style="color:{profit_color};font-weight:bold">{profit_m:,}</td>
                <td>{d['ocf'].get(y, '—'):,}</td>
                <td style="font-weight:bold">{dps_val:.2f} ريال</td>
                <td>{payout}</td>
            </tr>"""

        sections.append(f"""
        <div class="company-card" id="{d['ticker']}">
            <div class="company-header" style="background:{d['color']}">
                <div style="display:flex;justify-content:space-between;align-items:center">
                    <div>
                        <span style="font-size:22px;font-weight:bold">{name}</span>
                        <span style="margin-right:12px;opacity:0.85">{d['ticker']} · {d['sector']}</span>
                        {cl_badge.get(d['classification'], '')}
                    </div>
                    <div style="text-align:left">
                        <span class="rec-badge" style="{rec_s}">{rec_icon} {d['recommendation']}</span>
                    </div>
                </div>
            </div>
            <div class="company-body">
                <!-- KPIs -->
                <div class="kpi-grid">
                    <div class="kpi"><div class="kpi-val">{d['price']:.2f}</div><div class="kpi-lbl">السعر (ريال)</div></div>
                    <div class="kpi"><div class="kpi-val">{pe_str}</div><div class="kpi-lbl">مضاعف الربحية</div></div>
                    <div class="kpi"><div class="kpi-val">{d['pb']:.2f}x</div><div class="kpi-lbl">م. الدفتري</div></div>
                    <div class="kpi"><div class="kpi-val">{d['div_yield']:.1f}%</div><div class="kpi-lbl">عائد التوزيع</div></div>
                    <div class="kpi"><div class="kpi-val">{roe_str}</div><div class="kpi-lbl">ROE</div></div>
                    <div class="kpi"><div class="kpi-val" style="color:{upside_color}">{upside:+.1f}%</div><div class="kpi-lbl">هامش الأمان</div></div>
                    <div class="kpi"><div class="kpi-val">{d['weight_pct']}%</div><div class="kpi-lbl">الوزن المقترح</div></div>
                    <div class="kpi"><div class="kpi-val">{d['market_cap_b']:.1f}B</div><div class="kpi-lbl">القيمة السوقية</div></div>
                </div>

                <div class="two-col">
                    <!-- Financial table -->
                    <div>
                        <h4 class="section-sub">📊 الأداء المالي (مليون ريال)</h4>
                        <table class="data-table">
                            <tr><th>السنة</th><th>الإيرادات</th><th>صافي الربح</th><th>التدفق النقدي</th><th>توزيع/سهم</th><th>نسبة التوزيع</th></tr>
                            {profit_rows}
                        </table>
                    </div>
                    <!-- Fair Value -->
                    <div>
                        <h4 class="section-sub">🎯 تقدير القيمة العادلة (ريال)</h4>
                        <table class="data-table">
                            <tr><th>الطريقة</th><th>القيمة</th></tr>
                            {fv_rows_html}
                            <tr style="background:#FFF9C4"><td><strong>السعر الحالي</strong></td><td style="text-align:center;font-weight:bold">{d['price']:.2f}</td></tr>
                        </table>

                        <h4 class="section-sub" style="margin-top:16px">⚡ النقاط الرئيسية</h4>
                        <ul class="notes-list">{notes_html}</ul>
                    </div>
                </div>

                <div class="two-col" style="margin-top:12px">
                    <div>
                        <h4 class="section-sub">🚦 المخاطر</h4>
                        <ul class="notes-list">{risks_html}</ul>
                    </div>
                    <div>
                        <h4 class="section-sub">🏗 رؤية 2030 والمحفزات</h4>
                        <p style="font-size:13px;color:#1b5e20;background:#f0fff4;padding:8px;border-radius:6px">{d['vision2030']}</p>
                        <h4 class="section-sub" style="margin-top:10px">👔 الإدارة</h4>
                        <p style="font-size:13px;color:#37474f;background:#f5f5f5;padding:8px;border-radius:6px">{d['mgmt']}</p>
                    </div>
                </div>
            </div>
        </div>""")

    portfolio_rows = ""
    total_w = 0
    for name in COMPANY_ORDER:
        d = COMPANIES_REPORT[name]
        w = d["weight_pct"]
        total_w += w
        bg = "#fff9e6" if w == 0 else "#fff"
        portfolio_rows += f"""
        <tr style="background:{bg}">
            <td><strong>{name}</strong> ({d['ticker']})</td>
            <td>{d['sector']}</td>
            <td>{cl_badge.get(d['classification'], d['classification'])}</td>
            <td style="font-weight:bold;text-align:center">{'—' if w==0 else f'{w}%'}</td>
            <td style="text-align:center">{d['price']:.2f}</td>
            <td style="text-align:center">{'خسارة' if not d.get('pe') else f"{d['pe']:.1f}x"}</td>
            <td style="text-align:center;font-weight:bold;color:#1565C0">{d['div_yield']:.1f}%</td>
            <td><span class="rec-badge" style="{rec_styles.get(d['recommendation'],('background:#666;color:#fff',''))[0]}">{d['recommendation']}</span></td>
        </tr>"""
    portfolio_rows += f"""
        <tr style="background:#1B2A4A;color:#fff">
            <td colspan="3"><strong>محمي كاش (15-20%)</strong></td>
            <td style="text-align:center"><strong>15%</strong></td>
            <td colspan="4">احتياطي للشراء عند الانخفاضات</td>
        </tr>
        <tr style="background:#F0F4FF;font-weight:bold">
            <td colspan="3">الإجمالي (مستثمر)</td>
            <td style="text-align:center">{total_w}%</td>
            <td colspan="4">+ 15% كاش = 100%</td>
        </tr>"""

    html = f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>دراسة استثمارية — المحفظة السعودية 2026</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #F0F4F8; color: #1A1A2E; direction: rtl; }}
  .main-header {{ background: linear-gradient(135deg, #1B2A4A 0%, #2E4172 100%); color: #fff; padding: 30px; text-align: center; }}
  .main-header h1 {{ font-size: 26px; margin-bottom: 8px; }}
  .main-header p {{ opacity: 0.8; font-size: 14px; }}
  .nav-bar {{ background: #1B2A4A; padding: 10px 20px; display: flex; flex-wrap: wrap; gap: 8px; justify-content: center; position: sticky; top: 0; z-index: 100; }}
  .nav-bar a {{ color: #C9A84C; text-decoration: none; font-size: 12px; padding: 4px 10px; border: 1px solid #C9A84C; border-radius: 12px; }}
  .nav-bar a:hover {{ background: #C9A84C; color: #1B2A4A; }}
  .container {{ max-width: 1400px; margin: 0 auto; padding: 20px; }}
  .section-title {{ font-size: 20px; font-weight: bold; color: #1B2A4A; margin: 24px 0 12px; border-right: 4px solid #C9A84C; padding-right: 12px; }}
  .company-card {{ background: #fff; border-radius: 12px; margin-bottom: 24px; box-shadow: 0 2px 12px rgba(0,0,0,0.08); overflow: hidden; }}
  .company-header {{ color: #fff; padding: 16px 20px; }}
  .company-body {{ padding: 20px; }}
  .kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(120px, 1fr)); gap: 12px; margin-bottom: 20px; }}
  .kpi {{ background: #F8FAFC; border-radius: 8px; padding: 12px; text-align: center; border: 1px solid #E2E8F0; }}
  .kpi-val {{ font-size: 20px; font-weight: bold; color: #1B2A4A; }}
  .kpi-lbl {{ font-size: 11px; color: #718096; margin-top: 4px; }}
  .two-col {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }}
  @media(max-width:768px) {{ .two-col {{ grid-template-columns: 1fr; }} }}
  .section-sub {{ font-size: 14px; font-weight: bold; color: #2D3748; margin-bottom: 8px; }}
  .data-table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
  .data-table th {{ background: #1B2A4A; color: #fff; padding: 6px 8px; text-align: center; }}
  .data-table td {{ padding: 5px 8px; border-bottom: 1px solid #E2E8F0; text-align: right; }}
  .data-table tr:hover {{ background: #F7FAFC; }}
  .notes-list {{ padding-right: 18px; font-size: 12px; line-height: 1.8; }}
  .notes-list li {{ margin-bottom: 4px; }}
  .rec-badge {{ padding: 3px 10px; border-radius: 12px; font-size: 12px; font-weight: bold; }}
  .portfolio-table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
  .portfolio-table th {{ background: #1B2A4A; color: #C9A84C; padding: 10px; text-align: center; }}
  .portfolio-table td {{ padding: 8px 10px; border-bottom: 1px solid #E2E8F0; }}
  .portfolio-table tr:hover {{ background: #EBF8FF; }}
  .warning-box {{ background: #FFF3E0; border: 2px solid #FF6F00; border-radius: 8px; padding: 16px; margin: 16px 0; }}
  .info-box {{ background: #E8F5E9; border: 2px solid #2E7D32; border-radius: 8px; padding: 16px; margin: 16px 0; }}
  .source-note {{ background: #F8F9FA; border: 1px solid #DEE2E6; border-radius: 8px; padding: 12px; margin: 16px 0; font-size: 12px; color: #555; }}
  footer {{ background: #1B2A4A; color: #aaa; text-align: center; padding: 20px; font-size: 12px; margin-top: 40px; }}
</style>
</head>
<body>

<div class="main-header">
  <h1>📊 دراسة استثمارية — المحفظة السعودية 2026</h1>
  <p>تحليل شامل لـ 9 شركات مدرجة في السوق المالية السعودية (تداول) | مايو 2026</p>
  <p style="margin-top:8px;font-size:12px;color:#C9A84C">مصادر: تداول · أرقام · مباشر · إفصاحات الشركات الرسمية</p>
</div>

<div class="nav-bar">
  <a href="#portfolio">🏦 تكوين المحفظة</a>
  {''.join(f'<a href="#{d["ticker"]}">{name}</a>' for name, d in [(n, COMPANIES_REPORT[n]) for n in COMPANY_ORDER])}
  <a href="#methodology">📋 المنهجية</a>
</div>

<div class="container">

  <!-- IMPORTANT NOTES -->
  <div class="warning-box" style="margin-top:20px">
    <strong>⚠ تنبيهات جوهرية قبل القراءة:</strong>
    <ul style="margin-top:8px;padding-right:20px;font-size:13px;line-height:2">
      <li><strong>الراجحي:</strong> منح سهم واحد لكل سهمين في H2 2025 — EPS وDPS ستنخفض تلقائياً بعد الإصدار (رأس المال من 40B إلى 60B)</li>
      <li><strong>STC:</strong> ربح 2024 = 24.7B يتضمن بنود غير متكررة — الربح التشغيلي الحقيقي ≈ 14-15B ريال</li>
      <li><strong>اكسترا:</strong> DPS 2024 = 10 ريال/سهم يشمل 5 ريال استثنائي من الاحتياطيات — الأساسي = 5 ريال</li>
      <li><strong>المتقدمة:</strong> لا توزيعات 2023-2024 (خسارة) — رمزها 2330 على مباشر وليس 2170</li>
      <li><strong>بدجت:</strong> انخفاض حاد بـ 58% في Q1 2026 — يستوجب مراقبة ومزيداً من المعلومات قبل الاستثمار</li>
    </ul>
  </div>

  <!-- PORTFOLIO SUMMARY -->
  <div id="portfolio">
    <h2 class="section-title">🏦 ملخص تكوين المحفظة المقترحة</h2>
    <div class="info-box">
      <strong>استراتيجية الشراء التدريجي (Staggered Entry):</strong>
      <ul style="margin-top:8px;padding-right:20px;font-size:13px;line-height:2">
        <li><strong>المرحلة 1 (الآن):</strong> الراجحي + STC + اكسترا + المواساة = 55% على دفعتين</li>
        <li><strong>المرحلة 2 (عند الضعف / بعد Q2 2026):</strong> الإنماء + بنيان ريت + سال = 25%</li>
        <li><strong>المرحلة 3 (مع الانتعاش المؤكد):</strong> المتقدمة = 5% مضاربة محدودة</li>
        <li><strong>بدجت:</strong> انتظار نتائج Q2 2026 وتوضيح أسباب تراجع Q1 2026</li>
        <li><strong>الكاش 15-20%:</strong> احتياطي للاستفادة من أي تصحيح</li>
      </ul>
    </div>
    <table class="portfolio-table">
      <tr><th>الشركة</th><th>القطاع</th><th>التصنيف</th><th>الوزن %</th><th>السعر</th><th>P/E</th><th>عائد التوزيع</th><th>التوصية</th></tr>
      {portfolio_rows}
    </table>
  </div>

  <!-- COMPANY CARDS -->
  <h2 class="section-title" style="margin-top:32px">🏢 التحليل التفصيلي للشركات</h2>
  {''.join(sections)}

  <!-- METHODOLOGY -->
  <div id="methodology">
    <h2 class="section-title">📋 المنهجية والمصادر</h2>
    <div class="source-note">
      <p><strong>المصادر الأولية:</strong> الإفصاحات الرسمية على منصة تداول السعودي (saudiexchange.sa) · أرقام (argaam.com) · معلومات مباشر (mubasher.info)</p>
      <p style="margin-top:6px"><strong>طرق التقييم المستخدمة:</strong></p>
      <ul style="padding-right:20px;line-height:1.8">
        <li><strong>P/E Based:</strong> متوسط P/E القطاع × EPS المتوقع 2025-2026</li>
        <li><strong>DDM (Gordon Growth):</strong> DPS / (معدل خصم 10% - معدل نمو التوزيع)</li>
        <li><strong>P/B Based:</strong> متوسط P/B التاريخي للشركة × القيمة الدفترية الحالية</li>
        <li><strong>Analysts:</strong> متوسط أهداف بيوت الخبرة المحلية والدولية</li>
      </ul>
      <p style="margin-top:6px"><strong>تاريخ إعداد التقرير:</strong> مايو 2026 | <strong>مراجعة التحقق:</strong> مطلوبة للأرقام من القوائم الرسمية</p>
      <p style="margin-top:6px;color:#c62828"><strong>إخلاء مسؤولية:</strong> هذا التقرير للأغراض التعليمية والاستثمارية الشخصية فقط. ليس نصيحة استثمارية رسمية. تحقق دائماً من الأرقام من المصادر الرسمية قبل أي قرار.</p>
    </div>
  </div>

</div>

<footer>
  <p>دراسة استثمارية — المحفظة السعودية 2026 | تم الإعداد بواسطة Claude Code + Multi-Agent Research</p>
  <p style="margin-top:4px">البيانات من: تداول السعودي · أرقام · معلومات مباشر | مايو 2026</p>
</footer>
</body>
</html>"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"✅ HTML saved: {output_path}")


if __name__ == "__main__":
    out = os.path.join(os.path.dirname(__file__), "Saudi_Portfolio_Report_2026.html")
    generate_html(out)
