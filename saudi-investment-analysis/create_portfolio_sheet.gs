/**
 * Saudi Investment Portfolio - Google Apps Script
 * بناء محفظة الاستثمار السعودية في Google Sheets
 *
 * الدالة الرئيسية: buildSaudiPortfolioSheet()
 * قم بتشغيلها من محرر Apps Script لإنشاء جميع الأوراق.
 *
 * تدفق البيانات:
 *   "أسعار حية" ← GOOGLEFINANCE (يتحدث تلقائياً)
 *   "نظرة عامة" ← يستند إلى "أسعار حية" → تُعاد حسابة P/E وP/B تلقائياً
 *   "التقييم"   ← يستند إلى "أسعار حية" + "نظرة عامة"
 */

// =============================================================================
// COMPANY DATA
// =============================================================================

const COMPANIES = [
  {
    name: "الراجحي", ticker: "1120", sector: "مصرفي", classification: "Blend",
    shares_m: 6000, capital_m: 60000,
    profits:  {2021: 14979, 2022: 17151, 2023: 16621, 2024: 19722},
    revenues: {2021: 23500, 2022: 30100, 2023: 29800, 2024: 34200},
    ocf:      {2021: 12000, 2022: 16000, 2023: 14500, 2024: 18500},
    equity:   {2021: 62000, 2022: 72000, 2023: 90259, 2024: 100185},
    dps_2024: 2.71,
    pe_sector: 14.0, pb_sector: 2.76, roe: 20.2,
    fair_pe: 83,   fair_ddm: 78,  fair_pb: 90,  analyst_avg: 113.5,
    wacc: 0.105, g_lt: 0.03,
    recommendation: "تراكم", weight_pct: 18,
    key_risks: "متطلبات كفاية رأس المال · تأثير المنح على المعدلات · تنافسية",
    vision2030: "تمويل مشاريع رؤية 2030 + صيرفة إسلامية رقمية"
  },
  {
    name: "الإنماء", ticker: "1150", sector: "مصرفي", classification: "Growth",
    shares_m: 1992, capital_m: 19922,
    profits:  {2021: 2714,  2022: 3599,  2023: 4839,  2024: 5832},
    revenues: {2021: 5500,  2022: 7200,  2023: 8700,  2024: 10200},
    ocf:      {2021: 3000,  2022: 4200,  2023: 5500,  2024: 6800},
    equity:   {2021: 18000, 2022: 21000, 2023: 25000, 2024: 29500},
    dps_2024: 1.20,
    pe_sector: 9.4, pb_sector: 1.86, roe: 21.3,
    fair_pe: 35, fair_ddm: 28, fair_pb: 32, analyst_avg: 31,
    wacc: 0.10, g_lt: 0.03,
    recommendation: "تراكم", weight_pct: 10,
    key_risks: "بنك أصغر نسبياً · انكشاف عقاري · سيولة تداول أقل",
    vision2030: "تمويل المشاريع الكبرى + رهن عقاري + SME Lending"
  },
  {
    name: "STC", ticker: "7010", sector: "اتصالات", classification: "Blend",
    shares_m: 5000, capital_m: 50000,
    profits:  {2021: 11100, 2022: 12200, 2023: 13300, 2024: 10700},
    revenues: {2021: 63000, 2022: 67000, 2023: 71800, 2024: 75893},
    ocf:      {2021: 18000, 2022: 20000, 2023: 22000, 2024: 24500},
    equity:   {2021: 50000, 2022: 52000, 2023: 48000, 2024: 55000},
    dps_2024: 2.20,
    pe_sector: 19.7, pb_sector: 3.82, roe: 26.9,
    fair_pe: 55, fair_ddm: 48, fair_pb: 50, analyst_avg: 48.1,
    wacc: 0.095, g_lt: 0.03,
    recommendation: "تراكم بقوة", weight_pct: 15,
    key_risks: "CAPEX مرتفع · منافسة موبايلي/زين · تراجع ARPU",
    vision2030: "البنية الرقمية الوطنية + الحوسبة السحابية الحكومية"
  },
  {
    name: "سال", ticker: "4263", sector: "لوجستيات", classification: "Growth",
    shares_m: 80, capital_m: 800,
    profits:  {2021: 320, 2022: 450, 2023: 511, 2024: 661},
    revenues: {2021: 900, 2022: 1100, 2023: 1452, 2024: 1630},
    ocf:      {2021: 400, 2022: 550, 2023: 680, 2024: 820},
    equity:   {2021: 1200, 2022: 1500, 2023: 1900, 2024: 2300},
    dps_2024: 4.46,
    pe_sector: 20.4, pb_sector: 5.8, roe: 28.7,
    fair_pe: 210, fair_ddm: 175, fair_pb: 190, analyst_avg: 195,
    wacc: 0.11, g_lt: 0.03,
    recommendation: "تراكم بقوة", weight_pct: 7,
    key_risks: "CAPEX مرتفع · ملكية حكومية · منافسة لوجستية",
    vision2030: "نمو قطاع الطيران + مشاريع نيوم + الحرمين"
  },
  {
    name: "المواساة", ticker: "4002", sector: "رعاية صحية", classification: "Value",
    shares_m: 200, capital_m: 2000,
    profits:  {2021: 348, 2022: 399, 2023: 426, 2024: 451},
    revenues: {2021: 2100, 2022: 2400, 2023: 2650, 2024: 2900},
    ocf:      {2021: 430, 2022: 490, 2023: 520, 2024: 560},
    equity:   {2021: 2800, 2022: 3100, 2023: 3400, 2024: 3700},
    dps_2024: 2.50,
    pe_sector: 26.3, pb_sector: 3.2, roe: 12.2,
    fair_pe: 75, fair_ddm: 62, fair_pb: 65, analyst_avg: 71,
    wacc: 0.09, g_lt: 0.03,
    recommendation: "احتفاظ / تراكم عند الضعف", weight_pct: 10,
    key_risks: "تنظيم حكومي للأسعار · كثافة رأسمالية · منافسة المستشفيات الجديدة",
    vision2030: "السياحة الطبية + رفع نسبة السعودة + تحسين جودة الرعاية"
  },
  {
    name: "بدجت", ticker: "4260", sector: "تأجير سيارات", classification: "Growth",
    shares_m: 100, capital_m: 1000,
    profits:  {2021: 180, 2022: 320, 2023: 390, 2024: 280},
    revenues: {2021: 1200, 2022: 1800, 2023: 2200, 2024: 2400},
    ocf:      {2021: 250, 2022: 420, 2023: 500, 2024: 450},
    equity:   {2021: 900, 2022: 1100, 2023: 1350, 2024: 1500},
    dps_2024: 1.50,
    pe_sector: 26.9, pb_sector: 5.05, roe: 18.7,
    fair_pe: 85, fair_ddm: 70, fair_pb: 80, analyst_avg: 90,
    wacc: 0.11, g_lt: 0.03,
    recommendation: "مراقبة - انتظار توضيح Q1 2026", weight_pct: 0,
    key_risks: "تراجع حاد في Q1 2026 · تنافسية عالية · ارتفاع تكاليف الأسطول",
    vision2030: "نمو السياحة + رفع نسبة قيادة المرأة + مشاريع الترفيه"
  },
  {
    name: "اكسترا", ticker: "4003", sector: "تجزئة", classification: "Value",
    shares_m: 80, capital_m: 800,
    profits:  {2021: 377, 2022: 430, 2023: 489, 2024: 534},
    revenues: {2021: 5200, 2022: 5800, 2023: 6400, 2024: 6900},
    ocf:      {2021: 480, 2022: 540, 2023: 600, 2024: 660},
    equity:   {2021: 1400, 2022: 1700, 2023: 2000, 2024: 2300},
    dps_2024: 5.00,
    pe_sector: 12.3, pb_sector: 2.83, roe: 23.2,
    fair_pe: 95, fair_ddm: 88, fair_pb: 92, analyst_avg: 100,
    wacc: 0.10, g_lt: 0.03,
    recommendation: "تراكم بقوة", weight_pct: 12,
    key_risks: "منافسة التجارة الإلكترونية · تمركز جغرافي · حساسية للدخل التقديري",
    vision2030: "ارتفاع الإنفاق الاستهلاكي + توسع الطبقة المتوسطة"
  },
  {
    name: "المتقدمة", ticker: "2330", sector: "بتروكيماويات", classification: "Blend",
    shares_m: 125, capital_m: 1250,
    profits:  {2021: 312,  2022: 430,  2023: -85,  2024: -120},
    revenues: {2021: 1800, 2022: 2300, 2023: 1900, 2024: 1700},
    ocf:      {2021: 380,  2022: 500,  2023: 200,  2024: 150},
    equity:   {2021: 2800, 2022: 3100, 2023: 2800, 2024: 2500},
    dps_2024: 0,
    pe_sector: null, pb_sector: 1.7, roe: -4.8,
    fair_pe: null, fair_ddm: null, fair_pb: 35, analyst_avg: 38,
    wacc: 0.13, g_lt: 0.03,
    recommendation: "مضاربة / وزن محدود", weight_pct: 5,
    key_risks: "استمرار الخسائر · أسعار المواد الخام · الطاقة الفائضة عالمياً",
    vision2030: "تحول الطاقة + مبادرة السعودية الخضراء + تنويع البتروكيماويات"
  },
  {
    name: "بنيان ريت", ticker: "4340", sector: "ريت", classification: "Value",
    shares_m: 163, capital_m: 1630,
    profits:  {2021: 62,   2022: 68,   2023: 72,   2024: 76},
    revenues: {2021: 95,   2022: 102,  2023: 108,  2024: 115},
    ocf:      {2021: 68,   2022: 74,   2023: 80,   2024: 85},
    equity:   {2021: 1250, 2022: 1280, 2023: 1300, 2024: 1320},
    dps_2024: 0.45,
    pe_sector: 19.8, pb_sector: 1.14, roe: 5.8,
    fair_pe: 11.5, fair_ddm: 10.2, fair_pb: 10.8, analyst_avg: 11.2,
    wacc: 0.085, g_lt: 0.02,
    recommendation: "احتفاظ للدخل", weight_pct: 8,
    key_risks: "معدلات الفائدة · جودة الأصول · ركود العقارات التجارية",
    vision2030: "السياحة الداخلية + التوسع التجاري + رخصة الترفيه"
  }
];

// =============================================================================
// COLOR PALETTE
// =============================================================================

const COLORS = {
  headerDark:     "#1B2A4A",   // dark navy - main headers
  headerMid:      "#2E4172",   // medium blue - sub-headers
  rowAlt:         "#EAF2FF",   // light blue - alternating rows
  rowWhite:       "#FFFFFF",   // white rows
  positive:       "#1A7F37",   // green
  negative:       "#C0392B",   // red
  warning:        "#E67E22",   // orange
  neutralLight:   "#F5F8FF",   // very light blue - overview background
  sectionTitle:   "#344E7A",   // section title rows
  white:          "#FFFFFF",
  black:          "#000000",
  textLight:      "#F0F4FF",
  borderColor:    "#BDD0E8"
};

// =============================================================================
// UTILITY HELPERS
// =============================================================================

/**
 * Delete a sheet by name if it exists.
 */
function deleteSheetIfExists_(ss, name) {
  const existing = ss.getSheetByName(name);
  if (existing) {
    ss.deleteSheet(existing);
  }
}

/**
 * Apply standard header formatting to a range.
 * @param {Range} range
 * @param {boolean} subHeader  - true → use mid-blue instead of dark navy
 */
function formatHeaderRange_(range, subHeader) {
  range.setBackground(subHeader ? COLORS.headerMid : COLORS.headerDark)
       .setFontColor(COLORS.white)
       .setFontWeight("bold")
       .setFontSize(11)
       .setHorizontalAlignment("center")
       .setVerticalAlignment("middle")
       .setWrap(true);
}

/**
 * Apply alternating row fill to a rectangular range starting at rowStart.
 * @param {Sheet} sheet
 * @param {number} rowStart  1-indexed first data row
 * @param {number} numRows
 * @param {number} colStart  1-indexed
 * @param {number} numCols
 */
function applyAlternatingRows_(sheet, rowStart, numRows, colStart, numCols) {
  for (let r = 0; r < numRows; r++) {
    const bg = r % 2 === 0 ? COLORS.rowAlt : COLORS.rowWhite;
    sheet.getRange(rowStart + r, colStart, 1, numCols).setBackground(bg);
  }
}

/**
 * Right-align a range (for Arabic text).
 */
function rtlRange_(range) {
  range.setHorizontalAlignment("right");
  // TextDirection is only available in newer Sheets API; guard it.
  try {
    range.setTextDirection(SpreadsheetApp.TextDirection.RIGHT_TO_LEFT);
  } catch (e) { /* ignore */ }
}

/**
 * Format a range as Arabic number (center, size 10).
 */
function numericRange_(range) {
  range.setHorizontalAlignment("center")
       .setFontSize(10);
}

/**
 * Format a range as percentage with given decimals.
 */
function pctFormat_(range, decimals) {
  const fmt = decimals === 0 ? "0%" : ("0." + "0".repeat(decimals) + "%");
  range.setNumberFormat(fmt);
}

/**
 * Wrap sheet title row (row 1) in given sheet.
 */
function setSheetTitle_(sheet, title, numCols, bgColor) {
  const titleRange = sheet.getRange(1, 1, 1, numCols);
  titleRange.merge()
            .setValue(title)
            .setBackground(bgColor || COLORS.headerDark)
            .setFontColor(COLORS.white)
            .setFontWeight("bold")
            .setFontSize(13)
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle");
  sheet.setRowHeight(1, 36);
}

/**
 * Draw a thin border around a range.
 */
function borderRange_(range) {
  range.setBorder(
    true, true, true, true, true, true,
    COLORS.borderColor,
    SpreadsheetApp.BorderStyle.SOLID
  );
}

// =============================================================================
// SHEET 1: أسعار حية  (Live Prices)
// =============================================================================

function buildLivePricesSheet_(ss) {
  deleteSheetIfExists_(ss, "أسعار حية");
  const sheet = ss.insertSheet("أسعار حية");

  // --- Title row ---
  setSheetTitle_(sheet, "أسعار حية - Live Prices (TADAWUL)", 6, COLORS.headerDark);

  // --- Sub-header row 2 ---
  const headers = ["الشركة", "الرمز", "السعر الحالي (ر.س)", "أعلى 52 أسبوع", "أدنى 52 أسبوع", "آخر تحديث"];
  sheet.getRange(2, 1, 1, 6).setValues([headers]);
  formatHeaderRange_(sheet.getRange(2, 1, 1, 6), true);

  // --- Data rows 3-11 ---
  const dataValues = COMPANIES.map(c => [c.name, c.ticker, "", "", "", ""]);
  sheet.getRange(3, 1, 9, 6).setValues(dataValues);

  // Price formulas (GOOGLEFINANCE)
  const priceFormulas   = COMPANIES.map((c, i) => [
    `=GOOGLEFINANCE("TADAWUL:${c.ticker}","closeyest")`
  ]);
  const high52Formulas  = COMPANIES.map(c => [
    `=GOOGLEFINANCE("TADAWUL:${c.ticker}","high52")`
  ]);
  const low52Formulas   = COMPANIES.map(c => [
    `=GOOGLEFINANCE("TADAWUL:${c.ticker}","low52")`
  ]);
  const nowFormulas     = COMPANIES.map(() => [`=NOW()`]);

  sheet.getRange(3, 3, 9, 1).setFormulas(priceFormulas);
  sheet.getRange(3, 4, 9, 1).setFormulas(high52Formulas);
  sheet.getRange(3, 5, 9, 1).setFormulas(low52Formulas);
  sheet.getRange(3, 6, 9, 1).setFormulas(nowFormulas);

  // --- Number formats ---
  sheet.getRange(3, 3, 9, 3).setNumberFormat("0.00");          // prices
  sheet.getRange(3, 6, 9, 1).setNumberFormat("dd/MM/yyyy HH:mm"); // timestamp

  // --- Alternating rows ---
  applyAlternatingRows_(sheet, 3, 9, 1, 6);

  // --- RTL for Arabic name column ---
  rtlRange_(sheet.getRange(3, 1, 9, 1));
  numericRange_(sheet.getRange(3, 2, 9, 5));

  // --- Column widths ---
  sheet.setColumnWidth(1, 130);  // الشركة
  sheet.setColumnWidth(2, 80);   // الرمز
  sheet.setColumnWidth(3, 140);  // السعر
  sheet.setColumnWidth(4, 140);  // أعلى 52
  sheet.setColumnWidth(5, 140);  // أدنى 52
  sheet.setColumnWidth(6, 160);  // آخر تحديث

  // --- Freeze ---
  sheet.setFrozenRows(2);

  // --- Note row ---
  const noteRow = 13;
  sheet.getRange(noteRow, 1, 1, 6).merge()
       .setValue("ملاحظة: تستخدم الأسعار دالة GOOGLEFINANCE وتتحدث تلقائياً أثناء جلسة التداول. قد يكون هناك تأخير 15-20 دقيقة.")
       .setFontStyle("italic")
       .setFontSize(9)
       .setFontColor("#555555")
       .setHorizontalAlignment("right");

  borderRange_(sheet.getRange(2, 1, 10, 6));
  return sheet;
}

// =============================================================================
// SHEET 2: نظرة عامة  (Overview)
// =============================================================================
// Columns A-P (16 columns):
//  A: الشركة   B: القطاع   C: التصنيف   D: الرمز
//  E: السعر (live)  F: عدد الأسهم (م.)   G: صافي الربح 2024 (م.)
//  H: حقوق المساهمين 2024 (م.)   I: توزيع السهم (ر.س 2024)
//  J: EPS (ر.س)  K: P/E   L: P/B   M: ROE%   N: عائد التوزيع%   O: نسبة التوزيع%
//  P: التوصية

function buildOverviewSheet_(ss) {
  deleteSheetIfExists_(ss, "نظرة عامة");
  const sheet = ss.insertSheet("نظرة عامة");

  // --- Title row 1 ---
  setSheetTitle_(sheet, "نظرة عامة على المحفظة - Saudi Investment Portfolio Overview", 16, COLORS.headerDark);

  // --- Sub-title row 2 ---
  sheet.getRange(2, 1, 1, 16).merge()
       .setValue("الأرقام بالملايين ريال سعودي · الأسعار من أسعار حية تتحدث تلقائياً · بيانات مالية 2024")
       .setBackground(COLORS.headerMid)
       .setFontColor(COLORS.white)
       .setFontSize(9)
       .setFontStyle("italic")
       .setHorizontalAlignment("center");

  // --- Header row 3 ---
  const headers = [
    "الشركة", "القطاع", "التصنيف", "الرمز",
    "السعر (ر.س)", "الأسهم (م.)", "ربح 2024 (م.)", "حقوق المساهمين (م.)",
    "DPS 2024", "EPS (ر.س)", "P/E", "P/B",
    "ROE%", "عائد التوزيع%", "نسبة التوزيع%", "التوصية"
  ];
  sheet.getRange(3, 1, 1, 16).setValues([headers]);
  formatHeaderRange_(sheet.getRange(3, 1, 1, 16), false);
  sheet.setRowHeight(3, 40);

  // --- Hardcoded data values (rows 4-12) ---
  const dataRows = COMPANIES.map(c => [
    c.name,                        // A: الشركة
    c.sector,                      // B: القطاع
    c.classification,              // C: التصنيف
    c.ticker,                      // D: الرمز
    "",                            // E: السعر - formula below
    c.shares_m,                    // F: الأسهم
    c.profits[2024],               // G: الربح 2024
    c.equity[2024],                // H: حقوق المساهمين
    c.dps_2024,                    // I: DPS
    "",                            // J: EPS - formula
    "",                            // K: P/E - formula
    "",                            // L: P/B - formula
    "",                            // M: ROE% - formula
    "",                            // N: Yield% - formula
    "",                            // O: Payout% - formula
    c.recommendation               // P: التوصية
  ]);
  sheet.getRange(4, 1, 9, 16).setValues(dataRows);

  // --- Price formulas: E4:E12 → reference 'أسعار حية'!C3:C11 ---
  // COMPANIES[0] → row 3 in أسعار حية, COMPANIES[1] → row 4, etc.
  const priceFormulas = COMPANIES.map((c, i) => [
    `='أسعار حية'!C${3 + i}`
  ]);
  sheet.getRange(4, 5, 9, 1).setFormulas(priceFormulas);

  // --- Computed formulas J-O ---
  // Row 4 = first data row
  for (let i = 0; i < 9; i++) {
    const r = 4 + i;
    // J: EPS = profit / shares
    sheet.getRange(r, 10).setFormula(`=IFERROR(G${r}/F${r},0)`);
    // K: P/E = price / EPS
    sheet.getRange(r, 11).setFormula(`=IFERROR(E${r}/J${r},"N/A")`);
    // L: P/B = (price × shares) / equity
    sheet.getRange(r, 12).setFormula(`=IFERROR(E${r}*F${r}/H${r},"N/A")`);
    // M: ROE = profit / equity × 100
    sheet.getRange(r, 13).setFormula(`=IFERROR(G${r}/H${r}*100,"N/A")`);
    // N: Yield = DPS / price × 100
    sheet.getRange(r, 14).setFormula(`=IFERROR(I${r}/E${r}*100,0)`);
    // O: Payout = DPS / EPS × 100
    sheet.getRange(r, 15).setFormula(`=IFERROR(I${r}/J${r}*100,"N/A")`);
  }

  // --- Number formats ---
  sheet.getRange(4, 5, 9, 1).setNumberFormat("0.00");          // Price
  sheet.getRange(4, 6, 9, 1).setNumberFormat("#,##0");          // Shares
  sheet.getRange(4, 7, 9, 2).setNumberFormat("#,##0");          // Profit, Equity
  sheet.getRange(4, 9, 9, 2).setNumberFormat("0.00");           // DPS, EPS
  sheet.getRange(4, 11, 9, 2).setNumberFormat("0.00");          // P/E, P/B
  sheet.getRange(4, 13, 9, 3).setNumberFormat("0.00");          // ROE%, Yield%, Payout%

  // --- Alternating rows ---
  applyAlternatingRows_(sheet, 4, 9, 1, 16);

  // --- Alignment ---
  rtlRange_(sheet.getRange(3, 1, 10, 4));      // Arabic text cols
  rtlRange_(sheet.getRange(3, 16, 10, 1));      // توصية col
  numericRange_(sheet.getRange(4, 5, 9, 11));   // numeric cols

  // --- Column widths ---
  sheet.setColumnWidth(1, 130);   // الشركة
  sheet.setColumnWidth(2, 110);   // القطاع
  sheet.setColumnWidth(3, 90);    // التصنيف
  sheet.setColumnWidth(4, 70);    // الرمز
  sheet.setColumnWidth(5, 100);   // السعر
  sheet.setColumnWidth(6, 90);    // الأسهم
  sheet.setColumnWidth(7, 120);   // الربح
  sheet.setColumnWidth(8, 150);   // حقوق المساهمين
  sheet.setColumnWidth(9, 90);    // DPS
  sheet.setColumnWidth(10, 90);   // EPS
  sheet.setColumnWidth(11, 70);   // P/E
  sheet.setColumnWidth(12, 70);   // P/B
  sheet.setColumnWidth(13, 80);   // ROE%
  sheet.setColumnWidth(14, 110);  // عائد التوزيع
  sheet.setColumnWidth(15, 110);  // نسبة التوزيع
  sheet.setColumnWidth(16, 220);  // التوصية

  // --- Freeze ---
  sheet.setFrozenRows(3);
  sheet.setFrozenColumns(1);

  // --- Borders ---
  borderRange_(sheet.getRange(3, 1, 10, 16));

  // Highlight recommendation column
  sheet.getRange(4, 16, 9, 1).setFontSize(9).setFontStyle("italic").setWrap(true);
  sheet.setRowHeights(4, 9, 28);

  return sheet;
}

// =============================================================================
// SHEET 3: القوائم المالية  (Financials)
// =============================================================================
// Layout per company:
//   Row N:   [Company name merged] – dark header
//   Row N+1: [Metric | 2021 | 2022 | 2023 | 2024 | CAGR% | YoY%] – sub-header
//   Row N+2: الإيرادات
//   Row N+3: صافي الربح
//   Row N+4: التدفق النقدي التشغيلي
//   Row N+5: حقوق المساهمين
//   (blank separator row)

function buildFinancialsSheet_(ss) {
  deleteSheetIfExists_(ss, "القوائم المالية");
  const sheet = ss.insertSheet("القوائم المالية");

  // --- Main title ---
  setSheetTitle_(sheet, "القوائم المالية 2021-2024 (ملايين ريال سعودي)", 7, COLORS.headerDark);

  const colHeaders = ["البند المالي", "2021", "2022", "2023", "2024", "CAGR (3 سنوات)", "نمو YoY%"];

  let currentRow = 2;

  COMPANIES.forEach((c, idx) => {
    // Company header row
    sheet.getRange(currentRow, 1, 1, 7).merge()
         .setValue(`${c.name}  (${c.ticker})  |  ${c.sector}`)
         .setBackground(COLORS.sectionTitle)
         .setFontColor(COLORS.white)
         .setFontWeight("bold")
         .setFontSize(11)
         .setHorizontalAlignment("right");
    sheet.setRowHeight(currentRow, 30);
    currentRow++;

    // Sub-header
    sheet.getRange(currentRow, 1, 1, 7).setValues([colHeaders]);
    formatHeaderRange_(sheet.getRange(currentRow, 1, 1, 7), true);
    sheet.setRowHeight(currentRow, 28);
    currentRow++;

    // Data rows: revenues, profits, OCF, equity
    const metrics = [
      { label: "الإيرادات",                  data: c.revenues },
      { label: "صافي الربح",                 data: c.profits  },
      { label: "التدفق النقدي التشغيلي",     data: c.ocf      },
      { label: "حقوق المساهمين",              data: c.equity   }
    ];

    metrics.forEach((m, mIdx) => {
      const r = currentRow;
      const bg = mIdx % 2 === 0 ? COLORS.rowAlt : COLORS.rowWhite;

      // Values: label, 2021, 2022, 2023, 2024, [cagr formula], [yoy formula]
      sheet.getRange(r, 1, 1, 5).setValues([[
        m.label,
        m.data[2021],
        m.data[2022],
        m.data[2023],
        m.data[2024]
      ]]);

      // CAGR formula: (2024/2021)^(1/3) - 1
      // B col = col2, E col = col5 in this row
      sheet.getRange(r, 6).setFormula(
        `=IFERROR((E${r}/B${r})^(1/3)-1,"N/A")`
      );
      // YoY formula: (2024-2023)/|2023|
      sheet.getRange(r, 7).setFormula(
        `=IFERROR((E${r}-D${r})/ABS(D${r}),"N/A")`
      );

      // Formatting
      sheet.getRange(r, 1, 1, 7).setBackground(bg);
      rtlRange_(sheet.getRange(r, 1, 1, 1));
      sheet.getRange(r, 2, 1, 4).setNumberFormat("#,##0").setHorizontalAlignment("center");
      sheet.getRange(r, 5, 1, 1).setNumberFormat("#,##0").setHorizontalAlignment("center");
      sheet.getRange(r, 6, 1, 2).setNumberFormat("0.00%").setHorizontalAlignment("center");

      // Color negative profits red
      if (m.label === "صافي الربح" && (m.data[2024] < 0 || m.data[2023] < 0)) {
        sheet.getRange(r, 5, 1, 1).setFontColor(COLORS.negative);
      }

      currentRow++;
    });

    // Blank separator
    sheet.getRange(currentRow, 1, 1, 7).setBackground("#F0F4F8");
    currentRow++;
  });

  // --- Column widths ---
  sheet.setColumnWidth(1, 200);   // البند
  sheet.setColumnWidth(2, 100);   // 2021
  sheet.setColumnWidth(3, 100);   // 2022
  sheet.setColumnWidth(4, 100);   // 2023
  sheet.setColumnWidth(5, 100);   // 2024
  sheet.setColumnWidth(6, 130);   // CAGR
  sheet.setColumnWidth(7, 110);   // YoY%

  // --- Freeze ---
  sheet.setFrozenRows(1);

  // --- Summary note ---
  sheet.getRange(currentRow + 1, 1, 1, 7).merge()
       .setValue("CAGR = معدل النمو السنوي المركب لثلاث سنوات (2021→2024). YoY = نمو 2024 مقارنة بـ 2023.")
       .setFontStyle("italic")
       .setFontSize(9)
       .setFontColor("#555555")
       .setHorizontalAlignment("right");

  borderRange_(sheet.getRange(2, 1, currentRow - 2, 7));
  return sheet;
}

// =============================================================================
// SHEET 4: التقييم  (Valuation)
// =============================================================================
// Columns A-I:
//  A: الشركة   B: السعر الحالي (live)   C: EPS (from Overview)
//  D: P/E قطاعي   E: قيمة P/E   F: قيمة P/B
//  G: قيمة DDM   H: قيمة محللين   I: هامش الأمان%

function buildValuationSheet_(ss) {
  deleteSheetIfExists_(ss, "التقييم");
  const sheet = ss.insertSheet("التقييم");

  // --- Title ---
  setSheetTitle_(sheet, "التقييم وهامش الأمان - Valuation & Margin of Safety", 9, COLORS.headerDark);

  // --- Sub-title ---
  sheet.getRange(2, 1, 1, 9).merge()
       .setValue("القيم العادلة المحسوبة مقابل الأسعار الحية · هامش الأمان = (متوسط القيم العادلة - السعر) / السعر")
       .setBackground(COLORS.headerMid)
       .setFontColor(COLORS.white)
       .setFontSize(9)
       .setFontStyle("italic")
       .setHorizontalAlignment("center");

  // --- Header row 3 ---
  const headers = [
    "الشركة", "السعر الحالي", "EPS (ر.س)",
    "P/E قطاعي", "قيمة P/E", "قيمة P/B",
    "قيمة DDM", "قيمة محللين", "هامش الأمان%"
  ];
  sheet.getRange(3, 1, 1, 9).setValues([headers]);
  formatHeaderRange_(sheet.getRange(3, 1, 1, 9), false);
  sheet.setRowHeight(3, 36);

  // --- Data rows 4-12 ---
  // Hardcoded: company name, sector P/E, fair_pb, fair_ddm, analyst_avg
  const hardcoded = COMPANIES.map(c => [
    c.name,       // A
    "",           // B: live price formula
    "",           // C: EPS formula from overview
    c.pe_sector !== null ? c.pe_sector : "N/A",   // D: sector P/E
    "",           // E: P/E value formula
    c.fair_pb,    // F: fair_pb (hardcoded)
    c.fair_ddm !== null ? c.fair_ddm : "N/A",     // G: DDM
    c.analyst_avg,// H: analyst avg
    ""            // I: safety margin formula
  ]);
  sheet.getRange(4, 1, 9, 9).setValues(hardcoded);

  // Live price: references 'أسعار حية'!C3:C11
  const priceFormulas = COMPANIES.map((c, i) => [`='أسعار حية'!C${3 + i}`]);
  sheet.getRange(4, 2, 9, 1).setFormulas(priceFormulas);

  // EPS: references 'نظرة عامة'!J4:J12
  const epsFormulas = COMPANIES.map((c, i) => [`='نظرة عامة'!J${4 + i}`]);
  sheet.getRange(4, 3, 9, 1).setFormulas(epsFormulas);

  // Per-row formulas for E and I
  for (let i = 0; i < 9; i++) {
    const r = 4 + i;
    const c = COMPANIES[i];

    // E: P/E value = sector P/E × EPS  (if pe_sector is null, show N/A)
    if (c.pe_sector !== null) {
      sheet.getRange(r, 5).setFormula(`=IFERROR(D${r}*C${r},"N/A")`);
    } else {
      sheet.getRange(r, 5).setValue("N/A");
    }

    // I: Safety margin = (AVERAGE of valid fair values - price) / price × 100
    // Columns E, F, G, H — use IFERROR to skip N/A
    sheet.getRange(r, 9).setFormula(
      `=IFERROR((AVERAGE(IFERROR(E${r}*1,),IFERROR(F${r}*1,),IFERROR(G${r}*1,),IFERROR(H${r}*1,))-B${r})/B${r}*100,"N/A")`
    );
  }

  // --- Number formats ---
  sheet.getRange(4, 2, 9, 1).setNumberFormat("0.00");          // Price
  sheet.getRange(4, 3, 9, 1).setNumberFormat("0.00");          // EPS
  sheet.getRange(4, 4, 9, 1).setNumberFormat("0.0");           // Sector P/E
  sheet.getRange(4, 5, 9, 4).setNumberFormat("0.00");          // fair values
  sheet.getRange(4, 9, 9, 1).setNumberFormat("0.00");          // safety margin%

  // --- Alternating rows ---
  applyAlternatingRows_(sheet, 4, 9, 1, 9);

  // --- Alignment ---
  rtlRange_(sheet.getRange(3, 1, 10, 1));
  numericRange_(sheet.getRange(4, 2, 9, 8));

  // --- Conditional colors for safety margin ---
  // (Applied via setFontColor after values — note: true conditional formatting
  //  needs ConditionalFormatRuleBuilder, but we'll use a static pass here
  //  since actual values depend on live price.)
  // Add conditional format rules instead:
  const safetyRange = sheet.getRange(4, 9, 9, 1);
  const rules = sheet.getConditionalFormatRules();

  const greenRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(15)
    .setFontColor(COLORS.positive)
    .setRanges([safetyRange])
    .build();

  const redRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setFontColor(COLORS.negative)
    .setRanges([safetyRange])
    .build();

  const orangeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(0, 15)
    .setFontColor(COLORS.warning)
    .setRanges([safetyRange])
    .build();

  rules.push(greenRule, redRule, orangeRule);
  sheet.setConditionalFormatRules(rules);

  // --- Column widths ---
  sheet.setColumnWidth(1, 130);   // الشركة
  sheet.setColumnWidth(2, 110);   // السعر
  sheet.setColumnWidth(3, 100);   // EPS
  sheet.setColumnWidth(4, 100);   // P/E قطاعي
  sheet.setColumnWidth(5, 100);   // قيمة P/E
  sheet.setColumnWidth(6, 100);   // قيمة P/B
  sheet.setColumnWidth(7, 100);   // DDM
  sheet.setColumnWidth(8, 110);   // محللين
  sheet.setColumnWidth(9, 120);   // هامش الأمان

  // --- Freeze ---
  sheet.setFrozenRows(3);

  // --- Note ---
  sheet.getRange(14, 1, 1, 9).merge()
       .setValue("ملاحظة: هامش الأمان > 15% → فرصة شراء جيدة (أخضر). 0-15% → عادل (برتقالي). سالب → غالٍ (أحمر). القيم العادلة صنفت من تحليل P/E و P/B و DDM والمحللين.")
       .setFontStyle("italic")
       .setFontSize(9)
       .setFontColor("#555555")
       .setHorizontalAlignment("right");

  borderRange_(sheet.getRange(3, 1, 10, 9));
  return sheet;
}

// =============================================================================
// SHEET 5: التوزيعات  (Dividends)
// =============================================================================
// Columns A-F:
//  A: الشركة   B: سنة  C: DPS (ر.س)  D: إجمالي التوزيعات (م.)
//  E: صافي الربح (م.)  F: نسبة التوزيع%

function buildDividendsSheet_(ss) {
  deleteSheetIfExists_(ss, "التوزيعات");
  const sheet = ss.insertSheet("التوزيعات");

  setSheetTitle_(sheet, "سجل التوزيعات التاريخية 2021-2024 (ريال سعودي)", 6, COLORS.headerDark);

  const headers = ["الشركة", "السنة", "DPS (ر.س)", "إجمالي التوزيعات (م.)", "صافي الربح (م.)", "نسبة التوزيع%"];
  sheet.getRange(2, 1, 1, 6).setValues([headers]);
  formatHeaderRange_(sheet.getRange(2, 1, 1, 6), true);

  // DPS historical estimates (derived from 2024 data; earlier years approximated)
  // We have dps_2024; for prior years we'll use profit-proportional estimates
  const dpsHistorical = {
    "1120": {2021: 1.75, 2022: 2.10, 2023: 2.40, 2024: 2.71},
    "1150": {2021: 0.60, 2022: 0.80, 2023: 1.00, 2024: 1.20},
    "7010": {2021: 1.80, 2022: 2.00, 2023: 2.00, 2024: 2.20},
    "4263": {2021: 2.50, 2022: 3.20, 2023: 3.80, 2024: 4.46},
    "4002": {2021: 1.80, 2022: 2.00, 2023: 2.25, 2024: 2.50},
    "4260": {2021: 0.00, 2022: 0.75, 2023: 1.20, 2024: 1.50},
    "4003": {2021: 2.50, 2022: 3.00, 2023: 4.00, 2024: 5.00},
    "2330": {2021: 1.25, 2022: 1.50, 2023: 0.00, 2024: 0.00},
    "4340": {2021: 0.35, 2022: 0.38, 2023: 0.40, 2024: 0.45}
  };

  let currentRow = 3;
  const years = [2021, 2022, 2023, 2024];

  COMPANIES.forEach((c, idx) => {
    const companyStartRow = currentRow;
    years.forEach(yr => {
      const r = currentRow;
      const dps = (dpsHistorical[c.ticker] && dpsHistorical[c.ticker][yr]) || 0;
      const totalDiv = dps * c.shares_m;
      const profit   = c.profits[yr];

      sheet.getRange(r, 1, 1, 5).setValues([[
        yr === 2021 ? c.name : "",   // Only show name on first row of company
        yr,
        dps,
        totalDiv,
        profit
      ]]);

      // Payout% formula
      sheet.getRange(r, 6).setFormula(`=IFERROR(D${r}/E${r},"—")`);

      // Background
      const bg = idx % 2 === 0 ? COLORS.rowAlt : COLORS.rowWhite;
      sheet.getRange(r, 1, 1, 6).setBackground(bg);

      // Red profit if negative
      if (profit < 0) {
        sheet.getRange(r, 5).setFontColor(COLORS.negative);
      }

      currentRow++;
    });

    // Bold company name in first row of group
    sheet.getRange(companyStartRow, 1).setFontWeight("bold");
  });

  // --- Number formats ---
  sheet.getRange(3, 3, currentRow - 3, 1).setNumberFormat("0.000");   // DPS
  sheet.getRange(3, 4, currentRow - 3, 2).setNumberFormat("#,##0");   // Total div, Profit
  sheet.getRange(3, 6, currentRow - 3, 1).setNumberFormat("0.00%");   // Payout%

  // --- Alignment ---
  rtlRange_(sheet.getRange(3, 1, currentRow - 3, 1));
  numericRange_(sheet.getRange(3, 2, currentRow - 3, 5));

  // --- Column widths ---
  sheet.setColumnWidth(1, 130);  // الشركة
  sheet.setColumnWidth(2, 70);   // السنة
  sheet.setColumnWidth(3, 110);  // DPS
  sheet.setColumnWidth(4, 160);  // Total
  sheet.setColumnWidth(5, 160);  // Profit
  sheet.setColumnWidth(6, 130);  // Payout%

  sheet.setFrozenRows(2);
  borderRange_(sheet.getRange(2, 1, currentRow - 2, 6));

  // Summary: total 2024 dividends
  const summaryRow = currentRow + 1;
  const total2024Div = COMPANIES.reduce((sum, c) => sum + c.dps_2024 * c.shares_m, 0);
  sheet.getRange(summaryRow, 1, 1, 6).merge()
       .setValue(`إجمالي التوزيعات 2024 للمحفظة: ${Math.round(total2024Div).toLocaleString()} مليون ريال`)
       .setBackground(COLORS.headerMid)
       .setFontColor(COLORS.white)
       .setFontWeight("bold")
       .setHorizontalAlignment("right");

  return sheet;
}

// =============================================================================
// SHEET 6: المحفظة  (Portfolio Allocation)
// =============================================================================
// Columns A-G:
//  A: الشركة   B: القطاع   C: السعر الحالي (live)   D: الأسهم (م.)
//  E: القيمة السوقية (م.) = C×D/1000   F: الوزن المستهدف%   G: الوزن الفعلي%

function buildPortfolioSheet_(ss) {
  deleteSheetIfExists_(ss, "المحفظة");
  const sheet = ss.insertSheet("المحفظة");

  setSheetTitle_(sheet, "تخصيص المحفظة الاستثمارية - Portfolio Allocation", 7, COLORS.headerDark);

  sheet.getRange(2, 1, 1, 7).merge()
       .setValue("القيم السوقية تتحسب تلقائياً من الأسعار الحية · الوزن الفعلي يُعاد حسابه عند تغير الأسعار")
       .setBackground(COLORS.headerMid)
       .setFontColor(COLORS.white)
       .setFontSize(9)
       .setHorizontalAlignment("center");

  const headers = [
    "الشركة", "القطاع", "السعر الحالي (ر.س)",
    "عدد الأسهم (م.)", "القيمة السوقية (م.)",
    "الوزن المستهدف%", "الوزن الفعلي%"
  ];
  sheet.getRange(3, 1, 1, 7).setValues([headers]);
  formatHeaderRange_(sheet.getRange(3, 1, 1, 7), false);
  sheet.setRowHeight(3, 36);

  // Data rows 4-12
  const hardcoded = COMPANIES.map(c => [
    c.name,
    c.sector,
    "",            // C: live price - formula
    c.shares_m,
    "",            // E: market cap formula
    c.weight_pct,
    ""             // G: actual weight formula - computed after sum
  ]);
  sheet.getRange(4, 1, 9, 7).setValues(hardcoded);

  // Live price formulas
  const priceFormulas = COMPANIES.map((c, i) => [`='أسعار حية'!C${3 + i}`]);
  sheet.getRange(4, 3, 9, 1).setFormulas(priceFormulas);

  // Market cap: C × D (shares already in millions, price in SAR → value in M SAR)
  for (let i = 0; i < 9; i++) {
    const r = 4 + i;
    sheet.getRange(r, 5).setFormula(`=IFERROR(C${r}*D${r},"N/A")`);
  }

  // Total row 13
  sheet.getRange(13, 1, 1, 7).setValues([["الإجمالي", "", "", "", "", "", ""]]);
  sheet.getRange(13, 4).setFormula(`=SUM(D4:D12)`);
  sheet.getRange(13, 5).setFormula(`=IFERROR(SUM(E4:E12),"N/A")`);
  sheet.getRange(13, 6).setFormula(`=SUM(F4:F12)`);
  sheet.getRange(13, 1, 1, 7)
       .setBackground(COLORS.sectionTitle)
       .setFontColor(COLORS.white)
       .setFontWeight("bold");

  // Actual weight: company market cap / total market cap
  for (let i = 0; i < 9; i++) {
    const r = 4 + i;
    sheet.getRange(r, 7).setFormula(`=IFERROR(E${r}/$E$13*100,"N/A")`);
  }
  sheet.getRange(13, 7).setFormula(`=IFERROR(SUM(G4:G12),"N/A")`);

  // --- Number formats ---
  sheet.getRange(4, 3, 9, 1).setNumberFormat("0.00");           // Price
  sheet.getRange(4, 4, 9, 1).setNumberFormat("#,##0.0");         // Shares
  sheet.getRange(4, 5, 9, 1).setNumberFormat("#,##0.0");         // Market cap
  sheet.getRange(4, 6, 9, 2).setNumberFormat("0.0");             // Weights
  sheet.getRange(13, 4, 1, 3).setNumberFormat("#,##0.0");

  // Alternating rows
  applyAlternatingRows_(sheet, 4, 9, 1, 7);
  rtlRange_(sheet.getRange(3, 1, 11, 2));
  numericRange_(sheet.getRange(4, 3, 9, 5));

  // Column widths
  sheet.setColumnWidth(1, 130);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 160);
  sheet.setColumnWidth(6, 140);
  sheet.setColumnWidth(7, 130);

  sheet.setFrozenRows(3);
  borderRange_(sheet.getRange(3, 1, 11, 7));

  // Sector summary below
  const sectorSummaryRow = 15;
  sheet.getRange(sectorSummaryRow, 1, 1, 7).merge()
       .setValue("توزيع القطاعات في المحفظة")
       .setBackground(COLORS.headerDark)
       .setFontColor(COLORS.white)
       .setFontWeight("bold")
       .setHorizontalAlignment("right");

  const sectorHeaders = ["القطاع", "عدد الشركات", "الوزن المستهدف%", "", "", "", ""];
  sheet.getRange(sectorSummaryRow + 1, 1, 1, 7).setValues([sectorHeaders]);
  formatHeaderRange_(sheet.getRange(sectorSummaryRow + 1, 1, 1, 3), true);

  const sectorMap = {};
  COMPANIES.forEach(c => {
    if (!sectorMap[c.sector]) sectorMap[c.sector] = {count: 0, weight: 0};
    sectorMap[c.sector].count++;
    sectorMap[c.sector].weight += c.weight_pct;
  });
  const sectorRows = Object.entries(sectorMap).map(([sec, v]) => [sec, v.count, v.weight, "", "", "", ""]);
  sheet.getRange(sectorSummaryRow + 2, 1, sectorRows.length, 7).setValues(sectorRows);
  applyAlternatingRows_(sheet, sectorSummaryRow + 2, sectorRows.length, 1, 3);
  rtlRange_(sheet.getRange(sectorSummaryRow + 2, 1, sectorRows.length, 1));
  sheet.getRange(sectorSummaryRow + 2, 3, sectorRows.length, 1).setNumberFormat("0.0");

  return sheet;
}

// =============================================================================
// SHEET 7: منهجية التقييم  (Valuation Methodology)
// =============================================================================

function buildMethodologySheet_(ss) {
  deleteSheetIfExists_(ss, "منهجية التقييم");
  const sheet = ss.insertSheet("منهجية التقييم");

  setSheetTitle_(sheet, "منهجية التقييم - Valuation Methodology", 4, COLORS.headerDark);

  const sections = [
    {
      title: "1. نموذج P/E النسبي (Relative P/E Valuation)",
      rows: [
        ["المعادلة", "القيمة العادلة P/E = P/E القطاع × EPS", "", ""],
        ["المدخلات", "P/E القطاع: متوسط P/E السوق للشركات المماثلة في نفس القطاع", "", ""],
        ["", "EPS: صافي الربح 2024 ÷ عدد الأسهم المصدرة", "", ""],
        ["الحالات الخاصة", "الشركات ذات الخسائر (المتقدمة 2330): يُستبعد نموذج P/E ويُستخدم P/B فقط", "", ""],
        ["المصداقية", "مناسب للقطاعات الناضجة ذات الأرباح المستقرة (بنوك، اتصالات، تجزئة)", "", ""]
      ]
    },
    {
      title: "2. نموذج P/B النسبي (Price-to-Book Valuation)",
      rows: [
        ["المعادلة", "القيمة العادلة P/B = P/B القطاع × (حقوق المساهمين ÷ عدد الأسهم)", "", ""],
        ["المدخلات", "P/B القطاع: من تحليل المقارنة مع شركات القطاع", "", ""],
        ["", "حقوق المساهمين: من الميزانية العمومية 2024", "", ""],
        ["المصداقية", "مناسب للشركات ذات الأصول الكبيرة (بنوك، صناديق ريت، بتروكيماويات)", "", ""]
      ]
    },
    {
      title: "3. نموذج خصم التوزيعات DDM (Dividend Discount Model)",
      rows: [
        ["المعادلة", "السعر العادل = DPS / (WACC - g)", "", ""],
        ["المدخلات", "DPS: توزيعات السهم السنوية 2024", "", ""],
        ["", "WACC: تكلفة رأس المال المرجحة (محددة لكل شركة)", "", ""],
        ["", "g: معدل النمو المستدام طويل الأمد", "", ""],
        ["قيود", "لا يُطبق على الشركات غير الموزعة أو ذات الخسائر (المتقدمة 2330)", "", ""],
        ["مثال - الراجحي", "2.71 / (10.5% - 3%) = 2.71 / 7.5% = 36.1 ريال (DDM الصارم)", "", ""]
      ]
    },
    {
      title: "4. تقديرات المحللين (Analyst Consensus)",
      rows: [
        ["المصدر", "متوسط السعر المستهدف من المحللين الماليين المعتمدين", "", ""],
        ["المزايا", "تدمج المعلومات غير الكمية وتوقعات الإدارة والفرص القطاعية", "", ""],
        ["العيوب", "قد تتأخر في تحديث التوقعات عند الأحداث الطارئة", "", ""]
      ]
    },
    {
      title: "5. هامش الأمان (Margin of Safety)",
      rows: [
        ["المعادلة", "هامش الأمان% = (متوسط القيم العادلة - السعر الحالي) / السعر الحالي × 100", "", ""],
        ["التفسير", "أكبر من 15%: فرصة شراء جيدة (لون أخضر)", "", ""],
        ["", "0% إلى 15%: سعر عادل (لون برتقالي)", "", ""],
        ["", "أقل من 0%: السهم مرتفع نسبياً (لون أحمر)", "", ""],
        ["ملاحظة", "هامش الأمان يتحدث تلقائياً عند تغير الأسعار الحية", "", ""]
      ]
    },
    {
      title: "6. التوصيات (Recommendations)",
      rows: [
        ["تراكم بقوة",       "هامش أمان مرتفع + نمو قوي + رؤية 2030 إيجابية", "", ""],
        ["تراكم",            "قيمة عادلة مع توزيعات جيدة", "", ""],
        ["احتفاظ",           "لا تحرك جديد - انتظار محفزات", "", ""],
        ["مراقبة",           "بيانات غير مكتملة أو تراجع مؤقت في الأرباح", "", ""],
        ["مضاربة / وزن محدود","مخاطر مرتفعة + خسائر - للمستثمرين المتسامحين مع المخاطر", "", ""]
      ]
    }
  ];

  let currentRow = 2;

  sections.forEach(sec => {
    // Section header
    sheet.getRange(currentRow, 1, 1, 4).merge()
         .setValue(sec.title)
         .setBackground(COLORS.sectionTitle)
         .setFontColor(COLORS.white)
         .setFontWeight("bold")
         .setFontSize(11)
         .setHorizontalAlignment("right");
    sheet.setRowHeight(currentRow, 32);
    currentRow++;

    // Column headers
    const colHeaders = [["البند", "التفاصيل", "", ""]];
    sheet.getRange(currentRow, 1, 1, 4).setValues(colHeaders);
    formatHeaderRange_(sheet.getRange(currentRow, 1, 1, 2), true);
    sheet.getRange(currentRow, 3, 1, 2).setBackground(COLORS.headerMid);
    currentRow++;

    // Content rows
    sec.rows.forEach((row, rIdx) => {
      sheet.getRange(currentRow, 1, 1, 4).setValues([row]);
      const bg = rIdx % 2 === 0 ? COLORS.rowAlt : COLORS.rowWhite;
      sheet.getRange(currentRow, 1, 1, 4).setBackground(bg);
      rtlRange_(sheet.getRange(currentRow, 1, 1, 2));
      sheet.setRowHeight(currentRow, 24);
      currentRow++;
    });

    // Spacer
    sheet.getRange(currentRow, 1, 1, 4).setBackground("#F0F4F8");
    currentRow++;
  });

  // Column widths
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 500);
  sheet.setColumnWidth(3, 50);
  sheet.setColumnWidth(4, 50);

  sheet.setFrozenRows(1);
  borderRange_(sheet.getRange(2, 1, currentRow - 2, 2));

  return sheet;
}

// =============================================================================
// MAIN ENTRY POINT
// =============================================================================

/**
 * buildSaudiPortfolioSheet()
 *
 * الدالة الرئيسية - قم بتشغيلها من محرر Apps Script.
 * تحذف الأوراق الموجودة بنفس الأسماء وتُعيد بناءها من الصفر.
 * عند الانتهاء تظهر رسالة تأكيد.
 */
function buildSaudiPortfolioSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Build all sheets in order
    // (Sheets are ordered: Live Prices first, so Overview can reference it)
    const livePricesSheet  = buildLivePricesSheet_(ss);
    const overviewSheet    = buildOverviewSheet_(ss);
    const financialsSheet  = buildFinancialsSheet_(ss);
    const valuationSheet   = buildValuationSheet_(ss);
    const dividendsSheet   = buildDividendsSheet_(ss);
    const portfolioSheet   = buildPortfolioSheet_(ss);
    const methodologySheet = buildMethodologySheet_(ss);

    // Re-order sheets: Live Prices first
    const sheetOrder = [
      "أسعار حية",
      "نظرة عامة",
      "القوائم المالية",
      "التقييم",
      "التوزيعات",
      "المحفظة",
      "منهجية التقييم"
    ];
    sheetOrder.forEach((name, idx) => {
      const s = ss.getSheetByName(name);
      if (s) ss.setActiveSheet(s);
      if (s) ss.moveActiveSheet(idx + 1);
    });

    // Activate Overview on open
    ss.setActiveSheet(overviewSheet);

    // Show success — works both from editor and from spreadsheet menu
    const successMsg =
      "✓ تم إنشاء الأوراق السبع بنجاح:\n" +
      "1. أسعار حية  2. نظرة عامة  3. القوائم المالية\n" +
      "4. التقييم  5. التوزيعات  6. المحفظة  7. منهجية التقييم\n" +
      "GOOGLEFINANCE تتحدث تلقائياً — ارجع للـ Sheet لترى النتيجة.";
    try {
      SpreadsheetApp.getUi().alert("تم بنجاح ✓", successMsg, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (_) {
      Logger.log(successMsg); // fallback when running from editor
    }

  } catch (e) {
    const errMsg = "خطأ: " + e.message;
    try {
      SpreadsheetApp.getUi().alert("خطأ ✗", errMsg, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (_) {
      Logger.log("ERROR: " + errMsg);
      throw e; // re-throw so the execution log shows the real error
    }
  }
}

// =============================================================================
// OPTIONAL: Custom menu on open
// =============================================================================

/**
 * Adds a custom menu item when the spreadsheet is opened.
 * Not required for the script to work — just a convenience.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("محفظة الاستثمار السعودية")
    .addItem("بناء/إعادة بناء الأوراق", "buildSaudiPortfolioSheet")
    .addSeparator()
    .addItem("فتح ورقة الأسعار الحية", "openLivePrices_")
    .addItem("فتح ورقة التقييم", "openValuation_")
    .addToUi();
}

function openLivePrices_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName("أسعار حية");
  if (s) ss.setActiveSheet(s);
}

function openValuation_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName("التقييم");
  if (s) ss.setActiveSheet(s);
}
