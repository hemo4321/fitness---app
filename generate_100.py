from datetime import date

PORTFOLIO = '100'
OUT = '/tmp/portfolio_100.csv'

def q(v):
    if v is None or str(v).strip() == '': return ''
    return '"' + str(v).replace('"','""') + '"'

def n(v):
    if v is None: return ''
    f = float(v)
    if f == int(f): return str(int(f))
    return f'{f:.10g}'

def ds(d): return d.strftime('%Y-%m-%d') + ' GMT+0300'

HEADER = 'Id,Symbol,Name,Display Symbol,Exchange,Portfolio,Currency,Shares Owned,Cost Per Share,Commission,Transaction Date,Transaction Time,Purchase Exchange Rate,Type,Accounting,Accounting Execution Ids,Notes,OutgoingCashLink'

P = PORTFOLIO

def cash_row(rid, d, typ, amount, notes=''):
    return ','.join([q(rid),q('SAR=CASH'),'','','',q(P),q('SAR'),
        q(n(amount)),q('1'),q('0'),q(ds(d)),q('10:00:00'),'',q(typ),'','',
        q(notes) if notes else '',''])

def stock_header(rid, sym, en, ar):
    return ','.join([q(rid),q(sym),q(en),q(ar),q('SAU'),q(P),q('SAR'),
        '','','','','','','','','','',''])

def stock_buy(rid, d, sym, en, ar, shares, amount):
    price = round(amount / shares, 4) if shares > 0 else 0
    return ','.join([q(rid),q(sym),q(en),q(ar),q('SAU'),q(P),q('SAR'),
        q(n(shares)),q(n(price)),q('0'),q(ds(d)),q('10:00:00'),'',q('Buy'),
        '','','',''])

def stock_sell(rid, d, sym, en, ar, shares, amount):
    price = round(amount / shares, 4) if shares > 0 else 0
    return ','.join([q(rid),q(sym),q(en),q(ar),q('SAU'),q(P),q('SAR'),
        q(n(shares)),q(n(price)),q('0'),q(ds(d)),q('10:00:00'),'',q('Sell'),
        q('Weighted Average'),'','',''])

def stock_div(rid, d, sym, en, ar, amount, dps, notes):
    return ','.join([q(rid),q(sym),q(en),q(ar),q('SAU'),q(P),q('SAR'),
        q(n(amount)),q(n(dps)),q('0'),q(ds(d)),q('10:00:00'),'',q('Dividend'),
        q('FIFO'),'',q(notes),''])

# Stock master data
S = {
    '4002': ('4002.SR', 'Mouwasat Medical Services Co.',             'المواساة'),
    '7202': ('7202.SR', 'Arabian Internet And Communications Co.',   'سلوشنز'),
    '7203': ('7203.SR', 'Elm Co.',                                   'علم'),
    '6004': ('6004.SR', 'Saudi Airlines Catering Co.',               'التموين'),
    '4263': ('4263.SR', 'SAL Saudi Logistics Services Co.',          'سال'),
    '4003': ('4003.SR', 'United Electronics Co.',                    'اكسترا'),
    '4260': ('4260.SR', 'Budget Saudi Co.',                          'بدجت'),
}

lines = [HEADER, '']
rid = 1

# SAR=CASH header row (id=1)
lines.append(','.join([q(1),q('SAR=CASH'),'','','',q(P),q('SAR'),'','','','','','','','','','','']))
rid = 2

# ─── CASH SECTION ─────────────────────────────────────────────────────────────

# Deposit 84,000 – start of portfolio "100" (row 9 of PDF, 2026-02-23)
lines.append(cash_row(rid, date(2026,2,23), 'Buy',  84000.00)); rid+=1

# 2026-02-24: buy MOUWASAT 88 sh
sym,en,ar = S['4002']
lines.append(cash_row(rid, date(2026,2,24), 'Sell', 5672.64,  f'Purchased {sym} - {en}')); rid+=1

# 2026-02-24: buy SOLUTIONS 410 sh
sym,en,ar = S['7202']
lines.append(cash_row(rid, date(2026,2,24), 'Sell', 78326.86, f'Purchased {sym} - {en}')); rid+=1

# 2026-03-30: sell SOLUTIONS 370 sh
lines.append(cash_row(rid, date(2026,3,30), 'Buy',  72024.64, f'Sold {sym} - {en}')); rid+=1

# 2026-04-02: buy Elm 13 sh
sym,en,ar = S['7203']
lines.append(cash_row(rid, date(2026,4, 2), 'Sell', 6901.97,  f'Purchased {sym} - {en}')); rid+=1

# 2026-04-19: buy CATRION 100 sh
sym,en,ar = S['6004']
lines.append(cash_row(rid, date(2026,4,19), 'Sell', 7132.38,  f'Purchased {sym} - {en}')); rid+=1

# 2026-04-19: buy SAL 20 sh
sym,en,ar = S['4263']
lines.append(cash_row(rid, date(2026,4,19), 'Sell', 3371.85,  f'Purchased {sym} - {en}')); rid+=1

# 2026-04-19: buy EXTRA 80 sh
sym,en,ar = S['4003']
lines.append(cash_row(rid, date(2026,4,19), 'Sell', 6571.40,  f'Purchased {sym} - {en}')); rid+=1

# 2026-04-20: buy MOUWASAT 60 sh
sym,en,ar = S['4002']
lines.append(cash_row(rid, date(2026,4,20), 'Sell', 4303.47,  f'Purchased {sym} - {en}')); rid+=1

# 2026-04-30: buy BUDGET 170 sh
sym,en,ar = S['4260']
lines.append(cash_row(rid, date(2026,4,30), 'Sell', 7067.26,  f'Purchased {sym} - {en}')); rid+=1

# 2026-05-10: buy BUDGET 180 sh
lines.append(cash_row(rid, date(2026,5,10), 'Sell', 6321.76,  f'Purchased {sym} - {en}')); rid+=1

# 2026-05-13: buy SAL 20 sh
sym,en,ar = S['4263']
lines.append(cash_row(rid, date(2026,5,13), 'Sell', 3287.71,  f'Purchased {sym} - {en}')); rid+=1

# 2026-05-14: MOUWASAT dividend 166.50 SAR
sym,en,ar = S['4002']
lines.append(cash_row(rid, date(2026,5,14), 'Buy',  166.50,   f'Dividends from {sym} - {en}')); rid+=1

# ─── STOCK SECTIONS ───────────────────────────────────────────────────────────

# 4002 المواساة
sym,en,ar = S['4002']
lines.append('')
lines.append(stock_header(rid, sym, en, ar)); rid+=1
lines.append(stock_buy(rid, date(2026,2,24), sym, en, ar, 88,  5672.64)); rid+=1
lines.append(stock_buy(rid, date(2026,4,20), sym, en, ar, 60,  4303.47)); rid+=1
lines.append(stock_div(rid, date(2026,5,14), sym, en, ar, 166.50, 1.125, '1.125 (H2-2025)')); rid+=1

# 7202 سلوشنز  (bought 410, sold 370 → 40 remain)
sym,en,ar = S['7202']
lines.append('')
lines.append(stock_header(rid, sym, en, ar)); rid+=1
lines.append(stock_buy(rid,  date(2026,2,24), sym, en, ar, 410, 78326.86)); rid+=1
lines.append(stock_sell(rid, date(2026,3,30), sym, en, ar, 370, 72024.64)); rid+=1

# 7203 علم
sym,en,ar = S['7203']
lines.append('')
lines.append(stock_header(rid, sym, en, ar)); rid+=1
lines.append(stock_buy(rid, date(2026,4, 2), sym, en, ar, 13, 6901.97)); rid+=1

# 6004 التموين
sym,en,ar = S['6004']
lines.append('')
lines.append(stock_header(rid, sym, en, ar)); rid+=1
lines.append(stock_buy(rid, date(2026,4,19), sym, en, ar, 100, 7132.38)); rid+=1

# 4263 سال
sym,en,ar = S['4263']
lines.append('')
lines.append(stock_header(rid, sym, en, ar)); rid+=1
lines.append(stock_buy(rid, date(2026,4,19), sym, en, ar, 20, 3371.85)); rid+=1
lines.append(stock_buy(rid, date(2026,5,13), sym, en, ar, 20, 3287.71)); rid+=1

# 4003 اكسترا
sym,en,ar = S['4003']
lines.append('')
lines.append(stock_header(rid, sym, en, ar)); rid+=1
lines.append(stock_buy(rid, date(2026,4,19), sym, en, ar, 80, 6571.40)); rid+=1

# 4260 بدجت
sym,en,ar = S['4260']
lines.append('')
lines.append(stock_header(rid, sym, en, ar)); rid+=1
lines.append(stock_buy(rid, date(2026,4,30), sym, en, ar, 170, 7067.26)); rid+=1
lines.append(stock_buy(rid, date(2026,5,10), sym, en, ar, 180, 6321.76)); rid+=1

with open(OUT, 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines))

print(f"Portfolio 100: {len(lines)} lines, IDs 1-{rid-1}")

# ─── Cash balance verification ────────────────────────────────────────────────
cash_in  = 84000.00 + 72024.64 + 166.50
cash_out = (5672.64 + 78326.86 + 6901.97 + 7132.38 + 3371.85
           + 6571.40 + 4303.47 + 7067.26 + 6321.76 + 3287.71)
balance  = cash_in - cash_out
print(f"\nCash verification:")
print(f"  In  : {cash_in:,.2f}")
print(f"  Out : {cash_out:,.2f}")
print(f"  Net : {balance:,.2f}  (PDF closing balance: 27,233.84)")

# ─── Preview ──────────────────────────────────────────────────────────────────
print("\nFirst 30 lines:")
for i, l in enumerate(lines[:30]):
    print(f"  {i+1:3}: {l[:100]}")
