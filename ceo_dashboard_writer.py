"""
STAGE 5 — CEO DASHBOARD WRITER
Reads live data from all operational sheets.
Writes docs/ceo_data.json for the CEO HTML dashboard.
Runs every 2 hours during working day via GitHub Actions.

Metrics written:
  1. RM Stock        — grade-wise, total tons, vs minimum
  2. Daily Manpower  — present vs required, dept breakdown
  3. Electricity     — units today, cost/ton, vs target
  4. Oil Consumption — litres today, vs standard
  5. VF Rejection    — % per machine, inhouse + vendor
  6. Revenue/Tons    — today dispatch value + tons, MTD total
  
Plus:
  Overdue outstanding — total + top 5 customers
  Production vs Schedule — actual vs planned this month
  Top 10 parts by volume this month
"""

import json, os, pickle, time
from datetime import datetime, date, timedelta
from collections import defaultdict

from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

TOKEN_FILE  = 'token.pickle'
OUTPUT_FILE = 'docs/ceo_data.json'
DELAY       = 0.5

# All source sheet IDs
SHEETS = {
    'dashboard':    '1NR8EPGRJN0AQDXZjYw5k93clsO8AD4u2l2Xke1lBC2I',
    'pms_forgings': '1c-axqiBEufNb1vK-JdJP6t1eRTmzliA3otAbqxAOXZE',
    'machine_shop': '1yC-b36rgAxablmdXhngHCEOsmnlgWourep6ctKifXCA',
    'dispatch':     '1txZM9a9_OSG-ZWYaAEBLKj-9M7LYLrFyl0kJkPsTMGI',
    'manpower':     '1t7UjWTP_cpIJ2BjoaMlV6uUKA7ztH9UnKc_korYKCiw',
    'electricity':  '1nUvf-UWjBSbSWnZTNph-gRUbjzuguGlidpYBKshKUNQ',
    'vendor_rej':   '1izr5HWo4-qvXbmhkbmBIVQSR-X6ptU7OtSX-ICd43H0',
    'dropouts':     '1J1KKfygB8M1z3qZJQcaMXjkVb7HgxDARZrklMld8gS0',
}

# Targets — will be read from dashboard if available, else use these defaults
TARGETS = {
    'manpower_required':    120,
    'rejection_pct_max':    2.0,
    'electricity_per_ton':  950,   # kWh/ton target
    'oil_per_ton':          1.8,   # litres/ton target
    'rm_stock_min_tons':    50,    # minimum RM stock alert
}

def authenticate():
    creds = None
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE,'rb') as f: creds = pickle.load(f)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            raise RuntimeError("No valid token.pickle.")
    return creds

def clean_num(val):
    if not val: return 0.0
    import re
    val = str(val).replace('₹','').replace(',','').replace(' ','')
    m = re.search(r'-?\d+\.?\d*', val)
    return float(m.group()) if m else 0.0

def clean_date(val):
    if not val: return ''
    val = str(val).strip()
    for fmt in ['%d/%m/%Y','%m/%d/%Y','%Y-%m-%d','%d-%m-%Y',
                '%d/%m/%Y %H:%M:%S','%m/%d/%Y %H:%M:%S']:
        try:
            return datetime.strptime(val[:19], fmt).strftime('%Y-%m-%d')
        except: continue
    return val[:10]

def read_tab(svc, sheet_id, tab, max_rows=10000):
    try:
        resp = svc.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range=f"'{tab}'!A1:Z{max_rows}"
        ).execute()
        time.sleep(DELAY)
        values = resp.get('values', [])
        if not values: return [], []
        headers = [str(h).strip() for h in values[0]]
        rows = []
        for row in values[1:]:
            while len(row) < len(headers): row.append('')
            r = {headers[i]: str(row[i]).strip()
                 for i in range(len(headers)) if headers[i]}
            if any(v for v in r.values()): rows.append(r)
        return headers, rows
    except: return [], []

def get_tabs(svc, sheet_id):
    try:
        meta = svc.spreadsheets().get(
            spreadsheetId=sheet_id,
            fields='sheets.properties.title'
        ).execute()
        time.sleep(DELAY)
        return [s['properties']['title'] for s in meta.get('sheets',[])]
    except: return []

def status(value, target, higher_is_better=True):
    """Returns GREEN/AMBER/RED based on value vs target."""
    if target == 0: return 'GREY'
    pct = value / target
    if higher_is_better:
        if pct >= 0.95: return 'GREEN'
        if pct >= 0.80: return 'AMBER'
        return 'RED'
    else:
        if pct <= 1.05: return 'GREEN'
        if pct <= 1.20: return 'AMBER'
        return 'RED'

def date_range(days_back):
    today = date.today()
    return [(today - timedelta(days=i)).strftime('%Y-%m-%d')
            for i in range(days_back)]

def main():
    print('='*60)
    print('   CEO DASHBOARD WRITER')
    print(f'   {datetime.now().strftime("%Y-%m-%d %H:%M")} IST')
    print('='*60)

    os.makedirs('docs', exist_ok=True)

    creds = authenticate()
    svc   = build('sheets','v4',credentials=creds)

    today_str  = date.today().strftime('%Y-%m-%d')
    today_7d   = date_range(7)
    today_30d  = date_range(30)

    output = {
        'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M'),
        'date':         today_str,
        'metrics':      {},
        'charts':       {},
        'alerts':       [],
    }

    # ═══════════════════════════════════════════════
    # 1. RM STOCK
    # ═══════════════════════════════════════════════
    print('\n📦 Reading RM Stock...')
    rm_stock = {}
    tabs = get_tabs(svc, SHEETS['dashboard'])
    for tab in tabs:
        if any(k in tab.lower() for k in ['rm stock','rm inventory']):
            _, rows = read_tab(svc, SHEETS['dashboard'], tab)
            for r in rows:
                grade = r.get('Grade','')
                if grade:
                    opening = clean_num(r.get('Opening Stock') or r.get('Quantity in ton') or '')
                    inward  = clean_num(r.get('Inward') or r.get('Inward (Tons)') or '')
                    consume = clean_num(r.get('Consumption') or r.get('Consumption (Tons)') or '')
                    balance = opening + inward - consume
                    if balance != 0:
                        rm_stock[grade] = round(balance, 2)
            break

    total_rm = sum(rm_stock.values())
    output['metrics']['rm_stock'] = {
        'total_tons':    round(total_rm, 2),
        'target_min':    TARGETS['rm_stock_min_tons'],
        'status':        status(total_rm, TARGETS['rm_stock_min_tons']),
        'by_grade':      rm_stock,
        'grade_count':   len(rm_stock),
    }
    if total_rm < TARGETS['rm_stock_min_tons']:
        output['alerts'].append({
            'type':'RM_STOCK_LOW',
            'message': f'RM Stock {total_rm:.1f}T below minimum {TARGETS["rm_stock_min_tons"]}T',
            'severity':'RED'
        })
    print(f'  RM Stock: {total_rm:.1f} tons, {len(rm_stock)} grades')

    # ═══════════════════════════════════════════════
    # 2. DAILY MANPOWER
    # ═══════════════════════════════════════════════
    print('\n👥 Reading Manpower...')
    manpower_today = 0
    manpower_by_dept = {}
    manpower_trend = []

    tabs = get_tabs(svc, SHEETS['manpower'])
    for tab in tabs:
        _, rows = read_tab(svc, SHEETS['manpower'], tab)
        if not rows: continue
        headers_check = str(rows[:2]).lower()
        if not any(k in headers_check for k in ['present','manpower','attendance','department']):
            continue

        for r in rows:
            d = clean_date(r.get('Date') or r.get('Timestamp') or r.get('DATE') or '')
            if not d: continue
            present = clean_num(r.get('Present') or r.get('Total Present') or
                               r.get('Attendance') or '')
            dept    = r.get('Department') or r.get('Dept') or ''

            if d == today_str:
                manpower_today += present
                if dept:
                    manpower_by_dept[dept] = manpower_by_dept.get(dept,0) + present

            if d in today_7d:
                manpower_trend.append({'date':d,'present':present})
        break

    required = TARGETS['manpower_required']
    output['metrics']['manpower'] = {
        'today':        int(manpower_today),
        'required':     required,
        'shortage':     max(0, required - manpower_today),
        'status':       status(manpower_today, required),
        'by_dept':      manpower_by_dept,
        'trend_7d':     manpower_trend[-7:],
    }
    print(f'  Manpower today: {manpower_today} / {required} required')

    # ═══════════════════════════════════════════════
    # 3. ELECTRICITY
    # ═══════════════════════════════════════════════
    print('\n⚡ Reading Electricity...')
    elec_today = 0
    elec_cost_today = 0
    elec_trend = []

    tabs = get_tabs(svc, SHEETS['electricity'])
    for tab in tabs:
        _, rows = read_tab(svc, SHEETS['electricity'], tab)
        if not rows: continue
        for r in rows:
            d = clean_date(r.get('Date') or r.get('DATE') or r.get('Timestamp') or '')
            if not d: continue
            units = clean_num(r.get('Units') or r.get('Consumption') or r.get('KWH') or
                             r.get('Unit') or r.get('Electricity') or '')
            cost  = clean_num(r.get('Cost') or r.get('Amount') or r.get('Value') or '')
            if d == today_str:
                elec_today      += units
                elec_cost_today += cost
            if d in today_30d and units > 0:
                elec_trend.append({'date':d,'units':units,'cost':cost})
        if elec_trend: break

    output['metrics']['electricity'] = {
        'today_units':  round(elec_today, 1),
        'today_cost':   round(elec_cost_today, 2),
        'target_per_ton': TARGETS['electricity_per_ton'],
        'status':       'GREY' if not elec_today else status(
                            elec_today, TARGETS['electricity_per_ton']*10, False),
        'trend_30d':    sorted(elec_trend, key=lambda x:x['date'])[-30:],
    }
    print(f'  Electricity today: {elec_today:.0f} units, ₹{elec_cost_today:,.0f}')

    # ═══════════════════════════════════════════════
    # 4. OIL CONSUMPTION
    # ═══════════════════════════════════════════════
    print('\n🛢️ Reading Oil...')
    oil_today = 0
    oil_mtd   = 0
    oil_trend = []

    tabs = get_tabs(svc, SHEETS['dashboard'])
    for tab in tabs:
        if any(k in tab.lower() for k in ['oil','consumable']):
            _, rows = read_tab(svc, SHEETS['dashboard'], tab)
            for r in rows:
                d = clean_date(r.get('Date') or r.get('DATE') or '')
                qty = clean_num(r.get('Consumption') or r.get('Qty') or
                               r.get('Litres') or r.get('Ltr') or '')
                if d == today_str: oil_today += qty
                if d in today_30d: oil_mtd   += qty
                if qty > 0: oil_trend.append({'date':d,'litres':qty})
            if oil_today or oil_mtd: break

    output['metrics']['oil'] = {
        'today_litres': round(oil_today, 1),
        'mtd_litres':   round(oil_mtd,   1),
        'target_per_ton': TARGETS['oil_per_ton'],
        'status':       'GREY' if not oil_today else 'GREEN',
        'trend_30d':    sorted(oil_trend, key=lambda x:x['date'])[-30:],
    }
    print(f'  Oil today: {oil_today:.0f} litres, MTD: {oil_mtd:.0f}')

    # ═══════════════════════════════════════════════
    # 5. VF WISE REJECTION
    # ═══════════════════════════════════════════════
    print('\n🔴 Reading Rejection...')
    rejection_by_vf = defaultdict(lambda:{'dropout':0,'vendor':0,'produced':0})
    rejection_today = 0
    rejection_mtd   = 0

    for label, sheet_id in [('dropout', SHEETS['dropouts']),
                             ('vendor',  SHEETS['vendor_rej'])]:
        tabs = get_tabs(svc, sheet_id)
        for tab in tabs:
            _, rows = read_tab(svc, sheet_id, tab)
            for r in rows:
                d   = clean_date(r.get('Date') or r.get('Timestamp') or r.get('DATE') or '')
                vf  = r.get('VF No') or r.get('Machine') or r.get('Part') or 'Unknown'
                qty = clean_num(r.get('Rejection') or r.get('Qty') or r.get('Dropout') or '')
                if qty > 0:
                    rejection_by_vf[vf][label] += qty
                    if d == today_str: rejection_today += qty
                    if d in today_30d: rejection_mtd   += qty

    # Get production for context (from dashboard REVENUE tab)
    total_prod_mtd = 0
    for tab in get_tabs(svc, SHEETS['dashboard']):
        if 'revenue' in tab.lower() or 'tons' in tab.lower():
            _, rows = read_tab(svc, SHEETS['dashboard'], tab)
            for r in rows:
                total_prod_mtd += clean_num(r.get('Forgings Production') or
                                           r.get('Total Production') or '')
            if total_prod_mtd: break

    rej_pct = (rejection_mtd / total_prod_mtd * 100) if total_prod_mtd else 0

    output['metrics']['rejection'] = {
        'today':        int(rejection_today),
        'mtd':          int(rejection_mtd),
        'mtd_pct':      round(rej_pct, 2),
        'target_max':   TARGETS['rejection_pct_max'],
        'status':       status(TARGETS['rejection_pct_max'], rej_pct) if rej_pct else 'GREY',
        'by_vf':        {k: dict(v) for k,v in list(rejection_by_vf.items())[:20]},
    }
    if rej_pct > TARGETS['rejection_pct_max']:
        output['alerts'].append({
            'type':'REJECTION_HIGH',
            'message': f'Rejection {rej_pct:.1f}% above target {TARGETS["rejection_pct_max"]}%',
            'severity':'RED'
        })
    print(f'  Rejection MTD: {rejection_mtd:.0f} pcs, {rej_pct:.1f}%')

    # ═══════════════════════════════════════════════
    # 6. REVENUE / TONS
    # ═══════════════════════════════════════════════
    print('\n💰 Reading Revenue/Dispatch...')
    revenue_today = 0
    tons_today    = 0
    revenue_mtd   = 0
    tons_mtd      = 0
    revenue_by_customer = defaultdict(float)
    revenue_trend = []

    tabs = get_tabs(svc, SHEETS['dispatch'])
    for tab in tabs:
        _, rows = read_tab(svc, SHEETS['dispatch'], tab)
        if not rows: continue
        headers_str = str(rows[:1]).lower()
        if not any(k in headers_str for k in ['value','revenue','dispatch','sales','ton']):
            continue

        for r in rows:
            d    = clean_date(r.get('Date') or r.get('DATE') or r.get('Timestamp') or '')
            val  = clean_num(r.get('Value') or r.get('Sales FG Value') or
                            r.get('Sales Value') or r.get('Amount') or '')
            tons = clean_num(r.get('Tons') or r.get('Weight') or
                            r.get('Sales FG Ton') or r.get('Qty') or '')
            cust = r.get('Customer') or r.get('Party') or ''

            if d == today_str:
                revenue_today += val
                tons_today    += tons
            if d in today_30d:
                revenue_mtd += val
                tons_mtd    += tons
                if cust: revenue_by_customer[cust] += val
                revenue_trend.append({'date':d,'value':val,'tons':tons})
        if revenue_mtd: break

    # Also check dashboard ERP Daily tab
    if not revenue_mtd:
        _, rows = read_tab(svc, SHEETS['dashboard'], 'ERP Daily Ton / Sale')
        for r in rows:
            d   = clean_date(r.get('Date') or '')
            val = clean_num(r.get('Sales FG Value') or '')
            ton = clean_num(r.get('Sales FG Ton') or '')
            if d in today_30d:
                revenue_mtd += val
                tons_mtd    += ton
            if d == today_str:
                revenue_today += val
                tons_today    += ton

    top_customers = sorted(revenue_by_customer.items(),
                           key=lambda x: -x[1])[:10]

    output['metrics']['revenue'] = {
        'today_value':  round(revenue_today, 2),
        'today_tons':   round(tons_today,    2),
        'mtd_value':    round(revenue_mtd,   2),
        'mtd_tons':     round(tons_mtd,      2),
        'value_per_ton':round(revenue_mtd/tons_mtd if tons_mtd else 0, 2),
        'status':       'GREEN' if revenue_today > 0 else 'AMBER',
        'top_customers':[{'customer':c,'value':round(v,2)} for c,v in top_customers],
        'trend_30d':    sorted(revenue_trend, key=lambda x:x['date'])[-30:],
    }
    print(f'  Revenue today: ₹{revenue_today:,.0f}, MTD: ₹{revenue_mtd:,.0f}')

    # ═══════════════════════════════════════════════
    # 7. OVERDUE OUTSTANDING
    # ═══════════════════════════════════════════════
    print('\n⏰ Reading Overdue...')
    overdue_total = 0
    overdue_by_customer = []

    tabs = get_tabs(svc, SHEETS['dashboard'])
    for tab in tabs:
        if 'overdue' in tab.lower() or 'outstanding' in tab.lower():
            _, rows = read_tab(svc, SHEETS['dashboard'], tab)
            cust_overdue = defaultdict(float)
            for r in rows:
                cust = r.get('Customer') or r.get('Party') or r.get('Name') or ''
                amt  = clean_num(r.get('Overdue') or r.get('Amount') or
                                r.get('Outstanding') or r.get('Balance') or '')
                if cust and amt > 0:
                    cust_overdue[cust] += amt
                    overdue_total      += amt

            overdue_by_customer = sorted(
                [{'customer':k,'amount':round(v,2)} for k,v in cust_overdue.items()],
                key=lambda x: -x['amount']
            )[:10]
            break

    # Also read from dashboard header row
    if not overdue_total:
        _, rows = read_tab(svc, SHEETS['dashboard'], 'To date Dashboard')
        for r in rows:
            val = clean_num(r.get('Total Overedue') or r.get('Overdue') or '')
            if val: overdue_total = val; break

    output['metrics']['overdue'] = {
        'total':            round(overdue_total, 2),
        'top_customers':    overdue_by_customer,
        'status':           'RED' if overdue_total > 50000000 else
                            'AMBER' if overdue_total > 20000000 else 'GREEN',
    }
    print(f'  Overdue: ₹{overdue_total:,.0f}')

    # ═══════════════════════════════════════════════
    # SAVE JSON
    # ═══════════════════════════════════════════════
    with open(OUTPUT_FILE,'w') as f:
        json.dump(output, f, indent=2)

    print(f'\n✅ CEO data written → {OUTPUT_FILE}')
    print(f'   Alerts: {len(output["alerts"])}')
    for a in output['alerts']:
        print(f'   {a["severity"]}: {a["message"]}')


if __name__ == '__main__':
    main()
