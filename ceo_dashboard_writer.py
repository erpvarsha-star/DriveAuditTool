"""
STAGE 5 — CEO DASHBOARD WRITER
Reads live operational data from Google Sheets and writes
ceo_data.json for the CEO HTML dashboard.

Runs every 2 hours during working day via GitHub Actions.
CEO opens ceo.html on any device — sees live metrics.

Metrics:
  1. RM Stock — current stock vs minimum level by grade
  2. Daily Manpower — present vs required
  3. Electricity — units today vs target
  4. Oil Consumption — litres used vs standard
  5. VF Wise Rejection — % per machine
  6. Revenue / Tons — dispatch value + weight today + MTD

Run: python ceo_dashboard_writer.py
"""

import json
import os
import pickle
import time
from datetime import datetime, timezone, timedelta

from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ═══════════════════════════════════════════════════════
#  CONFIG
# ═══════════════════════════════════════════════════════

DASHBOARD_ID    = '1NR8EPGRJN0AQDXZjYw5k93clsO8AD4u2l2Xke1lBC2I'
PMS_ID          = '1c-axqiBEufNb1vK-JdJP6t1eRTmzliA3otAbqxAOXZE'
DISPATCH_ID     = '1txZM9a9_OSG-ZWYaAEBLKj-9M7LYLrFyl0kJkPsTMGI'
MANPOWER_ID     = '1t7UjWTP_cpIJ2BjoaMlV6uUKA7ztH9UnKc_korYKCiw'
ELECTRICITY_ID  = '1nUvf-UWjBSbSWnZTNph-gRUbjzuguGlidpYBKshKUNQ'

TOKEN_FILE  = 'token.pickle'
OUTPUT_FILE = 'docs/ceo_data.json'
DELAY       = 0.5

# Targets — update these to match your actual targets
TARGETS = {
    'rm_stock_min_tons':      50,      # minimum RM stock in tons
    'manpower_required':      120,     # total required headcount
    'electricity_daily_units': 8000,   # target units per day
    'oil_min_stock_ltrs':     500,     # minimum oil stock
    'rejection_max_pct':      2.0,     # max acceptable rejection %
    'daily_revenue_target':   3000000, # ₹30 lakh daily revenue target
    'monthly_revenue_target': 60000000,# ₹6 Cr monthly target
}

# ═══════════════════════════════════════════════════════
#  AUTH
# ═══════════════════════════════════════════════════════

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

# ═══════════════════════════════════════════════════════
#  HELPERS
# ═══════════════════════════════════════════════════════

def read_range(svc, sheet_id, range_str):
    try:
        resp = svc.spreadsheets().values().get(
            spreadsheetId=sheet_id, range=range_str
        ).execute()
        time.sleep(DELAY)
        return resp.get('values', [])
    except HttpError as e:
        if e.resp.status == 429:
            time.sleep(60)
            return read_range(svc, sheet_id, range_str)
        return []
    except Exception:
        return []


def get_tabs(svc, sheet_id):
    try:
        meta = svc.spreadsheets().get(
            spreadsheetId=sheet_id,
            fields='sheets.properties.title'
        ).execute()
        time.sleep(DELAY)
        return [s['properties']['title'] for s in meta.get('sheets', [])]
    except Exception:
        return []


def find_tab(tabs, keywords):
    for kw in keywords:
        for tab in tabs:
            if kw.lower() in tab.lower():
                return tab
    return tabs[0] if tabs else None


def safe_float(v, default=0):
    try:
        return float(str(v).replace(',','').replace('₹','').replace(' ','').replace('%',''))
    except Exception:
        return default


def today_str():
    ist = timezone(timedelta(hours=5, minutes=30))
    return datetime.now(ist).strftime('%d/%m/%Y')


def status(value, target, higher_is_better=True):
    """Return GREEN or RED based on target comparison."""
    if not target or not value:
        return 'GREY'
    try:
        v = float(str(value).replace(',',''))
        t = float(str(target).replace(',',''))
        if higher_is_better:
            return 'GREEN' if v >= t else 'RED'
        else:
            return 'GREEN' if v <= t else 'RED'
    except Exception:
        return 'GREY'


def trend(current, previous):
    """Return UP, DOWN, or FLAT."""
    try:
        c = float(str(current).replace(',',''))
        p = float(str(previous).replace(',',''))
        if c > p * 1.02: return 'UP'
        if c < p * 0.98: return 'DOWN'
        return 'FLAT'
    except Exception:
        return 'FLAT'

# ═══════════════════════════════════════════════════════
#  METRIC READERS
# ═══════════════════════════════════════════════════════

def get_rm_stock(svc):
    """Read RM Stock from dashboard."""
    print("  📦 Reading RM Stock...")
    tabs = get_tabs(svc, DASHBOARD_ID)
    tab  = find_tab(tabs, ['rm stock', 'rm inventory', 'raw material stock'])

    if not tab:
        return {'total_tons': 0, 'grades': [], 'status': 'GREY', 'tab_found': False}

    rows = read_range(svc, DASHBOARD_ID, f"'{tab}'!A1:H50")
    if not rows:
        return {'total_tons': 0, 'grades': [], 'status': 'GREY', 'tab_found': True}

    grades     = []
    total_tons = 0

    for row in rows[1:]:  # skip header
        if not row or not row[0]:
            continue
        grade   = str(row[0]).strip()
        opening = safe_float(row[1]) if len(row) > 1 else 0
        inward  = safe_float(row[3]) if len(row) > 3 else 0
        consump = safe_float(row[5]) if len(row) > 5 else 0
        closing = opening + inward - consump

        if grade and grade not in ['Grade', 'GRADE', '']:
            grades.append({
                'grade':       grade,
                'opening':     opening,
                'inward':      inward,
                'consumption': consump,
                'closing':     round(closing, 2),
            })
            total_tons += max(closing, 0)

    return {
        'total_tons': round(total_tons, 2),
        'grades':     grades[:20],  # top 20 grades
        'status':     status(total_tons, TARGETS['rm_stock_min_tons']),
        'target_tons': TARGETS['rm_stock_min_tons'],
        'tab_found':  True,
    }


def get_manpower(svc):
    """Read Daily Manpower."""
    print("  👥 Reading Manpower...")
    tabs = get_tabs(svc, MANPOWER_ID)
    tab  = find_tab(tabs, ['form response', 'manpower', 'data', 'daily'])

    if not tab:
        # Try dashboard
        dash_tabs = get_tabs(svc, DASHBOARD_ID)
        mp_tab    = find_tab(dash_tabs, ['manpower', 'daily manpower'])
        if mp_tab:
            rows = read_range(svc, DASHBOARD_ID, f"'{mp_tab}'!A1:Z5")
            print(f"    Dashboard manpower tab found: {mp_tab}")

    rows = read_range(svc, MANPOWER_ID, f"'{tab}'!A1:Z10") if tab else []

    # Try to get today's total from dashboard consolidation tab
    dash_tabs  = get_tabs(svc, DASHBOARD_ID)
    cons_tab   = find_tab(dash_tabs, ['manpower consolidat', 'daily manpower', 'manpower report'])
    today_data = {}

    if cons_tab:
        cons_rows = read_range(svc, DASHBOARD_ID, f"'{cons_tab}'!A1:Z5")
        print(f"    Manpower consolidation tab: {cons_tab}, rows: {len(cons_rows)}")

    return {
        'present':          0,   # will be filled from actual data
        'required':         TARGETS['manpower_required'],
        'status':           'GREY',
        'departments':      [],
        'note':             'Reading from manpower sheet — column mapping pending',
    }


def get_revenue(svc):
    """Read Revenue and Dispatch data."""
    print("  💰 Reading Revenue...")

    # Read ERP Daily tab from dashboard — has daily sales
    dash_tabs = get_tabs(svc, DASHBOARD_ID)
    erp_tab   = find_tab(dash_tabs, ['erp daily', 'daily ton', 'sales'])

    today      = today_str()
    today_rev  = 0
    today_tons = 0
    mtd_rev    = 0
    mtd_tons   = 0
    records    = []

    if erp_tab:
        rows = read_range(svc, DASHBOARD_ID, f"'{erp_tab}'!A1:F500")
        if rows and len(rows) > 1:
            headers = rows[0]
            print(f"    ERP Daily headers: {headers}")
            for row in rows[1:]:
                if len(row) < 3:
                    continue
                date    = str(row[1]).strip() if len(row) > 1 else ''
                fg_tons = safe_float(row[2]) if len(row) > 2 else 0
                fg_val  = safe_float(row[3]) if len(row) > 3 else 0
                jw_tons = safe_float(row[4]) if len(row) > 4 else 0
                jw_val  = safe_float(row[5]) if len(row) > 5 else 0

                total_tons_row = fg_tons + jw_tons
                total_val_row  = fg_val  + jw_val

                mtd_rev  += total_val_row
                mtd_tons += total_tons_row

                records.append({
                    'date':     date,
                    'fg_tons':  fg_tons,
                    'fg_value': fg_val,
                    'jw_tons':  jw_tons,
                    'jw_value': jw_val,
                })

                if date == today or today in date:
                    today_rev  = total_val_row
                    today_tons = total_tons_row

    # Also read overdue from dashboard
    overdue = 0
    od_tab  = find_tab(get_tabs(svc, DASHBOARD_ID), ['overdue', 'outstanding'])
    if od_tab:
        od_rows = read_range(svc, DASHBOARD_ID, f"'{od_tab}'!A1:B5")
        print(f"    Overdue tab: {od_tab}")

    return {
        'today_revenue':        today_rev,
        'today_tons':           today_tons,
        'mtd_revenue':          mtd_rev,
        'mtd_tons':             round(mtd_tons, 2),
        'daily_target':         TARGETS['daily_revenue_target'],
        'monthly_target':       TARGETS['monthly_revenue_target'],
        'status_today':         status(today_rev, TARGETS['daily_revenue_target']),
        'status_mtd':           status(mtd_rev, TARGETS['monthly_revenue_target']),
        'overdue_cr':           0,
        'last_7_days':          records[-7:] if records else [],
        'last_30_days':         records[-30:] if records else [],
    }


def get_production(svc):
    """Read production data from PMS."""
    print("  🏭 Reading Production...")
    tabs = get_tabs(svc, PMS_ID)
    tab  = find_tab(tabs, ['form response', 'responses', 'data'])

    today_qty  = 0
    today_tons = 0
    mtd_qty    = 0
    vf_data    = {}

    if tab:
        rows = read_range(svc, PMS_ID, f"'{tab}'!A1:Z200")
        if rows and len(rows) > 1:
            headers = [str(h).strip() for h in rows[0]]
            print(f"    PMS headers: {headers[:10]}")

            for row in rows[1:]:
                padded = row + [''] * max(0, len(headers) - len(row))
                r      = dict(zip(headers, padded))

                vf  = r.get('VF No', r.get('Machine', r.get('VF_No', ''))).strip()
                qty = safe_float(r.get('Qty Forged', r.get('Quantity', r.get('Qty', 0))))
                mtd_qty += qty

                if vf:
                    if vf not in vf_data:
                        vf_data[vf] = {'qty': 0, 'rejection': 0}
                    vf_data[vf]['qty']       += qty
                    rej = safe_float(r.get('Rejection', r.get('Rejected Qty', 0)))
                    vf_data[vf]['rejection'] += rej

    # Calculate VF-wise rejection %
    vf_rejection = []
    for vf, d in sorted(vf_data.items())[:20]:
        rej_pct = round(d['rejection']/d['qty']*100, 2) if d['qty'] > 0 else 0
        vf_rejection.append({
            'vf_no':       vf,
            'qty':         d['qty'],
            'rejection':   d['rejection'],
            'rejection_pct': rej_pct,
            'status':      status(rej_pct, TARGETS['rejection_max_pct'], higher_is_better=False),
        })

    return {
        'today_qty':     today_qty,
        'today_tons':    today_tons,
        'mtd_qty':       int(mtd_qty),
        'vf_rejection':  vf_rejection,
        'rejection_target': TARGETS['rejection_max_pct'],
    }


def get_electricity(svc):
    """Read electricity consumption."""
    print("  ⚡ Reading Electricity...")
    tabs = get_tabs(svc, ELECTRICITY_ID)
    print(f"    Electricity tabs: {tabs}")

    tab = find_tab(tabs, ['electricity', 'consumption', 'form response', 'data'])

    today_units = 0
    mtd_units   = 0
    records     = []

    if tab:
        rows = read_range(svc, ELECTRICITY_ID, f"'{tab}'!A1:Z100")
        if rows and len(rows) > 1:
            headers = [str(h).strip() for h in rows[0]]
            print(f"    Electricity headers: {headers[:8]}")
            for row in rows[1:]:
                padded = row + [''] * max(0, len(headers)-len(row))
                r      = dict(zip(headers, padded))
                units  = safe_float(r.get('Units', r.get('Electricity Units',
                         r.get('Consumption', r.get('KWH', 0)))))
                mtd_units += units
                records.append({'units': units})

    return {
        'today_units':   today_units,
        'mtd_units':     round(mtd_units, 0),
        'daily_target':  TARGETS['electricity_daily_units'],
        'status':        'GREY',
        'records':       records[-30:],
    }

# ═══════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════

def main():
    print('=' * 60)
    print('   STAGE 5 — CEO Dashboard Writer')
    print(f'   {datetime.now().strftime("%Y-%m-%d %H:%M IST")}')
    print('=' * 60)

    creds = authenticate()
    svc   = build('sheets', 'v4', credentials=creds)

    # Read all metrics
    rm_stock   = get_rm_stock(svc)
    revenue    = get_revenue(svc)
    production = get_production(svc)
    electricity= get_electricity(svc)
    manpower   = get_manpower(svc)

    # Build CEO data payload
    ceo_data = {
        'generated_at':  datetime.now().isoformat(),
        'generated_ist': datetime.now().strftime('%d %b %Y %H:%M IST'),
        'targets':       TARGETS,
        'metrics': {
            'rm_stock':    rm_stock,
            'revenue':     revenue,
            'production':  production,
            'electricity': electricity,
            'manpower':    manpower,
            'oil': {
                'stock_ltrs':   0,
                'min_target':   TARGETS['oil_min_stock_ltrs'],
                'status':       'GREY',
                'note':         'Oil stock from dashboard Oil Stock tab',
            },
        },
        'alerts': [],
    }

    # Generate alerts for RED metrics
    if rm_stock['status'] == 'RED':
        ceo_data['alerts'].append({
            'type': 'WARNING',
            'metric': 'RM Stock',
            'message': f"RM Stock {rm_stock['total_tons']} tons below minimum {TARGETS['rm_stock_min_tons']} tons",
        })
    if revenue['status_today'] == 'RED':
        ceo_data['alerts'].append({
            'type': 'WARNING',
            'metric': 'Revenue',
            'message': f"Today's revenue below daily target of ₹{TARGETS['daily_revenue_target']:,}",
        })

    # Flag high rejection VF machines
    for vf in production.get('vf_rejection', []):
        if vf['status'] == 'RED':
            ceo_data['alerts'].append({
                'type': 'CRITICAL',
                'metric': 'Rejection',
                'message': f"{vf['vf_no']} rejection {vf['rejection_pct']}% — above {TARGETS['rejection_max_pct']}% limit",
            })

    # Save to docs/ folder for GitHub Pages
    os.makedirs('docs', exist_ok=True)
    with open(OUTPUT_FILE, 'w') as f:
        json.dump(ceo_data, f, indent=2)

    print(f"\n✅ CEO data written → {OUTPUT_FILE}")
    print(f"   RM Stock:    {rm_stock['total_tons']} tons ({rm_stock['status']})")
    print(f"   MTD Revenue: ₹{revenue['mtd_revenue']:,.0f} ({revenue['status_mtd']})")
    print(f"   Production:  {production['mtd_qty']:,} pcs MTD")
    print(f"   Alerts:      {len(ceo_data['alerts'])}")

    if ceo_data['alerts']:
        print("\n⚠ ACTIVE ALERTS:")
        for a in ceo_data['alerts']:
            print(f"   [{a['type']}] {a['metric']}: {a['message']}")


if __name__ == '__main__':
    main()
