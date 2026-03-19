"""
STAGE 4 — MASTER BUILDER
Reads all confirmed sheets via Google Sheets API.
Auto-detects what data is in each sheet from content.
Builds 5 master tables:
  MASTER_PARTS       — central part master (VF No as key)
  SALES_MASTER       — daily dispatch/sales data
  PRODUCTION_MASTER  — daily production per VF machine
  COST_MASTER        — RM cost + electricity + manpower per day
  MACHINE_MASTER     — machine utilisation per VF per day

Then calculates:
  Part-wise cost per kg
  Customer-wise profitability
  Machine utilisation %
  Scrap/rejection rate per part

Run: python master_builder.py
"""

import json, os, pickle, time, csv
from datetime import datetime, timezone
from collections import defaultdict

import gspread
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ═══════════════════════════════════════════════════════
#  CONFIG — all confirmed sheet IDs
# ═══════════════════════════════════════════════════════

TOKEN_FILE = 'token.pickle'
DELAY      = 0.5
BATCH_SIZE = 500
TAB_INIT   = 1

# Output sheet — your main dashboard
OUTPUT_SHEET_ID = '1NR8EPGRJN0AQDXZjYw5k93clsO8AD4u2l2Xke1lBC2I'

# All source sheets — names don't matter, content auto-detected
SOURCE_SHEETS = {
    'dashboard':    '1NR8EPGRJN0AQDXZjYw5k93clsO8AD4u2l2Xke1lBC2I',
    'pms_forgings': '1c-axqiBEufNb1vK-JdJP6t1eRTmzliA3otAbqxAOXZE',
    'machine_shop': '1yC-b36rgAxablmdXhngHCEOsmnlgWourep6ctKifXCA',
    'dispatch':     '1txZM9a9_OSG-ZWYaAEBLKj-9M7LYLrFyl0kJkPsTMGI',
    'manpower':     '1t7UjWTP_cpIJ2BjoaMlV6uUKA7ztH9UnKc_korYKCiw',
    'electricity':  '1nUvf-UWjBSbSWnZTNph-gRUbjzuguGlidpYBKshKUNQ',
    '57f4_inward':  '13hnkD-VCdu54yjMClMVkabR5iCZIzWo1goXboQMQJLE',
    'vendor_rej':   '1izr5HWo4-qvXbmhkbmBIVQSR-X6ptU7OtSX-ICd43H0',
    'dropouts':     '1J1KKfygB8M1z3qZJQcaMXjkVb7HgxDARZrklMld8gS0',
    'rm_inward':    '16SNs4WQgM--iwLg2wAQLhpmfzpTAc9K-VRSbIftOiPQ',
    'oil_inward':   '1jvJUviVgjPWkTO8DYkprTGAuazf6TMDoOT2GwUBdfNo',
    'crm':          '1wZlHu1iEzS1yMcGm2WDJcX9nz4uAHaHaPxNP8o8_HXE',
}

FMT = {
    'blue':   {'backgroundColor':{'red':0.1,'green':0.2,'blue':0.4},
               'textFormat':{'bold':True,'foregroundColor':{'red':1,'green':1,'blue':1}}},
    'green':  {'backgroundColor':{'red':0.0,'green':0.5,'blue':0.2},
               'textFormat':{'bold':True,'foregroundColor':{'red':1,'green':1,'blue':1}}},
    'purple': {'backgroundColor':{'red':0.4,'green':0.1,'blue':0.5},
               'textFormat':{'bold':True,'foregroundColor':{'red':1,'green':1,'blue':1}}},
    'orange': {'backgroundColor':{'red':0.8,'green':0.4,'blue':0.0},
               'textFormat':{'bold':True,'foregroundColor':{'red':1,'green':1,'blue':1}}},
    'teal':   {'backgroundColor':{'red':0.0,'green':0.4,'blue':0.4},
               'textFormat':{'bold':True,'foregroundColor':{'red':1,'green':1,'blue':1}}},
    'red':    {'backgroundColor':{'red':0.6,'green':0.1,'blue':0.1},
               'textFormat':{'bold':True,'foregroundColor':{'red':1,'green':1,'blue':1}}},
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
#  SHEET READER — reads ALL rows from a tab
# ═══════════════════════════════════════════════════════

def read_all_rows(svc, sheet_id, tab_name, max_rows=50000):
    """Read all rows from a tab. Returns (headers, rows)."""
    try:
        resp = svc.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range=f"'{tab_name}'!A1:Z{max_rows}"
        ).execute()
        time.sleep(DELAY)
        values = resp.get('values', [])
        if not values:
            return [], []
        headers = [str(h).strip() for h in values[0]]
        rows    = []
        for row in values[1:]:
            # Pad row to header length
            while len(row) < len(headers):
                row.append('')
            r = {headers[i]: str(row[i]).strip()
                 for i in range(len(headers)) if headers[i]}
            if any(v for v in r.values()):
                rows.append(r)
        return headers, rows
    except HttpError as e:
        if e.resp.status == 429:
            print(f"  ⏳ Quota — waiting 60s")
            time.sleep(60)
            return read_all_rows(svc, sheet_id, tab_name, max_rows)
        print(f"  ⚠ Cannot read {sheet_id}/{tab_name}: {e}")
        return [], []
    except Exception as e:
        print(f"  ⚠ Error reading {sheet_id}/{tab_name}: {e}")
        return [], []


def get_tabs(svc, sheet_id):
    """Get all tab names from a sheet."""
    try:
        meta = svc.spreadsheets().get(
            spreadsheetId=sheet_id,
            fields='sheets.properties.title'
        ).execute()
        time.sleep(DELAY)
        return [s['properties']['title'] for s in meta.get('sheets', [])]
    except Exception as e:
        print(f"  ⚠ Cannot get tabs for {sheet_id}: {e}")
        return []


def read_sheet_all_tabs(svc, sheet_id, label=''):
    """Read all tabs from a sheet. Returns dict of tab_name: (headers, rows)."""
    tabs = get_tabs(svc, sheet_id)
    result = {}
    for tab in tabs:
        print(f"  📋 Reading {label}/{tab}...")
        headers, rows = read_all_rows(svc, sheet_id, tab)
        if rows:
            result[tab] = (headers, rows)
            print(f"     → {len(rows):,} rows, cols: {headers[:6]}")
    return result

# ═══════════════════════════════════════════════════════
#  AUTO-DETECTOR — identifies what data is in a sheet
# ═══════════════════════════════════════════════════════

def detect_content(headers, sample_rows):
    """Auto-detect what type of data this sheet contains."""
    h = ' '.join(headers).lower()
    r = str(sample_rows[:3]).lower() if sample_rows else ''
    combined = h + ' ' + r

    if any(k in combined for k in ['vf no','vf_no','vf704','vf855','forging','press shop','forge shop']):
        return 'PRODUCTION'
    if any(k in combined for k in ['dispatch','customer','invoice','challan','delivery','sales fg']):
        return 'DISPATCH'
    if any(k in combined for k in ['manpower','present','absent','contractor','headcount','attendance']):
        return 'MANPOWER'
    if any(k in combined for k in ['electricity','kwh','unit','eb reading','power','consumption unit']):
        return 'ELECTRICITY'
    if any(k in combined for k in ['rm stock','grade','inward','steel','opening stock','consumption ton']):
        return 'RM_STOCK'
    if any(k in combined for k in ['supplier','vendor','rate','quantity in kg','rm inward','tmt','grade']):
        return 'RM_INWARD'
    if any(k in combined for k in ['oil','ltr','litre','furnace oil','lub oil','hydraulic']):
        return 'OIL'
    if any(k in combined for k in ['rejection','dropout','scrap','defect','ppm','ok qty']):
        return 'REJECTION'
    if any(k in combined for k in ['57f4','job work','jwk','outward','inward material']):
        return 'JOBWORK'
    if any(k in combined for k in ['inquiry','enquiry','crm','quotation','pipeline','stage','forecast']):
        return 'CRM'
    if any(k in combined for k in ['schedule','tons','unit price','schedule turnover']):
        return 'SCHEDULE'
    if any(k in combined for k in ['revenue','turnover','sales value','month','total dispatch']):
        return 'REVENUE'
    return 'UNKNOWN'

# ═══════════════════════════════════════════════════════
#  DATA NORMALISER — cleans values
# ═══════════════════════════════════════════════════════

def clean_num(val):
    """Extract numeric value from string like '₹1,23,456' or '1.5 Ton'."""
    if not val: return 0.0
    import re
    val = str(val).replace('₹','').replace(',','').replace(' ','')
    match = re.search(r'-?\d+\.?\d*', val)
    return float(match.group()) if match else 0.0


def clean_date(val):
    """Normalise various date formats to YYYY-MM-DD."""
    if not val: return ''
    val = str(val).strip()
    for fmt in ['%d/%m/%Y','%m/%d/%Y','%Y-%m-%d','%d-%m-%Y',
                '%d/%m/%Y %H:%M:%S','%m/%d/%Y %H:%M:%S',
                '%d-%m-%Y %H:%M:%S','%Y-%m-%dT%H:%M:%S']:
        try:
            return datetime.strptime(val[:19], fmt).strftime('%Y-%m-%d')
        except:
            continue
    # Try timestamp
    try:
        parts = val.split('/')
        if len(parts)==3:
            d,m,y = parts[0], parts[1], parts[2][:4]
            return f"{y}-{int(m):02d}-{int(d):02d}"
    except:
        pass
    return val[:10]

# ═══════════════════════════════════════════════════════
#  MASTER BUILDERS
# ═══════════════════════════════════════════════════════

def build_master_parts(svc, dashboard_id):
    """Build MASTER_PARTS from dashboard Master Part & Weight Data tab."""
    print('\n🔨 Building MASTER_PARTS...')
    tabs = get_tabs(svc, dashboard_id)

    # Find the part master tab
    part_tab = None
    for t in tabs:
        if any(k in t.lower() for k in ['master part','part & weight','part weight','master']):
            part_tab = t
            break

    if not part_tab:
        print('  ⚠ Part master tab not found in dashboard')
        return []

    headers, rows = read_all_rows(svc, dashboard_id, part_tab)
    print(f'  Found {len(rows):,} parts in {part_tab}')
    print(f'  Columns: {headers[:10]}')

    master = []
    for r in rows:
        # Auto-map columns by content
        vf_no = (r.get('VF No') or r.get('VF_No') or r.get('VF NO') or
                 r.get('Part No') or r.get('Part Number') or '')
        if not vf_no: continue

        master.append({
            'VF_No':          vf_no.strip(),
            'Input_Weight':   clean_num(r.get('Input Weight') or r.get('Input Wt') or ''),
            'Furnace_Weight': clean_num(r.get('Oil Furnace Weight') or r.get('Furnace Wt') or ''),
            'Finish_Weight':  clean_num(r.get('Finish Weight') or r.get('Finish Wt') or ''),
            'Grade':          r.get('VF Used Grade') or r.get('Grade') or r.get('Material') or '',
            'Unit_Price':     clean_num(r.get('Unit Price') or r.get('Price') or r.get('Rate') or ''),
            'IBH_Weight':     clean_num(r.get('IBH Weight') or ''),
            'Customer':       r.get('Customer Name') or r.get('Customer') or '',
            'Machine':        r.get('Ideal Unit') or r.get('Machine') or '',
        })

    print(f'  ✅ {len(master):,} parts built')
    return master


def build_sales_master(svc, dispatch_id):
    """Build SALES_MASTER from dispatch sheet."""
    print('\n🔨 Building SALES_MASTER...')
    data = read_sheet_all_tabs(svc, dispatch_id, 'Dispatch')

    sales = []
    for tab, (headers, rows) in data.items():
        content = detect_content(headers, rows[:3])
        if content not in ('DISPATCH','REVENUE','UNKNOWN'):
            continue
        print(f'  Processing tab: {tab} ({len(rows):,} rows, type={content})')

        for r in rows:
            # Find date
            date = clean_date(
                r.get('Date') or r.get('DATE') or r.get('Timestamp') or
                r.get('Dispatch Date') or r.get('Invoice Date') or ''
            )
            if not date: continue

            # Find part/VF number
            vf_no = (r.get('VF No') or r.get('VF_No') or r.get('Part No') or
                     r.get('Part Number') or r.get('Item') or '')

            # Find customer
            customer = (r.get('Customer') or r.get('Customer Name') or
                       r.get('Party') or r.get('Buyer') or '')

            # Find quantities and values
            qty  = clean_num(r.get('Quantity') or r.get('Qty') or r.get('Nos') or '')
            tons = clean_num(r.get('Tons') or r.get('Weight') or r.get('Net Wt') or '')
            value= clean_num(r.get('Value') or r.get('Amount') or r.get('Sales Value') or
                            r.get('Sales FG Value') or '')
            rate = clean_num(r.get('Rate') or r.get('Price') or r.get('Unit Price') or '')

            if not (qty or tons or value): continue

            sales.append({
                'Date':      date,
                'VF_No':     vf_no,
                'Customer':  customer,
                'Qty':       qty,
                'Tons':      tons,
                'Rate':      rate,
                'Value':     value,
                'Source_Tab':tab,
            })

    print(f'  ✅ {len(sales):,} sales records built')
    return sales


def build_production_master(svc, pms_id, machine_shop_id):
    """Build PRODUCTION_MASTER from PMS Forgings + Machine Shop sheets."""
    print('\n🔨 Building PRODUCTION_MASTER...')
    production = []

    for label, sheet_id in [('PMS Forgings', pms_id), ('Machine Shop', machine_shop_id)]:
        data = read_sheet_all_tabs(svc, sheet_id, label)
        for tab, (headers, rows) in data.items():
            content = detect_content(headers, rows[:3])
            print(f'  {label}/{tab}: {len(rows):,} rows, type={content}')
            if content not in ('PRODUCTION','UNKNOWN'):
                continue

            for r in rows:
                date = clean_date(
                    r.get('Date') or r.get('DATE') or r.get('Timestamp') or
                    r.get('Production Date') or ''
                )
                if not date: continue

                vf_no = (r.get('VF No') or r.get('Machine') or r.get('VF') or
                         r.get('Part No') or '')
                shift = (r.get('Shift') or r.get('SHIFT') or '')
                operator = (r.get('Operator') or r.get('Name') or r.get('Employee') or '')
                qty_forged = clean_num(
                    r.get('Qty Forged') or r.get('Quantity') or r.get('Production') or
                    r.get('OK Qty') or r.get('Total Forged') or ''
                )
                scrap = clean_num(
                    r.get('Scrap') or r.get('Rejection') or r.get('Dropout') or ''
                )
                tons = clean_num(
                    r.get('Tons') or r.get('Production Tons') or r.get('Weight') or ''
                )

                if not (qty_forged or tons): continue

                production.append({
                    'Date':       date,
                    'VF_No':      vf_no,
                    'Shift':      shift,
                    'Operator':   operator,
                    'Qty_Forged': qty_forged,
                    'Scrap':      scrap,
                    'Tons':       tons,
                    'Source':     label,
                    'Tab':        tab,
                })

    print(f'  ✅ {len(production):,} production records built')
    return production


def build_cost_master(svc, rm_inward_id, electricity_id, manpower_id, oil_id, dashboard_id):
    """Build COST_MASTER — RM + electricity + manpower cost per day."""
    print('\n🔨 Building COST_MASTER...')
    cost_by_date = defaultdict(lambda: {
        'RM_Inward_Tons': 0, 'RM_Cost': 0,
        'Electricity_Units': 0, 'Electricity_Cost': 0,
        'Oil_Litres': 0, 'Oil_Cost': 0,
        'Manpower_Present': 0, 'Manpower_Cost': 0,
    })

    # RM Inward
    print('  Reading RM Inward...')
    for sheet_id in [rm_inward_id, dashboard_id]:
        tabs = get_tabs(svc, sheet_id)
        for tab in tabs:
            headers, rows = read_all_rows(svc, sheet_id, tab)
            content = detect_content(headers, rows[:3])
            if content != 'RM_INWARD': continue
            print(f'  Found RM Inward in {sheet_id}/{tab}: {len(rows):,} rows')
            for r in rows:
                date = clean_date(r.get('Date') or r.get('DATE') or '')
                if not date: continue
                qty  = clean_num(r.get('Quantity in KG') or r.get('Qty') or r.get('Quantity') or '')
                rate = clean_num(r.get('Rate') or r.get('Price') or '')
                val  = clean_num(r.get('Total Rate') or r.get('Total') or r.get('Value') or '')
                if not val and qty and rate: val = qty * rate / 1000
                cost_by_date[date]['RM_Inward_Tons'] += qty / 1000
                cost_by_date[date]['RM_Cost']        += val
            break

    # Electricity
    print('  Reading Electricity...')
    tabs = get_tabs(svc, electricity_id)
    for tab in tabs:
        headers, rows = read_all_rows(svc, electricity_id, tab)
        content = detect_content(headers, rows[:3])
        if content != 'ELECTRICITY': continue
        print(f'  Found Electricity in {tab}: {len(rows):,} rows')
        for r in rows:
            date  = clean_date(r.get('Date') or r.get('DATE') or '')
            units = clean_num(r.get('Units') or r.get('Consumption') or
                             r.get('KWH') or r.get('Unit') or '')
            cost  = clean_num(r.get('Cost') or r.get('Amount') or r.get('Value') or '')
            if date:
                cost_by_date[date]['Electricity_Units'] += units
                cost_by_date[date]['Electricity_Cost']  += cost
        break

    # Manpower
    print('  Reading Manpower...')
    tabs = get_tabs(svc, manpower_id)
    for tab in tabs:
        headers, rows = read_all_rows(svc, manpower_id, tab)
        content = detect_content(headers, rows[:3])
        if content != 'MANPOWER': continue
        print(f'  Found Manpower in {tab}: {len(rows):,} rows')
        for r in rows:
            date    = clean_date(r.get('Date') or r.get('DATE') or r.get('Timestamp') or '')
            present = clean_num(r.get('Present') or r.get('Attendance') or
                               r.get('Total Present') or '')
            if date:
                cost_by_date[date]['Manpower_Present'] += present
        break

    # Oil
    print('  Reading Oil Inward...')
    tabs = get_tabs(svc, oil_id)
    for tab in tabs:
        headers, rows = read_all_rows(svc, oil_id, tab)
        content = detect_content(headers, rows[:3])
        if content not in ('OIL', 'RM_INWARD'): continue
        for r in rows:
            date  = clean_date(r.get('Date') or r.get('DATE') or '')
            qty   = clean_num(r.get('Qty') or r.get('Litres') or r.get('Ltr') or
                             r.get('Quantity') or '')
            cost  = clean_num(r.get('Cost') or r.get('Amount') or r.get('Value') or '')
            if date:
                cost_by_date[date]['Oil_Litres'] += qty
                cost_by_date[date]['Oil_Cost']   += cost
        if cost_by_date: break

    cost_rows = []
    for date, vals in sorted(cost_by_date.items()):
        if date:
            row = {'Date': date}
            row.update(vals)
            cost_rows.append(row)

    print(f'  ✅ {len(cost_rows):,} daily cost records built')
    return cost_rows


def build_rejection_master(svc, dropout_id, vendor_rej_id):
    """Build rejection master from Drop Outs + Vendor Rejection sheets."""
    print('\n🔨 Building REJECTION_MASTER...')
    rejections = []

    for label, sheet_id in [('Inhouse', dropout_id), ('Vendor', vendor_rej_id)]:
        tabs = get_tabs(svc, sheet_id)
        for tab in tabs:
            headers, rows = read_all_rows(svc, sheet_id, tab)
            if not rows: continue
            print(f'  {label}/{tab}: {len(rows):,} rows, headers: {headers[:6]}')
            for r in rows:
                date  = clean_date(r.get('Date') or r.get('DATE') or r.get('Timestamp') or '')
                vf_no = r.get('VF No') or r.get('Part No') or r.get('Part') or ''
                qty   = clean_num(r.get('Rejection') or r.get('Qty') or
                                  r.get('Rejected') or r.get('Dropout') or '')
                reason= r.get('Reason') or r.get('Defect') or r.get('Remarks') or ''
                if qty > 0:
                    rejections.append({
                        'Date':   date,
                        'VF_No':  vf_no,
                        'Qty':    qty,
                        'Type':   label,
                        'Reason': reason,
                    })

    print(f'  ✅ {len(rejections):,} rejection records built')
    return rejections

# ═══════════════════════════════════════════════════════
#  PROFITABILITY ENGINE
# ═══════════════════════════════════════════════════════

def calculate_profitability(master_parts, sales, production, cost_master):
    """Calculate part-wise and customer-wise profitability."""
    print('\n📊 Calculating profitability...')

    # Index part master by VF_No
    parts_idx = {p['VF_No']: p for p in master_parts}

    # Sales by VF_No
    sales_by_part = defaultdict(lambda: {'qty':0,'tons':0,'value':0,'customers':set()})
    for s in sales:
        vf = s['VF_No']
        sales_by_part[vf]['qty']      += s['Qty']
        sales_by_part[vf]['tons']     += s['Tons']
        sales_by_part[vf]['value']    += s['Value']
        if s['Customer']:
            sales_by_part[vf]['customers'].add(s['Customer'])

    # Production by VF_No
    prod_by_part = defaultdict(lambda: {'qty':0,'tons':0,'scrap':0})
    for p in production:
        vf = p['VF_No']
        prod_by_part[vf]['qty']   += p['Qty_Forged']
        prod_by_part[vf]['tons']  += p['Tons']
        prod_by_part[vf]['scrap'] += p['Scrap']

    # Total costs
    total_rm_cost    = sum(c['RM_Cost']           for c in cost_master)
    total_elec_cost  = sum(c['Electricity_Cost']  for c in cost_master)
    total_oil_cost   = sum(c['Oil_Cost']          for c in cost_master)
    total_prod_tons  = sum(c['RM_Inward_Tons']    for c in cost_master) or 1

    # Cost per ton (total overhead ÷ total production)
    overhead_per_ton = (total_elec_cost + total_oil_cost) / total_prod_tons

    results = []
    for vf_no, part in parts_idx.items():
        s = sales_by_part.get(vf_no, {})
        p = prod_by_part.get(vf_no,  {})

        finish_wt  = part['Finish_Weight'] or 0
        unit_price = part['Unit_Price']    or 0
        grade      = part['Grade']         or ''

        prod_qty   = p.get('qty',   0)
        prod_tons  = p.get('tons',  0)
        scrap_qty  = p.get('scrap', 0)
        sales_val  = s.get('value', 0)
        sales_qty  = s.get('qty',   0)
        customers  = ', '.join(list(s.get('customers', set()))[:3])

        # RM cost per part (finish weight × grade rate — simplified)
        rm_cost_per_part = finish_wt * 65  # ₹65/kg average — will refine with grade rates
        overhead_per_part= finish_wt * (overhead_per_ton / 1000)
        total_cost_part  = rm_cost_per_part + overhead_per_part
        margin_per_part  = unit_price - total_cost_part if unit_price else 0
        margin_pct       = (margin_per_part / unit_price * 100) if unit_price else 0

        efficiency = ((prod_qty - scrap_qty) / prod_qty * 100) if prod_qty else 0

        results.append({
            'VF_No':              vf_no,
            'Customer':           customers or part['Customer'],
            'Grade':              grade,
            'Finish_Weight_KG':   finish_wt,
            'Machine':            part['Machine'],
            'Unit_Price':         unit_price,
            'RM_Cost_Per_Part':   round(rm_cost_per_part, 2),
            'Overhead_Per_Part':  round(overhead_per_part, 2),
            'Total_Cost_Per_Part':round(total_cost_part, 2),
            'Margin_Per_Part':    round(margin_per_part, 2),
            'Margin_Pct':         round(margin_pct, 1),
            'Qty_Produced':       prod_qty,
            'Qty_Sold':           sales_qty,
            'Sales_Value':        round(sales_val, 2),
            'Scrap_Qty':          scrap_qty,
            'Efficiency_Pct':     round(efficiency, 1),
        })

    # Sort by margin descending
    results.sort(key=lambda x: x['Margin_Per_Part'], reverse=True)
    print(f'  ✅ Profitability calculated for {len(results):,} parts')
    return results


def calculate_customer_profitability(sales, profitability):
    """Roll up profitability by customer."""
    prof_idx = {p['VF_No']: p for p in profitability}
    cust = defaultdict(lambda: {'revenue':0,'cost':0,'tons':0,'parts':set()})

    for s in sales:
        customer = s['Customer']
        if not customer: continue
        vf  = s['VF_No']
        val = s['Value']
        p   = prof_idx.get(vf, {})

        cust[customer]['revenue'] += val
        cust[customer]['tons']    += s['Tons']
        cust[customer]['parts'].add(vf)

        if p.get('Total_Cost_Per_Part') and s['Qty']:
            cust[customer]['cost'] += p['Total_Cost_Per_Part'] * s['Qty']

    result = []
    for customer, data in sorted(cust.items(), key=lambda x: -x[1]['revenue']):
        profit = data['revenue'] - data['cost']
        margin = (profit / data['revenue'] * 100) if data['revenue'] else 0
        result.append({
            'Customer':       customer,
            'Revenue':        round(data['revenue'],  2),
            'Est_Cost':       round(data['cost'],     2),
            'Est_Profit':     round(profit,           2),
            'Margin_Pct':     round(margin,           1),
            'Tons':           round(data['tons'],     2),
            'Parts_Count':    len(data['parts']),
        })

    return result

# ═══════════════════════════════════════════════════════
#  SHEET WRITERS
# ═══════════════════════════════════════════════════════

def get_or_create_tab(spreadsheet, title, cols=20):
    try:
        ws = spreadsheet.worksheet(title)
        print(f"  📋 Reusing: {title}")
        return ws
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=title, rows=TAB_INIT, cols=cols)
        print(f"  ➕ Created: {title}")
        return ws

def safe_clear(ws):
    ws.clear(); time.sleep(0.5)

def write_in_batches(ws, rows):
    total = len(rows)
    if not total: print("    (no rows)"); return
    for i in range(0, total, BATCH_SIZE):
        ws.append_rows(rows[i:i+BATCH_SIZE], value_input_option='USER_ENTERED')
        print(f"    ↳ {min(i+BATCH_SIZE,total):,}/{total:,}")
        time.sleep(1)

def write_master(spreadsheet, title, headers, data, fmt_key='blue'):
    print(f'\nWriting {title} ({len(data):,} rows)...')
    ws = get_or_create_tab(spreadsheet, title, cols=len(headers)+2)
    safe_clear(ws)
    ws.append_row(headers)
    col_letter = chr(64 + len(headers))
    ws.format(f'A1:{col_letter}1', FMT[fmt_key])
    rows = [[str(r.get(h,'')) for h in headers] for r in data]
    write_in_batches(ws, rows)

# ═══════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════

def main():
    print('='*60)
    print('   STAGE 4 — MASTER BUILDER')
    print('='*60)

    creds  = authenticate()
    svc    = build('sheets','v4',credentials=creds)
    gc     = gspread.authorize(creds)
    output = gc.open_by_key(OUTPUT_SHEET_ID)

    # ── Build all masters ──────────────────────────────
    master_parts = build_master_parts(svc, SOURCE_SHEETS['dashboard'])

    sales = build_sales_master(svc, SOURCE_SHEETS['dispatch'])

    production = build_production_master(
        svc, SOURCE_SHEETS['pms_forgings'], SOURCE_SHEETS['machine_shop']
    )

    cost_master = build_cost_master(
        svc,
        SOURCE_SHEETS['rm_inward'],
        SOURCE_SHEETS['electricity'],
        SOURCE_SHEETS['manpower'],
        SOURCE_SHEETS['oil_inward'],
        SOURCE_SHEETS['dashboard'],
    )

    rejections = build_rejection_master(
        svc, SOURCE_SHEETS['dropouts'], SOURCE_SHEETS['vendor_rej']
    )

    # ── Profitability ──────────────────────────────────
    profitability    = calculate_profitability(master_parts, sales, production, cost_master)
    customer_profit  = calculate_customer_profitability(sales, profitability)

    # ── Write to Google Sheet ──────────────────────────
    print('\n📊 Writing masters to Google Sheet...')

    if master_parts:
        write_master(output, '⚙️ Master Parts',
            ['VF_No','Customer','Grade','Finish_Weight_KG','Input_Weight',
             'Furnace_Weight','IBH_Weight','Unit_Price','Machine'],
            master_parts, 'blue')

    if sales:
        write_master(output, '💰 Sales Master',
            ['Date','VF_No','Customer','Qty','Tons','Rate','Value','Source_Tab'],
            sales, 'green')

    if production:
        write_master(output, '🏭 Production Master',
            ['Date','VF_No','Shift','Operator','Qty_Forged','Scrap','Tons','Source'],
            production, 'teal')

    if cost_master:
        write_master(output, '💸 Cost Master',
            ['Date','RM_Inward_Tons','RM_Cost','Electricity_Units',
             'Electricity_Cost','Oil_Litres','Oil_Cost','Manpower_Present'],
            cost_master, 'orange')

    if rejections:
        write_master(output, '⚠️ Rejection Master',
            ['Date','VF_No','Qty','Type','Reason'],
            rejections, 'red')

    if profitability:
        write_master(output, '📈 Part Profitability',
            ['VF_No','Customer','Grade','Machine','Unit_Price',
             'RM_Cost_Per_Part','Overhead_Per_Part','Total_Cost_Per_Part',
             'Margin_Per_Part','Margin_Pct','Qty_Produced','Qty_Sold',
             'Sales_Value','Scrap_Qty','Efficiency_Pct'],
            profitability, 'purple')

    if customer_profit:
        write_master(output, '🏆 Customer Profitability',
            ['Customer','Revenue','Est_Cost','Est_Profit','Margin_Pct',
             'Tons','Parts_Count'],
            customer_profit, 'green')

    # ── Summary ────────────────────────────────────────
    total_revenue = sum(s['Value'] for s in sales)
    total_cost    = sum(c['RM_Cost']+c['Electricity_Cost']+c['Oil_Cost']
                       for c in cost_master)
    top_parts     = profitability[:5]
    top_customers = customer_profit[:5]

    print(f'\n✅ MASTER BUILD COMPLETE')
    print(f'   Parts master:    {len(master_parts):,} parts')
    print(f'   Sales records:   {len(sales):,}')
    print(f'   Production recs: {len(production):,}')
    print(f'   Cost records:    {len(cost_master):,} days')
    print(f'   Rejections:      {len(rejections):,}')
    print(f'\n💰 Total Revenue:  ₹{total_revenue:,.0f}')
    print(f'💸 Total Cost:     ₹{total_cost:,.0f}')
    print(f'\n🏆 Top 5 Parts by Margin:')
    for p in top_parts:
        print(f'   {p["VF_No"]:10} ₹{p["Margin_Per_Part"]:>8,.2f}/part  {p["Margin_Pct"]:>5.1f}%')
    print(f'\n👥 Top 5 Customers by Revenue:')
    for c in top_customers:
        print(f'   {c["Customer"][:30]:30} ₹{c["Revenue"]:>12,.0f}')


if __name__ == '__main__':
    main()
