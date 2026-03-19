"""
material_tracker.py — Two-Stage RM Material Consumption Tracker

Material Flow:
  Stage 1 — CUTTING SHOP
    Bar Stock → Cut blanks per VF No
    RM Consumed = Qty Cut × Cut Weight (Input Weight from Part Master)
    RM Cost     = RM Consumed × Grade Rate (from RM Inward)

  Stage 2 — FORGING SHOP
    Cut blanks → Forged parts
    Tracked in PMS as Qty Forged + OK Qty

  RECONCILIATION per VF No per Month:
    Cut Qty vs Forged Qty → difference = WIP blanks in store
    Forged Qty vs OK Qty  → difference = rejection
    If Cut > Forged        → WIP blanks sitting (not a problem)
    If Cut < Forged        → Data entry issue or opening stock used
    If Forged > OK Qty     → Rejection

Output tabs in MASTER_SHEET:
  ✂️ Cutting vs Forging   — monthly reconciliation per VF No
  📦 RM Consumption       — RM KG + cost per VF No per month
  ⚠️ WIP & Rejection      — variances flagged per VF No per month
  📊 Monthly RM Summary   — total RM, cost, avg cost per piece

Run: python material_tracker.py
"""

import json
import os
import pickle
import time
from datetime import datetime
from collections import defaultdict

import gspread
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ═══════════════════════════════════════════════════════
#  CONFIG
# ═══════════════════════════════════════════════════════

PMS_ID          = '1c-axqiBEufNb1vK-JdJP6t1eRTmzliA3otAbqxAOXZE'
DASHBOARD_ID    = '1NR8EPGRJN0AQDXZjYw5k93clsO8AD4u2l2Xke1lBC2I'
RM_INWARD_ID    = '1OCDff85Tqop2yrtc4tn2kXAKj7TdRwvNmdq21K-5GPE'
MASTER_SHEET_ID = os.environ.get('MASTER_SHEET_ID','10Zjxy3mGKP6G3j7uuak3FTXl90JHEQoC0RDgl4tJJXc')

TOKEN_FILE  = 'token.pickle'
DELAY       = 0.5
BATCH_SIZE  = 500
TAB_INIT    = 1

# Keywords to auto-detect cutting shop entries
CUTTING_KEYWORDS  = ['cut', 'cutting', 'blank', 'cutting shop', 'cut shop',
                     'bar cut', 'billets', 'shearing']
FORGING_KEYWORDS  = ['forg', 'press', 'hammer', 'forge', 'forging shop',
                     'press shop', 'hammer shop']

FMT = {
    'blue':   {'backgroundColor':{'red':0.1,'green':0.2,'blue':0.4},
               'textFormat':{'bold':True,'foregroundColor':{'red':1,'green':1,'blue':1}}},
    'green':  {'backgroundColor':{'red':0.0,'green':0.5,'blue':0.2},
               'textFormat':{'bold':True,'foregroundColor':{'red':1,'green':1,'blue':1}}},
    'red':    {'backgroundColor':{'red':0.6,'green':0.1,'blue':0.1},
               'textFormat':{'bold':True,'foregroundColor':{'red':1,'green':1,'blue':1}}},
    'purple': {'backgroundColor':{'red':0.4,'green':0.1,'blue':0.5},
               'textFormat':{'bold':True,'foregroundColor':{'red':1,'green':1,'blue':1}}},
    'teal':   {'backgroundColor':{'red':0.0,'green':0.4,'blue':0.4},
               'textFormat':{'bold':True,'foregroundColor':{'red':1,'green':1,'blue':1}}},
    'orange': {'backgroundColor':{'red':0.8,'green':0.4,'blue':0.0},
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
#  HELPERS
# ═══════════════════════════════════════════════════════

def get_tabs(svc, sheet_id):
    try:
        meta = svc.spreadsheets().get(
            spreadsheetId=sheet_id,
            fields='sheets.properties.title'
        ).execute()
        time.sleep(DELAY)
        return [s['properties']['title'] for s in meta.get('sheets',[])]
    except Exception as e:
        print(f"  ⚠ Cannot get tabs: {e}")
        return []


def read_tab(svc, sheet_id, tab, max_rows=50000):
    try:
        resp = svc.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range=f"'{tab}'!A1:Z{max_rows}"
        ).execute()
        time.sleep(DELAY)
        rows = resp.get('values',[])
        if not rows or len(rows) < 2:
            return [], []
        headers = [str(h).strip() for h in rows[0]]
        data = []
        for row in rows[1:]:
            padded = row + [''] * max(0, len(headers)-len(row))
            data.append(dict(zip(headers, padded[:len(headers)])))
        return headers, data
    except HttpError as e:
        if e.resp.status == 429:
            time.sleep(60)
            return read_tab(svc, sheet_id, tab, max_rows)
        return [], []


def find_tab(tabs, keywords):
    for kw in keywords:
        for tab in tabs:
            if kw.lower() in tab.lower():
                return tab
    return tabs[0] if tabs else None


def safe_float(v, default=0):
    try:
        return float(str(v).replace(',','').replace('₹','').replace(' ',''))
    except Exception:
        return default


def safe_int(v):
    try:
        return int(float(str(v).replace(',','')))
    except Exception:
        return 0


def parse_month(date_str):
    for fmt in ['%d/%m/%Y','%Y-%m-%d','%m/%d/%Y','%d-%m-%Y',
                '%d/%m/%Y %H:%M:%S']:
        try:
            return datetime.strptime(str(date_str).strip()[:10], fmt).strftime('%Y-%m')
        except Exception:
            continue
    return str(date_str)[:7]


def is_cutting(value):
    """Check if a field value indicates cutting shop."""
    v = str(value).lower().strip()
    return any(kw in v for kw in CUTTING_KEYWORDS)


def is_forging(value):
    """Check if a field value indicates forging shop."""
    v = str(value).lower().strip()
    return any(kw in v for kw in FORGING_KEYWORDS)


def get_or_create_tab(spreadsheet, title, cols=12):
    try:
        ws = spreadsheet.worksheet(title)
        print(f"  📋 Reusing: {title}")
        return ws
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=title, rows=TAB_INIT, cols=cols)
        print(f"  ➕ Created: {title}")
        return ws


def safe_clear(ws):
    ws.clear()
    time.sleep(0.5)


def write_in_batches(ws, rows):
    total = len(rows)
    if not total:
        print("    (no rows)")
        return
    for i in range(0, total, BATCH_SIZE):
        ws.append_rows(rows[i:i+BATCH_SIZE], value_input_option='USER_ENTERED')
        print(f"    ↳ {min(i+BATCH_SIZE,total):,} / {total:,}")
        time.sleep(1)

# ═══════════════════════════════════════════════════════
#  STEP 1 — READ & ANALYSE PMS STRUCTURE
# ═══════════════════════════════════════════════════════

def read_and_analyse_pms(svc):
    """
    Read all PMS data. Auto-detect:
    - Which column separates Cutting vs Forging
    - Column names for VF No, Date, Qty, OK Qty, Shift, Machine
    """
    print("\n📋 Reading & analysing PMS structure...")
    tabs = get_tabs(svc, PMS_ID)
    print(f"  All tabs: {tabs}")

    all_data    = []
    tab_summary = {}

    for tab in tabs:
        headers, data = read_tab(svc, PMS_ID, tab)
        if not data:
            continue

        tab_summary[tab] = {
            'rows':    len(data),
            'headers': headers[:15],
            'sample':  [list(data[0].values())[:10]] if data else [],
        }
        all_data.extend(data)
        print(f"  [{tab}] {len(data):,} rows | Headers: {headers[:12]}")

    if not all_data:
        print("  ⚠ No data found in PMS")
        return [], {}, None

    # ── Auto-detect stage column ──────────────────────────
    # Look for a column whose values contain cutting/forging keywords
    stage_col    = None
    stage_values = set()

    # Sample first 100 rows to find stage column
    sample = all_data[:100]
    if sample:
        first_row = sample[0]
        for col, val in first_row.items():
            # Check all values in this column across sample
            col_values = [str(r.get(col,'')).lower() for r in sample]
            cutting_found = any(any(kw in v for kw in CUTTING_KEYWORDS) for v in col_values)
            forging_found = any(any(kw in v for kw in FORGING_KEYWORDS) for v in col_values)

            if cutting_found or forging_found:
                stage_col    = col
                stage_values = set(str(r.get(col,'')) for r in sample if r.get(col,''))
                break

    print(f"\n  🔍 Auto-detection results:")
    print(f"  Stage column found: '{stage_col}'")
    if stage_col:
        print(f"  Stage values: {stage_values}")
    else:
        print(f"  ⚠ Could not auto-detect cutting vs forging column")
        print(f"  Available columns: {list(all_data[0].keys()) if all_data else []}")
        print(f"  → Will treat all rows as forging data")
        print(f"  → To enable two-stage tracking, confirm which column")
        print(f"    separates 'Cutting' from 'Forging' in PMS")

    return all_data, tab_summary, stage_col


# ═══════════════════════════════════════════════════════
#  STEP 2 — READ PART MASTER
# ═══════════════════════════════════════════════════════

def read_part_master(svc):
    print("\n📋 Reading Part Master (Cut Weights)...")
    tabs = get_tabs(svc, DASHBOARD_ID)
    tab  = find_tab(tabs, ['master part','part & weight','part weight','master'])

    if not tab:
        print("  ⚠ Part master tab not found")
        return {}

    headers, data = read_tab(svc, DASHBOARD_ID, tab)
    print(f"  Tab: '{tab}' | {len(data):,} rows | Cols: {headers[:10]}")

    part_master = {}
    for r in data:
        vf_no = (r.get('VF No','') or r.get('VF NO','')).strip()
        if not vf_no:
            continue

        cut_wt = safe_float(
            r.get('Input Weight','') or
            r.get('Input Wt','') or
            r.get('Cut Weight','') or
            r.get('Input Weight (KG)','') or 0
        )

        grade = (
            r.get('VF Uced Grade','') or
            r.get('Grade','') or
            r.get('Material Grade','')
        ).strip()

        if cut_wt > 0:
            part_master[vf_no] = {
                'cut_weight_kg': cut_wt,
                'grade':         grade,
            }

    print(f"  Parts with cut weight: {len(part_master):,}")
    return part_master

# ═══════════════════════════════════════════════════════
#  STEP 3 — READ RM RATES
# ═══════════════════════════════════════════════════════

def read_rm_rates(svc):
    print("\n📋 Reading RM Grade Rates...")

    dash_tabs = get_tabs(svc, DASHBOARD_ID)
    rm_tab    = find_tab(dash_tabs, ['rm inward','raw material inward','steel inward','inward'])

    data = []
    if rm_tab:
        headers, data = read_tab(svc, DASHBOARD_ID, rm_tab)
        print(f"  Dashboard RM tab: '{rm_tab}' | {len(data):,} rows | Cols: {headers[:8]}")

    if not data:
        tabs = get_tabs(svc, RM_INWARD_ID)
        tab  = find_tab(tabs, ['form response','data','inward'])
        if tab:
            headers, data = read_tab(svc, RM_INWARD_ID, tab)
            print(f"  RM Inward sheet tab: '{tab}' | {len(data):,} rows")

    if not data:
        print("  ⚠ No RM rate data — using default ₹55/KG")
        return {'monthly': {}, 'overall': {}}

    grade_month_rates = defaultdict(lambda: defaultdict(list))
    for r in data:
        date  = (r.get('Date','') or r.get('DATE','')).strip()
        grade = (r.get('Grade','') or r.get('GRADE','')).strip()
        rate  = safe_float(
            r.get('Rate','') or r.get('Rate/KG','') or
            r.get('Price','') or r.get('Per KG','') or 0
        )
        if rate > 1000:   # convert per ton to per KG
            rate = rate / 1000

        month = parse_month(date)
        if grade and rate > 0:
            grade_month_rates[grade][month].append(rate)

    monthly = {}
    overall = {}
    for grade, months in grade_month_rates.items():
        monthly[grade] = {m: round(sum(r)/len(r),2) for m,r in months.items()}
        all_r = [r for rates in months.values() for r in rates]
        overall[grade] = round(sum(all_r)/len(all_r),2) if all_r else 55

    print(f"  Grades with rates: {list(overall.keys())[:8]}")
    return {'monthly': monthly, 'overall': overall}

# ═══════════════════════════════════════════════════════
#  STEP 4 — PROCESS TWO-STAGE DATA
# ═══════════════════════════════════════════════════════

def process_two_stage(pms_data, stage_col, part_master, rm_rates):
    """
    Separate cutting and forging entries.
    Calculate RM consumption from cutting stage.
    Reconcile cutting vs forging quantities.
    """
    print(f"\n⚙️  Processing two-stage material flow...")
    print(f"  Stage column: '{stage_col}'")

    cutting_data = []
    forging_data = []
    unknown_data = []

    for r in pms_data:
        vf_no   = (r.get('VF No','') or r.get('Part No','') or
                   r.get('VF_No','')).strip()
        date    = (r.get('Date','') or r.get('Production Date','')).strip()
        machine = (r.get('Machine','') or r.get('Machine Name','')).strip()
        shift   = (r.get('Shift','') or '').strip()

        # Qty columns — try multiple names
        qty_total = safe_int(
            r.get('Qty','') or r.get('Quantity','') or
            r.get('Total Qty','') or r.get('Qty Forged','') or
            r.get('Qty Cut','') or r.get('Cut Qty','') or 0
        )
        qty_ok = safe_int(
            r.get('OK Qty','') or r.get('Good Qty','') or
            r.get('Accepted Qty','') or qty_total
        )
        rejection = safe_int(
            r.get('Rejection','') or r.get('Rejected Qty','') or
            r.get('Scrap Qty','') or 0
        )

        if not vf_no or qty_total <= 0:
            continue

        month = parse_month(date)
        entry = {
            'vf_no':     vf_no,
            'date':      date,
            'month':     month,
            'machine':   machine,
            'shift':     shift,
            'qty_total': qty_total,
            'qty_ok':    qty_ok,
            'rejection': rejection,
        }

        # Classify as cutting or forging
        if stage_col:
            stage_val = str(r.get(stage_col,'')).strip()
            if is_cutting(stage_val):
                entry['stage'] = 'CUTTING'
                cutting_data.append(entry)
            elif is_forging(stage_val):
                entry['stage'] = 'FORGING'
                forging_data.append(entry)
            else:
                entry['stage'] = stage_val or 'UNKNOWN'
                unknown_data.append(entry)
        else:
            # No stage column — check tab name or machine name
            machine_lower = machine.lower()
            if is_cutting(machine_lower):
                entry['stage'] = 'CUTTING'
                cutting_data.append(entry)
            else:
                entry['stage'] = 'FORGING'
                forging_data.append(entry)

    print(f"  Cutting entries:  {len(cutting_data):,}")
    print(f"  Forging entries:  {len(forging_data):,}")
    if unknown_data:
        print(f"  Unknown entries:  {len(unknown_data):,}")
        unique_unknown = set(e['stage'] for e in unknown_data)
        print(f"  Unknown values:   {unique_unknown}")

    # ── Calculate RM consumption from cutting stage ──────
    consumption = []
    for entry in cutting_data:
        vf_no  = entry['vf_no']
        month  = entry['month']
        qty    = entry['qty_total']

        part   = part_master.get(vf_no, {})
        cut_wt = part.get('cut_weight_kg', 0)
        grade  = part.get('grade', '')

        if not cut_wt:
            continue

        rm_kg = round(qty * cut_wt, 3)

        # Get rate
        rate = 55  # default
        if rm_rates:
            rate = (rm_rates.get('monthly',{}).get(grade,{}).get(month, 0) or
                    rm_rates.get('overall',{}).get(grade, 55))

        rm_cost = round(rm_kg * rate, 2)
        consumption.append({**entry, 'cut_weight_kg': cut_wt, 'grade': grade,
                            'rm_kg': rm_kg, 'rate': rate, 'rm_cost': rm_cost})

    # If no cutting data found — calculate from forging data instead
    if not cutting_data and forging_data:
        print("  ℹ No cutting stage found — calculating RM from forging qty")
        for entry in forging_data:
            vf_no  = entry['vf_no']
            month  = entry['month']
            qty    = entry['qty_total']
            part   = part_master.get(vf_no, {})
            cut_wt = part.get('cut_weight_kg', 0)
            grade  = part.get('grade', '')
            if not cut_wt:
                continue
            rm_kg   = round(qty * cut_wt, 3)
            rate    = 55
            if rm_rates:
                rate = (rm_rates.get('monthly',{}).get(grade,{}).get(month,0) or
                        rm_rates.get('overall',{}).get(grade, 55))
            rm_cost = round(rm_kg * rate, 2)
            consumption.append({**entry, 'cut_weight_kg': cut_wt, 'grade': grade,
                                'rm_kg': rm_kg, 'rate': rate, 'rm_cost': rm_cost})

    return cutting_data, forging_data, consumption

# ═══════════════════════════════════════════════════════
#  STEP 5 — BUILD RECONCILIATION
# ═══════════════════════════════════════════════════════

def build_reconciliation(cutting_data, forging_data):
    """
    Per VF No per Month:
      Cut Qty vs Forged Qty → WIP blanks
      Forged Qty vs OK Qty  → Rejection
    """
    cut_by_part_month = defaultdict(int)
    for e in cutting_data:
        cut_by_part_month[(e['vf_no'], e['month'])] += e['qty_total']

    forge_totals = defaultdict(lambda: {'forged':0,'ok':0,'rejection':0})
    for e in forging_data:
        key = (e['vf_no'], e['month'])
        forge_totals[key]['forged']    += e['qty_total']
        forge_totals[key]['ok']        += e['qty_ok']
        forge_totals[key]['rejection'] += e['rejection']

    # All keys
    all_keys = set(cut_by_part_month.keys()) | set(forge_totals.keys())

    recon = []
    for (vf_no, month) in sorted(all_keys):
        cut_qty    = cut_by_part_month.get((vf_no, month), 0)
        forged_qty = forge_totals[(vf_no, month)]['forged']
        ok_qty     = forge_totals[(vf_no, month)]['ok']
        rejection  = forge_totals[(vf_no, month)]['rejection']

        # WIP = cut but not yet forged
        wip = cut_qty - forged_qty if cut_qty > 0 else 0

        # Rejection % of forged
        rej_pct = round(rejection/forged_qty*100, 2) if forged_qty > 0 else 0

        # Status
        if cut_qty == 0:
            status = '⚙️ Forging only'
        elif wip > 0:
            status = f'📦 {wip} blanks in WIP'
        elif wip < 0:
            status = f'⚠️ {abs(wip)} over-forged'
        else:
            status = '✅ Balanced'

        recon.append({
            'month':      month,
            'vf_no':      vf_no,
            'cut_qty':    cut_qty,
            'forged_qty': forged_qty,
            'ok_qty':     ok_qty,
            'rejection':  rejection,
            'rej_pct':    rej_pct,
            'wip_blanks': wip,
            'status':     status,
        })

    return recon

# ═══════════════════════════════════════════════════════
#  STEP 6 — WRITE TABS
# ═══════════════════════════════════════════════════════

def write_all_tabs(spreadsheet, recon, consumption, tab_summary):

    # ── Tab 1: Cutting vs Forging Reconciliation ─────────
    RECON_HEADERS = ['Month','VF No','Cut Qty','Forged Qty',
                     'OK Qty','Rejection','Rejection %',
                     'WIP Blanks','Status']
    recon_rows = [RECON_HEADERS]
    for r in recon:
        recon_rows.append([
            r['month'], r['vf_no'],
            r['cut_qty'], r['forged_qty'],
            r['ok_qty'], r['rejection'], r['rej_pct'],
            r['wip_blanks'], r['status'],
        ])

    print(f"\n✍ Writing Cutting vs Forging Reconciliation ({len(recon_rows)-1:,} rows)...")
    ws1 = get_or_create_tab(spreadsheet, '✂️ Cut vs Forging', cols=9)
    safe_clear(ws1)
    write_in_batches(ws1, recon_rows)
    ws1.format('A1:I1', FMT['blue'])

    # ── Tab 2: RM Consumption by Part Month ─────────────
    part_month = defaultdict(lambda: {'qty':0,'rm_kg':0,'rm_cost':0,'grade':''})
    for r in consumption:
        key = (r['vf_no'], r['month'])
        part_month[key]['qty']     += r['qty_total']
        part_month[key]['rm_kg']   += r['rm_kg']
        part_month[key]['rm_cost'] += r['rm_cost']
        part_month[key]['grade']    = r['grade']

    RM_HEADERS = ['Month','VF No','Grade','Qty',
                  'Cut Weight (KG)','RM Consumed (KG)',
                  'Rate/KG','RM Cost (₹)','Cost/Piece (₹)']
    rm_rows = [RM_HEADERS]
    for (vf_no, month), d in sorted(part_month.items()):
        cut_wt = round(d['rm_kg']/d['qty'],3) if d['qty'] else 0
        rate   = round(d['rm_cost']/d['rm_kg'],2) if d['rm_kg'] else 0
        cpp    = round(d['rm_cost']/d['qty'],2) if d['qty'] else 0
        rm_rows.append([
            month, vf_no, d['grade'], d['qty'],
            cut_wt, round(d['rm_kg'],2),
            rate, round(d['rm_cost'],2), cpp,
        ])

    print(f"✍ Writing RM Consumption ({len(rm_rows)-1:,} rows)...")
    ws2 = get_or_create_tab(spreadsheet, '📦 RM Consumption', cols=9)
    safe_clear(ws2)
    write_in_batches(ws2, rm_rows)
    ws2.format('A1:I1', FMT['teal'])

    # ── Tab 3: WIP & Rejection flags ────────────────────
    flags = [r for r in recon if r['wip_blanks'] != 0 or r['rejection'] > 0]
    FLAG_HEADERS = ['Month','VF No','WIP Blanks','Rejection','Rejection %','Status']
    flag_rows = [FLAG_HEADERS]
    for r in flags:
        flag_rows.append([
            r['month'], r['vf_no'],
            r['wip_blanks'], r['rejection'],
            r['rej_pct'], r['status'],
        ])

    print(f"✍ Writing WIP & Rejection flags ({len(flag_rows)-1:,} rows)...")
    ws3 = get_or_create_tab(spreadsheet, '⚠️ WIP & Rejection', cols=6)
    safe_clear(ws3)
    write_in_batches(ws3, flag_rows)
    ws3.format('A1:F1', FMT['orange'])

    # ── Tab 4: Monthly RM Summary ────────────────────────
    monthly = defaultdict(lambda: {'qty':0,'rm_kg':0,'rm_cost':0,'parts':set()})
    for r in consumption:
        m = r['month']
        monthly[m]['qty']     += r['qty_total']
        monthly[m]['rm_kg']   += r['rm_kg']
        monthly[m]['rm_cost'] += r['rm_cost']
        monthly[m]['parts'].add(r['vf_no'])

    SUM_HEADERS = ['Month','Total Qty','Unique Parts',
                   'Total RM (KG)','Total RM (Tons)',
                   'Total RM Cost (₹)','Avg Cost/KG','Avg Cost/Piece (₹)']
    sum_rows = [SUM_HEADERS]
    for month, d in sorted(monthly.items()):
        avg_kg  = round(d['rm_cost']/d['rm_kg'],2) if d['rm_kg'] else 0
        avg_pce = round(d['rm_cost']/d['qty'],2)   if d['qty']   else 0
        sum_rows.append([
            month, d['qty'], len(d['parts']),
            round(d['rm_kg'],2), round(d['rm_kg']/1000,3),
            round(d['rm_cost'],2), avg_kg, avg_pce,
        ])

    print(f"✍ Writing Monthly RM Summary ({len(sum_rows)-1:,} rows)...")
    ws4 = get_or_create_tab(spreadsheet, '📊 Monthly RM Summary', cols=8)
    safe_clear(ws4)
    write_in_batches(ws4, sum_rows)
    ws4.format('A1:H1', FMT['purple'])

    # ── Tab 5: PMS Structure Report (first run diagnostic) ─
    STRUCT_HEADERS = ['Tab Name','Rows','Columns Found']
    struct_rows = [STRUCT_HEADERS]
    for tab, info in tab_summary.items():
        struct_rows.append([tab, info['rows'], ', '.join(info['headers'])])

    ws5 = get_or_create_tab(spreadsheet, '🔍 PMS Structure', cols=3)
    safe_clear(ws5)
    write_in_batches(ws5, struct_rows)
    ws5.format('A1:C1', FMT['blue'])

    print(f"\n✅ Material Tracker complete")
    wip_count = sum(1 for r in recon if r['wip_blanks'] > 0)
    rej_count = sum(1 for r in recon if r['rejection'] > 0)
    print(f"   Reconciliation rows: {len(recon):,}")
    print(f"   WIP flags:           {wip_count:,}")
    print(f"   Rejection flags:     {rej_count:,}")
    print(f"   RM records:          {len(consumption):,}")

# ═══════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════

def main():
    print('=' * 60)
    print('   MATERIAL CONSUMPTION TRACKER (Two-Stage)')
    print(f'   {datetime.now().strftime("%Y-%m-%d %H:%M")}')
    print('=' * 60)

    creds = authenticate()
    svc   = build('sheets','v4',credentials=creds)
    gc    = gspread.authorize(creds)

    try:
        spreadsheet = gc.open_by_key(MASTER_SHEET_ID)
        print(f"Writing to: {spreadsheet.title}")
    except Exception as e:
        raise SystemExit(f"❌ Cannot open master sheet: {e}")

    pms_data, tab_summary, stage_col = read_and_analyse_pms(svc)
    if not pms_data:
        raise SystemExit("❌ No PMS data found")

    part_master  = read_part_master(svc)
    rm_rates     = read_rm_rates(svc)

    cutting_data, forging_data, consumption = process_two_stage(
        pms_data, stage_col, part_master, rm_rates
    )

    recon = build_reconciliation(cutting_data, forging_data)
    write_all_tabs(spreadsheet, recon, consumption, tab_summary)


if __name__ == '__main__':
    main()
