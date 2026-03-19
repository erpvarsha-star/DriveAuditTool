"""
profitability_engine.py — Part-wise Profitability Calculator

Calculates for each VF No per month:
  Revenue        = Qty Dispatched × Selling Price (Unit Price from Part Master)
  RM Cost        = Qty Cut × Cut Weight × Grade Rate (from material_tracker logic)
  Contribution   = Revenue - RM Cost
  Margin %       = Contribution / Revenue × 100

Also calculates:
  Customer-wise profitability
  Machine-wise profitability
  Monthly trend

Sources:
  - MASTER_SHEET: RM by Part tab (from material_tracker)
  - Dashboard:    ERP Daily tab (sales/dispatch)
  - Dashboard:    VF Schedule tab (selling prices)
  - Dashboard:    Master Part & Weight Data (cut weight, grade, customer)

Output tabs in MASTER_SHEET:
  💰 Part Profitability    — per VF No per month
  👥 Customer Profitability — per customer per month
  ⚙️ Machine Profitability  — per machine per month
  📈 Monthly P&L Summary   — overall monthly profit/loss

Run: python profitability_engine.py
"""

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

DASHBOARD_ID    = '1NR8EPGRJN0AQDXZjYw5k93clsO8AD4u2l2Xke1lBC2I'
DISPATCH_ID     = '1txZM9a9_OSG-ZWYaAEBLKj-9M7LYLrFyl0kJkPsTMGI'
MASTER_SHEET_ID = os.environ.get('MASTER_SHEET_ID','10Zjxy3mGKP6G3j7uuak3FTXl90JHEQoC0RDgl4tJJXc')

TOKEN_FILE  = 'token.pickle'
DELAY       = 0.5
BATCH_SIZE  = 500
TAB_INIT    = 1

FMT = {
    'blue':   {'backgroundColor':{'red':0.1,'green':0.2,'blue':0.4},
               'textFormat':{'bold':True,'foregroundColor':{'red':1,'green':1,'blue':1}}},
    'green':  {'backgroundColor':{'red':0.0,'green':0.5,'blue':0.2},
               'textFormat':{'bold':True,'foregroundColor':{'red':1,'green':1,'blue':1}}},
    'red':    {'backgroundColor':{'red':0.6,'green':0.1,'blue':0.1},
               'textFormat':{'bold':True,'foregroundColor':{'red':1,'green':1,'blue':1}}},
    'purple': {'backgroundColor':{'red':0.4,'green':0.1,'blue':0.5},
               'textFormat':{'bold':True,'foregroundColor':{'red':1,'green':1,'blue':1}}},
    'gold':   {'backgroundColor':{'red':0.8,'green':0.6,'blue':0.0},
               'textFormat':{'bold':True,'foregroundColor':{'red':1,'green':1,'blue':1}}},
    'teal':   {'backgroundColor':{'red':0.0,'green':0.4,'blue':0.4},
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
    except Exception:
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
            return datetime.strptime(str(date_str).strip()[:10],
                                     fmt).strftime('%Y-%m')
        except Exception:
            continue
    return str(date_str)[:7]


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
#  STEP 1 — READ PART MASTER (prices + customers)
# ═══════════════════════════════════════════════════════

def read_part_master(svc):
    print("\n📋 Reading Part Master...")
    tabs = get_tabs(svc, DASHBOARD_ID)
    tab  = find_tab(tabs, ['master part','part & weight','part weight'])

    if not tab:
        print("  ⚠ Part master not found")
        return {}

    headers, data = read_tab(svc, DASHBOARD_ID, tab)
    print(f"  Tab: '{tab}' | {len(data):,} rows | Cols: {headers[:10]}")

    parts = {}
    for r in data:
        vf_no = (r.get('VF No','') or r.get('VF NO','')).strip()
        if not vf_no:
            continue

        cut_wt    = safe_float(r.get('Input Weight','') or
                               r.get('Cut Weight','') or 0)
        finish_wt = safe_float(r.get('Finish Weight','') or 0)
        grade     = (r.get('VF Uced Grade','') or r.get('Grade','')).strip()
        price     = safe_float(r.get('Unit Price','') or
                               r.get('Price','') or 0)
        customer  = (r.get('Customer Name','') or
                     r.get('Customer','')).strip()
        machine   = (r.get('Ideal Unit','') or
                     r.get('Machine','')).strip()

        parts[vf_no] = {
            'cut_weight':  cut_wt,
            'finish_weight': finish_wt,
            'grade':       grade,
            'unit_price':  price,
            'customer':    customer,
            'machine':     machine,
        }

    print(f"  Parts loaded: {len(parts):,}")
    with_price    = sum(1 for p in parts.values() if p['unit_price'] > 0)
    with_customer = sum(1 for p in parts.values() if p['customer'])
    print(f"  With price:    {with_price:,}")
    print(f"  With customer: {with_customer:,}")
    return parts

# ═══════════════════════════════════════════════════════
#  STEP 2 — READ DISPATCH DATA (revenue)
# ═══════════════════════════════════════════════════════

def read_dispatch(svc):
    print("\n💰 Reading Dispatch/Sales data...")

    # Try ERP Daily tab in dashboard first — has daily sales totals
    dash_tabs = get_tabs(svc, DASHBOARD_ID)
    erp_tab   = find_tab(dash_tabs, ['erp daily','daily ton','erp'])

    sales = []

    if erp_tab:
        headers, data = read_tab(svc, DASHBOARD_ID, erp_tab)
        print(f"  ERP Daily tab: '{erp_tab}' | {len(data):,} rows | Cols: {headers[:6]}")
        for r in data:
            date    = (r.get('Date','') or '').strip()
            fg_tons = safe_float(r.get('Sales FG Ton','') or 0)
            fg_val  = safe_float(r.get('Sales FG Value','') or 0)
            jw_tons = safe_float(r.get('Sales Jobwork Ton','') or 0)
            jw_val  = safe_float(r.get('Sales Jobwork Value','') or 0)
            if not date:
                continue
            month = parse_month(date)
            sales.append({
                'date':     date,
                'month':    month,
                'vf_no':    '',
                'customer': '',
                'qty':      0,
                'fg_tons':  fg_tons,
                'fg_value': fg_val,
                'jw_tons':  jw_tons,
                'jw_value': jw_val,
                'total_value': fg_val + jw_val,
                'total_tons':  fg_tons + jw_tons,
            })

    # Also try actual dispatch sheet for part-wise data
    disp_tabs = get_tabs(svc, DISPATCH_ID)
    disp_tab  = find_tab(disp_tabs, ['form response','dispatch','data','sales'])
    if disp_tab:
        headers, data = read_tab(svc, DISPATCH_ID, disp_tab)
        print(f"  Dispatch sheet: '{disp_tab}' | {len(data):,} rows | Cols: {headers[:10]}")
        for r in data:
            date     = (r.get('Date','') or r.get('DATE','')).strip()
            vf_no    = (r.get('VF No','') or r.get('Part No','')).strip()
            customer = (r.get('Customer','') or
                        r.get('Customer Name','')).strip()
            qty      = safe_int(r.get('Qty','') or r.get('Quantity','') or 0)
            value    = safe_float(r.get('Value','') or
                                  r.get('Total Value','') or
                                  r.get('Amount','') or 0)
            tons     = safe_float(r.get('Weight Tons','') or
                                  r.get('Tons','') or 0)
            if not date:
                continue
            month = parse_month(date)
            sales.append({
                'date':        date,
                'month':       month,
                'vf_no':       vf_no,
                'customer':    customer,
                'qty':         qty,
                'fg_tons':     tons,
                'fg_value':    value,
                'jw_tons':     0,
                'jw_value':    0,
                'total_value': value,
                'total_tons':  tons,
            })

    print(f"  Total sales records: {len(sales):,}")
    return sales

# ═══════════════════════════════════════════════════════
#  STEP 3 — READ RM COSTS (from material_tracker output)
# ═══════════════════════════════════════════════════════

def read_rm_costs(svc, gc):
    """Read RM Consumption tab written by material_tracker."""
    print("\n📦 Reading RM Costs from master sheet...")
    try:
        spreadsheet = gc.open_by_key(MASTER_SHEET_ID)
        tabs        = [ws.title for ws in spreadsheet.worksheets()]
        rm_tab      = find_tab(tabs, ['rm consumption','rm by part'])

        if not rm_tab:
            print("  ⚠ RM Consumption tab not found — run material_tracker.py first")
            return {}

        ws      = spreadsheet.worksheet(rm_tab)
        values  = ws.get_all_values()
        time.sleep(DELAY)

        if not values or len(values) < 2:
            return {}

        headers = values[0]
        rm_costs = {}
        for row in values[1:]:
            padded = row + [''] * max(0, len(headers)-len(row))
            r      = dict(zip(headers, padded))
            vf_no  = r.get('VF No','').strip()
            month  = r.get('Month','').strip()
            cost   = safe_float(r.get('RM Cost (₹)','') or 0)
            qty    = safe_int(r.get('Qty','') or 0)
            rm_kg  = safe_float(r.get('RM Consumed (KG)','') or 0)

            if vf_no and month:
                rm_costs[(vf_no, month)] = {
                    'rm_cost': cost,
                    'qty':     qty,
                    'rm_kg':   rm_kg,
                }

        print(f"  RM cost records loaded: {len(rm_costs):,}")
        return rm_costs

    except Exception as e:
        print(f"  ⚠ Cannot read RM costs: {e}")
        return {}

# ═══════════════════════════════════════════════════════
#  STEP 4 — CALCULATE PROFITABILITY
# ═══════════════════════════════════════════════════════

def calculate_profitability(sales, rm_costs, parts):
    """
    Per VF No per Month:
      Revenue      = Qty × Unit Price  (or from dispatch value directly)
      RM Cost      = from material_tracker
      Contribution = Revenue - RM Cost
      Margin %     = Contribution / Revenue × 100
    """
    print("\n📊 Calculating profitability...")

    # Aggregate sales by vf_no + month
    sales_by_part = defaultdict(lambda: {
        'revenue':0, 'qty':0, 'tons':0, 'customer':''
    })
    sales_by_month = defaultdict(lambda: {'revenue':0, 'tons':0})

    for s in sales:
        month    = s['month']
        vf_no    = s.get('vf_no','')
        value    = s['total_value']
        tons     = s['total_tons']
        qty      = s.get('qty', 0)
        customer = s.get('customer','')

        sales_by_month[month]['revenue'] += value
        sales_by_month[month]['tons']    += tons

        if vf_no:
            key = (vf_no, month)
            sales_by_part[key]['revenue']  += value
            sales_by_part[key]['qty']      += qty
            sales_by_part[key]['tons']     += tons
            if customer and not sales_by_part[key]['customer']:
                sales_by_part[key]['customer'] = customer

    # Build profitability per part per month
    part_profit = []
    missing_price = set()

    # Get all unique vf_no + month combinations
    all_keys = set(sales_by_part.keys()) | set(rm_costs.keys())

    for (vf_no, month) in sorted(all_keys):
        part     = parts.get(vf_no, {})
        customer = (sales_by_part[(vf_no,month)].get('customer','') or
                    part.get('customer',''))
        machine  = part.get('machine','')

        # Revenue
        revenue = sales_by_part[(vf_no,month)]['revenue']
        qty     = sales_by_part[(vf_no,month)]['qty']
        tons    = sales_by_part[(vf_no,month)]['tons']

        # If no dispatch value but have unit price + qty from production
        if revenue == 0 and qty == 0:
            rm_rec = rm_costs.get((vf_no,month),{})
            qty    = rm_rec.get('qty',0)

        if revenue == 0 and qty > 0:
            unit_price = part.get('unit_price',0)
            if unit_price > 0:
                # Unit price is per KG — convert
                finish_wt = part.get('finish_weight',0)
                if finish_wt > 0:
                    revenue = qty * finish_wt * unit_price
                else:
                    revenue = qty * unit_price
            else:
                missing_price.add(vf_no)

        # RM Cost
        rm_rec  = rm_costs.get((vf_no,month),{})
        rm_cost = rm_rec.get('rm_cost',0)

        # Contribution
        contribution = revenue - rm_cost
        margin_pct   = round(contribution/revenue*100,1) if revenue > 0 else 0

        # Cost per piece
        cpp_rm  = round(rm_cost/qty,2)  if qty > 0 else 0
        rev_pce = round(revenue/qty,2)  if qty > 0 else 0
        con_pce = round(contribution/qty,2) if qty > 0 else 0

        # Status
        if margin_pct >= 30:
            status = '🏆 High Margin'
        elif margin_pct >= 15:
            status = '✅ Good'
        elif margin_pct >= 0:
            status = '⚠️ Low Margin'
        else:
            status = '❌ Loss Making'

        part_profit.append({
            'month':        month,
            'vf_no':        vf_no,
            'customer':     customer,
            'machine':      machine,
            'qty':          qty,
            'tons':         round(tons,3),
            'revenue':      round(revenue,2),
            'rm_cost':      round(rm_cost,2),
            'contribution': round(contribution,2),
            'margin_pct':   margin_pct,
            'rev_per_piece':rev_pce,
            'rm_per_piece': cpp_rm,
            'con_per_piece':con_pce,
            'status':       status,
        })

    if missing_price:
        print(f"  ⚠ No unit price for: {list(missing_price)[:10]}")
        print(f"    Fill Unit Price in Master Part sheet for these VF Nos")

    print(f"  Profitability records: {len(part_profit):,}")
    return part_profit

# ═══════════════════════════════════════════════════════
#  STEP 5 — WRITE TABS
# ═══════════════════════════════════════════════════════

def write_profitability(spreadsheet, part_profit):

    # ── Tab 1: Part-wise Profitability ───────────────────
    PART_HEADERS = ['Month','VF No','Customer','Machine',
                    'Qty','Tons','Revenue (₹)','RM Cost (₹)',
                    'Contribution (₹)','Margin %',
                    'Revenue/Piece','RM/Piece','Contribution/Piece',
                    'Status']
    part_rows = [PART_HEADERS]
    for r in sorted(part_profit, key=lambda x:(x['month'],x['vf_no'])):
        part_rows.append([
            r['month'], r['vf_no'], r['customer'], r['machine'],
            r['qty'], r['tons'], r['revenue'], r['rm_cost'],
            r['contribution'], r['margin_pct'],
            r['rev_per_piece'], r['rm_per_piece'], r['con_per_piece'],
            r['status'],
        ])

    print(f"\n✍ Writing Part Profitability ({len(part_rows)-1:,} rows)...")
    ws1 = get_or_create_tab(spreadsheet, '💰 Part Profitability', cols=14)
    safe_clear(ws1)
    write_in_batches(ws1, part_rows)
    ws1.format('A1:N1', FMT['green'])

    # ── Tab 2: Customer Profitability ────────────────────
    cust_month = defaultdict(lambda: {'revenue':0,'rm_cost':0,'qty':0,'parts':set()})
    for r in part_profit:
        if not r['customer']:
            continue
        key = (r['customer'], r['month'])
        cust_month[key]['revenue']  += r['revenue']
        cust_month[key]['rm_cost']  += r['rm_cost']
        cust_month[key]['qty']      += r['qty']
        cust_month[key]['parts'].add(r['vf_no'])

    CUST_HEADERS = ['Month','Customer','Parts Count','Total Qty',
                    'Revenue (₹)','RM Cost (₹)',
                    'Contribution (₹)','Margin %','Status']
    cust_rows = [CUST_HEADERS]
    for (customer, month), d in sorted(cust_month.items()):
        contrib    = d['revenue'] - d['rm_cost']
        margin     = round(contrib/d['revenue']*100,1) if d['revenue'] else 0
        status     = ('🏆 High' if margin>=30 else
                      '✅ Good' if margin>=15 else
                      '⚠️ Low'  if margin>=0  else '❌ Loss')
        cust_rows.append([
            month, customer, len(d['parts']), d['qty'],
            round(d['revenue'],2), round(d['rm_cost'],2),
            round(contrib,2), margin, status,
        ])

    print(f"✍ Writing Customer Profitability ({len(cust_rows)-1:,} rows)...")
    ws2 = get_or_create_tab(spreadsheet, '👥 Customer Profitability', cols=9)
    safe_clear(ws2)
    write_in_batches(ws2, cust_rows)
    ws2.format('A1:I1', FMT['purple'])

    # ── Tab 3: Machine Profitability ─────────────────────
    mach_month = defaultdict(lambda: {'revenue':0,'rm_cost':0,'qty':0,'parts':set()})
    for r in part_profit:
        if not r['machine']:
            continue
        key = (r['machine'], r['month'])
        mach_month[key]['revenue']  += r['revenue']
        mach_month[key]['rm_cost']  += r['rm_cost']
        mach_month[key]['qty']      += r['qty']
        mach_month[key]['parts'].add(r['vf_no'])

    MACH_HEADERS = ['Month','Machine','Parts Count','Total Qty',
                    'Revenue (₹)','RM Cost (₹)',
                    'Contribution (₹)','Margin %','Status']
    mach_rows = [MACH_HEADERS]
    for (machine, month), d in sorted(mach_month.items()):
        contrib = d['revenue'] - d['rm_cost']
        margin  = round(contrib/d['revenue']*100,1) if d['revenue'] else 0
        status  = ('🏆 High' if margin>=30 else
                   '✅ Good' if margin>=15 else
                   '⚠️ Low'  if margin>=0  else '❌ Loss')
        mach_rows.append([
            month, machine, len(d['parts']), d['qty'],
            round(d['revenue'],2), round(d['rm_cost'],2),
            round(contrib,2), margin, status,
        ])

    print(f"✍ Writing Machine Profitability ({len(mach_rows)-1:,} rows)...")
    ws3 = get_or_create_tab(spreadsheet, '⚙️ Machine Profitability', cols=9)
    safe_clear(ws3)
    write_in_batches(ws3, mach_rows)
    ws3.format('A1:I1', FMT['teal'])

    # ── Tab 4: Monthly P&L Summary ───────────────────────
    monthly = defaultdict(lambda: {'revenue':0,'rm_cost':0,'qty':0,
                                   'parts':set(),'customers':set()})
    for r in part_profit:
        m = r['month']
        monthly[m]['revenue']   += r['revenue']
        monthly[m]['rm_cost']   += r['rm_cost']
        monthly[m]['qty']       += r['qty']
        monthly[m]['parts'].add(r['vf_no'])
        if r['customer']:
            monthly[m]['customers'].add(r['customer'])

    PL_HEADERS = ['Month','Total Qty','Active Parts','Customers',
                  'Revenue (₹)','RM Cost (₹)',
                  'Gross Contribution (₹)','Gross Margin %','Status']
    pl_rows = [PL_HEADERS]
    for month, d in sorted(monthly.items()):
        contrib = d['revenue'] - d['rm_cost']
        margin  = round(contrib/d['revenue']*100,1) if d['revenue'] else 0
        status  = ('🏆 Strong' if margin>=30 else
                   '✅ Healthy' if margin>=15 else
                   '⚠️ Thin'   if margin>=0  else '❌ Loss Month')
        pl_rows.append([
            month, d['qty'],
            len(d['parts']), len(d['customers']),
            round(d['revenue'],2), round(d['rm_cost'],2),
            round(contrib,2), margin, status,
        ])

    print(f"✍ Writing Monthly P&L ({len(pl_rows)-1:,} rows)...")
    ws4 = get_or_create_tab(spreadsheet, '📈 Monthly P&L', cols=9)
    safe_clear(ws4)
    write_in_batches(ws4, pl_rows)
    ws4.format('A1:I1', FMT['gold'])

    # Summary stats
    loss_parts = [r for r in part_profit if r['margin_pct'] < 0]
    high_parts = [r for r in part_profit if r['margin_pct'] >= 30]
    print(f"\n✅ Profitability Engine complete")
    print(f"   Total records:    {len(part_profit):,}")
    print(f"   Loss-making:      {len(loss_parts):,} part-months")
    print(f"   High margin:      {len(high_parts):,} part-months")

# ═══════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════

def main():
    print('=' * 60)
    print('   PROFITABILITY ENGINE')
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

    parts       = read_part_master(svc)
    sales       = read_dispatch(svc)
    rm_costs    = read_rm_costs(svc, gc)
    part_profit = calculate_profitability(sales, rm_costs, parts)

    if not part_profit:
        raise SystemExit("❌ No profitability data calculated")

    write_profitability(spreadsheet, part_profit)


if __name__ == '__main__':
    main()
