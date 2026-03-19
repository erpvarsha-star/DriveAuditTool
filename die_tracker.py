"""
die_tracker.py — Die Life Tracker
Reads PMS Forgings data, detects consecutive die runs per
Machine + VF No, calculates die life achieved vs estimated.

Logic:
  - Sort PMS by Machine + Date + Shift
  - Consecutive rows with same Machine + same VF No = one die run
  - Sum pieces in that run = Die Life Achieved
  - Average all historical runs = Estimated Die Life
  - Monthly: current run ÷ estimated = achievement %

Status:
  < 90%  → BAD  (die wore out early)
  90-104% → GOOD (normal life)
  ≥ 105% → GREAT (exceeded life)

Output: Die Life Tracker tab in MASTER_SHEET
Run: python die_tracker.py
"""

import json
import os
import pickle
import time
from datetime import datetime, timedelta
from collections import defaultdict

import gspread
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ═══════════════════════════════════════════════════════
#  CONFIG
# ═══════════════════════════════════════════════════════

PMS_ID         = '1c-axqiBEufNb1vK-JdJP6t1eRTmzliA3otAbqxAOXZE'
MASTER_SHEET_ID = os.environ.get('MASTER_SHEET_ID', '10Zjxy3mGKP6G3j7uuak3FTXl90JHEQoC0RDgl4tJJXc')

TOKEN_FILE  = 'token.pickle'
DELAY       = 0.5
BATCH_SIZE  = 500
TAB_INIT    = 1

FMT = {
    'blue':   {'backgroundColor': {'red':0.1, 'green':0.2, 'blue':0.4},
               'textFormat': {'bold':True, 'foregroundColor': {'red':1,'green':1,'blue':1}}},
    'green':  {'backgroundColor': {'red':0.0, 'green':0.5, 'blue':0.2},
               'textFormat': {'bold':True, 'foregroundColor': {'red':1,'green':1,'blue':1}}},
    'red':    {'backgroundColor': {'red':0.6, 'green':0.1, 'blue':0.1},
               'textFormat': {'bold':True, 'foregroundColor': {'red':1,'green':1,'blue':1}}},
    'gold':   {'backgroundColor': {'red':0.8, 'green':0.6, 'blue':0.0},
               'textFormat': {'bold':True, 'foregroundColor': {'red':1,'green':1,'blue':1}}},
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
        rows = resp.get('values', [])
        if not rows or len(rows) < 2:
            return [], []
        headers = [str(h).strip() for h in rows[0]]
        data = []
        for row in rows[1:]:
            padded = row + [''] * max(0, len(headers) - len(row))
            data.append(dict(zip(headers, padded[:len(headers)])))
        return headers, data
    except HttpError as e:
        if e.resp.status == 429:
            print(f"  ⏳ Quota — waiting 60s")
            time.sleep(60)
            return read_tab(svc, sheet_id, tab, max_rows)
        print(f"  ⚠ Cannot read {tab}: {e}")
        return [], []


def find_tab(tabs, keywords):
    for kw in keywords:
        for tab in tabs:
            if kw.lower() in tab.lower():
                return tab
    return tabs[0] if tabs else None


def safe_int(v):
    try:
        return int(float(str(v).replace(',','')))
    except Exception:
        return 0


def parse_date(v):
    """Parse date string to date object."""
    for fmt in ['%d/%m/%Y','%Y-%m-%d','%m/%d/%Y','%d-%m-%Y',
                '%d/%m/%Y %H:%M:%S','%Y-%m-%dT%H:%M:%S']:
        try:
            return datetime.strptime(str(v).strip()[:19], fmt).date()
        except Exception:
            continue
    return None


def shift_order(shift_val):
    """Convert shift to sortable integer."""
    s = str(shift_val).strip().upper()
    if s in ['1','A','MORNING','FIRST']:   return 1
    if s in ['2','B','AFTERNOON','SECOND']: return 2
    if s in ['3','C','NIGHT','THIRD']:      return 3
    try:
        return int(float(s))
    except Exception:
        return 99


def get_or_create_tab(spreadsheet, title, cols=10):
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
#  STEP 1 — READ PMS DATA
# ═══════════════════════════════════════════════════════

def read_pms(svc):
    """Read all production data from PMS Forgings sheet."""
    print("\n📋 Reading PMS Forgings data...")
    tabs = get_tabs(svc, PMS_ID)
    print(f"  Tabs: {tabs}")

    # Collect from all Form Response tabs
    all_data = []
    for tab in tabs:
        if any(kw in tab.lower() for kw in ['form response', 'responses', 'data']):
            headers, data = read_tab(svc, PMS_ID, tab)
            if data:
                print(f"  {tab}: {len(data):,} rows | Cols: {headers[:10]}")
                all_data.extend(data)

    if not all_data:
        # Try first tab
        tab = tabs[0] if tabs else None
        if tab:
            headers, all_data = read_tab(svc, PMS_ID, tab)
            print(f"  {tab}: {len(all_data):,} rows | Cols: {headers[:10]}")

    print(f"  Total PMS rows: {len(all_data):,}")
    return all_data

# ═══════════════════════════════════════════════════════
#  STEP 2 — DETECT DIE RUNS
# ═══════════════════════════════════════════════════════

def detect_die_runs(pms_data):
    """
    Detect consecutive die runs per Machine + VF No.
    
    A run = consecutive shifts on same Machine with same VF No.
    Break = VF No changes on the same machine = die change.
    """
    print("\n🔍 Detecting die runs...")

    # Parse and sort entries
    entries = []
    for r in pms_data:
        # Find machine column
        machine = (r.get('Machine', '') or
                   r.get('Machine Name', '') or
                   r.get('VF Machine', '') or
                   r.get('Hammer', '') or
                   r.get('Press', '')).strip()

        # Find VF No (part number)
        vf_no = (r.get('VF No', '') or
                 r.get('Part No', '') or
                 r.get('Part Number', '') or
                 r.get('VF_No', '')).strip()

        # Find date
        date_str = (r.get('Date', '') or
                    r.get('Production Date', '') or
                    r.get('DATE', '')).strip()

        # Find shift
        shift = (r.get('Shift', '') or
                 r.get('Shift No', '') or
                 r.get('SHIFT', '')).strip()

        # Find qty
        qty = safe_int(r.get('Qty Forged', '') or
                       r.get('Total Qty', '') or
                       r.get('Qty', '') or
                       r.get('Quantity', '') or
                       r.get('OK Qty', '') or 0)

        # Parse timestamp for sorting
        ts = r.get('Timestamp', '')
        date_obj = parse_date(date_str) or parse_date(ts)

        if not machine or not vf_no or qty <= 0:
            continue

        entries.append({
            'machine':   machine,
            'vf_no':     vf_no,
            'date':      date_obj,
            'date_str':  str(date_obj) if date_obj else date_str,
            'shift':     shift_order(shift),
            'qty':       qty,
        })

    if not entries:
        print("  ⚠ No valid entries found — check column names in PMS")
        return []

    # Sort by machine, date, shift
    entries.sort(key=lambda x: (
        x['machine'],
        x['date'] or datetime.min.date(),
        x['shift']
    ))

    print(f"  Valid entries: {len(entries):,}")
    print(f"  Machines found: {sorted(set(e['machine'] for e in entries))}")
    print(f"  Sample VF Nos: {sorted(set(e['vf_no'] for e in entries))[:10]}")

    # Detect runs — consecutive same machine + same vf_no
    runs     = []
    i        = 0
    while i < len(entries):
        current  = entries[i]
        machine  = current['machine']
        vf_no    = current['vf_no']
        run_qty  = current['qty']
        run_start= current['date_str']
        run_end  = current['date_str']
        shifts   = 1

        # Look ahead for consecutive same machine + vf_no
        j = i + 1
        while j < len(entries):
            nxt = entries[j]
            if nxt['machine'] == machine and nxt['vf_no'] == vf_no:
                run_qty += nxt['qty']
                run_end  = nxt['date_str']
                shifts  += 1
                j       += 1
            else:
                break

        runs.append({
            'machine':    machine,
            'vf_no':      vf_no,
            'start_date': run_start,
            'end_date':   run_end,
            'shifts':     shifts,
            'qty':        run_qty,
        })

        i = j if j > i else i + 1

    print(f"  Die runs detected: {len(runs):,}")
    return runs

# ═══════════════════════════════════════════════════════
#  STEP 3 — CALCULATE DIE LIFE ESTIMATES
# ═══════════════════════════════════════════════════════

def calculate_die_life(runs):
    """
    For each Machine + VF No combination:
    - Collect all historical runs
    - Average pieces = Estimated Die Life
    """
    print("\n📊 Calculating die life estimates...")

    # Group runs by machine + vf_no
    grouped = defaultdict(list)
    for run in runs:
        key = (run['machine'], run['vf_no'])
        grouped[key].append(run['qty'])

    estimates = {}
    for (machine, vf_no), qtys in grouped.items():
        estimates[(machine, vf_no)] = {
            'machine':        machine,
            'vf_no':          vf_no,
            'run_count':      len(qtys),
            'estimated_life': round(sum(qtys) / len(qtys)),
            'min_run':        min(qtys),
            'max_run':        max(qtys),
            'total_pieces':   sum(qtys),
            'all_runs':       qtys,
        }

    print(f"  Machine+Part combinations: {len(estimates):,}")
    return estimates

# ═══════════════════════════════════════════════════════
#  STEP 4 — CURRENT MONTH STATUS
# ═══════════════════════════════════════════════════════

def get_current_month_status(runs, estimates):
    """
    For each machine+part, find the most recent run this month
    and compare against estimated die life.
    """
    now       = datetime.now()
    this_month= now.strftime('%Y-%m')

    # Find latest run per machine+vf_no this month
    current_runs = {}
    for run in runs:
        key = (run['machine'], run['vf_no'])
        # Check if run is in current month
        if this_month in str(run['end_date']):
            if key not in current_runs:
                current_runs[key] = run['qty']
            else:
                current_runs[key] += run['qty']

    # Build status rows
    status_rows = []
    for key, est in estimates.items():
        current_qty  = current_runs.get(key, 0)
        estimated    = est['estimated_life']

        if estimated > 0 and current_qty > 0:
            pct = round(current_qty / estimated * 100, 1)
        else:
            pct = 0

        if pct >= 105:
            status = '🏆 GREAT'
            color  = 'gold'
        elif pct >= 90:
            status = '✅ GOOD'
            color  = 'green'
        elif pct > 0:
            status = '❌ BAD'
            color  = 'red'
        else:
            status = '⏳ Running'
            color  = 'blue'

        status_rows.append({
            'machine':        est['machine'],
            'vf_no':          est['vf_no'],
            'estimated_life': estimated,
            'current_qty':    current_qty,
            'achievement_pct':pct,
            'status':         status,
            'run_count':      est['run_count'],
            'min_run':        est['min_run'],
            'max_run':        est['max_run'],
            'total_pieces':   est['total_pieces'],
        })

    # Sort by machine then vf_no
    status_rows.sort(key=lambda x: (x['machine'], x['vf_no']))
    return status_rows

# ═══════════════════════════════════════════════════════
#  STEP 5 — WRITE TO MASTER SHEET
# ═══════════════════════════════════════════════════════

def write_die_life(spreadsheet, runs, status_rows):
    """Write die life data to master sheet."""

    # Tab 1 — Die Life Estimates (all historical)
    EST_HEADERS = [
        'Machine', 'VF No', 'Est. Die Life (Pcs)',
        'Min Run', 'Max Run', 'Total Runs',
        'Total Pieces', 'Current Month Qty',
        'Achievement %', 'Status',
    ]

    est_rows = [EST_HEADERS]
    for r in status_rows:
        est_rows.append([
            r['machine'], r['vf_no'], r['estimated_life'],
            r['min_run'], r['max_run'], r['run_count'],
            r['total_pieces'], r['current_qty'],
            r['achievement_pct'], r['status'],
        ])

    print(f"\n✍ Writing Die Life Tracker ({len(est_rows)-1:,} rows)...")
    ws = get_or_create_tab(spreadsheet, '🔩 Die Life Tracker', cols=10)
    safe_clear(ws)
    write_in_batches(ws, est_rows)
    ws.format('A1:J1', FMT['blue'])

    # Colour code status column
    for i, r in enumerate(status_rows, start=2):
        if '🏆' in r['status']:
            ws.format(f'J{i}', FMT['gold'])
        elif '✅' in r['status']:
            ws.format(f'J{i}', FMT['green'])
        elif '❌' in r['status']:
            ws.format(f'J{i}', FMT['red'])
        if i > 200:  # limit formatting calls
            break

    # Tab 2 — All Historical Runs
    RUN_HEADERS = [
        'Machine', 'VF No', 'Start Date', 'End Date',
        'Shifts', 'Qty Produced',
    ]
    run_rows = [RUN_HEADERS]
    for r in sorted(runs, key=lambda x: (x['machine'], x['vf_no'], x['start_date'])):
        run_rows.append([
            r['machine'], r['vf_no'],
            r['start_date'], r['end_date'],
            r['shifts'], r['qty'],
        ])

    print(f"✍ Writing All Die Runs ({len(run_rows)-1:,} rows)...")
    ws2 = get_or_create_tab(spreadsheet, '📋 Die Run History', cols=6)
    safe_clear(ws2)
    write_in_batches(ws2, run_rows)
    ws2.format('A1:F1', FMT['blue'])

    print(f"\n✅ Die Life Tracker complete")
    print(f"   Machine+Part combinations: {len(status_rows):,}")
    print(f"   Total historical runs:     {len(runs):,}")

    # Summary
    great = sum(1 for r in status_rows if '🏆' in r['status'])
    good  = sum(1 for r in status_rows if '✅' in r['status'])
    bad   = sum(1 for r in status_rows if '❌' in r['status'])
    print(f"   This month: 🏆 {great} GREAT | ✅ {good} GOOD | ❌ {bad} BAD")

# ═══════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════

def main():
    print('=' * 60)
    print('   DIE LIFE TRACKER')
    print(f'   {datetime.now().strftime("%Y-%m-%d %H:%M")}')
    print('=' * 60)

    creds = authenticate()
    svc   = build('sheets', 'v4', credentials=creds)
    gc    = gspread.authorize(creds)

    try:
        spreadsheet = gc.open_by_key(MASTER_SHEET_ID)
        print(f"Writing to: {spreadsheet.title}")
    except Exception as e:
        raise SystemExit(f"❌ Cannot open master sheet: {e}")

    # Run pipeline
    pms_data    = read_pms(svc)
    if not pms_data:
        raise SystemExit("❌ No PMS data found — check PMS sheet access")

    runs        = detect_die_runs(pms_data)
    if not runs:
        raise SystemExit("❌ No die runs detected — check Machine and VF No columns in PMS")

    estimates   = calculate_die_life(runs)
    status_rows = get_current_month_status(runs, estimates)
    write_die_life(spreadsheet, runs, status_rows)


if __name__ == '__main__':
    main()
