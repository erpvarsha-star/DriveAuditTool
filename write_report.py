"""
STAGE 2 — REPORT WRITER  (cell-limit safe, v3)
Reads audit_data.json and writes into Google Sheets without hitting
the 10,000,000-cell limit.

Root cause fix:
  Never pass rows=len(files) to add_worksheet().
  Google pre-allocates every row×col at creation time.
  27,097 files × 12 cols × 12 tabs = way over 10M cells.
  Fix: always create tabs with rows=1 and let append_rows() grow them.

Cell budget for 27,097 files:
  Summary tab        ~30  rows × 3  cols =      90 cells
  ERP Priority       ~500 rows × 12 cols =   6,000 cells  (high-score only)
  Duplicates         varies    × 8  cols   (slim columns)
  Full Inventory  → written to a SEPARATE Sheets file (INVENTORY_SHEET_ID)
"""

import json
import os
import pickle
import time
from datetime import datetime

import gspread
from google.auth.transport.requests import Request

# ═══════════════════════════════════════════════════════
#  CONFIG
# ═══════════════════════════════════════════════════════

SHEET_ID           = '13gHBHZz1MbvDMGtE3fAHJWyhTbnoNCcnZit5Rmh-L80'
INVENTORY_SHEET_ID = os.environ.get('INVENTORY_SHEET_ID', '')  # separate sheet for full data

DATA_FILE  = 'audit_data.json'
TOKEN_FILE = 'token.pickle'
BATCH_SIZE = 500

TAB_INIT_ROWS = 1    # KEY FIX: never pre-allocate — append_rows grows the sheet
TAB_INIT_COLS = 12

FULL_HEADERS = [
    'File Name', 'Status', 'Score', 'ERP Ready', 'Category',
    'File Type', 'Last Modified', 'Owner', 'Dependencies',
    'Folder Path', 'Duplicate Notes', 'Link',
]

DUP_HEADERS = [
    'File Name', 'Status', 'Category', 'File Type',
    'Last Modified', 'Owner', 'Duplicate Notes', 'Folder Path',
]

MIME_LABELS = {
    'application/vnd.google-apps.spreadsheet':  'Sheet',
    'application/vnd.google-apps.document':     'Doc',
    'application/vnd.google-apps.presentation': 'Slides',
    'application/vnd.google-apps.folder':       'Folder',
    'application/pdf':                           'PDF',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':         'Excel',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document':   'Word',
    'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'PPT',
    'image/jpeg': 'JPEG', 'image/png': 'PNG',
}

FMT = {
    'blue':  {'backgroundColor': {'red': 0.1,  'green': 0.2, 'blue': 0.4},
              'textFormat': {'bold': True, 'foregroundColor': {'red':1,'green':1,'blue':1}}},
    'green': {'backgroundColor': {'red': 0.0,  'green': 0.5, 'blue': 0.2},
              'textFormat': {'bold': True, 'foregroundColor': {'red':1,'green':1,'blue':1}}},
    'red':   {'backgroundColor': {'red': 0.55, 'green': 0.1, 'blue': 0.1},
              'textFormat': {'bold': True, 'foregroundColor': {'red':1,'green':1,'blue':1}}},
}

# ═══════════════════════════════════════════════════════
#  AUTH
# ═══════════════════════════════════════════════════════

def authenticate():
    creds = None
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, 'rb') as f:
            creds = pickle.load(f)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            raise RuntimeError("No valid token.pickle — run token_refresh.py locally.")
    return creds

# ═══════════════════════════════════════════════════════
#  SHEET HELPERS
# ═══════════════════════════════════════════════════════

def get_or_create_tab(spreadsheet, title, cols=TAB_INIT_COLS):
    """
    ROOT CAUSE FIX: create with rows=1, not rows=len(data).
    Google pre-allocates rows*cols cells at sheet creation.
    append_rows() grows the sheet dynamically at zero upfront cost.
    """
    try:
        ws = spreadsheet.worksheet(title)
        print(f"  📋 Reusing: {title}")
        return ws
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=title, rows=TAB_INIT_ROWS, cols=cols)
        print(f"  ➕ Created: {title}")
        return ws


def safe_clear(ws):
    ws.clear()
    time.sleep(0.5)


def write_in_batches(ws, rows):
    total = len(rows)
    if total == 0:
        print("    (no rows)")
        return
    for i in range(0, total, BATCH_SIZE):
        chunk = rows[i: i + BATCH_SIZE]
        ws.append_rows(chunk, value_input_option='USER_ENTERED')
        print(f"    ↳ {min(i + BATCH_SIZE, total):,} / {total:,}")
        time.sleep(1)

# ═══════════════════════════════════════════════════════
#  ROW CONVERTERS
# ═══════════════════════════════════════════════════════

def _mime(r):
    raw = r.get('Type', r.get('mimeType', r.get('file_type', '')))
    return MIME_LABELS.get(raw, raw.split('/')[-1] if raw else '')


def to_full_row(r):
    return [
        r.get('Name',         r.get('name',         '')),
        r.get('Status',       r.get('status',        '')),
        r.get('Score',        r.get('score',          0)),
        r.get('ERP_Ready',    r.get('erp_ready',    'NO')),
        r.get('Category',     r.get('category',      '')),
        _mime(r),
        r.get('Modified',     r.get('modified',      ''))[:10],
        r.get('Owner',        r.get('owner',         '')),
        r.get('Dependencies', r.get('dependencies',  '')),
        r.get('Folder_Path',  r.get('folder_path',   '')),
        r.get('Notes',        r.get('duplicate_of',  '')),
        r.get('Link',         r.get('link',          '')),
    ]


def to_dup_row(r):
    return [
        r.get('Name',         r.get('name',         '')),
        r.get('Status',       r.get('status',        '')),
        r.get('Category',     r.get('category',      '')),
        _mime(r),
        r.get('Modified',     r.get('modified',      ''))[:10],
        r.get('Owner',        r.get('owner',         '')),
        r.get('Notes',        r.get('duplicate_of',  '')),
        r.get('Folder_Path',  r.get('folder_path',   '')),
    ]

# ═══════════════════════════════════════════════════════
#  CELL BUDGET REPORT
# ═══════════════════════════════════════════════════════

def print_cell_budget(files):
    total = len(files)
    erp   = sum(1 for f in files if f.get('ERP_Ready','NO') == 'YES')
    dups  = sum(1 for f in files if f.get('Status','')      == 'DUPLICATE')

    summary_cells = 30 * 3
    erp_cells     = (erp  + 1) * 12
    dup_cells     = (dups + 1) * 8
    dash_total    = summary_cells + erp_cells + dup_cells
    inv_cells     = (total + 1) * 12

    print("\n📐 Cell budget")
    print(f"   Summary     : {summary_cells:>10,}")
    print(f"   ERP list    : {erp_cells:>10,}  ({erp:,} files × 12 cols)")
    print(f"   Duplicates  : {dup_cells:>10,}  ({dups:,} files × 8 cols)")
    print(f"   ─────────────────────────────────────")
    print(f"   Dashboard   : {dash_total:>10,}  {'✅ safe' if dash_total < 10_000_000 else '❌ over limit'}")
    print(f"   Full inventory → separate file: {inv_cells:,} cells")
    print()

# ═══════════════════════════════════════════════════════
#  TAB WRITERS
# ═══════════════════════════════════════════════════════

def write_summary(spreadsheet, files):
    total     = len(files)
    dups      = sum(1 for f in files if f.get('Status','')      == 'DUPLICATE')
    erp_ready = sum(1 for f in files if f.get('ERP_Ready','NO') == 'YES')
    bloat     = round(dups / total * 100, 1) if total else 0

    cats = {}
    for f in files:
        c = f.get('Category', f.get('category', 'Other'))
        cats[c] = cats.get(c, 0) + 1

    rows = [
        ['MIGRATION READINESS SUMMARY', '', datetime.now().strftime('%Y-%m-%d %H:%M')],
        [''],
        ['METRIC',                 'COUNT',              'ACTION'],
        ['Total files scanned',    total,                'Full inventory → separate sheet'],
        ['Duplicates found',       dups,                 f'Dedup potential: {bloat} %'],
        ['ERP-ready (score > 65)', erp_ready,            'Migrate by April 1st'],
        ['Low value / archive',    total-erp_ready-dups, 'Move to 2025 Archive folder'],
        [''],
        ['CATEGORY BREAKDOWN', 'FILES', ''],
    ] + [[f'  {c}', n, ''] for c, n in sorted(cats.items(), key=lambda x: -x[1])]

    print('\nWriting Executive Summary...')
    ws = get_or_create_tab(spreadsheet, '📊 Summary', cols=3)
    safe_clear(ws)
    ws.append_rows(rows, value_input_option='USER_ENTERED')
    ws.format('A1:C1', FMT['blue'])
    ws.format('A3:C3', FMT['blue'])
    ws.format('A9:C9', FMT['blue'])


def write_erp_priority(spreadsheet, files):
    priority = sorted(
        [f for f in files if f.get('ERP_Ready','NO') == 'YES'],
        key=lambda x: x.get('Score', x.get('score', 0)),
        reverse=True,
    )
    print(f'\nWriting ERP Priority ({len(priority):,} files)...')
    ws = get_or_create_tab(spreadsheet, '🚀 ERP Migration List')
    safe_clear(ws)
    ws.append_row(FULL_HEADERS)
    ws.format('A1:L1', FMT['green'])
    write_in_batches(ws, [to_full_row(f) for f in priority])


def write_duplicates(spreadsheet, files):
    dupes = [f for f in files if f.get('Status','') == 'DUPLICATE']
    print(f'\nWriting Duplicates ({len(dupes):,} files, 8-col slim format)...')
    ws = get_or_create_tab(spreadsheet, '🔁 Duplicates', cols=8)
    safe_clear(ws)
    ws.append_row(DUP_HEADERS)
    ws.format('A1:H1', FMT['red'])
    write_in_batches(ws, [to_dup_row(f) for f in dupes])


def write_full_inventory(gc, files):
    """Full inventory → separate Sheet so dashboard never hits cell limit."""
    if not INVENTORY_SHEET_ID:
        print('\nℹ️  INVENTORY_SHEET_ID not set — skipping full inventory.')
        print('   Steps to enable:')
        print('   1. Create a new blank Google Sheet')
        print('   2. Copy its ID from the URL')
        print('   3. Add GitHub secret: INVENTORY_SHEET_ID=<id>')
        return

    print(f'\nWriting Full Inventory ({len(files):,} files) → separate sheet...')
    try:
        inv = gc.open_by_key(INVENTORY_SHEET_ID)
    except gspread.exceptions.SpreadsheetNotFound:
        print(f'   ❌ Sheet not found: {INVENTORY_SHEET_ID}')
        return

    ws = get_or_create_tab(inv, 'Full Inventory')
    safe_clear(ws)
    ws.append_row(FULL_HEADERS)
    ws.format('A1:L1', FMT['blue'])
    write_in_batches(ws, [to_full_row(f) for f in files])
    print('   ✅ Full inventory written.')

# ═══════════════════════════════════════════════════════
#  NOTIFICATION
# ═══════════════════════════════════════════════════════

def notify(message, status='success'):
    import urllib.request
    webhook = os.environ.get('SLACK_WEBHOOK', '')
    if not webhook:
        return
    icon    = '✅' if status == 'success' else '❌'
    payload = json.dumps({'text': f'{icon} Drive Audit — {message}'}).encode()
    try:
        req = urllib.request.Request(
            webhook, data=payload,
            headers={'Content-Type': 'application/json'},
        )
        urllib.request.urlopen(req, timeout=10)
        print('📣 Slack notified.')
    except Exception as e:
        print(f'⚠ Slack failed (non-fatal): {e}')

# ═══════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════

def main():
    print('=' * 60)
    print('   STAGE 2 — Report Writer')

    if not os.path.exists(DATA_FILE):
        raise SystemExit(f"❌ {DATA_FILE} not found — run drive_audit.py first.")

    with open(DATA_FILE) as f:
        all_files = json.load(f)
    print(f'📊 Loaded {len(all_files):,} files from {DATA_FILE}')

    print_cell_budget(all_files)

    creds = authenticate()
    gc    = gspread.authorize(creds)

    try:
        spreadsheet = gc.open_by_key(SHEET_ID)
        print(f'Writing to: {spreadsheet.title}')
        print('=' * 60)

        write_summary(spreadsheet, all_files)
        write_erp_priority(spreadsheet, all_files)
        write_duplicates(spreadsheet, all_files)
        write_full_inventory(gc, all_files)

        total     = len(all_files)
        erp_ready = sum(1 for f in all_files if f.get('ERP_Ready','NO') == 'YES')
        dups      = sum(1 for f in all_files if f.get('Status','')      == 'DUPLICATE')
        msg = (f"{total:,} files | {erp_ready:,} ERP-ready | "
               f"{dups:,} duplicates | {datetime.now().strftime('%Y-%m-%d')}")

        print(f'\n✅ Done — {msg}')
        notify(msg, 'success')

    except Exception as e:
        print(f'\n❌ Failed: {e}')
        notify(f'FAILED: {e}', 'error')
        raise


if __name__ == '__main__':
    main()
