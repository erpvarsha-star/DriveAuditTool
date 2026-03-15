"""
STAGE 2 — REPORT WRITER
Reads audit_data.json and writes into your
existing 'Drive Audit System' Google Sheet.
Run AFTER drive_audit.py completes.
Run: python write_report.py
"""

import json
import pickle
import time
import os
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import gspread
from datetime import datetime

# ═══════════════════════════════════════════════════════
#  YOUR EXISTING SHEET ID
# ═══════════════════════════════════════════════════════
SHEET_ID   = '13gHBHZz1MbvDMGtE3fAHJWyhTbnoNCcnZit5Rmh-L80'
DATA_FILE  = 'audit_data.json'
TOKEN_FILE = 'token.pickle'
BATCH_SIZE = 500   # Write 500 rows at a time to avoid API limits

# ═══════════════════════════════════════════════════════
#  TAB DEFINITIONS
# ═══════════════════════════════════════════════════════
TABS = [
    '📁 Files',
    '🔁 Duplicates',
    '🔗 Dependencies',
    '📋 Forms',
    '💤 Inactive',
    '🙋 Personal',
    '🏢 Varsha',
    '🏢 Stego',
    '🏢 Prima',
    '📁 Forgetech',
    '📁 RFQ',
    '📁 Master',
    '📊 Summary',
    '❓ Uncategorized'
]

HEADERS = [
    'File Name', 'File Type', 'Company', 'Project', 'Category',
    'Sheet Role', 'Dependencies', 'Size', 'Created', 'Last Modified',
    'Status', 'Owner', 'Folder Path', 'Duplicate Of', 'Link'
]

HEADER_FMT = {
    'backgroundColor': {'red': 0.13, 'green': 0.29, 'blue': 0.53},
    'textFormat': {
        'bold': True,
        'foregroundColor': {'red': 1.0, 'green': 1.0, 'blue': 1.0},
        'fontSize': 10
    }
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
    return creds


# ═══════════════════════════════════════════════════════
#  BATCH WRITER
# ═══════════════════════════════════════════════════════
def write_in_batches(ws, rows):
    """Write rows in batches of BATCH_SIZE with retry on error"""
    total    = len(rows)
    written  = 0
    for i in range(0, total, BATCH_SIZE):
        batch = rows[i:i+BATCH_SIZE]
        retry = 0
        while retry < 3:
            try:
                ws.append_rows(batch, value_input_option='USER_ENTERED')
                written += len(batch)
                print(f'    ✅ Written {written}/{total} rows...')
                time.sleep(1.5)  # Respect API rate limits
                break
            except Exception as e:
                retry += 1
                print(f'    ⚠️  Retry {retry}/3 after error: {e}')
                time.sleep(5)


# ═══════════════════════════════════════════════════════
#  SETUP TABS
# ═══════════════════════════════════════════════════════
def setup_tabs(spreadsheet):
    """Create or clear all required tabs"""
    existing = {ws.title: ws for ws in spreadsheet.worksheets()}

    worksheets = {}
    for tab_name in TABS:
        if tab_name in existing:
            ws = existing[tab_name]
            ws.clear()
            print(f'  🧹 Cleared: {tab_name}')
        else:
            ws = spreadsheet.add_worksheet(title=tab_name, rows=50000, cols=15)
            print(f'  ➕ Created: {tab_name}')
        worksheets[tab_name] = ws
        time.sleep(0.5)

    return worksheets


# ═══════════════════════════════════════════════════════
#  CONVERT RECORD TO ROW
# ═══════════════════════════════════════════════════════
def to_row(r):
    return [
        r.get('name', ''),
        r.get('file_type', ''),
        r.get('company', ''),
        r.get('project', ''),
        r.get('category', ''),
        r.get('sheet_role', ''),
        r.get('dependencies', ''),
        r.get('size', ''),
        r.get('created', ''),
        r.get('modified', ''),
        r.get('status', ''),
        r.get('owner', ''),
        r.get('folder_path', ''),
        r.get('duplicate_of', ''),
        r.get('link', '')
    ]


# ═══════════════════════════════════════════════════════
#  WRITE SUMMARY TAB
# ═══════════════════════════════════════════════════════
def write_summary(ws, files, timestamp):
    ws.clear()
    rows = []
    rows.append(['DRIVE AUDIT SUMMARY', '', f'Generated: {timestamp}'])
    rows.append([])

    def section(title, counts):
        rows.append([title, 'Count'])
        for k, v in sorted(counts.items(), key=lambda x: -x[1]):
            rows.append([k, v])
        rows.append([])

    company_c  = {}
    project_c  = {}
    category_c = {}
    ftype_c    = {}
    status_c   = {}
    dep_c      = {}

    for f in files:
        company_c[f['company']]   = company_c.get(f['company'], 0) + 1
        project_c[f['project']]   = project_c.get(f['project'], 0) + 1
        category_c[f['category']] = category_c.get(f['category'], 0) + 1
        ftype_c[f['file_type']]   = ftype_c.get(f['file_type'], 0) + 1
        status_c[f['status']]     = status_c.get(f['status'], 0) + 1
        if f.get('dependencies'):
            for d in f['dependencies'].split(', '):
                dep_c[d] = dep_c.get(d, 0) + 1

    section('📌 BY COMPANY',      company_c)
    section('📁 BY PROJECT',      project_c)
    section('🏷️ BY CATEGORY',     category_c)
    section('📄 BY FILE TYPE',    ftype_c)
    section('⚡ BY STATUS',       status_c)
    if dep_c:
        section('🔗 DEPENDENCIES FOUND', dep_c)

    dups = sum(1 for f in files if f.get('duplicate_of'))
    rows.append(['TOTALS', ''])
    rows.append(['Total Files',      len(files)])
    rows.append(['Total Duplicates', dups])
    rows.append(['Active Files',     status_c.get('Active', 0)])
    rows.append(['Inactive Files',   status_c.get('Inactive', 0)])

    ws.append_rows(rows)
    ws.format('A1:C1', {
        'backgroundColor': {'red': 0.13, 'green': 0.29, 'blue': 0.53},
        'textFormat': {'bold': True, 'foregroundColor': {'red': 1, 'green': 1, 'blue': 1}}
    })
    print(f'  ✅ Summary written')


# ═══════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════
def main():
    print('=' * 65)
    print('   STAGE 2 — Report Writer')
    print('   Writing to: Drive Audit System')
    print('=' * 65)

    # Load scan data
    if not os.path.exists(DATA_FILE):
        print(f'\n❌ {DATA_FILE} not found! Run drive_audit.py first.')
        return

    with open(DATA_FILE, 'r', encoding='utf-8') as f:
        data = json.load(f)

    files     = data.get('files', [])
    saved_at  = data.get('saved_at', 'Unknown')
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M')

    print(f'\n📊 Loaded {len(files)} files from {DATA_FILE}')
    print(f'🕐 Scan was saved at: {saved_at}\n')

    # Auth
    creds       = authenticate()
    gc          = gspread.authorize(creds)
    spreadsheet = gc.open_by_key(SHEET_ID)

    print('🗂️  Setting up tabs...')
    worksheets = setup_tabs(spreadsheet)

    # ── Filter sets ──────────────────────────────────
    all_rows        = [to_row(f) for f in files]
    duplicates      = [to_row(f) for f in files if f.get('duplicate_of')]
    dependencies    = [to_row(f) for f in files if f.get('dependencies')]
    forms           = [to_row(f) for f in files if 'Form' in f.get('sheet_role', '') or
                       f.get('file_type') == 'Google Form' or
                       'form' in f.get('dependencies', '').lower()]
    inactive        = [to_row(f) for f in files if f.get('status') == 'Inactive']
    personal        = [to_row(f) for f in files if f.get('company') == 'Personal']
    varsha          = [to_row(f) for f in files if f.get('company') == 'Varsha (VFL/VFPL)']
    stego           = [to_row(f) for f in files if f.get('company') == 'Stego']
    prima           = [to_row(f) for f in files if f.get('company') == 'Prima']
    forgetech       = [to_row(f) for f in files if f.get('project') == 'Forgetech']
    rfq             = [to_row(f) for f in files if f.get('project') == 'RFQ']
    master          = [to_row(f) for f in files if f.get('project') == 'Master']
    uncategorized   = [to_row(f) for f in files if f.get('category') == 'Other / Uncategorized']

    # ── Write each tab ───────────────────────────────
    tab_data = {
        '📁 Files':        all_rows,
        '🔁 Duplicates':   duplicates,
        '🔗 Dependencies': dependencies,
        '📋 Forms':        forms,
        '💤 Inactive':     inactive,
        '🙋 Personal':     personal,
        '🏢 Varsha':       varsha,
        '🏢 Stego':        stego,
        '🏢 Prima':        prima,
        '📁 Forgetech':    forgetech,
        '📁 RFQ':          rfq,
        '📁 Master':       master,
        '❓ Uncategorized': uncategorized,
    }

    for tab_name, rows in tab_data.items():
        ws = worksheets[tab_name]
        print(f'\n📝 Writing {tab_name} ({len(rows)} rows)...')
        ws.append_row(HEADERS)
        ws.format('A1:O1', HEADER_FMT)
        ws.freeze(rows=1)
        if rows:
            write_in_batches(ws, rows)
        else:
            ws.append_row(['No data for this category'])

    # ── Summary tab ──────────────────────────────────
    print(f'\n📊 Writing Summary...')
    write_summary(worksheets['📊 Summary'], files, timestamp)

    print('\n' + '=' * 65)
    print('🎉 REPORT COMPLETE!')
    print(f'📋 Open: https://docs.google.com/spreadsheets/d/{SHEET_ID}')
    print(f'\n📊 Quick Stats:')
    print(f'   Total Files:    {len(files)}')
    print(f'   Duplicates:     {len(duplicates)}')
    print(f'   Dependencies:   {len(dependencies)}')
    print(f'   Forms:          {len(forms)}')
    print(f'   Inactive:       {len(inactive)}')
    print(f'   Personal:       {len(personal)}')
    print(f'   Varsha:         {len(varsha)}')
    print(f'   Stego:          {len(stego)}')
    print(f'   Uncategorized:  {len(uncategorized)}')
    print('=' * 65)

if __name__ == '__main__':
    main()