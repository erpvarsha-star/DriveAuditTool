"""
STAGE 2 — REPORT WRITER (v5)

New tabs added:
  📎 Form Attachments   — images/PDFs from forms, linked to parent form
  📂 Drive Reorg Plan   — current path → suggested ERP module path
  🗂 Master Data Map    — all detected master data files by type
  📋 Form-linked Sheets — with ACTIVE/INACTIVE status column
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

SHEET_ID           = '1yrjuBo9LJyR41cO0AZzkWJQQ_BAtYhTJ0Y0VgLr4eog'
INVENTORY_SHEET_ID = os.environ.get('INVENTORY_SHEET_ID', '')
DATA_FILE          = 'audit_data.json'
TOKEN_FILE         = 'token.pickle'
BATCH_SIZE         = 500
TAB_INIT_ROWS      = 1

FULL_HEADERS = [
    'File Name', 'Status', 'Score', 'ERP Ready', 'Has Form', 'Form Status',
    'Master Data Type', 'Category', 'File Type', 'Last Modified', 'Owner',
    'Dependencies', 'Folder Path', 'Duplicate Notes', 'Link',
]

DUP_HEADERS = [
    'File Name', 'Status', 'Category', 'File Type',
    'Last Modified', 'Owner', 'Duplicate Notes', 'Link',
]

FORM_HEADERS = [
    'File Name', 'Form Status', 'Form Detail', 'Score', 'ERP Ready',
    'Category', 'Last Modified', 'Owner', 'Dependencies', 'Folder Path', 'Link',
]

ATT_HEADERS = [
    'File Name', 'File Type', 'Parent Form Sheet', 'Suggested Folder',
    'Last Modified', 'Owner', 'Folder Path', 'Link',
]

REORG_HEADERS = [
    'File Name', 'Current Folder Path', 'Suggested ERP Folder',
    'Suggested New Name', 'ERP Module', 'Master Data Type', 'Link',
]

MASTER_HEADERS = [
    'File Name', 'Master Data Type', 'Category', 'Score',
    'ERP Ready', 'Last Modified', 'Owner', 'Folder Path', 'Link',
]

MIME_LABELS = {
    'application/vnd.google-apps.spreadsheet':  'Sheet',
    'application/vnd.google-apps.document':     'Doc',
    'application/vnd.google-apps.presentation': 'Slides',
    'application/vnd.google-apps.form':         'Form',
    'application/pdf':                           'PDF',
    'image/jpeg': 'JPEG', 'image/png': 'PNG',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':         'Excel',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document':   'Word',
    'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'PPT',
}

ERP_MODULE_MAP = {
    'Plant Operations': 'Production / Planning (PPC)',
    'Quality & QMS':    'Quality Management (QMS)',
    'Commercial':       'Purchase / Procurement',
    'Finance':          'Finance & Accounts',
    'HR & Admin':       'HR & Payroll',
    'Sales':            'Sales / Dispatch',
    'Production':       'Production / Planning (PPC)',
}

FMT = {
    'blue':   {'backgroundColor': {'red':0.1,'green':0.2,'blue':0.4},
               'textFormat':{'bold':True,'foregroundColor':{'red':1,'green':1,'blue':1}}},
    'green':  {'backgroundColor': {'red':0.0,'green':0.5,'blue':0.2},
               'textFormat':{'bold':True,'foregroundColor':{'red':1,'green':1,'blue':1}}},
    'red':    {'backgroundColor': {'red':0.55,'green':0.1,'blue':0.1},
               'textFormat':{'bold':True,'foregroundColor':{'red':1,'green':1,'blue':1}}},
    'purple': {'backgroundColor': {'red':0.4,'green':0.1,'blue':0.5},
               'textFormat':{'bold':True,'foregroundColor':{'red':1,'green':1,'blue':1}}},
    'orange': {'backgroundColor': {'red':0.8,'green':0.4,'blue':0.0},
               'textFormat':{'bold':True,'foregroundColor':{'red':1,'green':1,'blue':1}}},
    'teal':   {'backgroundColor': {'red':0.0,'green':0.4,'blue':0.4},
               'textFormat':{'bold':True,'foregroundColor':{'red':1,'green':1,'blue':1}}},
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
            raise RuntimeError("No valid token.pickle.")
    return creds

# ═══════════════════════════════════════════════════════
#  SHEET HELPERS
# ═══════════════════════════════════════════════════════

def get_or_create_tab(spreadsheet, title, cols=15):
    try:
        ws = spreadsheet.worksheet(title)
        print(f"  📋 Reusing: {title}")
        return ws
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=title, rows=TAB_INIT_ROWS, cols=cols)
        print(f"  ➕ Created: {title}")
        return ws

def safe_clear(ws):
    ws.clear(); time.sleep(0.5)

def write_in_batches(ws, rows):
    total = len(rows)
    if not total:
        print("    (no rows)"); return
    for i in range(0, total, BATCH_SIZE):
        ws.append_rows(rows[i:i+BATCH_SIZE], value_input_option='USER_ENTERED')
        print(f"    ↳ {min(i+BATCH_SIZE,total):,} / {total:,}")
        time.sleep(1)

# ═══════════════════════════════════════════════════════
#  ROW CONVERTERS
# ═══════════════════════════════════════════════════════

def _mime(r):
    raw = r.get('Type', r.get('mimeType', ''))
    return MIME_LABELS.get(raw, raw.split('/')[-1] if raw else '')

def _erp_module(r):
    return ERP_MODULE_MAP.get(r.get('Category',''), r.get('Category',''))

def _suggested_folder(r):
    module  = _erp_module(r)
    mdt     = r.get('Master_Data_Type', '')
    cat_map = {
        'Quotation':    'Purchase / Procurement/Quotations',
        'Purchase Order':'Purchase / Procurement/POs',
        'Cost Sheet':   'Finance & Accounts/Cost Sheets',
        'Customer Master':'Sales / Dispatch/Customer Master',
        'Vendor Master':'Purchase / Procurement/Vendor Master',
        'Item Master':  'Production/Item Master',
        'HR Master':    'HR & Payroll/HR Master',
        'Sales Order':  'Sales / Dispatch/Orders',
        'Production':   'Production / Planning (PPC)/Batches',
        'Finance':      'Finance & Accounts/Reports',
    }
    if mdt in cat_map:
        return f"/ERP/{cat_map[mdt]}/"
    return f"/ERP/{module}/" if module else '/ERP/Other/'

def _suggested_name(r):
    """Suggest rename if file name is too generic."""
    name = r.get('Name', '')
    generic = ['sheet', 'copy of', 'untitled', 'new spreadsheet', 'book1',
               'document', 'presentation', 'form responses']
    if any(g in name.lower() for g in generic):
        mdt  = r.get('Master_Data_Type', '')
        cat  = r.get('Category', '')
        mod  = f.get('Modified', '')[:7] if (f := r) else ''
        hint = mdt or cat or 'Review'
        return f"{hint} — {mod} (rename needed)"
    return ''

def to_full_row(r):
    return [
        r.get('Name',''), r.get('Status',''), r.get('Score',0),
        r.get('ERP_Ready','NO'), r.get('Has_Form',''),
        r.get('Form_Status',''), r.get('Master_Data_Type',''),
        r.get('Category',''), _mime(r),
        r.get('Modified','')[:10], r.get('Owner',''),
        r.get('Dependencies',''), r.get('Folder_Path',''),
        r.get('Notes',''), r.get('Link',''),
    ]

def to_dup_row(r):
    return [
        r.get('Name',''), r.get('Status',''), r.get('Category',''),
        _mime(r), r.get('Modified','')[:10], r.get('Owner',''),
        r.get('Notes',''), r.get('Link',''),
    ]

def to_form_row(r):
    return [
        r.get('Name',''), r.get('Form_Status',''), r.get('Form_Detail',''),
        r.get('Score',0), r.get('ERP_Ready','NO'), r.get('Category',''),
        r.get('Modified','')[:10], r.get('Owner',''),
        r.get('Dependencies',''), r.get('Folder_Path',''), r.get('Link',''),
    ]

def to_att_row(r):
    suggested = f"/ERP/Forms/{r.get('Parent_Form_Sheet','Unknown')}/Attachments/"
    return [
        r.get('Name',''), _mime(r), r.get('Parent_Form_Sheet',''),
        suggested, r.get('Modified','')[:10], r.get('Owner',''),
        r.get('Folder_Path',''), r.get('Link',''),
    ]

def to_reorg_row(r):
    return [
        r.get('Name',''), r.get('Folder_Path',''),
        _suggested_folder(r), _suggested_name(r),
        _erp_module(r), r.get('Master_Data_Type',''), r.get('Link',''),
    ]

def to_master_row(r):
    return [
        r.get('Name',''), r.get('Master_Data_Type',''), r.get('Category',''),
        r.get('Score',0), r.get('ERP_Ready','NO'),
        r.get('Modified','')[:10], r.get('Owner',''),
        r.get('Folder_Path',''), r.get('Link',''),
    ]

# ═══════════════════════════════════════════════════════
#  TAB WRITERS
# ═══════════════════════════════════════════════════════

def write_summary(spreadsheet, files):
    total     = len(files)
    dups      = sum(1 for f in files if f.get('Status')          == 'DUPLICATE')
    erp       = sum(1 for f in files if f.get('ERP_Ready')       == 'YES')
    forms     = sum(1 for f in files if f.get('Has_Form')        == 'YES')
    active_f  = sum(1 for f in files if f.get('Form_Status')     == 'ACTIVE')
    inactive_f= sum(1 for f in files if f.get('Form_Status')     == 'INACTIVE')
    atts      = sum(1 for f in files if f.get('Is_Form_Attachment')=='YES')
    mdt       = sum(1 for f in files if f.get('Master_Data_Type'))
    deps      = sum(1 for f in files if f.get('Dependencies'))
    bloat     = round(dups/total*100,1) if total else 0

    cats = {}
    for f in files:
        c = f.get('Category','Other')
        cats[c] = cats.get(c,0)+1

    mdt_counts = {}
    for f in files:
        m = f.get('Master_Data_Type','')
        if m: mdt_counts[m] = mdt_counts.get(m,0)+1

    rows = [
        ['MIGRATION READINESS SUMMARY','',datetime.now().strftime('%Y-%m-%d %H:%M')],
        [''],
        ['METRIC','COUNT','ACTION'],
        ['Total files scanned',         total,    'Full inventory → separate sheet'],
        ['Duplicates found',            dups,     f'Dedup potential: {bloat}%'],
        ['ERP-ready (all criteria)',    erp,      'Migrate by April 1st'],
        ['  Form-linked sheets',        forms,    f'{active_f} active / {inactive_f} inactive'],
        ['  Connected hub files',       deps,     'IMPORTRANGE / QUERY dependencies'],
        ['  Form attachments',          atts,     'Images/PDFs from form uploads'],
        ['  Master data files',         mdt,      'Quotations, POs, Masters etc.'],
        ['Low value / archive',         total-erp-dups, 'Move to 2025 Archive'],
        [''],
        ['CATEGORY BREAKDOWN','FILES',''],
    ] + [[f'  {c}',n,''] for c,n in sorted(cats.items(),key=lambda x:-x[1])] + [
        [''],
        ['MASTER DATA TYPES FOUND','FILES',''],
    ] + [[f'  {m}',n,''] for m,n in sorted(mdt_counts.items(),key=lambda x:-x[1])]

    print('\nWriting Summary...')
    ws = get_or_create_tab(spreadsheet,'📊 Summary',cols=3)
    safe_clear(ws)
    ws.append_rows(rows,value_input_option='USER_ENTERED')
    ws.format('A1:C1',FMT['blue'])
    ws.format('A3:C3',FMT['blue'])
    ws.format('A13:C13',FMT['blue'])


def write_erp_priority(spreadsheet, files):
    priority = sorted(
        [f for f in files if f.get('ERP_Ready')=='YES'],
        key=lambda x:(x.get('Has_Form')=='YES', bool(x.get('Dependencies')),
                      x.get('Score',0)),
        reverse=True,
    )
    print(f'\nWriting ERP Priority ({len(priority):,})...')
    ws = get_or_create_tab(spreadsheet,'🚀 ERP Migration List',cols=15)
    safe_clear(ws)
    ws.append_row(FULL_HEADERS)
    ws.format('A1:O1',FMT['green'])
    write_in_batches(ws,[to_full_row(f) for f in priority])


def write_form_sheets(spreadsheet, files):
    form_files = [f for f in files if f.get('Has_Form')=='YES']
    active   = [f for f in form_files if f.get('Form_Status')=='ACTIVE']
    inactive = [f for f in form_files if f.get('Form_Status')=='INACTIVE']
    ordered  = active + inactive
    print(f'\nWriting Form-linked Sheets ({len(ordered):,} — {len(active)} active / {len(inactive)} inactive)...')
    ws = get_or_create_tab(spreadsheet,'📋 Form-linked Sheets',cols=11)
    safe_clear(ws)
    ws.append_row(FORM_HEADERS)
    ws.format('A1:K1',FMT['purple'])
    write_in_batches(ws,[to_form_row(f) for f in ordered])


def write_form_attachments(spreadsheet, files):
    atts = [f for f in files if f.get('Is_Form_Attachment')=='YES']
    print(f'\nWriting Form Attachments ({len(atts):,})...')
    ws = get_or_create_tab(spreadsheet,'📎 Form Attachments',cols=8)
    safe_clear(ws)
    ws.append_row(ATT_HEADERS)
    ws.format('A1:H1',FMT['orange'])
    write_in_batches(ws,[to_att_row(f) for f in atts])


def write_drive_reorg(spreadsheet, files):
    reorg = [f for f in files
             if f.get('ERP_Ready')=='YES' and f.get('Status')=='ORIGINAL']
    print(f'\nWriting Drive Reorg Plan ({len(reorg):,})...')
    ws = get_or_create_tab(spreadsheet,'📂 Drive Reorg Plan',cols=7)
    safe_clear(ws)
    ws.append_row(REORG_HEADERS)
    ws.format('A1:G1',FMT['teal'])
    write_in_batches(ws,[to_reorg_row(f) for f in reorg])


def write_master_data_map(spreadsheet, files):
    masters = [f for f in files if f.get('Master_Data_Type')]
    masters.sort(key=lambda x:(x.get('Master_Data_Type',''),
                               x.get('Category','')))
    print(f'\nWriting Master Data Map ({len(masters):,})...')
    ws = get_or_create_tab(spreadsheet,'🗂 Master Data Map',cols=9)
    safe_clear(ws)
    ws.append_row(MASTER_HEADERS)
    ws.format('A1:I1',FMT['teal'])
    write_in_batches(ws,[to_master_row(f) for f in masters])


def write_duplicates(spreadsheet, files):
    dupes = [f for f in files if f.get('Status')=='DUPLICATE']
    print(f'\nWriting Duplicates ({len(dupes):,})...')
    ws = get_or_create_tab(spreadsheet,'🔁 Duplicates',cols=8)
    safe_clear(ws)
    ws.append_row(DUP_HEADERS)
    ws.format('A1:H1',FMT['red'])
    write_in_batches(ws,[to_dup_row(f) for f in dupes])


def write_full_inventory(gc, files):
    if not INVENTORY_SHEET_ID:
        print('\nℹ️  INVENTORY_SHEET_ID not set — skipping full inventory.')
        return
    print(f'\nWriting Full Inventory ({len(files):,}) → separate sheet...')
    try:
        inv = gc.open_by_key(INVENTORY_SHEET_ID)
    except gspread.exceptions.SpreadsheetNotFound:
        print(f'   ❌ Inventory sheet not found.')
        return
    ws = get_or_create_tab(inv,'Full Inventory',cols=15)
    safe_clear(ws)
    ws.append_row(FULL_HEADERS)
    ws.format('A1:O1',FMT['blue'])
    write_in_batches(ws,[to_full_row(f) for f in files])
    print('   ✅ Full inventory written.')

# ═══════════════════════════════════════════════════════
#  NOTIFICATION
# ═══════════════════════════════════════════════════════

def notify(message, status='success'):
    import urllib.request
    webhook = os.environ.get('SLACK_WEBHOOK','')
    if not webhook: return
    icon    = '✅' if status=='success' else '❌'
    payload = json.dumps({'text':f'{icon} Drive Audit — {message}'}).encode()
    try:
        req = urllib.request.Request(webhook,data=payload,
              headers={'Content-Type':'application/json'})
        urllib.request.urlopen(req,timeout=10)
    except Exception as e:
        print(f'⚠ Slack failed: {e}')

# ═══════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════

def main():
    print('='*60)
    print('   STAGE 2 — Report Writer v5')

    if not os.path.exists(DATA_FILE):
        raise SystemExit(f"❌ {DATA_FILE} not found.")

    with open(DATA_FILE) as f:
        all_files = json.load(f)
    print(f'📊 Loaded {len(all_files):,} files')

    creds = authenticate()
    gc    = gspread.authorize(creds)

    try:
        spreadsheet = gc.open_by_key(SHEET_ID)
        print(f'Writing to: {spreadsheet.title}')
        print('='*60)

        write_summary(spreadsheet, all_files)
        write_erp_priority(spreadsheet, all_files)
        write_form_sheets(spreadsheet, all_files)
        write_form_attachments(spreadsheet, all_files)
        write_drive_reorg(spreadsheet, all_files)
        write_master_data_map(spreadsheet, all_files)
        write_duplicates(spreadsheet, all_files)
        write_full_inventory(gc, all_files)

        total  = len(all_files)
        erp    = sum(1 for f in all_files if f.get('ERP_Ready')=='YES')
        dups   = sum(1 for f in all_files if f.get('Status')=='DUPLICATE')
        forms  = sum(1 for f in all_files if f.get('Has_Form')=='YES')
        msg    = (f"{total:,} files | {erp:,} ERP-ready | "
                  f"{forms:,} form-linked | {dups:,} duplicates | "
                  f"{datetime.now().strftime('%Y-%m-%d')}")

        print(f'\n✅ Done — {msg}')
        notify(msg,'success')

    except Exception as e:
        print(f'\n❌ Failed: {e}')
        notify(f'FAILED: {e}','error')
        raise


if __name__ == '__main__':
    main()
