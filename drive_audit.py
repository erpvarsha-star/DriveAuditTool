"""
STAGE 1 — DRIVE AUDIT ENGINE (v5)

Changes:
  [FIX-1]  Form activity: ACTIVE/INACTIVE based on any of 4 signals
           - Sheet modified < 60 days
           - Sheet modified < 180 days  
           - Last form response < 60 days
           - Last form response < 180 days
  [FIX-2]  Form attachments: images/PDFs in same folder as form sheet
           flagged as FORM_ATTACHMENT and linked to parent form
  [FIX-3]  Master data pre-detection via keywords before AI sees it
           Quotation, RFQ, PO, Cost Sheet, Vendor, Customer, Item, BOM, HR
  [FIX-4]  Dependency files auto-qualify as ERP-ready
  [FIX-5]  Form-linked sheets auto-qualify as ERP-ready
  [FIX-6]  Real clickable Drive URLs for every file
  [FIX-7]  Has_Form, Form_Status, Master_Data_Type columns added
"""

import os
import json
import pickle
import time
import csv
from datetime import datetime, timezone

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ═══════════════════════════════════════════════════════
#  CONFIG
# ═══════════════════════════════════════════════════════

TOKEN_FILE    = 'token.pickle'
DATA_FILE     = 'audit_data.json'
HASHES_FILE   = 'seen_hashes.json'
SCANNED_FILE  = 'scanned_folders.json'
REPORT_FILE   = 'ERP_Migration_Final_Report.csv'

SCOPES = [
    'https://www.googleapis.com/auth/drive.readonly',
    'https://www.googleapis.com/auth/spreadsheets.readonly',
]

SHEETS_CALL_DELAY = 0.5
ACTIVE_DAYS_SHORT = 60
ACTIVE_DAYS_LONG  = 180

# Drive URL templates
DRIVE_URL = {
    'application/vnd.google-apps.spreadsheet':  'https://docs.google.com/spreadsheets/d/{id}/edit',
    'application/vnd.google-apps.document':     'https://docs.google.com/document/d/{id}/edit',
    'application/vnd.google-apps.presentation': 'https://docs.google.com/presentation/d/{id}/edit',
    'application/vnd.google-apps.form':         'https://docs.google.com/forms/d/{id}/edit',
}
DRIVE_URL_DEFAULT = 'https://drive.google.com/file/d/{id}/view'

# Category keywords
CATEGORY_KEYWORDS = {
    'Plant Operations': ['plant', 'production', 'ppc', 'dispatch',
                         'press shop', 'hammer', 'die shop', 'stores'],
    'Quality & QMS':   ['qms', 'quality', 'control plan', 'audit',
                         'iso', 'iatf', 'calibration', 'vfqa'],
    'HR & Admin':      ['hr', 'salary', 'leave', 'skill matrix',
                         'attendance', 'payroll', 'recruit', 'joining'],
    'Commercial':      ['purchase', 'procurement', 'rfq', 'quotation',
                         'vendor', 'po', 'quote', 'cost', 'price'],
    'Finance':         ['accounts', 'ledger', 'balance sheet',
                         'gst', '2025-2026', '2024-2025', 'invoice',
                         'expense', 'investment'],
    'Sales':           ['sales', 'order', 'customer', 'dispatch',
                         'delivery', 'revenue', 'invoice'],
    'Production':      ['forging', 'batch', 'machine', 'vf', 'shift',
                         'operator', 'scrap', 'rejection', 'bom'],
}

# Master data type keywords [FIX-3]
MASTER_DATA_KEYWORDS = {
    'Quotation':       ['quotation', 'quote', 'rfq', 'enquiry', 'enq'],
    'Purchase Order':  ['purchase order', ' po ', 'p.o.', 'po no', 'po#'],
    'Cost Sheet':      ['cost sheet', 'costing', 'cost analysis', 'rate analysis'],
    'Customer Master': ['customer master', 'customer list', 'client master'],
    'Vendor Master':   ['vendor master', 'supplier master', 'vendor list'],
    'Item Master':     ['item master', 'part master', 'material master', 'bom'],
    'HR Master':       ['employee master', 'manpower', 'headcount', 'salary master',
                        'skill matrix', 'recruiting', 'joining'],
    'Sales Order':     ['sales order', 'so no', 'customer order', 'dispatch order'],
    'Production':      ['production plan', 'batch card', 'forge batch', 'job card'],
    'Finance':         ['balance sheet', 'p&l', 'profit loss', 'trial balance',
                        'cash flow', 'expense sheet'],
}

# Image/PDF mime types that could be form attachments [FIX-2]
ATTACHMENT_MIMES = {
    'image/jpeg', 'image/png', 'image/gif', 'image/webp',
    'application/pdf',
    'image/heic', 'image/heif',
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
            print("ERROR: No valid token.pickle found.")
            return None
    return creds

# ═══════════════════════════════════════════════════════
#  HELPERS
# ═══════════════════════════════════════════════════════

def make_link(file_id, mime):
    template = DRIVE_URL.get(mime, DRIVE_URL_DEFAULT)
    return template.format(id=file_id)


def days_since(date_str):
    if not date_str:
        return 9999
    try:
        dt = datetime.fromisoformat(date_str.replace('Z', '+00:00'))
        return (datetime.now(timezone.utc) - dt).days
    except Exception:
        return 9999


def detect_master_data_type(name, folder_path):
    """[FIX-3] Detect master data type from file name and folder path."""
    text = (name + ' ' + folder_path).lower()
    for dtype, keywords in MASTER_DATA_KEYWORDS.items():
        if any(k in text for k in keywords):
            return dtype
    return ''

# ═══════════════════════════════════════════════════════
#  FORM INTELLIGENCE  [FIX-1]
# ═══════════════════════════════════════════════════════

def get_last_response_date(sheets_svc, file_id, form_tab_names):
    """Read timestamp column from Form Response tabs to find last response."""
    latest = ''
    for tab in form_tab_names[:3]:
        try:
            resp = sheets_svc.spreadsheets().values().get(
                spreadsheetId=file_id,
                range=f"'{tab}'!A2:A500"
            ).execute()
            time.sleep(SHEETS_CALL_DELAY)
            values = resp.get('values', [])
            for row in reversed(values):
                if row and row[0].strip():
                    for fmt in ['%d/%m/%Y %H:%M:%S', '%m/%d/%Y %H:%M:%S',
                                '%Y-%m-%d %H:%M:%S', '%d-%m-%Y %H:%M:%S',
                                '%d/%m/%Y', '%m/%d/%Y']:
                        try:
                            dt = datetime.strptime(row[0].strip(), fmt)
                            ts = dt.isoformat()
                            if not latest or ts > latest:
                                latest = ts
                            break
                        except ValueError:
                            continue
                    if latest:
                        break
        except Exception:
            pass
    return latest


def check_form_activity(modified_date, last_response_date):
    """
    Returns (is_active, status_string, signals_list)
    Any one signal = ACTIVE. All absent = INACTIVE.
    """
    signals    = []
    sheet_age  = days_since(modified_date)
    resp_age   = days_since(last_response_date)

    if sheet_age  <= ACTIVE_DAYS_SHORT: signals.append(f'Sheet modified <{ACTIVE_DAYS_SHORT}d')
    if sheet_age  <= ACTIVE_DAYS_LONG:  signals.append(f'Sheet modified <{ACTIVE_DAYS_LONG}d')
    if resp_age   <= ACTIVE_DAYS_SHORT: signals.append(f'Response <{ACTIVE_DAYS_SHORT}d')
    if resp_age   <= ACTIVE_DAYS_LONG:  signals.append(f'Response <{ACTIVE_DAYS_LONG}d')

    is_active = len(signals) > 0
    status    = 'ACTIVE' if is_active else 'INACTIVE'
    detail    = ' | '.join(signals) if signals else 'No activity in 6+ months — consider reactivating'
    return is_active, status, detail


def get_sheet_dependencies(sheets_svc, file_id, modified_date):
    """
    Scan all tabs for IMPORTRANGE/QUERY and form links.
    Returns: (deps_list, has_form, form_status, form_detail, form_tab_names)
    """
    found          = set()
    has_form       = False
    form_tab_names = []

    try:
        meta = sheets_svc.spreadsheets().get(
            spreadsheetId=file_id,
            fields='sheets.properties.title,sheets.properties.sheetType'
        ).execute()
        time.sleep(SHEETS_CALL_DELAY)

        sheets = meta.get('sheets', [])
        for s in sheets:
            props = s.get('properties', {})
            title = props.get('title', '')
            stype = props.get('sheetType', '')
            if 'form response' in title.lower() or stype == 'OBJECT':
                has_form = True
                form_tab_names.append(title)

        tabs = [s['properties']['title'] for s in sheets]
        for tab in tabs[:10]:
            try:
                resp = sheets_svc.spreadsheets().values().get(
                    spreadsheetId=file_id,
                    range=f"'{tab}'!A1:Z100"
                ).execute()
                time.sleep(SHEETS_CALL_DELAY)
                flat = str(resp.get('values', [])).lower()
                if 'importrange' in flat: found.add('IMPORTRANGE')
                if 'query('      in flat: found.add('QUERY')
            except HttpError as e:
                if e.resp.status == 429:
                    print(f"  ⏳ Quota — waiting 60s")
                    time.sleep(60)
            except Exception:
                pass

    except HttpError as e:
        if e.resp.status == 429:
            print(f"  ⏳ Quota meta — waiting 60s")
            time.sleep(60)
        elif e.resp.status != 400:
            print(f"  ⚠ Sheet error ({e.resp.status}) {file_id}")
    except Exception as e:
        print(f"  ⚠ Sheet error {file_id}: {e}")

    # Get form activity if form detected [FIX-1]
    form_status = form_detail = ''
    if has_form:
        last_resp = get_last_response_date(sheets_svc, file_id, form_tab_names)
        _, form_status, form_detail = check_form_activity(modified_date, last_resp)

    return list(found), has_form, form_status, form_detail

# ═══════════════════════════════════════════════════════
#  SCORING ENGINE
# ═══════════════════════════════════════════════════════

def calculate_utility(name, modified_date, deps, category, has_form=False):
    score    = 0
    name_low = name.lower()

    # Recency
    if days_since(modified_date) < 60:  score += 35

    # Name keywords
    if any(k in name_low for k in ['master', 'database', 'final', '2026', '25-26']):
        score += 30

    # Dependencies — raised to auto-qualify
    if deps:     score += 35

    # Form linked — active data collection
    if has_form: score += 40

    # Domain priority
    if category in ['Plant Operations', 'Finance']:
        score += 10

    return min(score, 100)


def is_erp_ready(score, deps, has_form):
    if has_form: return True   # always — active or inactive, form = data collection
    if deps:     return True   # hub file other sheets depend on
    return score > 65

# ═══════════════════════════════════════════════════════
#  PERSISTENCE
# ═══════════════════════════════════════════════════════

def save_progress(all_files, seen_hashes, scanned_folders):
    with open(DATA_FILE,    'w') as f: json.dump(all_files,             f)
    with open(HASHES_FILE,  'w') as f: json.dump(seen_hashes,           f)
    with open(SCANNED_FILE, 'w') as f: json.dump(list(scanned_folders), f)


def load_progress():
    all_files = []; seen_hashes = {}; scanned_folders = set()
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE)    as f: all_files       = json.load(f)
        print(f"♻️  Resuming: {len(all_files):,} files loaded.")
    if os.path.exists(HASHES_FILE):
        with open(HASHES_FILE)  as f: seen_hashes     = json.load(f)
        print(f"♻️  Resuming: {len(seen_hashes):,} fingerprints loaded.")
    if os.path.exists(SCANNED_FILE):
        with open(SCANNED_FILE) as f: scanned_folders = set(json.load(f))
        print(f"♻️  Resuming: {len(scanned_folders):,} folders already scanned.")
    return all_files, seen_hashes, scanned_folders

# ═══════════════════════════════════════════════════════
#  FOLDER CONTEXT — track form sheets per folder [FIX-2]
# ═══════════════════════════════════════════════════════

# Map folder_id → list of form-linked sheet names in that folder
# Used to detect if an image/PDF in the same folder is a form attachment
form_folders = {}   # folder_id: [sheet_name, ...]

# ═══════════════════════════════════════════════════════
#  RECURSIVE CORE
# ═══════════════════════════════════════════════════════

def scan_folder(drive_svc, sheets_svc, folder_id, folder_path,
                all_files, seen_hashes, scanned_folders, depth=0):

    if folder_id in scanned_folders:
        return

    try:
        # First pass — collect all files in this folder
        folder_files = []
        page_token   = None
        while True:
            resp = drive_svc.files().list(
                q=f"'{folder_id}' in parents and trashed=false",
                fields='nextPageToken, files(id,name,mimeType,size,modifiedTime,owners,md5Checksum)',
                pageSize=100,
                pageToken=page_token
            ).execute()
            folder_files.extend(resp.get('files', []))
            page_token = resp.get('nextPageToken')
            if not page_token:
                break

        # Identify form-linked sheets in this folder [FIX-2]
        form_sheet_names = set()

        # Process each file
        for f in folder_files:
            fid      = f['id']
            name     = f['name']
            mime     = f['mimeType']

            # Recurse into sub-folders
            if mime == 'application/vnd.google-apps.folder':
                scan_folder(drive_svc, sheets_svc, fid,
                            f"{folder_path} > {name}",
                            all_files, seen_hashes, scanned_folders, depth + 1)
                continue

            # ── Duplicate detection ──────────────────────────
            checksum    = f.get('md5Checksum')
            fingerprint = checksum if checksum else f"{name}_{f.get('size',0)}"
            if fingerprint in seen_hashes:
                status    = "DUPLICATE"
                dup_notes = f"Master in: {seen_hashes[fingerprint]['path']}"
            else:
                status    = "ORIGINAL"
                dup_notes = ""
                seen_hashes[fingerprint] = {'name': name, 'path': folder_path}

            # ── Categorise ───────────────────────────────────
            category = "Other"
            for cat, keys in CATEGORY_KEYWORDS.items():
                if any(k in (name + folder_path).lower() for k in keys):
                    category = cat
                    break

            # ── Master data type detection [FIX-3] ───────────
            master_data_type = detect_master_data_type(name, folder_path)

            # ── Deep scan for spreadsheets ────────────────────
            deps = []; has_form = False; form_status = ''; form_detail = ''
            if status == "ORIGINAL" and 'spreadsheet' in mime:
                deps, has_form, form_status, form_detail = get_sheet_dependencies(
                    sheets_svc, fid, f.get('modifiedTime', '')
                )
                if has_form:
                    form_sheet_names.add(name)

            # ── Check if this is a form attachment [FIX-2] ────
            is_form_attachment = False
            parent_form_sheet  = ''
            if mime in ATTACHMENT_MIMES and status == "ORIGINAL":
                # If there are form sheets in this folder, this attachment belongs to them
                if form_sheet_names:
                    is_form_attachment = True
                    parent_form_sheet  = ', '.join(form_sheet_names)

            # ── Score ─────────────────────────────────────────
            score     = calculate_utility(name, f.get('modifiedTime',''), deps, category, has_form) if status == "ORIGINAL" else 0
            erp_ready = is_erp_ready(score, deps, has_form) if status == "ORIGINAL" else False

            # ── Override: form attachments are ERP-ready ──────
            if is_form_attachment:
                erp_ready = True

            # ── Build link ────────────────────────────────────
            link = make_link(fid, mime)

            all_files.append({
                'Name':              name,
                'Category':          category,
                'Status':            status,
                'Score':             score,
                'Type':              mime,
                'Has_Form':          'YES' if has_form else '',
                'Form_Status':       form_status,
                'Form_Detail':       form_detail,
                'Is_Form_Attachment':  'YES' if is_form_attachment else '',
                'Parent_Form_Sheet': parent_form_sheet,
                'Master_Data_Type':  master_data_type,
                'Folder_Path':       folder_path,
                'Dependencies':      ', '.join(deps),
                'Modified':          f.get('modifiedTime', '')[:10],
                'Owner':             (f.get('owners') or [{}])[0].get('emailAddress', ''),
                'Notes':             dup_notes,
                'ERP_Ready':         'YES' if erp_ready else 'NO',
                'Link':              link,
            })

            icon      = '✅' if status == 'ORIGINAL' else '🔁'
            form_flag = ' 📋FORM'   if has_form          else ''
            dep_flag  = ' 🔗DEP'    if deps               else ''
            att_flag  = ' 📎ATT'    if is_form_attachment else ''
            mdt_flag  = f' [{master_data_type}]' if master_data_type else ''
            print(f"{'  '*depth}{icon} {score:02}{form_flag}{dep_flag}{att_flag}{mdt_flag} | {name[:50]}")

        scanned_folders.add(folder_id)
        save_progress(all_files, seen_hashes, scanned_folders)

    except HttpError as e:
        print(f"Drive API error in '{folder_path}': {e}")
    except Exception as e:
        print(f"Unexpected error in '{folder_path}': {e}")

# ═══════════════════════════════════════════════════════
#  CSV EXPORT
# ═══════════════════════════════════════════════════════

def export_csv(all_files):
    if not all_files:
        print("⚠ No files to export.")
        return
    with open(REPORT_FILE, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=all_files[0].keys())
        writer.writeheader()
        writer.writerows(all_files)
    print(f"📄 CSV saved → {REPORT_FILE}")

# ═══════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════

def main():
    print("🔐 Authenticating...")
    creds = authenticate()
    if not creds:
        raise SystemExit("Authentication failed.")

    drive_svc  = build('drive',  'v3', credentials=creds)
    sheets_svc = build('sheets', 'v4', credentials=creds)

    all_files, seen_hashes, scanned_folders = load_progress()

    print("🚀 Starting Drive Audit v5...")

    results = drive_svc.files().list(
        q="sharedWithMe=true and mimeType='application/vnd.google-apps.folder'",
        fields='files(id,name)'
    ).execute()

    for folder in results.get('files', []):
        print(f"\n📁 {folder['name']}")
        scan_folder(drive_svc, sheets_svc,
                    folder['id'], folder['name'],
                    all_files, seen_hashes, scanned_folders)

    export_csv(all_files)

    total = len(all_files)
    dups  = sum(1 for f in all_files if f['Status']             == 'DUPLICATE')
    ready = sum(1 for f in all_files if f['ERP_Ready']          == 'YES')
    forms = sum(1 for f in all_files if f.get('Has_Form')       == 'YES')
    atts  = sum(1 for f in all_files if f.get('Is_Form_Attachment') == 'YES')
    mdt   = sum(1 for f in all_files if f.get('Master_Data_Type'))
    print(f"\n✅ Done — {total:,} files | {dups:,} duplicates | {ready:,} ERP-ready")
    print(f"   {forms:,} form-linked | {atts:,} form attachments | {mdt:,} master data files")


if __name__ == '__main__':
    main()
