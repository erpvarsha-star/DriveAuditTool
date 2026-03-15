"""
STAGE 1 — DRIVE SCANNER
Scans all shared folders recursively.
Saves every file instantly to audit_data.json.
Can resume if interrupted.
Run: python drive_audit.py
"""

import os
import json
import pickle
import time
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from datetime import datetime, timezone

# ═══════════════════════════════════════════════════════
#  SETTINGS
# ═══════════════════════════════════════════════════════
CREDENTIALS_FILE  = 'credentials.json'
TOKEN_FILE        = 'token.pickle'
DATA_FILE         = 'audit_data.json'      # All scanned data saved here
SCANNED_FILE      = 'scanned_folders.json' # Tracks which folders already done
INACTIVE_MONTHS   = 6                      # Files not modified in 6+ months = Inactive

SCOPES = [
    'https://www.googleapis.com/auth/drive.readonly',
    'https://www.googleapis.com/auth/documents.readonly',
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.file'
]

# ═══════════════════════════════════════════════════════
#  PERSONAL FOLDERS → Personal tab
# ═══════════════════════════════════════════════════════
PERSONAL_FOLDERS = [
    'photos', 'photo', 'pictures', 'pics',
    'neha', 'cvs', 'cv', 'ids', 'id card',
    'healthcraft', 'eo', 'personal', 'private', 'family'
]

# ═══════════════════════════════════════════════════════
#  COMPANY — ORDER MATTERS (Stego/Prima checked BEFORE Varsha)
# ═══════════════════════════════════════════════════════
COMPANY_KEYWORDS = {
    'Stego': ['stego', 'stg'],
    'Prima': ['prima', 'prm'],
    'Varsha (VFL/VFPL)': [
        'vfl', 'vfpl', 'varsha', 'vfqa',
        'forgetech', 'prompt',
        'press shop', 'hammer shop', 'die shop',
        'cutting shop', 'heat treatment', 'forge',
        'ppc', 'qms', 'stores', 'marathi',
        'naresh', 'kavadu', 'lata', 'subhash',
        'master', 'yash', 'management', 'purchase',
        'bung', 'annual dinner', 'greeting',
        'appsheet', 'delegation', 'authority',
        'controlled', 'interactive', 'instant form',
        'form data', 'box sheet', 'project management',
        'die shop', 'hammer', 'level 1', 'level 2', 'level 3',
        'salary', 'hr ', 'leave', 'skill matrix',
        'part drawing', 'presentation'
    ]
}

# ═══════════════════════════════════════════════════════
#  PROJECT — "master" in filename = Master project
# ═══════════════════════════════════════════════════════
PROJECT_KEYWORDS = {
    'RFQ': [
        'rfq', 'request for quotation', 'request for quote',
        'rfq-', 'rfq_', ' rfq '
    ],
    'Forgetech': [
        'forgetech', 'forge tech', 'forgetech 2023',
        'booking confirmation dme', 'ftl'
    ],
    'Prompt': ['prompt', 'prmt'],
    'Master': [
        'master'  # ANY file with master in name
    ]
}

# ═══════════════════════════════════════════════════════
#  CATEGORY
# ═══════════════════════════════════════════════════════
CATEGORY_KEYWORDS = {
    'Plant Operations': [
        'plant', 'production', 'ppc', 'operations',
        'press shop', 'press', 'hammer shop', 'hammer',
        'die shop', 'die', 'cutting shop', 'cutting',
        'forge shop', 'forging', 'heat treatment',
        'lab', 'laboratory', 'stores', 'store',
        'dispatch', 'maintenance', 'machine', 'equipment',
        'shift', 'downtime', 'shop floor', 'final shop',
        'departmentwise', 'department', 'bung',
        'project management', 'part drawing'
    ],
    'Quality & QMS': [
        'qms', 'quality', 'control plan', 'controlled copy',
        'controlled', 'level 1', 'level 2', 'level 3',
        'level1', 'level2', 'level3', 'manual',
        'procedure', 'procedures', 'format', 'formats',
        'inspection', 'audit', 'sop', 'work instruction',
        'ncr', 'corrective', 'preventive', 'iso', 'iatf',
        'calibration', 'document control', 'vfqa',
        'interactive table', 'control copy'
    ],
    'HR & Admin': [
        'hr', 'human resource', 'salary', 'leave',
        'skill matrix', 'skill matric', 'evaluation',
        'authority letter', 'authority', 'delegation',
        'annual dinner', 'joining', 'resignation',
        'attendance', 'payroll', 'employee', 'staff',
        'recruitment', 'training', 'appraisal',
        'greeting', 'greetings', 'marathi', 'f-hr',
        'admin', 'ceo', 'management', 'web script',
        'webscript', 'dont delete', 'presentation', 'cv'
    ],
    'Commercial & Purchase': [
        'purchase', 'purchasing', 'procurement',
        'crm', 'customer', 'marketing', 'market',
        'booking', 'confirmation', 'invoice',
        'quotation', 'quote', 'rfq', 'proforma',
        'price list', 'pricing', 'offer', 'proposal',
        'bid', 'tender', 'supplier', 'vendor',
        'po', 'purchase order', 'contract', 'agreement',
        'payment', 'receipt', 'bill', 'advance',
        'development', 'developement'
    ],
    'Finance & Accounts': [
        'salary sheet', 'accounts', 'finance', 'ledger',
        'balance sheet', 'profit', 'loss', 'tax', 'gst',
        'annual', 'yearly', 'monthly',
        '2022-2023', '2023-2024', '2024-2025', '2025-2026'
    ],
    'IT & Digital Tools': [
        'appsheet', 'web script', 'webscript', 'form data',
        'instant form', 'interactive', 'box sheet',
        'database', 'format required', 'script',
        'automation', 'tool', 'updator', 'form links', 'app'
    ]
}

# ═══════════════════════════════════════════════════════
#  FILE TYPE LABELS
# ═══════════════════════════════════════════════════════
MIME_LABELS = {
    'application/vnd.google-apps.spreadsheet':  'Google Sheet',
    'application/vnd.google-apps.document':     'Google Doc',
    'application/vnd.google-apps.presentation': 'Google Slides',
    'application/vnd.google-apps.form':         'Google Form',
    'application/vnd.google-apps.folder':       'Folder',
    'application/vnd.google-apps.script':       'Apps Script',
    'application/vnd.google-apps.drawing':      'Google Drawing',
    'application/pdf':                          'PDF',
    'application/vnd.ms-excel':                 'Excel (.xls)',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'Excel (.xlsx)',
    'application/vnd.ms-powerpoint':            'PowerPoint (.ppt)',
    'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'PowerPoint (.pptx)',
    'application/msword':                       'Word (.doc)',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'Word (.docx)',
    'text/plain':                               'Text File',
    'text/csv':                                 'CSV',
    'image/jpeg':                               'Image (JPG)',
    'image/png':                                'Image (PNG)',
    'image/gif':                                'Image (GIF)',
    'image/bmp':                                'Image (BMP)',
    'image/svg+xml':                            'SVG',
    'video/mp4':                                'Video (MP4)',
    'audio/mpeg':                               'Audio (MP3)',
    'application/zip':                          'ZIP Archive',
}

NO_SIZE_MIMES = {
    'application/vnd.google-apps.spreadsheet',
    'application/vnd.google-apps.document',
    'application/vnd.google-apps.presentation',
    'application/vnd.google-apps.form',
    'application/vnd.google-apps.folder',
    'application/vnd.google-apps.script',
    'application/vnd.google-apps.drawing',
}


# ═══════════════════════════════════════════════════════
#  AUTHENTICATION
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
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'wb') as f:
            pickle.dump(creds, f)
    return creds


# ═══════════════════════════════════════════════════════
#  HELPERS
# ═══════════════════════════════════════════════════════
def is_personal(folder_path):
    path_lower = folder_path.lower()
    return any(pf in path_lower for pf in PERSONAL_FOLDERS)

def detect_company(name, folder_path):
    text = (name + ' ' + folder_path).lower()
    for company, keywords in COMPANY_KEYWORDS.items():
        if any(kw in text for kw in keywords):
            return company
    return 'Unknown'

def detect_project(name, folder_path):
    text = (name + ' ' + folder_path).lower()
    for project, keywords in PROJECT_KEYWORDS.items():
        if any(kw in text for kw in keywords):
            return project
    return 'General'

def classify_category(name, folder_path):
    text = (name + ' ' + folder_path).lower()
    for category, keywords in CATEGORY_KEYWORDS.items():
        if any(kw in text for kw in keywords):
            return category
    return 'Other / Uncategorized'

def get_file_type(mime):
    return MIME_LABELS.get(mime, mime.split('.')[-1].upper() if '.' in mime else 'Unknown')

def format_size(size_str, mime):
    if mime in NO_SIZE_MIMES:
        return 'N/A (Google File)' if mime != 'application/vnd.google-apps.folder' else 'N/A (Folder)'
    try:
        size = int(size_str)
        if size < 1024:       return f'{size} B'
        elif size < 1024**2:  return f'{size/1024:.1f} KB'
        elif size < 1024**3:  return f'{size/1024**2:.1f} MB'
        else:                 return f'{size/1024**3:.1f} GB'
    except:
        return 'N/A'

def is_inactive(modified_date_str):
    try:
        modified = datetime.fromisoformat(modified_date_str.replace('Z', '+00:00'))
        now      = datetime.now(timezone.utc)
        months   = (now - modified).days / 30
        return months >= INACTIVE_MONTHS
    except:
        return False

def detect_sheet_role(name):
    """Detect if a sheet is a form, master, report or linked sheet"""
    n = name.lower()
    if any(k in n for k in ['form', 'response', 'appsheet', 'instant form', 'form data', 'updator']):
        return 'Form / Input Sheet'
    if any(k in n for k in ['master', 'database', 'master data', 'source', 'raw data']):
        return 'Master / Source Data'
    if any(k in n for k in ['report', 'summary', 'dashboard', 'tracker', 'consolidated',
                             'analysis', 'interactive', 'crm', 'salary sheet',
                             'skill matrix', 'control plan']):
        return 'Report / Linked Sheet'
    if any(k in n for k in ['download', 'export', 'backup', 'archive', 'controlled copy']):
        return 'Export / Backup'
    return ''

def get_sheet_dependencies(sheets_service, file_id):
    """
    Check if a Google Sheet has IMPORTRANGE, QUERY, or form links.
    Returns list of dependency types found.
    """
    deps = []
    try:
        result = sheets_service.spreadsheets().get(
            spreadsheetId=file_id,
            fields='sheets.properties.title'
        ).execute()
        sheet_names = [s['properties']['title'] for s in result.get('sheets', [])]

        # Check cell values for formulas in first sheet
        if sheet_names:
            values = sheets_service.spreadsheets().values().get(
                spreadsheetId=file_id,
                range=f"'{sheet_names[0]}'!A1:Z100"
            ).execute().get('values', [])

            flat = ' '.join(str(cell) for row in values for cell in row).lower()
            if 'importrange' in flat:  deps.append('IMPORTRANGE')
            if 'query('      in flat:  deps.append('QUERY')
            if 'importdata'  in flat:  deps.append('IMPORTDATA')
            if 'form'        in flat:  deps.append('Form Linked')

        # Check if it has a linked form
        meta = sheets_service.spreadsheets().get(
            spreadsheetId=file_id,
            fields='sheets.properties'
        ).execute()
        for sheet in meta.get('sheets', []):
            if sheet.get('properties', {}).get('sheetType') == 'OBJECT':
                deps.append('Embedded Object')

    except:
        pass
    return deps


# ═══════════════════════════════════════════════════════
#  SAVE / LOAD JSON
# ═══════════════════════════════════════════════════════
def save_data(all_files, seen_checksums):
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump({
            'files':     all_files,
            'checksums': seen_checksums,
            'saved_at':  datetime.now().isoformat()
        }, f, ensure_ascii=False, indent=2)

def load_scanned_folders():
    if os.path.exists(SCANNED_FILE):
        with open(SCANNED_FILE, 'r') as f:
            return set(json.load(f))
    return set()

def save_scanned_folder(folder_id):
    scanned = load_scanned_folders()
    scanned.add(folder_id)
    with open(SCANNED_FILE, 'w') as f:
        json.dump(list(scanned), f)


# ═══════════════════════════════════════════════════════
#  RECURSIVE FOLDER SCANNER
# ═══════════════════════════════════════════════════════
def scan_folder(drive_svc, sheets_svc, folder_id, folder_path,
                all_files, seen_checksums, name_index, scanned_folders, depth=0):

    if folder_id in scanned_folders:
        print(f'  {"  "*depth}⏭️  Skipping (already scanned): {folder_path}')
        return

    try:
        page_token = None
        while True:
            resp = drive_svc.files().list(
                q=f"'{folder_id}' in parents and trashed=false",
                spaces='drive',
                fields='nextPageToken, files(id,name,mimeType,size,createdTime,modifiedTime,owners,md5Checksum,webViewLink)',
                pageSize=100,
                pageToken=page_token
            ).execute()

            for f in resp.get('files', []):
                fid      = f.get('id', '')
                name     = f.get('name', 'Unknown')
                mime     = f.get('mimeType', '')
                size     = format_size(f.get('size', '0'), mime)
                created  = f.get('createdTime', '')[:10]
                modified = f.get('modifiedTime', '')
                mod_date = modified[:10]
                owner    = f.get('owners', [{}])[0].get('emailAddress', 'Unknown')
                checksum = f.get('md5Checksum', '')
                link     = f.get('webViewLink', '')

                # Recurse into subfolders
                if mime == 'application/vnd.google-apps.folder':
                    sub_path = f'{folder_path} > {name}'
                    print(f'  {"  "*depth}📂 {sub_path}')
                    scan_folder(drive_svc, sheets_svc, fid, sub_path,
                                all_files, seen_checksums, name_index,
                                scanned_folders, depth+1)
                    continue

                # Personal check — based on folder path
                personal = is_personal(folder_path)

                # Classifications
                file_type  = get_file_type(mime)
                company    = 'Personal' if personal else detect_company(name, folder_path)
                project    = 'Personal' if personal else detect_project(name, folder_path)
                category   = 'Personal' if personal else classify_category(name, folder_path)
                sheet_role = detect_sheet_role(name) if 'spreadsheet' in mime else ''
                inactive   = 'Inactive' if is_inactive(modified) else 'Active'

                # Dependency check for Google Sheets
                dependencies = []
                if mime == 'application/vnd.google-apps.spreadsheet':
                    dependencies = get_sheet_dependencies(sheets_svc, fid)
                dep_str = ', '.join(dependencies) if dependencies else ''

                # Duplicate detection — across ALL folders by checksum AND name
                dup_of = ''
                if checksum and checksum in seen_checksums:
                    dup_of = f'Duplicate of: {seen_checksums[checksum]["name"]} in {seen_checksums[checksum]["path"]}'
                elif not checksum and name.lower() in name_index:
                    dup_of = f'Same name as: {name_index[name.lower()]}'

                if checksum and checksum not in seen_checksums:
                    seen_checksums[checksum] = {'name': name, 'path': folder_path}
                if name.lower() not in name_index:
                    name_index[name.lower()] = folder_path

                record = {
                    'name':         name,
                    'file_type':    file_type,
                    'mime':         mime,
                    'company':      company,
                    'project':      project,
                    'category':     category,
                    'sheet_role':   sheet_role,
                    'dependencies': dep_str,
                    'size':         size,
                    'created':      created,
                    'modified':     mod_date,
                    'status':       inactive,
                    'owner':        owner,
                    'folder_path':  folder_path,
                    'duplicate_of': dup_of,
                    'link':         link
                }

                all_files.append(record)

                flag = '🔁' if dup_of else ('💤' if inactive == 'Inactive' else '✅')
                dep  = f' [{dep_str}]' if dep_str else ''
                print(f'  {"  "*depth}{flag} {name[:35]:<35} | {file_type:<16} | {company:<18} | {category}{dep}')

            page_token = resp.get('nextPageToken')
            if not page_token:
                break

        # Mark folder as done and save progress
        scanned_folders.add(folder_id)
        save_scanned_folder(folder_id)

        # Save data after every folder — so no data is ever lost
        save_data(all_files, seen_checksums)

    except Exception as e:
        print(f'  ⚠️  Error scanning {folder_path}: {e}')
        time.sleep(2)


# ═══════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════
def main():
    print('=' * 65)
    print('   STAGE 1 — Google Drive Scanner')
    print('   Varsha / Forgetech Drive Audit Tool')
    print('=' * 65)

    # Resume check
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            existing = json.load(f)
        all_files      = existing.get('files', [])
        seen_checksums = existing.get('checksums', {})
        print(f'\n♻️  Resuming previous scan — {len(all_files)} files already saved')
    else:
        all_files      = []
        seen_checksums = {}
        print('\n🆕 Starting fresh scan')

    name_index      = {r['name'].lower(): r['folder_path'] for r in all_files}
    scanned_folders = load_scanned_folders()

    creds      = authenticate()
    drive_svc  = build('drive',  'v3', credentials=creds)
    docs_svc   = build('docs',   'v1', credentials=creds)
    sheets_svc = build('sheets', 'v4', credentials=creds)

    print('\n🔍 Finding shared folders...\n')

    # Get all shared items
    page_token     = None
    shared_folders = []
    shared_files   = []

    while True:
        resp = drive_svc.files().list(
            q='sharedWithMe=true and trashed=false',
            spaces='drive',
            fields='nextPageToken, files(id,name,mimeType,size,createdTime,modifiedTime,owners,md5Checksum,webViewLink)',
            pageSize=100,
            pageToken=page_token
        ).execute()

        for f in resp.get('files', []):
            if f.get('mimeType') == 'application/vnd.google-apps.folder':
                shared_folders.append(f)
            else:
                shared_files.append(f)

        page_token = resp.get('nextPageToken')
        if not page_token:
            break

    print(f'📂 {len(shared_folders)} shared folders found')
    print(f'📄 {len(shared_files)} files directly shared\n')

    # Scan each shared folder recursively
    for folder in shared_folders:
        fname = folder.get('name', 'Unknown')
        fid   = folder.get('id', '')
        if fid in scanned_folders:
            print(f'⏭️  Already scanned: {fname}')
            continue
        print(f'\n📂 Scanning folder: {fname}')
        scan_folder(drive_svc, sheets_svc, fid, fname,
                    all_files, seen_checksums, name_index, scanned_folders)

    # Process directly shared files
    print(f'\n📄 Processing directly shared files...')
    for f in shared_files:
        fid      = f.get('id', '')
        name     = f.get('name', 'Unknown')
        mime     = f.get('mimeType', '')
        size     = format_size(f.get('size', '0'), mime)
        created  = f.get('createdTime', '')[:10]
        modified = f.get('modifiedTime', '')
        mod_date = modified[:10]
        owner    = f.get('owners', [{}])[0].get('emailAddress', 'Unknown')
        checksum = f.get('md5Checksum', '')
        link     = f.get('webViewLink', '')

        personal   = False
        file_type  = get_file_type(mime)
        company    = detect_company(name, '')
        project    = detect_project(name, '')
        category   = classify_category(name, '')
        sheet_role = detect_sheet_role(name) if 'spreadsheet' in mime else ''
        inactive   = 'Inactive' if is_inactive(modified) else 'Active'

        dependencies = []
        if mime == 'application/vnd.google-apps.spreadsheet':
            dependencies = get_sheet_dependencies(sheets_svc, fid)
        dep_str = ', '.join(dependencies) if dependencies else ''

        dup_of = ''
        if checksum and checksum in seen_checksums:
            dup_of = f'Duplicate of: {seen_checksums[checksum]["name"]} in {seen_checksums[checksum]["path"]}'
        if checksum and checksum not in seen_checksums:
            seen_checksums[checksum] = {'name': name, 'path': 'Directly Shared'}

        all_files.append({
            'name': name, 'file_type': file_type, 'mime': mime,
            'company': company, 'project': project, 'category': category,
            'sheet_role': sheet_role, 'dependencies': dep_str,
            'size': size, 'created': created, 'modified': mod_date,
            'status': inactive, 'owner': owner,
            'folder_path': 'Directly Shared', 'duplicate_of': dup_of, 'link': link
        })
        print(f'  ✅ {name[:40]}')

    # Final save
    save_data(all_files, seen_checksums)

    print('\n' + '=' * 65)
    print(f'🎉 SCAN COMPLETE!')
    print(f'📊 Total files: {len(all_files)}')
    print(f'💾 Data saved to: {DATA_FILE}')
    print(f'\n👉 Now run: python write_report.py')
    print('=' * 65)

if __name__ == '__main__':
    main()