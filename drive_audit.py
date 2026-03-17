import os
import json
import pickle
import time
import csv
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from datetime import datetime, timezone

# ═══════════════════════════════════════════════════════
#  SETTINGS & FILTERS (Your Original Logic)
# ═══════════════════════════════════════════════════════
CREDENTIALS_FILE = 'credentials.json'
TOKEN_FILE      = 'token.pickle'
DATA_FILE       = 'audit_data.json'
REPORT_FILE     = 'ERP_Migration_Final_Report.csv'
SCANNED_FILE    = 'scanned_folders.json'

# Category Keywords (Your specific Plant/QMS/Finance list)
CATEGORY_KEYWORDS = {
    'Plant Operations': ['plant', 'production', 'ppc', 'dispatch', 'press shop', 'hammer', 'die shop', 'stores'],
    'Quality & QMS': ['qms', 'quality', 'control plan', 'audit', 'iso', 'iatf', 'calibration', 'vfqa'],
    'HR & Admin': ['hr', 'salary', 'leave', 'skill matrix', 'attendance', 'payroll'],
    'Commercial': ['purchase', 'procurement', 'rfq', 'quotation', 'vendor', 'po'],
    'Finance': ['accounts', 'ledger', 'balance sheet', 'gst', '2025-2026', '2024-2025']
}

SCOPES = [
    'https://www.googleapis.com/auth/drive.readonly',
    'https://www.googleapis.com/auth/spreadsheets.readonly'
]

# ═══════════════════════════════════════════════════════
#  INTELLIGENCE ENGINE (The Addition)
# ═══════════════════════════════════════════════════════

def calculate_utility(name, modified_date, deps, category):
    """The Potato Principle: Identifying the 20% high-value data."""
    score = 0
    name_low = name.lower()
    
    # 1. Recency (Last 60 days)
    try:
        mod = datetime.fromisoformat(modified_date.replace('Z', '+00:00'))
        if (datetime.now(timezone.utc) - mod).days < 60: score += 35
    except: pass

    # 2. Functional Criticality (Is it a Master or a Hub?)
    if any(k in name_low for k in ['master', 'database', 'final', '2026', '25-26']): score += 30
    if deps: score += 25 # Has IMPORTRANGE/QUERY

    # 3. Domain Priority
    if category in ['Plant Operations', 'Finance & Accounts']: score += 10
    
    return min(score, 100)

def get_sheet_dependencies(sheets_svc, file_id):
    """Deep scan for logic hubs."""
    try:
        res = sheets_svc.spreadsheets().get(spreadsheetId=file_id, fields='sheets.properties.title').execute()
        title = res['sheets'][0]['properties']['title']
        vals = sheets_svc.spreadsheets().values().get(spreadsheetId=file_id, range=f"'{title}'!A1:Z50").execute().get('values', [])
        flat = str(vals).lower()
        found = []
        if 'importrange' in flat: found.append('IMPORTRANGE')
        if 'query(' in flat: found.append('QUERY')
        return found
    except: return []

# ═══════════════════════════════════════════════════════
#  RECURSIVE CORE (Maintained from your script)
# ═══════════════════════════════════════════════════════

def scan_folder(drive_svc, sheets_svc, folder_id, folder_path, all_files, seen_hashes, scanned_folders, depth=0):
    if folder_id in scanned_folders: return

    try:
        page_token = None
        while True:
            resp = drive_svc.files().list(
                q=f"'{folder_id}' in parents and trashed=false",
                fields='nextPageToken, files(id,name,mimeType,size,modifiedTime,owners,md5Checksum)',
                pageSize=100, pageToken=page_token).execute()

            for f in resp.get('files', []):
                fid = f['id']
                name = f['name']
                mime = f['mimeType']
                checksum = f.get('md5Checksum')
                
                if mime == 'application/vnd.google-apps.folder':
                    scan_folder(drive_svc, sheets_svc, fid, f"{folder_path} > {name}", all_files, seen_hashes, scanned_folders, depth+1)
                    continue

                # --- DUPLICATE & UTILITY LOGIC ---
                status = "ORIGINAL"
                dup_notes = ""
                fingerprint = checksum if checksum else f"{name}_{f.get('size',0)}"
                
                if fingerprint in seen_hashes:
                    status = "DUPLICATE"
                    master = seen_hashes[fingerprint]
                    dup_notes = f"Master exists in: {master['path']}"
                else:
                    seen_hashes[fingerprint] = {'name': name, 'path': folder_path}

                # Categorize (Your logic)
                category = "Other"
                for cat, keys in CATEGORY_KEYWORDS.items():
                    if any(k in (name + folder_path).lower() for k in keys):
                        category = cat
                        break

                # Deep Audit (Only for Potential Masters)
                deps = []
                if status == "ORIGINAL" and 'spreadsheet' in mime:
                    deps = get_sheet_dependencies(sheets_svc, fid)
                
                score = calculate_utility(name, f.get('modifiedTime',''), deps, category) if status == "ORIGINAL" else 0

                all_files.append({
                    'Name': name,
                    'Category': category,
                    'Status': status,
                    'Score': score,
                    'Folder_Path': folder_path,
                    'Dependencies': ", ".join(deps),
                    'Modified': f.get('modifiedTime', '')[:10],
                    'Owner': f.get('owners', [{}])[0].get('emailAddress'),
                    'Notes': dup_notes,
                    'ERP_Ready': "YES" if score > 65 else "NO"
                })
                print(f"{'  '*depth} {'✅' if status=='ORIGINAL' else '🔁'} {score:02} | {name[:40]}")

            page_token = resp.get('nextPageToken')
            if not page_token: break
        
        scanned_folders.add(folder_id)
        # Save progress after every folder (Your Resume Feature)
        with open(DATA_FILE, 'w') as f: json.dump(all_files, f)
            
    except Exception as e:
        print(f"Error in {folder_path}: {e}")

# ═══════════════════════════════════════════════════════
#  MAIN EXECUTION
# ═══════════════════════════════════════════════════════

def main():
    # ... (Authentication code same as your script) ...
    creds = authenticate() # Assumes your authenticate() function is present
    drive_svc = build('drive', 'v3', credentials=creds)
    sheets_svc = build('sheets', 'v4', credentials=creds)

    all_files = []
    seen_hashes = {}
    scanned_folders = set()

    # Load Resume Data
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, 'r') as f: all_files = json.load(f)
        print(f"♻️ Resuming: {len(all_files)} files loaded.")

    print("🚀 Starting Deep Industrial Audit...")
    
    # Get Shared Folders
    results = drive_svc.files().list(q="sharedWithMe=true and mimeType='application/vnd.google-apps.folder'").execute()
    for folder in results.get('files', []):
        scan_folder(drive_svc, sheets_svc, folder['id'], folder['name'], all_files, seen_hashes, scanned_folders)

    # Output to CSV for your March 31st review
    if all_files:
        keys = all_files[0].keys()
        with open(REPORT_FILE, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=keys)
            writer.writeheader()
            writer.writerows(all_files)
    
    print(f"📊 Full Audit Complete. Report: {REPORT_FILE}")

if __name__ == '__main__':
    main()
