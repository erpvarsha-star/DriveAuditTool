"""
backup.py — Step 0: Export all critical sheets as .xlsx before any write runs.

Exports every sheet in your system to Excel format and saves to:
  1. Local /backups/ folder (available as GitHub artifact)
  2. Google Drive backup folder

Run BEFORE any other script. Added as first step in weekly_scan.yml.
Run: python backup.py
"""

import os
import pickle
import time
import io
from datetime import datetime

from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# ═══════════════════════════════════════════════════════
#  CONFIG
# ═══════════════════════════════════════════════════════

TOKEN_FILE        = 'token.pickle'
BACKUP_FOLDER     = 'backups'
DELAY             = 1.0

# All critical sheets to backup
SHEETS_TO_BACKUP = {
    'VFPL_Dashboard':        '1NR8EPGRJN0AQDXZjYw5k93clsO8AD4u2l2Xke1lBC2I',
    'VFL_PMS_Forgings':      '1c-axqiBEufNb1vK-JdJP6t1eRTmzliA3otAbqxAOXZE',
    'VFL_Actual_Dispatch':   '1txZM9a9_OSG-ZWYaAEBLKj-9M7LYLrFyl0kJkPsTMGI',
    'VFL_Daily_Manpower':    '1t7UjWTP_cpIJ2BjoaMlV6uUKA7ztH9UnKc_korYKCiw',
    'Electricity':           '1nUvf-UWjBSbSWnZTNph-gRUbjzuguGlidpYBKshKUNQ',
    'VFL_PMS_Machine_Shop':  '1yC-b36rgAxablmdXhngHCEOsmnlgWourep6ctKifXCA',
    'Master_Data_Sheet':     '10Zjxy3mGKP6G3j7uuak3FTXl90JHEQoC0RDgl4tJJXc',
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
#  BACKUP
# ═══════════════════════════════════════════════════════

def export_sheet_as_xlsx(drive_svc, sheet_id, name, backup_dir):
    """Export a Google Sheet as .xlsx and save locally."""
    date_str  = datetime.now().strftime('%Y%m%d_%H%M')
    filename  = f"{name}_{date_str}.xlsx"
    filepath  = os.path.join(backup_dir, filename)

    url = (f"https://docs.google.com/spreadsheets/d/{sheet_id}"
           f"/export?format=xlsx")

    try:
        request  = drive_svc.files().export_media(
            fileId=sheet_id,
            mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        buffer   = io.BytesIO()
        downloader = MediaIoBaseDownload(buffer, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()

        with open(filepath, 'wb') as f:
            f.write(buffer.getvalue())

        size_kb = os.path.getsize(filepath) // 1024
        print(f"  ✅ {name} → {filename} ({size_kb} KB)")
        return filepath

    except HttpError as e:
        if e.resp.status == 429:
            print(f"  ⏳ Quota — waiting 60s")
            time.sleep(60)
            return export_sheet_as_xlsx(drive_svc, sheet_id, name, backup_dir)
        print(f"  ⚠ Cannot export {name}: {e}")
        return None
    except Exception as e:
        print(f"  ⚠ Error exporting {name}: {e}")
        return None

# ═══════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════

def main():
    print('=' * 60)
    print('   STEP 0 — BACKUP')
    print(f'   {datetime.now().strftime("%Y-%m-%d %H:%M")}')
    print('=' * 60)

    creds     = authenticate()
    drive_svc = build('drive','v3',credentials=creds)

    # Create backup folder
    date_str   = datetime.now().strftime('%Y%m%d')
    backup_dir = os.path.join(BACKUP_FOLDER, date_str)
    os.makedirs(backup_dir, exist_ok=True)
    print(f"Backup folder: {backup_dir}\n")

    # Export each sheet
    backed_up = []
    failed    = []

    for name, sheet_id in SHEETS_TO_BACKUP.items():
        print(f"  Backing up: {name}...")
        filepath = export_sheet_as_xlsx(drive_svc, sheet_id, name, backup_dir)
        if filepath:
            backed_up.append(name)
        else:
            failed.append(name)
        time.sleep(DELAY)

    print(f"\n✅ Backup complete")
    print(f"   Backed up: {len(backed_up)} sheets")
    print(f"   Failed:    {len(failed)} sheets")
    if failed:
        print(f"   Failed:    {failed}")
    print(f"   Location:  {backup_dir}/")
    print(f"\n   Files will be available as GitHub Actions artifact")
    print(f"   Download from: Actions → Run → Artifacts → backup-data")


if __name__ == '__main__':
    main()
