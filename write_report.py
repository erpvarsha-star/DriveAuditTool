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
import gspread
from google.auth.transport.requests import Request
from datetime import datetime, timezone

# ═══════════════════════════════════════════════════════
#  CONFIG
# ═══════════════════════════════════════════════════════
SHEET_ID        = '13gHBHZz1MbvDMGtE3fAHJWyhTbnoNCcnZit5Rmh-L80'
DATA_FILE       = 'audit_data.json'
TOKEN_FILE      = 'token.pickle'
BATCH_SIZE      = 500

# Updated Headers to include Intelligence Data
FULL_HEADERS = [
    'File Name', 'Status', 'Score', 'ERP Ready', 'Category', 
    'File Type', 'Last Modified', 'Owner', 'Dependencies', 
    'Folder Path', 'Duplicate Notes', 'Link'
]

HEADER_FMT = {
    'backgroundColor': {'red': 0.1, 'green': 0.2, 'blue': 0.4},
    'textFormat': {'bold': True, 'foregroundColor': {'red': 1.0, 'green': 1.0, 'blue': 1.0}}
}

# ═══════════════════════════════════════════════════════
#  ROW CONVERTER
# ═══════════════════════════════════════════════════════
def to_audit_row(r):
    return [
        r.get('name', r.get('Name', '')), # Handles both naming conventions
        r.get('status', r.get('Status', '')),
        r.get('score', r.get('Score', 0)),
        r.get('erp_ready', r.get('ERP_Ready', 'NO')),
        r.get('category', r.get('Category', '')),
        r.get('file_type', r.get('Type', '')),
        r.get('modified', r.get('Modified', '')),
        r.get('owner', r.get('Owner', '')),
        r.get('dependencies', r.get('Dependencies', '')),
        r.get('folder_path', r.get('Folder_Path', '')),
        r.get('duplicate_of', r.get('Notes', '')),
        r.get('link', r.get('Link', ''))
    ]

# ═══════════════════════════════════════════════════════
#  NEW TAB: ERP PRIORITY
# ═══════════════════════════════════════════════════════
def write_erp_priority(spreadsheet, files):
    # The Potato Principle: Filter for top 20%
    priority = [f for f in files if f.get('ERP_Ready') == 'YES' or f.get('erp_ready') == 'YES']
    priority = sorted(priority, key=lambda x: x.get('score', x.get('Score', 0)), reverse=True)
    
    print(f'\nWriting ERP Priority Tab ({len(priority)} key files)...')
    ws = get_or_create_tab(spreadsheet, '🚀 ERP MIGRATION LIST', rows=len(priority)+10, cols=12)
    ws.append_row(FULL_HEADERS)
    ws.format('A1:L1', {'backgroundColor': {'red': 0.0, 'green': 0.5, 'blue': 0.2}, 'textFormat': {'bold': True, 'foregroundColor': {'red': 1.0, 'green': 1.0, 'blue': 1.0}}})
    
    rows = [to_audit_row(f) for f in priority]
    write_in_batches(ws, rows)

# ═══════════════════════════════════════════════════════
#  ENHANCED SUMMARY
# ═══════════════════════════════════════════════════════
def write_enhanced_summary(spreadsheet, files):
    ws = get_or_create_tab(spreadsheet, 'Executive Summary', rows=50, cols=3)
    
    total = len(files)
    dups = sum(1 for f in files if f.get('status', f.get('Status')) == 'DUPLICATE')
    erp_ready = sum(1 for f in files if f.get('erp_ready', f.get('ERP_Ready')) == 'YES')
    
    # Calculate Data Bloat
    bloat_percent = (dups / total) * 100 if total > 0 else 0
    
    summary_rows = [
        ['MIGRATION READINESS SUMMARY', '', datetime.now().strftime('%Y-%m-%d')],
        ['', '', ''],
        ['METRIC', 'COUNT', 'ACTION'],
        ['Total Files Scanned', total, 'Full Inventory'],
        ['Duplicates Found', dups, f'DEDUPLICATION POTENTIAL: {bloat_percent:.1f}%'],
        ['ERP Ready (High Value)', erp_ready, 'MOVE THESE BY APRIL 1ST'],
        ['Low Value / Archive', total - erp_ready - dups, 'Move to 2025 Archive Folder'],
        ['', '', ''],
        ['PLANT OPERATIONS FOCUS', '', ''],
    ]
    
    # Category Breakdown
    cats = {}
    for f in files:
        c = f.get('category', f.get('Category', 'Other'))
        cats[c] = cats.get(c, 0) + 1
    
    for c, count in sorted(cats.items(), key=lambda x: -x[1]):
        summary_rows.append([f'Category: {c}', count, ''])

    ws.append_rows(summary_rows)
    ws.format('A1:C1', HEADER_FMT)

# ... [Keep your existing authenticate, write_in_batches, and get_or_create_tab functions] ...

def main():
    print('🚀 Running Stage 2 Intelligence Report Writer...')
    
    with open(DATA_FILE, 'r') as f:
        all_files = json.load(f)

    creds = authenticate()
    gc = gspread.authorize(creds)
    spreadsheet = gc.open_by_key(SHEET_ID)

    # 1. New Executive Summary
    write_enhanced_summary(spreadsheet, all_files)

    # 2. ERP Priority Tab (The Actionable 20%)
    write_erp_priority(spreadsheet, all_files)

    # 3. Full Inventory (All Data)
    print("\nWriting Full Inventory...")
    ws_full = get_or_create_tab(spreadsheet, 'Full Inventory', rows=len(all_files)+10, cols=12)
    ws_full.append_row(FULL_HEADERS)
    ws_full.format('A1:L1', HEADER_FMT)
    full_rows = [to_audit_row(f) for f in all_files]
    write_in_batches(ws_full, full_rows)

    print(f'✅ Report Complete! Focus on the "ERP MIGRATION LIST" tab for April 1st.')

if __name__ == '__main__':
    main()
