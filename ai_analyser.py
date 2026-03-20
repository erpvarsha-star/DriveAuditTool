"""
STAGE 3 — AI CONTENT ANALYSER (v1)

For each ERP-ready Google Sheet, reads first 50 rows and asks Claude to return:
  1.  ERP module + why
  2.  Data quality score + reason
  3.  ERP field mapping
  4.  Master data type + customer name (for quotations)
  5.  Risk flag + detail
  6.  Suggested action + reason
  7.  Owner signoff
  8.  Form source chain (if form-linked)
  9.  Drive reorg suggestion
  10. Duplicate content flag

Resumes from ai_analysis.json if interrupted.
Only analyses files not already in ai_analysis.json.

Requires: ANTHROPIC_API_KEY GitHub secret
"""

import json
import os
import pickle
import time
import csv
import urllib.request
from datetime import datetime

import gspread
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ═══════════════════════════════════════════════════════
#  CONFIG
# ═══════════════════════════════════════════════════════

DATA_FILE          = 'audit_data.json'
AI_OUTPUT_FILE     = 'ai_analysis.json'
TOKEN_FILE         = 'token.pickle'
AI_REPORT_CSV      = 'AI_ERP_Intelligence_Report.csv'

SHEET_ID           = '1yrjuBo9LJyR41cO0AZzkWJQQ_BAtYhTJ0Y0VgLr4eog'
INVENTORY_SHEET_ID = os.environ.get('INVENTORY_SHEET_ID','')
ANTHROPIC_API_KEY  = os.environ.get('ANTHROPIC_API_KEY','')

ROWS_TO_READ       = 50
BATCH_DELAY        = 0.5
AI_DELAY           = 0.5
BATCH_SIZE         = 500
TAB_INIT_ROWS      = 1
MAX_FILES_PER_RUN  = 100   # Process max 100 files per run to stay under 6h limit
API_TIMEOUT        = 45    # Seconds before giving up on one file
SAVE_EVERY         = 1     # Save after EVERY file — never lose progress

ERP_MODULES = [
    'Production / Planning (PPC)',
    'Quality Management (QMS)',
    'Purchase / Procurement',
    'Finance & Accounts',
    'HR & Payroll',
    'Inventory / Stores',
    'Sales / Dispatch',
]

AI_HEADERS = [
    'File Name','ERP Module','ERP Value','ERP Field Mapping',
    'Master Data Type','Customer Name',
    'Data Quality','Quality Reason',
    'Duplicate Content','Risk Flag','Risk Detail',
    'Suggested Action','Action Reason','Owner Signoff',
    'Form Chain','Drive Suggested Folder','Suggested File Name',
    'Form Status','Has Form','Dependencies',
    'Score','Category','Owner','Last Modified','Link',
]

FMT = {
    'green':  {'backgroundColor':{'red':0.0,'green':0.5,'blue':0.2},
               'textFormat':{'bold':True,'foregroundColor':{'red':1,'green':1,'blue':1}}},
    'blue':   {'backgroundColor':{'red':0.1,'green':0.2,'blue':0.4},
               'textFormat':{'bold':True,'foregroundColor':{'red':1,'green':1,'blue':1}}},
    'orange': {'backgroundColor':{'red':0.8,'green':0.4,'blue':0.0},
               'textFormat':{'bold':True,'foregroundColor':{'red':1,'green':1,'blue':1}}},
    'red':    {'backgroundColor':{'red':0.6,'green':0.1,'blue':0.1},
               'textFormat':{'bold':True,'foregroundColor':{'red':1,'green':1,'blue':1}}},
    'purple': {'backgroundColor':{'red':0.4,'green':0.1,'blue':0.5},
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
#  SHEET CONTENT READER
# ═══════════════════════════════════════════════════════

def get_sheet_content(sheets_svc, file_id, rows=ROWS_TO_READ):
    try:
        meta = sheets_svc.spreadsheets().get(
            spreadsheetId=file_id,
            fields='sheets.properties.title'
        ).execute()
        time.sleep(BATCH_DELAY)
        tabs = meta.get('sheets',[])
        if not tabs: return []
        first_tab = tabs[0]['properties']['title']
        resp = sheets_svc.spreadsheets().values().get(
            spreadsheetId=file_id,
            range=f"'{first_tab}'!A1:Z{rows}"
        ).execute()
        time.sleep(BATCH_DELAY)
        return resp.get('values',[])
    except HttpError as e:
        if e.resp.status == 429:
            print(f"  ⏳ Quota — waiting 60s")
            time.sleep(60)
            return get_sheet_content(sheets_svc, file_id, rows)
        return []
    except Exception:
        return []


def format_for_ai(rows, name, folder, owner, has_form, deps, form_status):
    lines = [
        f"File name: {name}",
        f"Folder path: {folder}",
        f"Owner: {owner}",
        f"Has Google Form linked: {'Yes — ' + form_status if has_form else 'No'}",
        f"Has IMPORTRANGE/QUERY dependencies: {'Yes — ' + deps if deps else 'No'}",
        f"",
        f"Sheet content (first {len(rows)} rows):",
    ]
    for i, row in enumerate(rows[:ROWS_TO_READ]):
        clean = [str(c)[:80] for c in row]
        lines.append(f"Row {i+1}: {' | '.join(clean)}")
    return '\n'.join(lines)

# ═══════════════════════════════════════════════════════
#  AI ANALYSER
# ═══════════════════════════════════════════════════════

SYSTEM_PROMPT = f"""You are an ERP migration consultant for an Indian manufacturing / forging company.
The company is implementing ERP across 7 modules:
{chr(10).join(f'  - {m}' for m in ERP_MODULES)}

Analyse the Google Sheet content provided and return ONLY a valid JSON object with exactly these keys:

{{
  "erp_module": "one of the 7 modules or 'Not ERP Relevant'",
  "erp_value": "1-2 sentences on why this data matters for ERP",
  "erp_field_mapping": "comma-separated ERP masters or transactions e.g. Item Master, BOM, Work Order, Vendor Master",
  "master_data_type": "Quotation / Purchase Order / Cost Sheet / Customer Master / Vendor Master / Item Master / BOM / HR Master / Sales Order / Production / Finance / Other",
  "customer_name": "customer name if this is a quotation or sales order, else empty string",
  "data_quality": "Complete or Incomplete or Messy",
  "quality_reason": "1 sentence — what is good or bad about the data",
  "duplicate_content": "YES or NO",
  "duplicate_note": "if YES — what data is likely duplicated across files",
  "risk_flag": "NONE or SENSITIVE or FINANCIAL or PERSONAL or CONFIDENTIAL",
  "risk_detail": "1 sentence on the risk if any, else empty string",
  "suggested_action": "MIGRATE or ARCHIVE or DELETE or REVIEW",
  "action_reason": "1 sentence explaining the action",
  "owner_signoff": "job title who should approve e.g. Plant Manager, Finance Head, HR Manager",
  "form_chain": "if form linked: Form name → questions captured → Sheet → ERP module → ERP field. Else empty string.",
  "suggested_folder": "suggested Drive folder path e.g. /ERP/Purchase/Quotations/CustomerName/",
  "suggested_filename": "suggested file name if current name is unclear, else empty string"
}}

Return ONLY the JSON. No markdown. No explanation. No backticks."""


def analyse_with_ai(content_text, file_name):
    if not ANTHROPIC_API_KEY:
        raise RuntimeError("ANTHROPIC_API_KEY not set.")

    payload = json.dumps({
        "model":      "claude-sonnet-4-20250514",
        "max_tokens": 1000,
        "system":     SYSTEM_PROMPT,
        "messages":   [{"role":"user","content":f"Analyse this file:\n\n{content_text}"}]
    }).encode()

    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=payload,
        headers={
            "Content-Type":      "application/json",
            "x-api-key":         ANTHROPIC_API_KEY,
            "anthropic-version": "2023-06-01",
        }
    )

    try:
        with urllib.request.urlopen(req, timeout=API_TIMEOUT) as resp:
            data     = json.loads(resp.read())
            raw_text = data['content'][0]['text'].strip()
            if raw_text.startswith('```'):
                raw_text = raw_text.split('```')[1]
                if raw_text.startswith('json'): raw_text = raw_text[4:]
            return json.loads(raw_text.strip())
    except urllib.error.URLError as e:
        print(f"  ⏱ Timeout/network error for {file_name}: {e} — skipping")
        return None
    except Exception as e:
        print(f"  ⚠ AI error for '{file_name}': {e}")
        return {
            "erp_module":"Review needed","erp_value":"AI analysis failed",
            "erp_field_mapping":"","master_data_type":"","customer_name":"",
            "data_quality":"Unknown","quality_reason":str(e),
            "duplicate_content":"NO","duplicate_note":"",
            "risk_flag":"NONE","risk_detail":"",
            "suggested_action":"REVIEW","action_reason":"AI analysis failed",
            "owner_signoff":"","form_chain":"",
            "suggested_folder":"","suggested_filename":"",
        }

# ═══════════════════════════════════════════════════════
#  SHEET HELPERS
# ═══════════════════════════════════════════════════════

def get_or_create_tab(spreadsheet, title, cols=25):
    try:
        ws = spreadsheet.worksheet(title)
        print(f"  📋 Reusing: {title}"); return ws
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=title,rows=TAB_INIT_ROWS,cols=cols)
        print(f"  ➕ Created: {title}"); return ws

def safe_clear(ws):
    ws.clear(); time.sleep(0.5)

def write_in_batches(ws, rows):
    total = len(rows)
    if not total: return
    for i in range(0,total,BATCH_SIZE):
        ws.append_rows(rows[i:i+BATCH_SIZE],value_input_option='USER_ENTERED')
        print(f"    ↳ {min(i+BATCH_SIZE,total):,}/{total:,}")
        time.sleep(1)

# ═══════════════════════════════════════════════════════
#  ROW BUILDER
# ═══════════════════════════════════════════════════════

def to_ai_row(f, ai):
    return [
        f.get('Name',''),
        ai.get('erp_module',''),
        ai.get('erp_value',''),
        ai.get('erp_field_mapping',''),
        ai.get('master_data_type',''),
        ai.get('customer_name',''),
        ai.get('data_quality',''),
        ai.get('quality_reason',''),
        ai.get('duplicate_content','') + (' — '+ai.get('duplicate_note','') if ai.get('duplicate_content')=='YES' else ''),
        ai.get('risk_flag',''),
        ai.get('risk_detail',''),
        ai.get('suggested_action',''),
        ai.get('action_reason',''),
        ai.get('owner_signoff',''),
        ai.get('form_chain',''),
        ai.get('suggested_folder',''),
        ai.get('suggested_filename',''),
        f.get('Form_Status',''),
        f.get('Has_Form',''),
        f.get('Dependencies',''),
        f.get('Score',0),
        f.get('Category',''),
        f.get('Owner',''),
        f.get('Modified',''),
        f.get('Link',''),
    ]

# ═══════════════════════════════════════════════════════
#  PROGRESS
# ═══════════════════════════════════════════════════════

def load_ai_progress():
    if os.path.exists(AI_OUTPUT_FILE):
        with open(AI_OUTPUT_FILE) as f: data = json.load(f)
        print(f"♻️  Resuming AI: {len(data):,} files already done.")
        return data
    return {}

def save_ai_progress(results):
    with open(AI_OUTPUT_FILE,'w') as f: json.dump(results,f)

# ═══════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════

def main():
    print('='*60)
    print('   STAGE 3 — AI ERP Intelligence Analyser')
    print('='*60)

    if not ANTHROPIC_API_KEY:
        raise SystemExit("❌ ANTHROPIC_API_KEY not set — add as GitHub secret.")

    if not os.path.exists(DATA_FILE):
        raise SystemExit("❌ audit_data.json not found — run drive_audit.py first.")

    with open(DATA_FILE) as f: all_files = json.load(f)

    # Analyse ERP-ready original spreadsheets only
    to_analyse = [
        f for f in all_files
        if f.get('ERP_Ready','NO') == 'YES'
        and f.get('Status','')     == 'ORIGINAL'
        and 'spreadsheet' in f.get('Type', f.get('mimeType',''))
    ]

    print(f"📊 Total files      : {len(all_files):,}")
    print(f"🎯 Files to analyse : {len(to_analyse):,}")
    est = len(to_analyse) * 0.003
    print(f"💰 Estimated cost   : ${est:.2f}\n")

    creds      = authenticate()
    sheets_svc = build('sheets','v4',credentials=creds)
    gc         = gspread.authorize(creds)

    ai_results = load_ai_progress()
    done_names = set(ai_results.keys())
    total      = len(to_analyse)
    completed  = len(done_names)

    for f in to_analyse:
        name = f.get('Name','')
        fid  = f.get('id', f.get('file_id',''))

        if name in done_names:
            continue

        # Stop if we've hit max files for this run
        files_this_run = completed - len(done_names)
        if files_this_run >= MAX_FILES_PER_RUN:
            print(f"\n⏸ Reached {MAX_FILES_PER_RUN} files limit for this run.")
            print(f"   Remaining {total - completed} files will be processed next run.")
            break

        print(f"[{completed+1}/{total}] 🤖 {name[:55]}")

        content_rows = get_sheet_content(sheets_svc, fid) if fid else []
        content_text = format_for_ai(
            content_rows, name,
            f.get('Folder_Path',''), f.get('Owner',''),
            f.get('Has_Form','')=='YES',
            f.get('Dependencies',''),
            f.get('Form_Status',''),
        )

        ai = analyse_with_ai(content_text, name)

        # If AI returned None (timeout/error) — mark as skipped and continue
        if ai is None:
            ai = {
                'erp_module': 'Unknown',
                'suggested_action': 'REVIEW',
                'risk_flag': 'NONE',
                'customer_name': '',
                'master_data_type': 'Other',
                'data_quality': 'Unknown',
                'quality_reason': 'API timeout — needs manual review',
                'erp_value': '',
                'erp_field_mapping': '',
                'duplicate_content': 'NO',
                'duplicate_note': '',
                'risk_detail': '',
                'action_reason': 'Skipped due to API timeout',
                'owner_signoff': '',
                'form_chain': '',
                'suggested_folder': '',
                'suggested_filename': '',
            }

        ai_results[name] = {'file': f, 'ai': ai}

        module  = ai.get('erp_module','')
        action  = ai.get('suggested_action','')
        risk    = ai.get('risk_flag','NONE')
        cust    = ai.get('customer_name','')
        mdt     = ai.get('master_data_type','')
        risk_ic = ' ⚠️' if risk != 'NONE' else ''
        cust_ic = f' 👤{cust}' if cust else ''
        print(f"   → {module} | {action}{risk_ic}{cust_ic} | {mdt}")

        completed += 1

        # Save after EVERY file — never lose progress
        save_ai_progress(ai_results)

        time.sleep(AI_DELAY)

    save_ai_progress(ai_results)
    print(f"\n✅ AI analysis complete — {completed:,} files")

    # ── Categorise results ─────────────────────────────
    migrate_rows = []; review_rows = []; archive_rows = []
    risk_rows    = []; customer_map = {}; all_rows = []

    for entry in ai_results.values():
        f   = entry['file']
        ai  = entry['ai']
        row = to_ai_row(f, ai)
        all_rows.append(row)

        action = ai.get('suggested_action','')
        risk   = ai.get('risk_flag','NONE')
        cust   = ai.get('customer_name','')
        mdt    = ai.get('master_data_type','')

        if action == 'MIGRATE':  migrate_rows.append(row)
        elif action == 'REVIEW': review_rows.append(row)
        elif action == 'ARCHIVE':archive_rows.append(row)
        if risk != 'NONE':       risk_rows.append(row)

        # Build customer quotation map
        if cust and mdt in ('Quotation','Sales Order'):
            if cust not in customer_map:
                customer_map[cust] = []
            customer_map[cust].append(row)

    # ── Write to Google Sheet ──────────────────────────
    print('\nWriting AI tabs to Google Sheet...')
    spreadsheet = gc.open_by_key(SHEET_ID)

    def write_tab(title, rows, fmt_key):
        ws = get_or_create_tab(spreadsheet, title, cols=25)
        safe_clear(ws)
        ws.append_row(AI_HEADERS)
        ws.format(f'A1:Y1', FMT[fmt_key])
        write_in_batches(ws, rows)
        print(f"  ✅ {title}: {len(rows):,}")

    write_tab('🤖 Migrate Now',     migrate_rows, 'green')
    write_tab('👀 Review Required', review_rows,  'orange')
    write_tab('📦 Archive',         archive_rows, 'blue')
    write_tab('⚠️ Risk Flagged',    risk_rows,    'red')
    write_tab('🧠 Full AI Report',  all_rows,     'purple')

    # ── Customer Quotation Map ─────────────────────────
    print('\nWriting Customer Quotation Map...')
    cust_rows = []
    for cust_name, rows in sorted(customer_map.items()):
        for row in rows:
            cust_rows.append(row)
    ws_c = get_or_create_tab(spreadsheet,'👤 Customer Quotation Map',cols=25)
    safe_clear(ws_c)
    ws_c.append_row(AI_HEADERS)
    ws_c.format('A1:Y1',FMT['teal'])
    write_in_batches(ws_c, cust_rows)
    print(f"  ✅ {len(customer_map):,} customers | {len(cust_rows):,} files")

    # ── AI Summary tab ─────────────────────────────────
    print('\nWriting AI Summary...')
    modules = {}; actions = {}; risks = {}; quality = {}; customers = {}
    for entry in ai_results.values():
        ai = entry['ai']
        m  = ai.get('erp_module','Unknown')
        a  = ai.get('suggested_action','Unknown')
        r  = ai.get('risk_flag','NONE')
        q  = ai.get('data_quality','Unknown')
        c  = ai.get('customer_name','')
        modules[m]  = modules.get(m,0)+1
        actions[a]  = actions.get(a,0)+1
        risks[r]    = risks.get(r,0)+1
        quality[q]  = quality.get(q,0)+1
        if c: customers[c] = customers.get(c,0)+1

    sum_rows = [
        ['AI ERP INTELLIGENCE SUMMARY','',datetime.now().strftime('%Y-%m-%d %H:%M')],
        ['FILES ANALYSED', completed,''],[''],
        ['ACTION BREAKDOWN','COUNT',''],
    ] + [[f'  {a}',n,''] for a,n in sorted(actions.items(),key=lambda x:-x[1])] + [
        [''],['ERP MODULE MAPPING','COUNT',''],
    ] + [[f'  {m}',n,''] for m,n in sorted(modules.items(),key=lambda x:-x[1])] + [
        [''],['DATA QUALITY','COUNT',''],
    ] + [[f'  {q}',n,''] for q,n in sorted(quality.items(),key=lambda x:-x[1])] + [
        [''],['RISK FLAGS','COUNT',''],
    ] + [[f'  {r}',n,''] for r,n in sorted(risks.items(),key=lambda x:-x[1])] + [
        [''],['TOP CUSTOMERS (by quote count)','QUOTES',''],
    ] + [[f'  {c}',n,''] for c,n in sorted(customers.items(),key=lambda x:-x[1])[:20]]

    ws_s = get_or_create_tab(spreadsheet,'🤖 AI Summary',cols=3)
    safe_clear(ws_s)
    ws_s.append_rows(sum_rows,value_input_option='USER_ENTERED')
    ws_s.format('A1:C1',FMT['purple'])

    # ── CSV export ─────────────────────────────────────
    if all_rows:
        with open(AI_REPORT_CSV,'w',newline='',encoding='utf-8') as f:
            csv.writer(f).writerow(AI_HEADERS)
            csv.writer(f).writerows(all_rows)
        print(f"\n📄 CSV → {AI_REPORT_CSV}")

    print(f"\n🎯 Tabs written:")
    print(f"   🤖 Migrate Now     — {len(migrate_rows):,} files")
    print(f"   👀 Review Required — {len(review_rows):,} files")
    print(f"   📦 Archive         — {len(archive_rows):,} files")
    print(f"   ⚠️  Risk Flagged    — {len(risk_rows):,} files")
    print(f"   👤 Customer Map    — {len(customer_map):,} customers")


if __name__ == '__main__':
    main()
