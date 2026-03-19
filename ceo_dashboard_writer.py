"""
STAGE 5 — CEO DASHBOARD WRITER (Google Drive hosted version)
Reads live operational data from Google Sheets and writes
ceo_data.json + ceo.html to a Google Drive folder.

CEO opens the shared Drive folder link — clicks ceo.html — sees live dashboard.
No GitHub Pages needed. Repo stays private.

Drive folder: https://drive.google.com/drive/folders/1w7Vn-l7SihOaiJheKvJCbCHZ7WBeNdyt

Run: python ceo_dashboard_writer.py
"""

import json
import os
import pickle
import time
import io
from datetime import datetime, timezone, timedelta

from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseUpload

# ═══════════════════════════════════════════════════════
#  CONFIG
# ═══════════════════════════════════════════════════════

DASHBOARD_ID    = '1NR8EPGRJN0AQDXZjYw5k93clsO8AD4u2l2Xke1lBC2I'
PMS_ID          = '1c-axqiBEufNb1vK-JdJP6t1eRTmzliA3otAbqxAOXZE'
DISPATCH_ID     = '1txZM9a9_OSG-ZWYaAEBLKj-9M7LYLrFyl0kJkPsTMGI'
MANPOWER_ID     = '1t7UjWTP_cpIJ2BjoaMlV6uUKA7ztH9UnKc_korYKCiw'
ELECTRICITY_ID  = '1nUvf-UWjBSbSWnZTNph-gRUbjzuguGlidpYBKshKUNQ'
MACHINE_SHOP_ID = '1yC-b36rgAxablmdXhngHCEOsmnlgWourep6ctKifXCA'

# CEO Dashboard Drive folder
CEO_FOLDER_ID   = '1w7Vn-l7SihOaiJheKvJCbCHZ7WBeNdyt'

TOKEN_FILE      = 'token.pickle'
LOCAL_JSON      = 'ceo_data.json'
DELAY           = 0.5

# Targets — update these to match your actual targets
TARGETS = {
    'rm_stock_min_tons':       50,
    'manpower_required':       120,
    'electricity_daily_units': 8000,
    'oil_min_stock_ltrs':      500,
    'rejection_max_pct':       2.0,
    'daily_revenue_target':    3000000,
    'monthly_revenue_target':  60000000,
}

# ═══════════════════════════════════════════════════════
#  AUTH — needs drive scope for uploading
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
#  HELPERS
# ═══════════════════════════════════════════════════════

def read_range(svc, sheet_id, range_str):
    try:
        resp = svc.spreadsheets().values().get(
            spreadsheetId=sheet_id, range=range_str
        ).execute()
        time.sleep(DELAY)
        return resp.get('values', [])
    except HttpError as e:
        if e.resp.status == 429:
            time.sleep(60)
            return read_range(svc, sheet_id, range_str)
        return []
    except Exception:
        return []


def get_tabs(svc, sheet_id):
    try:
        meta = svc.spreadsheets().get(
            spreadsheetId=sheet_id,
            fields='sheets.properties.title'
        ).execute()
        time.sleep(DELAY)
        return [s['properties']['title'] for s in meta.get('sheets', [])]
    except Exception:
        return []


def find_tab(tabs, keywords):
    for kw in keywords:
        for tab in tabs:
            if kw.lower() in tab.lower():
                return tab
    return tabs[0] if tabs else None


def safe_float(v, default=0):
    try:
        return float(str(v).replace(',','').replace('₹','').replace(' ','').replace('%',''))
    except Exception:
        return default


def today_ist():
    ist = timezone(timedelta(hours=5, minutes=30))
    return datetime.now(ist)


def status(value, target, higher_is_better=True):
    if not target or value is None:
        return 'GREY'
    try:
        v = safe_float(value)
        t = safe_float(target)
        if higher_is_better:
            return 'GREEN' if v >= t else 'RED'
        else:
            return 'GREEN' if v <= t else 'RED'
    except Exception:
        return 'GREY'

# ═══════════════════════════════════════════════════════
#  METRIC READERS
# ═══════════════════════════════════════════════════════

def get_rm_stock(svc):
    print("  📦 Reading RM Stock...")
    tabs = get_tabs(svc, DASHBOARD_ID)
    tab  = find_tab(tabs, ['rm stock', 'rm inventory', 'raw material'])

    if not tab:
        return {'total_tons': 0, 'grades': [], 'status': 'GREY'}

    rows = read_range(svc, DASHBOARD_ID, f"'{tab}'!A1:H100")
    if not rows:
        return {'total_tons': 0, 'grades': [], 'status': 'GREY'}

    grades = []
    total  = 0

    for row in rows[1:]:
        if not row or not row[0].strip():
            continue
        grade   = str(row[0]).strip()
        opening = safe_float(row[1]) if len(row) > 1 else 0
        inward  = safe_float(row[3]) if len(row) > 3 else 0
        consump = safe_float(row[5]) if len(row) > 5 else 0
        closing = opening + inward - consump

        if grade and grade.lower() not in ['grade', '']:
            grades.append({
                'grade':       grade,
                'opening':     round(opening, 2),
                'inward':      round(inward, 2),
                'consumption': round(consump, 2),
                'closing':     round(closing, 2),
            })
            total += max(closing, 0)

    return {
        'total_tons':  round(total, 2),
        'grades':      grades[:30],
        'status':      status(total, TARGETS['rm_stock_min_tons']),
        'target_tons': TARGETS['rm_stock_min_tons'],
    }


def get_revenue(svc):
    print("  💰 Reading Revenue...")
    dash_tabs = get_tabs(svc, DASHBOARD_ID)
    erp_tab   = find_tab(dash_tabs, ['erp daily', 'daily ton', 'erp'])

    today_ist_dt = today_ist()
    today_rev    = 0
    today_tons   = 0
    mtd_rev      = 0
    mtd_tons     = 0
    records      = []

    if erp_tab:
        rows = read_range(svc, DASHBOARD_ID, f"'{erp_tab}'!A1:F2000")
        if rows and len(rows) > 1:
            headers = [str(h).strip() for h in rows[0]]
            print(f"    ERP Daily headers: {headers[:6]}")
            for row in rows[1:]:
                if len(row) < 3:
                    continue
                date    = str(row[1]).strip() if len(row) > 1 else ''
                fg_tons = safe_float(row[2]) if len(row) > 2 else 0
                fg_val  = safe_float(row[3]) if len(row) > 3 else 0
                jw_tons = safe_float(row[4]) if len(row) > 4 else 0
                jw_val  = safe_float(row[5]) if len(row) > 5 else 0

                t_tons = fg_tons + jw_tons
                t_val  = fg_val  + jw_val

                mtd_rev  += t_val
                mtd_tons += t_tons

                # Check if today
                today_str = today_ist_dt.strftime('%d/%m/%Y')
                if today_str in date or date == today_str:
                    today_rev  = t_val
                    today_tons = t_tons

                records.append({
                    'date':     date,
                    'fg_tons':  fg_tons,
                    'fg_value': fg_val,
                    'jw_tons':  jw_tons,
                    'jw_value': jw_val,
                    'total_value': t_val,
                    'total_tons':  t_tons,
                })

    # Read overdue
    overdue    = 0
    od_tab     = find_tab(get_tabs(svc, DASHBOARD_ID), ['overdue', 'outstanding'])
    if od_tab:
        od_rows = read_range(svc, DASHBOARD_ID, f"'{od_tab}'!A1:Z5")
        # Try to find overdue value
        for row in od_rows:
            for cell in row:
                if 'cr' in str(cell).lower() or ',' in str(cell):
                    try:
                        v = safe_float(str(cell).replace('Cr','').replace('cr',''))
                        if 0 < v < 100:
                            overdue = v
                            break
                    except Exception:
                        pass

    return {
        'today_revenue':    today_rev,
        'today_tons':       round(today_tons, 2),
        'mtd_revenue':      round(mtd_rev, 0),
        'mtd_tons':         round(mtd_tons, 2),
        'daily_target':     TARGETS['daily_revenue_target'],
        'monthly_target':   TARGETS['monthly_revenue_target'],
        'status_today':     status(today_rev, TARGETS['daily_revenue_target']),
        'status_mtd':       status(mtd_rev,   TARGETS['monthly_revenue_target']),
        'overdue_cr':       round(overdue, 2),
        'last_7_days':      records[-7:]  if records else [],
        'last_30_days':     records[-30:] if records else [],
    }


def get_production(svc):
    print("  🏭 Reading Production...")
    tabs = get_tabs(svc, PMS_ID)
    tab  = find_tab(tabs, ['form response', 'responses', 'data', 'production'])

    mtd_qty    = 0
    mtd_tons   = 0
    vf_data    = {}

    if tab:
        rows = read_range(svc, PMS_ID, f"'{tab}'!A1:Z500")
        if rows and len(rows) > 1:
            headers = [str(h).strip() for h in rows[0]]
            print(f"    PMS headers: {headers[:10]}")
            for row in rows[1:]:
                padded = row + [''] * max(0, len(headers)-len(row))
                r      = dict(zip(headers, padded))
                vf     = r.get('VF No', r.get('Machine', r.get('VF_No',''))).strip()
                qty    = safe_float(r.get('Qty Forged', r.get('Quantity', r.get('Qty', 0))))
                rej    = safe_float(r.get('Rejection', r.get('Rejected Qty', 0)))
                wt     = safe_float(r.get('Output Weight', r.get('Finish Weight', 0)))

                mtd_qty  += qty
                mtd_tons += wt / 1000 if wt > 100 else wt

                if vf:
                    if vf not in vf_data:
                        vf_data[vf] = {'qty': 0, 'rejection': 0}
                    vf_data[vf]['qty']       += qty
                    vf_data[vf]['rejection'] += rej

    vf_rejection = []
    for vf, d in sorted(vf_data.items())[:25]:
        rej_pct = round(d['rejection']/d['qty']*100, 2) if d['qty'] > 0 else 0
        vf_rejection.append({
            'vf_no':         vf,
            'qty':           int(d['qty']),
            'rejection':     int(d['rejection']),
            'rejection_pct': rej_pct,
            'status':        status(rej_pct, TARGETS['rejection_max_pct'],
                                   higher_is_better=False),
        })

    return {
        'mtd_qty':          int(mtd_qty),
        'mtd_tons':         round(mtd_tons, 2),
        'vf_rejection':     vf_rejection,
        'rejection_target': TARGETS['rejection_max_pct'],
    }


def get_electricity(svc):
    print("  ⚡ Reading Electricity...")
    tabs = get_tabs(svc, ELECTRICITY_ID)
    print(f"    Tabs: {tabs}")
    tab  = find_tab(tabs, ['electricity', 'consumption', 'form response', 'data'])

    mtd_units = 0
    records   = []

    if tab:
        rows = read_range(svc, ELECTRICITY_ID, f"'{tab}'!A1:Z200")
        if rows and len(rows) > 1:
            headers = [str(h).strip() for h in rows[0]]
            print(f"    Electricity headers: {headers[:8]}")
            for row in rows[1:]:
                padded = row + [''] * max(0, len(headers)-len(row))
                r      = dict(zip(headers, padded))
                units  = safe_float(r.get('Units', r.get('Electricity Units',
                         r.get('Consumption', r.get('KWH',
                         r.get('Unit', 0))))))
                date   = r.get('Date', r.get('Timestamp', ''))
                if units:
                    mtd_units += units
                    records.append({'date': str(date)[:10], 'units': units})

    return {
        'mtd_units':    round(mtd_units, 0),
        'daily_target': TARGETS['electricity_daily_units'],
        'status':       'GREY',
        'records':      records[-30:],
    }


def get_manpower(svc):
    print("  👥 Reading Manpower...")
    dash_tabs = get_tabs(svc, DASHBOARD_ID)
    mp_tab    = find_tab(dash_tabs, ['manpower consolidat', 'daily manpower',
                                      'manpower report', 'manpower'])
    present = 0
    depts   = []

    if mp_tab:
        rows = read_range(svc, DASHBOARD_ID, f"'{mp_tab}'!A1:Z10")
        print(f"    Manpower tab: {mp_tab}, rows: {len(rows)}")
        if rows:
            headers = [str(h).strip() for h in rows[0]] if rows else []
            print(f"    Headers: {headers[:8]}")

    # Also try manpower sheet directly
    mp_tabs2 = get_tabs(svc, MANPOWER_ID)
    mp_tab2  = find_tab(mp_tabs2, ['form response', 'data', 'manpower'])
    if mp_tab2:
        rows2 = read_range(svc, MANPOWER_ID, f"'{mp_tab2}'!A1:Z5")
        print(f"    Manpower sheet tab: {mp_tab2}")
        if rows2:
            print(f"    Headers: {[str(h) for h in rows2[0][:8]]}")

    return {
        'present':  present,
        'required': TARGETS['manpower_required'],
        'status':   'GREY' if not present else status(present, TARGETS['manpower_required']),
        'depts':    depts,
    }


def get_oil(svc):
    print("  🛢️  Reading Oil Stock...")
    dash_tabs = get_tabs(svc, DASHBOARD_ID)
    oil_tab   = find_tab(dash_tabs, ['oil stock', 'oil consumption', 'oil'])

    stock = 0
    if oil_tab:
        rows = read_range(svc, DASHBOARD_ID, f"'{oil_tab}'!A1:Z5")
        print(f"    Oil tab: {oil_tab}, rows: {len(rows)}")
        # Try to find stock value
        for row in rows:
            for cell in row:
                v = safe_float(cell)
                if 100 < v < 1000000:  # reasonable oil stock range
                    stock = v
                    break

    return {
        'stock_ltrs': round(stock, 0),
        'min_target': TARGETS['oil_min_stock_ltrs'],
        'status':     status(stock, TARGETS['oil_min_stock_ltrs']),
    }

# ═══════════════════════════════════════════════════════
#  DRIVE UPLOADER
# ═══════════════════════════════════════════════════════

def upload_to_drive(drive_svc, folder_id, filename, content, mime_type):
    """Upload or update a file in Google Drive folder."""
    # Check if file already exists
    results = drive_svc.files().list(
        q=f"name='{filename}' and '{folder_id}' in parents and trashed=false",
        fields='files(id,name)'
    ).execute()
    time.sleep(DELAY)

    existing = results.get('files', [])
    media    = MediaIoBaseUpload(
        io.BytesIO(content.encode('utf-8')),
        mimetype=mime_type,
        resumable=False
    )

    if existing:
        # Update existing file
        file_id = existing[0]['id']
        drive_svc.files().update(
            fileId=file_id,
            media_body=media
        ).execute()
        print(f"  ✅ Updated {filename} in Drive (ID: {file_id})")
    else:
        # Create new file
        meta = {
            'name':    filename,
            'parents': [folder_id],
        }
        f = drive_svc.files().create(
            body=meta,
            media_body=media,
            fields='id'
        ).execute()
        print(f"  ✅ Created {filename} in Drive (ID: {f['id']})")

# ═══════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════

def main():
    print('=' * 60)
    print('   STAGE 5 — CEO Dashboard Writer')
    print(f'   {datetime.now().strftime("%Y-%m-%d %H:%M")} IST')
    print('=' * 60)

    creds     = authenticate()
    sheets    = build('sheets', 'v4', credentials=creds)
    drive_svc = build('drive',  'v3', credentials=creds)

    # Read all metrics
    rm_stock    = get_rm_stock(sheets)
    revenue     = get_revenue(sheets)
    production  = get_production(sheets)
    electricity = get_electricity(sheets)
    manpower    = get_manpower(sheets)
    oil         = get_oil(sheets)

    # Build CEO data payload
    ceo_data = {
        'generated_at':  datetime.now().isoformat(),
        'generated_ist': datetime.now().strftime('%d %b %Y %H:%M IST'),
        'targets':       TARGETS,
        'metrics': {
            'rm_stock':    rm_stock,
            'revenue':     revenue,
            'production':  production,
            'electricity': electricity,
            'manpower':    manpower,
            'oil':         oil,
        },
        'alerts': [],
    }

    # Generate alerts
    if rm_stock['status'] == 'RED':
        ceo_data['alerts'].append({
            'type':    'WARNING',
            'metric':  'RM Stock',
            'message': f"RM Stock {rm_stock['total_tons']}T below minimum {TARGETS['rm_stock_min_tons']}T",
        })
    if revenue['status_today'] == 'RED':
        ceo_data['alerts'].append({
            'type':    'WARNING',
            'metric':  'Revenue',
            'message': f"Today's revenue below daily target",
        })
    if oil['status'] == 'RED':
        ceo_data['alerts'].append({
            'type':    'WARNING',
            'metric':  'Oil Stock',
            'message': f"Oil stock {oil['stock_ltrs']}L below minimum {TARGETS['oil_min_stock_ltrs']}L",
        })
    for vf in production.get('vf_rejection', []):
        if vf['status'] == 'RED':
            ceo_data['alerts'].append({
                'type':    'CRITICAL',
                'metric':  'Rejection',
                'message': f"{vf['vf_no']} rejection {vf['rejection_pct']}% above limit",
            })

    # Save locally
    json_str = json.dumps(ceo_data, indent=2)
    with open(LOCAL_JSON, 'w') as f:
        f.write(json_str)
    print(f"\n💾 Saved locally → {LOCAL_JSON}")

    # Upload to Google Drive
    print(f"\n📤 Uploading to Drive folder...")
    upload_to_drive(drive_svc, CEO_FOLDER_ID, 'ceo_data.json',
                    json_str, 'application/json')

    print(f"\n✅ CEO Dashboard updated")
    print(f"   RM Stock:    {rm_stock['total_tons']}T ({rm_stock['status']})")
    print(f"   MTD Revenue: ₹{revenue['mtd_revenue']:,.0f} ({revenue['status_mtd']})")
    print(f"   Production:  {production['mtd_qty']:,} pcs MTD")
    print(f"   Oil:         {oil['stock_ltrs']}L ({oil['status']})")
    print(f"   Alerts:      {len(ceo_data['alerts'])}")

    if ceo_data['alerts']:
        print("\n⚠ ACTIVE ALERTS:")
        for a in ceo_data['alerts']:
            print(f"   [{a['type']}] {a['metric']}: {a['message']}")

    print(f"\n🔗 CEO Dashboard folder:")
    print(f"   https://drive.google.com/drive/folders/{CEO_FOLDER_ID}")


if __name__ == '__main__':
    main()


def build_self_contained_html(ceo_data):
    """
    Build a self-contained HTML file with CEO data embedded directly.
    No external fetch needed — open the file and see live data.
    Works perfectly from Google Drive.
    """
    # Read the base ceo.html template
    html_path = 'ceo.html'
    if not os.path.exists(html_path):
        print(f"  ⚠ {html_path} not found — skipping self-contained build")
        return None

    with open(html_path, 'r') as f:
        html = f.read()

    # Embed data directly into HTML
    data_script = f"""
<script>
// Data embedded at build time: {ceo_data['generated_ist']}
window.CEO_DATA = {json.dumps(ceo_data, indent=2)};
</script>"""

    # Replace the fetch call with embedded data
    html = html.replace(
        'async function loadData() {',
        f'''{data_script}
async function loadData() {{
  try {{
    DATA = window.CEO_DATA;
    document.getElementById('loading').style.display = 'none';
    render();
    return;
  }} catch(e) {{}}
  try {{'''
    )

    return html
