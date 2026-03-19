"""
UTILITY — read_sheets.py
One-time deep reader. Connects to all your Google Sheets via API,
reads structure, data, IMPORTRANGE links, errors, and row counts.
Saves everything to sheet_map.json for analysis.

Run once locally or via GitHub Actions.
Upload sheet_map.json here after it completes.

Run: python read_sheets.py
"""

import json
import os
import pickle
import time
import re
from datetime import datetime

from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

TOKEN_FILE   = 'token.pickle'
OUTPUT_FILE  = 'sheet_map.json'
DELAY        = 0.5   # seconds between API calls

# ═══════════════════════════════════════════════════════
#  ALL YOUR SHEETS — name + ID
# ═══════════════════════════════════════════════════════

SHEETS = {
    # Production & Operations
    'VFL PMS Forgings 2025-2026':              '1NSXi3aHO-lhONXhU43hDEhkSJAsSydTwrF_AXf7EMEw',
    'VFL Actual Dispatch 2025-2026':           '1ImlfLUoFAPgQsTgteAV2A0S7IXDaZDXEyedURsa0uHA',
    'VFL Consumable Report 2025-2026':         '1R7lErrebtcNoVf-TxSzmEcZR5YB_u7gJNw51vm8l89Y',
    'VFL Daily Manpower 2025-2026':            '1dklVQB4ofrxlJxX4un8y4paUbPVkDdyyXkh6Hj3y5pQ',
    'VFL 57F4 Inward 2025-2026':               '1m4oTDyPiOl3_E_33L_WyRjCktHQIMvSkLHzgGw9q-zY',
    'VFL 57F4 Outward 2025-2026':              '13hnkD-VCdu54yjMClMVkabR5iCZIzWo1goXboQMQJLE',
    'Summary of 57F4 Inward & Outward':        '1Kyt3pQjrysEEnmXM2AIWKL4siODoVbDC9RlWtcQN9dc',

    # Raw Material
    'VFL RM Steel Inward Form Responses':      '16SNs4WQgM--iwLg2wAQLhpmfzpTAc9K-VRSbIftOiPQ',
    'VFL Oil Inward Form':                     '1jvJUviVgjPWkTO8DYkprTGAuazf6TMDoOT2GwUBdfNo',

    # Main Dashboard
    'VFPL Dashboard 2025-2026':                '1NR8EPGRJN0AQDXZjYw5k93clsO8AD4u2l2Xke1lBC2I',

    # HR & People
    'VFPL People Directory':                   '13t_8abt-bc6GjE4fbij9kd93K6uyRmSy39Je7xKTu3o',
    'VFPL Recruitment Tracker':                '1BV7fBiIt4DFnDL4r_X0A0Rc6uTWW0dSQEPA8IE2Qlvk',
    'VFPL Worker Salary Sheet v1':             '1tLyvhgFNRxflS7yTQNfOxhyxNIxQL9olFbiTq4nUagY',
    'Skill Matrix Responses':                  '1BNER3ugBsGyN6S6DKNFlej_hI_3uyvoGNipXQnhBuOo',
    'VFPL Monthly KRA KPI Evaluation':         '1a62b_-FBX_2wlrl4PiIzjhVM9lGcYEYew4W0Yt2zR3Y',
    'Training Attendance':                     '1GMsG08oI-Fp1ZSPhcrdGuw35mncYTLK-iooSzVcdbf0',
    'VFL Leave Application Responses':         '1h3yD9I_CtqNVa2RNfUd4Dcd9-PrNvASkPSe0296lTfM',
    'VFL Overtime Form Responses':             '11TquszkTamTBlDsfEQUIUu2bd9ZrAhQmXSYQhyH6qss',
    'VFPL Interview Feedback Responses':       '1jBqkWI6jocSDjmv93XhpEjRSGITOhWjS32U_gddsWwo',
    'VFL Late Attendance Form Responses':      '1pwVE0XKqAhAKHbyqtlF9GzfuGnidnZuw2zKbtMjUz9Q',
    'VFPL Attendance Form With Location':      '1Tf3t5i5BTUQCnodbGW_5OL6-9qHayV2iCSruJU62q14',
    'Daily Contractual Manpower Form':         '1wf0bLbsy1en7r61foC4hJaykXF01FuLPMFR9iltdbSU',
    'Exit Clearance Form Responses':           '1wTe9wWv8D-0WIZBhAe18rVYR4vYPEwJVz52Up-GIAXc',
    'VFPL Exit Interview Responses':           '1AssFUO5PJZLUzZCFINlqwGw9mckICIRnsYxHCmuhVkM',
    'PIP Responses':                           '12Ct4iA0ZnxoU1xq-Aa4AObPBV_26EAvqa1G7p2DNxPE',
    'VFPL Appointment Letter Generator':       '1-izyJt10gFID0jy7zyiVSk0EI3BaE1MMxXnkFCUxCrQ',
    'VFPL Confirmation Review Responses':      '1DVBA45-YkFaddci36lbO667BH7pJuXCkXDWO52JPqVo',
    'VFL Work With Us Responses':              '1CNCb5t7fGAaQGp_t1VFDqShdecXNS0FKRGx_o62VLoo',
    'VFPL Inductee Feedback Form 30 days':     '1m5TDKINIL-ff9BwiGdjz0R8lnNjMq_rYyQTuSW7Wz14',
    'VFPL Departments Introduction':           '1He3cZmRMbGoGVT98oJkbZeOUc_q6xc53ArT-u7M1Rtc',
    'VFPL Asset Equipment Issue Form':         '1qODK9GAI5SEzxl5d6I4Der4OrX_WaQp6dJjN7kWfakE',
    'VFPL Courier Form Responses':             '1jAVi5kpfhZ-zPg80S0Fy-G2wVHJKaXAJ9qBdJzD6U30',
    'Training Feedback Form Responses':        '1Vmi7gzDSsepr0_XsWlekVWEAe4MDO28xJvNhUm-oqPU',
    'VFPL Medicine Issue Form Responses':      '1NBgUNDgR92Mx2MXNprxu6wL784KsFoVVgTrW_2rIA8g',

    # Quality & Compliance
    'Weekly House Keeping Responses':          '1yNAojZR7qbIlJXOEhqP16RflZpnI8dsGb0meVQql6tg',
    'Weekly 5S System':                        '1KAW4PeHx5sesE64qef66ku3Jmn9nbkV01RNnaKsVKc4',
    'Usage of Safety PPE Form':                '145iIzbCKR6pT9TXVHdYklHaYMW7QtgC0hcq6_SMi2Ms',
    'Breakdown Slip':                          '1jCELd65TbSi-FDGBB2jvzSfgZu2txvAoo3RQEG_qxbc',
    'Die Machine Setting Form':                '1Nt43-MjLWIxsLLdWus8VoVIW8ek_M-AiZEeHe4DuFPg',
    'Die Failure Form':                        '1xNjusx4-r2fZ3XliCUll-HvBl6Qk-vey8soNdEAsxn4',
    'Drop Outs Form':                          '1J1KKfygB8M1z3qZJQcaMXjkVb7HgxDARZrklMld8gS0',
    'Vendor Rejection':                        '1izr5HWo4-qvXbmhkbmBIVQSR-X6ptU7OtSX-ICd43H0',

    # Commercial & Admin
    'VFPL CRM':                                '1wZlHu1iEzS1yMcGm2WDJcX9nz4uAHaHaPxNP8o8_HXE',
    'VFPL NPD Status':                         '1t0BRmrUHVgy1PYvoC2XD7ovUz75b7A-bjv8hdtM8lRI',
    'POA Tracker Form':                        '1mk3O5b8-XVsW1ol6F-EuOno1g7-Sdnq2J-nBiuAgcTQ',
    'VFL Gate Pass':                           '1sVmWRovbkRki825PmVC065C7u2fBvOlM5-_QvBaDnpA',
    'Consultant Attendance Sheet':             '1OXDSDfkVBjodoixd7zXWtviGX8hTeEbVyARUd0JQSZo',
    'Letter Inward Form':                      '1WQQeRaCg6HzwFfplwFaBIMnQkqAq0t3EW3KP0UJ88nw',
    'VFPL Vehicle Diesel Form Responses':      '1i8fwXoKp3RxJqMun4B2R2tAgegj65tk2aqR-iBIPQf8',
    'Diesel Form Forklift Generator':          '1OCDff85Tqop2yrtc4tn2kXAKj7TdRwvNmdq21K-5GPE',
    'VFL Diesel Brought Into Plant':           '1iP4Ikp-K3k3m7iwY-YJ51X0Sw6KCI5EWiYFEr22-wv0',
    'VFPL Advance Loan Form Responses':        '1uX7LqWXwhzA1sLu3TUS-WJ_JLeW1JdOmLeDgNhf2Zfg',
    'RM Consumption Form':                     '1yRsAs7zV0HvH63Ah64D0fcttNAhmrRI6AX1MHKvRyxg',
    'VFPL PMSM Machine Shop 2hr Report':       '1rLKj2r8-M0621sx_OVCfp-3n2hNlspEsrlRDk8mg_4g',
    'VFL Visitor Automated Greetings':         '1qfbnUhvtmyt2-J-pJzJaaQ-6xT2yOnneWrfvAFuELhU',
    'VFL Automated Festive Greetings':         '1txZM9a9_OSG-ZWYaAEBLKj-9M7LYLrFyl0kJkPsTMGI',
    'VFL Greetings Birthdays Anniversaries':   '1Jkk2SHtgN8azGVmgA_2QJ2QXNsPaJIFQ2RCvscxJJfU',
    'Rent Invoice Generator CEOIBOX':          '1QVX6chu4mR6gkWvYBQS1WsUvVBvTknfCzlPuj-TB59w',
    'VFPL Experience Letter':                  '1FQ2UqNnuP_WQ2-3KTnmtXhVNJk5BFysC32-ehbpayRs',
    'VFL Employee Salary Slip Generator':      '19bbRksXWokLUT7EdfBFFz3WY-stleVcMCEEQthP54MY',
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
#  IMPORTRANGE DETECTOR
# ═══════════════════════════════════════════════════════

def find_importrange(values):
    """Find all IMPORTRANGE URLs in cell values."""
    found = []
    flat  = str(values)
    urls  = re.findall(
        r'https://docs\.google\.com/spreadsheets/d/([a-zA-Z0-9_-]+)',
        flat
    )
    for uid in set(urls):
        found.append(f'https://docs.google.com/spreadsheets/d/{uid}')
    return found


def find_errors(values):
    """Find formula errors in cell values."""
    errors = []
    flat   = str(values)
    for err in ['#DIV/0!', '#REF!', '#N/A', '#VALUE!', '#NAME?', '#NULL!']:
        if err in flat:
            errors.append(err)
    return errors

# ═══════════════════════════════════════════════════════
#  SHEET READER
# ═══════════════════════════════════════════════════════

def read_sheet(service, sheet_id, sheet_name):
    """Read structure and sample data from a Google Sheet."""
    result = {
        'name':       sheet_name,
        'id':         sheet_id,
        'tabs':       [],
        'errors':     [],
        'importrange':[],
        'read_ok':    False,
        'error_msg':  '',
    }

    try:
        # Get all tab names
        meta = service.spreadsheets().get(
            spreadsheetId=sheet_id,
            fields='sheets.properties.title,sheets.properties.gridProperties'
        ).execute()
        time.sleep(DELAY)

        tabs = meta.get('sheets', [])

        for tab in tabs[:10]:   # max 10 tabs per sheet
            title    = tab['properties']['title']
            row_count= tab['properties'].get('gridProperties',{}).get('rowCount',0)

            tab_result = {
                'title':      title,
                'row_count':  row_count,
                'headers':    [],
                'sample':     [],
                'importrange':[],
                'errors':     [],
            }

            try:
                # Read first 5 rows
                resp = service.spreadsheets().values().get(
                    spreadsheetId=sheet_id,
                    range=f"'{title}'!A1:Z5"
                ).execute()
                time.sleep(DELAY)

                values = resp.get('values', [])
                if values:
                    # First non-empty row = headers
                    tab_result['headers'] = [str(c) for c in values[0]]
                    tab_result['sample']  = [
                        [str(c) for c in row]
                        for row in values[1:4]
                    ]

                # Check for IMPORTRANGE and errors in first 50 rows
                resp50 = service.spreadsheets().values().get(
                    spreadsheetId=sheet_id,
                    range=f"'{title}'!A1:Z50"
                ).execute()
                time.sleep(DELAY)

                vals50 = resp50.get('values', [])
                tab_result['importrange'] = find_importrange(vals50)
                tab_result['errors']      = find_errors(vals50)

                # Add to sheet-level lists
                result['importrange'].extend(tab_result['importrange'])
                result['errors'].extend(tab_result['errors'])

            except HttpError as e:
                tab_result['error'] = str(e)

            result['tabs'].append(tab_result)

        # Deduplicate
        result['importrange'] = list(set(result['importrange']))
        result['errors']      = list(set(result['errors']))
        result['read_ok']     = True

    except HttpError as e:
        result['error_msg'] = f"HTTP {e.resp.status}: {e}"
        if e.resp.status == 429:
            print(f"  ⏳ Quota — waiting 60s")
            time.sleep(60)
    except Exception as e:
        result['error_msg'] = str(e)

    return result

# ═══════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════

def main():
    print('=' * 60)
    print('   SHEET MAP READER')
    print(f'   Reading {len(SHEETS)} sheets')
    print('=' * 60)

    creds   = authenticate()
    service = build('sheets', 'v4', credentials=creds)

    results   = {}
    total     = len(SHEETS)
    completed = 0
    errors    = 0

    # Resume if interrupted
    if os.path.exists(OUTPUT_FILE):
        with open(OUTPUT_FILE) as f:
            results = json.load(f)
        completed = len(results)
        print(f"♻️  Resuming: {completed} sheets already read.")

    for name, sid in SHEETS.items():
        if name in results:
            continue

        print(f"\n[{completed+1}/{total}] 📋 {name}")

        data = read_sheet(service, sid, name)

        if data['read_ok']:
            tab_names = [t['title'] for t in data['tabs']]
            print(f"  Tabs: {', '.join(tab_names[:5])}")
            if data['errors']:
                print(f"  ⚠ Errors found: {data['errors']}")
            if data['importrange']:
                print(f"  🔗 IMPORTRANGE: {len(data['importrange'])} link(s)")
        else:
            print(f"  ❌ Could not read: {data['error_msg']}")
            errors += 1

        results[name] = data
        completed    += 1

        # Save after every sheet
        with open(OUTPUT_FILE, 'w') as f:
            json.dump(results, f, indent=2)

    print(f"\n{'='*60}")
    print(f"✅ Complete — {completed} sheets read, {errors} errors")
    print(f"📄 Saved to {OUTPUT_FILE}")
    print(f"\nNext step:")
    print(f"  Upload {OUTPUT_FILE} here and I will analyse everything.")

    # Print summary
    all_errors     = []
    all_importrange= []
    sheet_summary  = []

    for name, data in results.items():
        if data.get('errors'):
            all_errors.append(f"{name}: {data['errors']}")
        if data.get('importrange'):
            all_importrange.append(f"{name}: {len(data['importrange'])} links")
        tab_count = len(data.get('tabs', []))
        row_total = sum(t.get('row_count', 0) for t in data.get('tabs', []))
        sheet_summary.append(f"  {name[:45]:<45} {tab_count:>3} tabs  {row_total:>8,} rows")

    print(f"\n📊 SHEET SUMMARY:")
    for s in sheet_summary:
        print(s)

    if all_errors:
        print(f"\n⚠ FORMULA ERRORS FOUND:")
        for e in all_errors:
            print(f"  {e}")

    if all_importrange:
        print(f"\n🔗 IMPORTRANGE CONNECTIONS:")
        for i in all_importrange:
            print(f"  {i}")


if __name__ == '__main__':
    main()
