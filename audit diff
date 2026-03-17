"""
UTILITY — audit_diff.py
Compares two audit_data.json snapshots and shows:
  • New files added since last run
  • Files deleted / moved
  • Files whose ERP_Ready status changed
  • Score changes (±10 or more)

Usage:
  python audit_diff.py audit_data_prev.json audit_data.json

Or if you use the GitHub artifact naming convention:
  python audit_diff.py audit-data-42/audit_data.json audit-data-43/audit_data.json
"""

import json
import sys
import csv
from datetime import datetime

DIFF_REPORT = 'audit_diff_report.csv'


def load(path):
    with open(path, 'r') as f:
        records = json.load(f)
    # Key by (Name, Folder_Path) tuple for reliable matching
    return {
        (r.get('Name', r.get('name', '')),
         r.get('Folder_Path', r.get('folder_path', ''))): r
        for r in records
    }


def diff(prev_path, curr_path):
    prev = load(prev_path)
    curr = load(curr_path)

    prev_keys = set(prev.keys())
    curr_keys = set(curr.keys())

    added   = curr_keys - prev_keys
    removed = prev_keys - curr_keys
    common  = prev_keys & curr_keys

    rows = []

    for key in sorted(added):
        r = curr[key]
        rows.append({
            'Change':      'NEW FILE',
            'Name':         key[0],
            'Folder_Path':  key[1],
            'Old_Status':   '',
            'New_Status':   r.get('Status', ''),
            'Old_Score':    '',
            'New_Score':    r.get('Score', ''),
            'Old_ERP':      '',
            'New_ERP':      r.get('ERP_Ready', ''),
        })

    for key in sorted(removed):
        r = prev[key]
        rows.append({
            'Change':      'DELETED / MOVED',
            'Name':         key[0],
            'Folder_Path':  key[1],
            'Old_Status':   r.get('Status', ''),
            'New_Status':   '',
            'Old_Score':    r.get('Score', ''),
            'New_Score':    '',
            'Old_ERP':      r.get('ERP_Ready', ''),
            'New_ERP':      '',
        })

    for key in sorted(common):
        p = prev[key]
        c = curr[key]
        changes = []

        old_erp = p.get('ERP_Ready', 'NO')
        new_erp = c.get('ERP_Ready', 'NO')
        if old_erp != new_erp:
            changes.append(f'ERP {old_erp}→{new_erp}')

        old_score = int(p.get('Score', 0))
        new_score = int(c.get('Score', 0))
        if abs(new_score - old_score) >= 10:
            changes.append(f'Score {old_score:+d}→{new_score:+d}')

        if changes:
            rows.append({
                'Change':      ' | '.join(changes),
                'Name':         key[0],
                'Folder_Path':  key[1],
                'Old_Status':   p.get('Status', ''),
                'New_Status':   c.get('Status', ''),
                'Old_Score':    old_score,
                'New_Score':    new_score,
                'Old_ERP':      old_erp,
                'New_ERP':      new_erp,
            })

    print(f"\n📊 Diff summary")
    print(f"   New files:    {len(added)}")
    print(f"   Deleted:      {len(removed)}")
    print(f"   Changed:      {len([r for r in rows if '→' in r['Change']])}")
    print(f"   Total deltas: {len(rows)}")

    if rows:
        with open(DIFF_REPORT, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=rows[0].keys())
            writer.writeheader()
            writer.writerows(rows)
        print(f"\n💾 Diff saved → {DIFF_REPORT}")
    else:
        print("\n✅ No significant changes between runs.")


if __name__ == '__main__':
    if len(sys.argv) != 3:
        print("Usage: python audit_diff.py <prev_audit_data.json> <curr_audit_data.json>")
        sys.exit(1)
    diff(sys.argv[1], sys.argv[2])
