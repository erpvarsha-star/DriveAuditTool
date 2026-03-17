"""
UTILITY — validate_env.py
Pre-flight environment checker. Run this before any audit job to catch
missing secrets, expired tokens, and unreachable APIs early — not
mid-scan after 40 minutes of work.

Exit codes:
  0  All checks passed
  1  One or more checks failed

Usage (standalone):
  python validate_env.py

Usage in GitHub Actions (add as first step after token restore):
  - name: Validate environment
    run: python validate_env.py
"""

import json
import os
import pickle
import sys
from datetime import datetime, timezone

# ──────────────────────────────────────────────────────
REQUIRED_FILES = ['credentials.json', 'token.pickle']
REQUIRED_ENV   = []          # add e.g. 'SLACK_WEBHOOK' if you want to enforce it

SCOPES = [
    'https://www.googleapis.com/auth/drive.readonly',
    'https://www.googleapis.com/auth/spreadsheets.readonly',
]
# ──────────────────────────────────────────────────────

PASS = '  ✅'
FAIL = '  ❌'
WARN = '  ⚠️ '

errors   = []
warnings = []


def check(label, condition, error_msg, warn=False):
    if condition:
        print(f'{PASS} {label}')
    else:
        symbol = WARN if warn else FAIL
        print(f'{symbol} {label} — {error_msg}')
        (warnings if warn else errors).append(error_msg)


# ─── File presence ────────────────────────────────────
print('\n📋 Checking required files...')
for fname in REQUIRED_FILES:
    check(
        f'{fname} exists',
        os.path.exists(fname),
        f'{fname} not found in working directory',
    )

# ─── credentials.json structure ───────────────────────
print('\n🔑 Checking credentials.json...')
if os.path.exists('credentials.json'):
    try:
        with open('credentials.json') as f:
            cred_data = json.load(f)
        has_installed = 'installed' in cred_data or 'web' in cred_data
        check('credentials.json is valid OAuth JSON', has_installed,
              'Expected "installed" or "web" key — re-download from Google Cloud Console')
    except json.JSONDecodeError as e:
        check('credentials.json is valid JSON', False, str(e))

# ─── token.pickle validity ────────────────────────────
print('\n🎟  Checking token.pickle...')
if os.path.exists('token.pickle'):
    try:
        with open('token.pickle', 'rb') as f:
            creds = pickle.load(f)

        check('token.pickle loaded successfully', True, '')

        valid = creds is not None and creds.valid
        check('Token is currently valid', valid,
              'Token is expired or invalid — run token_refresh.py locally')

        if creds and creds.expiry:
            days_left = (creds.expiry.replace(tzinfo=timezone.utc)
                         - datetime.now(timezone.utc)).days
            check(
                f'Token expiry — {days_left} days remaining',
                days_left > 3,
                f'Token expires in {days_left} days — refresh soon',
                warn=(0 < days_left <= 3),
            )

        has_refresh = bool(getattr(creds, 'refresh_token', None))
        check('Refresh token present', has_refresh,
              'No refresh token — token cannot auto-renew; re-run token_refresh.py',
              warn=True)

        scope_ok = all(s in (creds.scopes or []) for s in SCOPES)
        check('Token has required scopes', scope_ok,
              f'Missing scopes. Expected: {SCOPES}')

    except Exception as e:
        check('token.pickle is readable', False, str(e))

# ─── Optional env vars ────────────────────────────────
print('\n🌍 Checking environment variables...')
for var in REQUIRED_ENV:
    check(f'{var} set', bool(os.environ.get(var)), f'{var} env var is missing')

slack = os.environ.get('SLACK_WEBHOOK', '')
check('SLACK_WEBHOOK configured (optional)', bool(slack),
      'SLACK_WEBHOOK not set — notifications will be skipped', warn=True)

# ─── Result ───────────────────────────────────────────
print('\n' + '═' * 50)
if errors:
    print(f'❌ {len(errors)} error(s) found — audit will likely fail:\n')
    for e in errors:
        print(f'   • {e}')
    print()
    sys.exit(1)
elif warnings:
    print(f'⚠️  {len(warnings)} warning(s) — audit can proceed but check these:\n')
    for w in warnings:
        print(f'   • {w}')
    print()
    sys.exit(0)
else:
    print('✅ All checks passed — ready to audit.')
    sys.exit(0)
