"""
UTILITY — token_refresh.py
Generates a fresh token.pickle via browser OAuth and prints the base64
string ready to paste as the GOOGLE_TOKEN GitHub secret.

Run locally (NOT in CI) whenever the token expires or scopes change:
  python token_refresh.py

Requirements:
  pip install google-auth google-auth-oauthlib
"""

import base64
import os
import pickle

from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

CREDENTIALS_FILE = 'credentials.json'
TOKEN_FILE       = 'token.pickle'

# Full read + write scopes needed for:
#   - drive          → backup.py uploads .xlsx to Drive
#   - drive          → ceo_dashboard_writer.py uploads to CEO folder
#   - spreadsheets   → all scripts read and write Google Sheets
SCOPES = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets',
]


def refresh_token():
    creds = None

    # Delete old token to force fresh login with new scopes
    if os.path.exists(TOKEN_FILE):
        print("🗑  Deleting old token to apply new scopes...")
        os.remove(TOKEN_FILE)

    print("🌐 Opening browser for Google OAuth login...")
    print("   Sign in with the Google account that has access to all your sheets.")
    print()

    if not os.path.exists(CREDENTIALS_FILE):
        raise FileNotFoundError(
            f"{CREDENTIALS_FILE} not found.\n"
            "Download from Google Cloud Console → APIs & Services → Credentials → OAuth 2.0"
        )

    flow  = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
    creds = flow.run_local_server(port=0)
    print("✅ New token generated with full read+write scopes.")

    with open(TOKEN_FILE, 'wb') as f:
        pickle.dump(creds, f)
    print(f"💾 Saved to {TOKEN_FILE}")

    with open(TOKEN_FILE, 'rb') as f:
        b64 = base64.b64encode(f.read()).decode()

    print("\n" + "═" * 60)
    print("Copy the value below and save it as the GitHub secret")
    print("  Settings → Secrets → Actions → GOOGLE_TOKEN → Update")
    print("═" * 60)
    print(b64)
    print("═" * 60)
    print(f"\nToken expiry: {creds.expiry}")
    print("\nScopes granted:")
    for s in creds.scopes or SCOPES:
        print(f"  ✅ {s}")


if __name__ == '__main__':
    refresh_token()
