"""
UTILITY — token_refresh.py
Generates a fresh token.pickle via browser OAuth and prints the base64
string ready to paste as the GOOGLE_TOKEN GitHub secret.

Run locally (NOT in CI) whenever the token expires:
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

SCOPES = [
    'https://www.googleapis.com/auth/drive.readonly',
    'https://www.googleapis.com/auth/spreadsheets.readonly',
]


def refresh_token():
    creds = None

    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, 'rb') as f:
            creds = pickle.load(f)

    if creds and creds.valid:
        print("✅ Existing token is still valid — no browser login needed.")
    elif creds and creds.expired and creds.refresh_token:
        print("🔄 Refreshing expired token...")
        creds.refresh(Request())
        print("✅ Token refreshed successfully.")
    else:
        print("🌐 Opening browser for Google OAuth login...")
        if not os.path.exists(CREDENTIALS_FILE):
            raise FileNotFoundError(
                f"{CREDENTIALS_FILE} not found. Download it from "
                "Google Cloud Console → APIs & Services → Credentials."
            )
        flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
        creds = flow.run_local_server(port=0)
        print("✅ New token generated.")

    with open(TOKEN_FILE, 'wb') as f:
        pickle.dump(creds, f)
    print(f"💾 Saved to {TOKEN_FILE}")

    with open(TOKEN_FILE, 'rb') as f:
        b64 = base64.b64encode(f.read()).decode()

    print("\n" + "═" * 60)
    print("Copy the value below and save it as the GitHub secret")
    print("  Settings → Secrets → Actions → GOOGLE_TOKEN")
    print("═" * 60)
    print(b64)
    print("═" * 60 + "\n")
    print("Token expiry:", creds.expiry)


if __name__ == '__main__':
    refresh_token()
