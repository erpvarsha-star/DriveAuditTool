"""
UTILITY — notify.py
Centralised notification helper.
Supports Slack (via webhook) and plain email (via SMTP / Gmail app password).
Both channels are opt-in via environment variables — missing vars = silent skip.

Environment variables:
  SLACK_WEBHOOK        Slack incoming webhook URL
  NOTIFY_EMAIL_TO      Recipient address (e.g. you@company.com)
  NOTIFY_EMAIL_FROM    Sender address (Gmail recommended)
  NOTIFY_EMAIL_PASS    Gmail App Password (not your main password)

Usage (standalone):
  python notify.py "Audit complete — 1,204 files, 87 ERP-ready"

Usage (from another script):
  from notify import send
  send("Audit complete", status="success", details={"ERP-ready": 87})
"""

import json
import os
import smtplib
import sys
import urllib.request
from email.mime.text import MIMEText
from datetime import datetime


# ═══════════════════════════════════════════════════════
#  SLACK
# ═══════════════════════════════════════════════════════

def slack(message: str, status: str = 'success') -> bool:
    webhook = os.environ.get('SLACK_WEBHOOK', '').strip()
    if not webhook:
        return False

    icon   = ':white_check_mark:' if status == 'success' else ':x:'
    colour = '#2eb886' if status == 'success' else '#cc0000'

    payload = json.dumps({
        'attachments': [{
            'color': colour,
            'text':  f'{icon}  *Drive Audit* — {message}',
            'footer': f'GitHub Actions · {datetime.now().strftime("%Y-%m-%d %H:%M IST")}',
        }]
    }).encode()

    try:
        req = urllib.request.Request(
            webhook, data=payload,
            headers={'Content-Type': 'application/json'},
        )
        urllib.request.urlopen(req, timeout=10)
        print('📣 Slack notification sent.')
        return True
    except Exception as e:
        print(f'⚠ Slack notification failed (non-fatal): {e}')
        return False


# ═══════════════════════════════════════════════════════
#  EMAIL  (Gmail SMTP)
# ═══════════════════════════════════════════════════════

def email(message: str, status: str = 'success',
          details: dict | None = None) -> bool:
    to_addr   = os.environ.get('NOTIFY_EMAIL_TO',   '').strip()
    from_addr = os.environ.get('NOTIFY_EMAIL_FROM', '').strip()
    password  = os.environ.get('NOTIFY_EMAIL_PASS', '').strip()

    if not (to_addr and from_addr and password):
        return False

    subject = (
        f"✅ Drive Audit complete — {datetime.now().strftime('%Y-%m-%d')}"
        if status == 'success'
        else f"❌ Drive Audit FAILED — {datetime.now().strftime('%Y-%m-%d')}"
    )

    body_lines = [message, '']
    if details:
        body_lines += [f'  {k}: {v}' for k, v in details.items()]
    body = '\n'.join(body_lines)

    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From']    = from_addr
    msg['To']      = to_addr

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(from_addr, password)
            server.sendmail(from_addr, [to_addr], msg.as_string())
        print(f'📧 Email notification sent to {to_addr}.')
        return True
    except Exception as e:
        print(f'⚠ Email notification failed (non-fatal): {e}')
        return False


# ═══════════════════════════════════════════════════════
#  UNIFIED ENTRY POINT
# ═══════════════════════════════════════════════════════

def send(message: str, status: str = 'success',
         details: dict | None = None) -> None:
    """Send via all configured channels. Fails silently per channel."""
    slack(message, status)
    email(message, status, details)


if __name__ == '__main__':
    msg = ' '.join(sys.argv[1:]) if len(sys.argv) > 1 else 'Test notification'
    send(msg, status='success')
