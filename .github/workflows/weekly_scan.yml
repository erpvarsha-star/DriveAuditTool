name: Weekly Drive Audit

on:
  schedule:
    - cron: '0 3 * * 0'      # Full scan every Sunday 08:30 IST
  workflow_dispatch:           # Manual trigger anytime

env:
  FORCE_JAVASCRIPT_ACTIONS_TO_NODE24: true

jobs:
  audit:
    runs-on: ubuntu-latest
    steps:
      - name: Checkout code
        uses: actions/checkout@v4

      - name: Set up Python
        uses: actions/setup-python@v5
        with:
          python-version: '3.11'

      - name: Install dependencies
        run: |
          pip install google-auth google-auth-oauthlib google-auth-httplib2 \
                      google-api-python-client gspread

      - name: Restore credentials
        run: echo '${{ secrets.GOOGLE_CREDENTIALS }}' > credentials.json

      - name: Restore token
        env:
          GOOGLE_TOKEN: ${{ secrets.GOOGLE_TOKEN }}
        run: |
          python - <<'EOF'
          import base64, os, sys
          token = os.environ.get('GOOGLE_TOKEN', '')
          if not token:
              sys.exit("ERROR: GOOGLE_TOKEN secret is missing or empty.")
          with open('token.pickle', 'wb') as f:
              f.write(base64.b64decode(token))
          print("token.pickle restored successfully.")
          EOF

      - name: Restore previous audit data
        uses: actions/cache@v4
        with:
          path: |
            audit_data.json
            seen_hashes.json
            scanned_folders.json
            ai_analysis.json
            master_data.json
          key: audit-resume-cache
          restore-keys: |
            audit-resume-cache

      - name: Run Drive Scanner
        env:
          SLACK_WEBHOOK: ${{ secrets.SLACK_WEBHOOK }}
        run: python drive_audit.py

      - name: Run Report Writer
        env:
          SLACK_WEBHOOK: ${{ secrets.SLACK_WEBHOOK }}
          INVENTORY_SHEET_ID: ${{ secrets.INVENTORY_SHEET_ID }}
        run: python write_report.py

      - name: Run AI Analyser
        env:
          ANTHROPIC_API_KEY: ${{ secrets.ANTHROPIC_API_KEY }}
          INVENTORY_SHEET_ID: ${{ secrets.INVENTORY_SHEET_ID }}
        run: python ai_analyser.py

      - name: Run Master Builder
        env:
          MASTER_SHEET_ID: ${{ secrets.MASTER_SHEET_ID }}
        run: python master_builder.py

      - name: Run CEO Dashboard Writer
        run: python ceo_dashboard_writer.py

      - name: Commit CEO data to GitHub Pages
        run: |
          git config user.name "github-actions[bot]"
          git config user.email "github-actions[bot]@users.noreply.github.com"
          git add docs/ceo_data.json docs/ceo.html || true
          git diff --staged --quiet || git commit -m "chore: update CEO dashboard data [skip ci]"
          git push || true

      - name: Upload audit artifacts
        if: always()
        uses: actions/upload-artifact@v4
        with:
          name: audit-data-${{ github.run_number }}
          path: |
            audit_data.json
            seen_hashes.json
            scanned_folders.json
            ai_analysis.json
            master_data.json
            ERP_Migration_Final_Report.csv
            AI_ERP_Intelligence_Report.csv
          retention-days: 30

      - name: Notify on failure
        if: failure()
        env:
          SLACK_WEBHOOK: ${{ secrets.SLACK_WEBHOOK }}
        run: |
          curl -s -X POST "$SLACK_WEBHOOK" \
            -H 'Content-type: application/json' \
            --data '{"text": ":x: VFPL Audit FAILED — check GitHub Actions"}'

# ─── CEO Dashboard — runs every 2 hours during working day ───
  ceo-refresh:
    runs-on: ubuntu-latest
    if: github.event_name == 'schedule' || github.event_name == 'workflow_dispatch'
    steps:
      - name: Checkout code
        uses: actions/checkout@v4

      - name: Set up Python
        uses: actions/setup-python@v5
        with:
          python-version: '3.11'

      - name: Install dependencies
        run: |
          pip install google-auth google-auth-oauthlib google-auth-httplib2 \
                      google-api-python-client

      - name: Restore credentials
        run: echo '${{ secrets.GOOGLE_CREDENTIALS }}' > credentials.json

      - name: Restore token
        env:
          GOOGLE_TOKEN: ${{ secrets.GOOGLE_TOKEN }}
        run: |
          python - <<'EOF'
          import base64, os
          with open('token.pickle', 'wb') as f:
              f.write(base64.b64decode(os.environ['GOOGLE_TOKEN']))
          EOF

      - name: Update CEO Dashboard
        run: python ceo_dashboard_writer.py

      - name: Commit updated data
        run: |
          git config user.name "github-actions[bot]"
          git config user.email "github-actions[bot]@users.noreply.github.com"
          git add docs/ceo_data.json || true
          git diff --staged --quiet || git commit -m "chore: CEO dashboard refresh [skip ci]"
          git push || true
