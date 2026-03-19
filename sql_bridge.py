"""
STAGE 7 — SQL ERP BRIDGE
Connects directly to your SQL ERP database on port 5533,
pulls master data and transaction history, writes to Google Sheets.

Replaces n8n — pure Python, zero extra cost, runs via GitHub Actions.

What it pulls:
  - Item Master (Part codes, descriptions, UOM, category)
  - Vendor Master (Supplier codes, names, GST, contact)
  - Customer Master (Customer codes, names, credit limit)
  - Purchase Order History (PO No, Date, Item, Qty, Rate, Vendor)
  - Sales Order History (SO No, Date, Item, Qty, Rate, Customer)
  - Stock Ledger (Current stock by item)

Run: python sql_bridge.py
Requires: pip install pymysql (or pyodbc for MS SQL)

Set these GitHub secrets:
  ERP_HOST     — your SQL server IP or hostname
  ERP_PORT     — 5533 (your server config)
  ERP_USER     — read-only database user
  ERP_PASSWORD — database password
  ERP_DATABASE — database name
  ERP_TYPE     — mysql or mssql
"""

import json
import os
import pickle
import time
from datetime import datetime

import gspread
from google.auth.transport.requests import Request

# ═══════════════════════════════════════════════════════
#  CONFIG
# ═══════════════════════════════════════════════════════

ERP_HOST     = os.environ.get('ERP_HOST',     '')
ERP_PORT     = int(os.environ.get('ERP_PORT', '5533'))
ERP_USER     = os.environ.get('ERP_USER',     '')
ERP_PASSWORD = os.environ.get('ERP_PASSWORD', '')
ERP_DATABASE = os.environ.get('ERP_DATABASE', '')
ERP_TYPE     = os.environ.get('ERP_TYPE',     'mysql').lower()

# Output sheet — master data goes here
MASTER_SHEET_ID = os.environ.get('MASTER_SHEET_ID',
                  '10Zjxy3mGKP6G3j7uuak3FTXl90JHEQoC0RDgl4tJJXc')

TOKEN_FILE  = 'token.pickle'
OUTPUT_FILE = 'erp_data.json'
BATCH_SIZE  = 500
TAB_INIT    = 1

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
#  SQL CONNECTION
# ═══════════════════════════════════════════════════════

def get_connection():
    """Connect to SQL ERP database. Read-only user only."""
    if not ERP_HOST:
        raise RuntimeError(
            "ERP_HOST not set. Add these GitHub secrets:\n"
            "  ERP_HOST, ERP_PORT, ERP_USER, ERP_PASSWORD, ERP_DATABASE, ERP_TYPE"
        )

    if ERP_TYPE == 'mysql':
        try:
            import pymysql
        except ImportError:
            raise RuntimeError("Run: pip install pymysql")

        conn = pymysql.connect(
            host=ERP_HOST,
            port=ERP_PORT,
            user=ERP_USER,
            password=ERP_PASSWORD,
            database=ERP_DATABASE,
            connect_timeout=30,
            read_timeout=60,
            charset='utf8mb4',
        )
        print(f"  ✅ Connected to MySQL ERP at {ERP_HOST}:{ERP_PORT}")
        return conn

    elif ERP_TYPE in ('mssql', 'sqlserver'):
        try:
            import pyodbc
        except ImportError:
            raise RuntimeError("Run: pip install pyodbc")

        conn_str = (
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={ERP_HOST},{ERP_PORT};"
            f"DATABASE={ERP_DATABASE};"
            f"UID={ERP_USER};"
            f"PWD={ERP_PASSWORD};"
            f"Timeout=30;"
        )
        conn = pyodbc.connect(conn_str)
        print(f"  ✅ Connected to MS SQL ERP at {ERP_HOST}:{ERP_PORT}")
        return conn

    else:
        raise RuntimeError(f"Unknown ERP_TYPE: {ERP_TYPE}. Use 'mysql' or 'mssql'")


def run_query(conn, sql, params=None):
    """Run a SELECT query and return list of dicts."""
    try:
        cursor = conn.cursor()
        cursor.execute(sql, params or [])
        columns = [desc[0] for desc in cursor.description]
        rows    = []
        for row in cursor.fetchall():
            rows.append(dict(zip(columns, [str(v) if v is not None else '' for v in row])))
        cursor.close()
        return rows
    except Exception as e:
        print(f"  ⚠ Query error: {e}")
        print(f"  SQL: {sql[:100]}")
        return []

# ═══════════════════════════════════════════════════════
#  ERP QUERIES
# Adjust table/column names to match YOUR ERP schema
# Common ERP table names are provided as comments
# ═══════════════════════════════════════════════════════

def get_item_master(conn):
    """
    Pull item master from ERP.
    Common table names: ITEM_MASTER, MST_ITEM, ITEMMST, tblItem
    """
    print("  📦 Pulling Item Master...")
    sql = """
        SELECT
            ItemCode        AS item_code,
            ItemName        AS item_name,
            ItemDesc        AS description,
            UOM             AS unit,
            Category        AS category,
            HSNCode         AS hsn_code,
            StandardCost    AS standard_cost,
            IsActive        AS is_active
        FROM ITEM_MASTER
        WHERE IsActive = 1
        ORDER BY ItemCode
        LIMIT 10000
    """
    # If above fails, try these alternative table names:
    # FROM MST_ITEM
    # FROM ITEMMST
    # FROM tblItemMaster
    return run_query(conn, sql)


def get_vendor_master(conn):
    """Pull vendor/supplier master from ERP."""
    print("  🏪 Pulling Vendor Master...")
    sql = """
        SELECT
            VendorCode      AS vendor_code,
            VendorName      AS vendor_name,
            GSTNo           AS gst_no,
            ContactPerson   AS contact,
            Phone           AS phone,
            Email           AS email,
            PaymentTerms    AS payment_terms,
            IsActive        AS is_active
        FROM VENDOR_MASTER
        WHERE IsActive = 1
        ORDER BY VendorCode
        LIMIT 5000
    """
    return run_query(conn, sql)


def get_customer_master(conn):
    """Pull customer master from ERP."""
    print("  👥 Pulling Customer Master...")
    sql = """
        SELECT
            CustomerCode    AS customer_code,
            CustomerName    AS customer_name,
            GSTNo           AS gst_no,
            ContactPerson   AS contact,
            Phone           AS phone,
            Email           AS email,
            CreditLimit     AS credit_limit,
            PaymentDays     AS payment_days,
            IsActive        AS is_active
        FROM CUSTOMER_MASTER
        WHERE IsActive = 1
        ORDER BY CustomerCode
        LIMIT 5000
    """
    return run_query(conn, sql)


def get_po_history(conn):
    """Pull purchase order history — last 24 months."""
    print("  📋 Pulling PO History...")
    sql = """
        SELECT
            PO.PONo         AS po_no,
            PO.PODate       AS po_date,
            PO.VendorCode   AS vendor_code,
            VM.VendorName   AS vendor_name,
            PD.ItemCode     AS item_code,
            IM.ItemName     AS item_name,
            PD.Quantity     AS quantity,
            PD.Rate         AS rate,
            PD.Amount       AS amount,
            PD.Unit         AS unit,
            PO.Status       AS status
        FROM PURCHASE_ORDER PO
        JOIN PURCHASE_ORDER_DETAIL PD ON PO.PONo = PD.PONo
        LEFT JOIN VENDOR_MASTER VM ON PO.VendorCode = VM.VendorCode
        LEFT JOIN ITEM_MASTER IM ON PD.ItemCode = IM.ItemCode
        WHERE PO.PODate >= DATE_SUB(NOW(), INTERVAL 24 MONTH)
        ORDER BY PO.PODate DESC
        LIMIT 50000
    """
    return run_query(conn, sql)


def get_sales_history(conn):
    """Pull sales/dispatch history — last 24 months."""
    print("  💰 Pulling Sales History...")
    sql = """
        SELECT
            SO.SONo         AS so_no,
            SO.SODate       AS so_date,
            SO.CustomerCode AS customer_code,
            CM.CustomerName AS customer_name,
            SD.ItemCode     AS item_code,
            IM.ItemName     AS item_name,
            SD.Quantity     AS quantity,
            SD.Rate         AS rate,
            SD.Amount       AS amount,
            SD.Unit         AS unit,
            SO.Status       AS status
        FROM SALES_ORDER SO
        JOIN SALES_ORDER_DETAIL SD ON SO.SONo = SD.SONo
        LEFT JOIN CUSTOMER_MASTER CM ON SO.CustomerCode = CM.CustomerCode
        LEFT JOIN ITEM_MASTER IM ON SD.ItemCode = IM.ItemCode
        WHERE SO.SODate >= DATE_SUB(NOW(), INTERVAL 24 MONTH)
        ORDER BY SO.SODate DESC
        LIMIT 50000
    """
    return run_query(conn, sql)


def get_stock_ledger(conn):
    """Pull current stock levels."""
    print("  📊 Pulling Stock Ledger...")
    sql = """
        SELECT
            ItemCode        AS item_code,
            ItemName        AS item_name,
            WarehouseCode   AS warehouse,
            CurrentStock    AS current_stock,
            Unit            AS unit,
            LastUpdated     AS last_updated
        FROM STOCK_LEDGER
        WHERE CurrentStock != 0
        ORDER BY ItemCode
        LIMIT 10000
    """
    return run_query(conn, sql)

# ═══════════════════════════════════════════════════════
#  SHEET WRITER
# ═══════════════════════════════════════════════════════

def write_to_sheet(gc, tab_name, headers, data, fmt_color='blue'):
    """Write data to a tab in the master sheet."""
    FMT_COLORS = {
        'blue':   {'backgroundColor':{'red':0.1,'green':0.2,'blue':0.4},
                   'textFormat':{'bold':True,'foregroundColor':{'red':1,'green':1,'blue':1}}},
        'green':  {'backgroundColor':{'red':0.0,'green':0.5,'blue':0.2},
                   'textFormat':{'bold':True,'foregroundColor':{'red':1,'green':1,'blue':1}}},
        'purple': {'backgroundColor':{'red':0.4,'green':0.1,'blue':0.5},
                   'textFormat':{'bold':True,'foregroundColor':{'red':1,'green':1,'blue':1}}},
        'teal':   {'backgroundColor':{'red':0.0,'green':0.4,'blue':0.4},
                   'textFormat':{'bold':True,'foregroundColor':{'red':1,'green':1,'blue':1}}},
        'orange': {'backgroundColor':{'red':0.8,'green':0.4,'blue':0.0},
                   'textFormat':{'bold':True,'foregroundColor':{'red':1,'green':1,'blue':1}}},
    }

    try:
        spreadsheet = gc.open_by_key(MASTER_SHEET_ID)
    except Exception as e:
        print(f"  ⚠ Cannot open master sheet: {e}")
        return

    try:
        ws = spreadsheet.worksheet(tab_name)
        ws.clear()
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(tab_name, rows=TAB_INIT, cols=len(headers))

    time.sleep(0.5)

    if not data:
        ws.append_row(headers)
        ws.format(f'A1:{chr(64+len(headers))}1', FMT_COLORS.get(fmt_color, FMT_COLORS['blue']))
        print(f"  ⚠ No data for {tab_name}")
        return

    rows = [headers]
    for r in data:
        rows.append([r.get(h.lower().replace(' ','_'), '') for h in headers])

    for i in range(0, len(rows), BATCH_SIZE):
        ws.append_rows(rows[i:i+BATCH_SIZE], value_input_option='USER_ENTERED')
        time.sleep(1)

    ws.format(f'A1:{chr(64+len(headers))}1',
              FMT_COLORS.get(fmt_color, FMT_COLORS['blue']))
    print(f"  ✅ {tab_name}: {len(data):,} records")

# ═══════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════

def main():
    print('=' * 60)
    print('   STAGE 7 — SQL ERP Bridge')
    print(f'   {datetime.now().strftime("%Y-%m-%d %H:%M")}')
    print(f'   ERP: {ERP_TYPE.upper()} @ {ERP_HOST}:{ERP_PORT}')
    print('=' * 60)

    if not ERP_HOST:
        print("\n⚠ ERP connection not configured.")
        print("  Add these GitHub secrets to enable:")
        print("  ERP_HOST, ERP_PORT, ERP_USER, ERP_PASSWORD, ERP_DATABASE, ERP_TYPE")
        print("\n  Script will exit gracefully — other stages will continue.")
        return

    # Install required driver
    print("\n📦 Installing SQL driver...")
    if ERP_TYPE == 'mysql':
        os.system('pip install pymysql -q')
    elif ERP_TYPE in ('mssql', 'sqlserver'):
        os.system('pip install pyodbc -q')

    # Connect
    print("\n🔌 Connecting to ERP...")
    try:
        conn = get_connection()
    except Exception as e:
        print(f"❌ Cannot connect to ERP: {e}")
        print("  Check ERP_HOST, ERP_PORT, ERP_USER, ERP_PASSWORD secrets")
        return

    # Pull data
    erp_data = {}
    try:
        erp_data['items']     = get_item_master(conn)
        erp_data['vendors']   = get_vendor_master(conn)
        erp_data['customers'] = get_customer_master(conn)
        erp_data['po_history']= get_po_history(conn)
        erp_data['sales']     = get_sales_history(conn)
        erp_data['stock']     = get_stock_ledger(conn)
    except Exception as e:
        print(f"\n⚠ Query error: {e}")
        print("  Table names may differ in your ERP.")
        print("  Update the SQL queries in sql_bridge.py to match your schema.")
    finally:
        conn.close()
        print("\n🔌 Connection closed")

    # Save locally
    erp_data['generated'] = datetime.now().isoformat()
    with open(OUTPUT_FILE, 'w') as f:
        json.dump(erp_data, f, indent=2, default=str)
    print(f"💾 Saved → {OUTPUT_FILE}")

    # Write to Google Sheets
    print("\n📊 Writing to Master Sheet...")
    creds = authenticate()
    gc    = gspread.authorize(creds)

    if erp_data.get('items'):
        write_to_sheet(gc, '🏷️ ERP Item Master',
            ['item_code','item_name','description','unit','category','hsn_code','standard_cost'],
            erp_data['items'], 'purple')

    if erp_data.get('vendors'):
        write_to_sheet(gc, '🏪 ERP Vendor Master',
            ['vendor_code','vendor_name','gst_no','contact','phone','payment_terms'],
            erp_data['vendors'], 'teal')

    if erp_data.get('customers'):
        write_to_sheet(gc, '👥 ERP Customer Master',
            ['customer_code','customer_name','gst_no','contact','credit_limit','payment_days'],
            erp_data['customers'], 'green')

    if erp_data.get('po_history'):
        write_to_sheet(gc, '📋 ERP PO History',
            ['po_no','po_date','vendor_code','vendor_name','item_code','item_name',
             'quantity','rate','amount','unit','status'],
            erp_data['po_history'], 'orange')

    if erp_data.get('sales'):
        write_to_sheet(gc, '💰 ERP Sales History',
            ['so_no','so_date','customer_code','customer_name','item_code','item_name',
             'quantity','rate','amount','unit','status'],
            erp_data['sales'], 'blue')

    if erp_data.get('stock'):
        write_to_sheet(gc, '📊 ERP Stock Ledger',
            ['item_code','item_name','warehouse','current_stock','unit','last_updated'],
            erp_data['stock'], 'teal')

    print(f"\n✅ SQL Bridge complete")
    print(f"   Items:     {len(erp_data.get('items',[]) ):,}")
    print(f"   Vendors:   {len(erp_data.get('vendors',[])):,}")
    print(f"   Customers: {len(erp_data.get('customers',[])):,}")
    print(f"   PO History:{len(erp_data.get('po_history',[])):,}")
    print(f"   Sales:     {len(erp_data.get('sales',[])):,}")
    print(f"   Stock:     {len(erp_data.get('stock',[])):,}")
    print(f"\n⚠ NOTE: SQL queries use generic table names.")
    print(f"  If you get errors, share your ERP table names and I'll update the queries.")


if __name__ == '__main__':
    main()
