from flask import Flask, request, redirect, url_for, render_template_string, session, flash, send_file, jsonify
import sqlite3
import pandas as pd
from datetime import datetime, date, timedelta
from functools import wraps
from io import BytesIO
from pathlib import Path
import hashlib
import math
import json
import os
import qrcode
from io import BytesIO as _BytesIO
import base64
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm

APP_TITLE = 'Koperasi Enterprise V4 Smart Control'
APP_SHORT = 'KOPERASI'
UI_UX_VERSION = '25.0.2-button-runtime-total-fix'
DB_NAME = 'koperasi_enterprise_v3.db'
SECRET_KEY = os.environ.get('KOPERASI_SECRET_KEY', 'dev-only-change-me')
LOAN_AUTO_APPROVE_LIMIT = 5000000
MANUAL_JOURNAL_APPROVE_LIMIT = 10000000
LOAN_VERIFICATION_REQUIRED = True
LATE_PENALTY_PERCENT = 0.005

app = Flask(__name__)
app.secret_key = SECRET_KEY
app.config.update(SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE='Lax', MAX_CONTENT_LENGTH=20 * 1024 * 1024)

@app.errorhandler(413)
def request_entity_too_large(e):
    flash('Upload terlalu besar. Maksimal 20 MB untuk bukti pembayaran.', 'error')
    return redirect(request.referrer or url_for('dashboard'))

def get_conn():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    conn.execute('PRAGMA foreign_keys = ON')
    return conn

def q_all(sql, params=None):
    conn = get_conn()
    rows = conn.execute(sql, params or []).fetchall()
    conn.close()
    return rows

def q_one(sql, params=None):
    conn = get_conn()
    row = conn.execute(sql, params or []).fetchone()
    conn.close()
    return row

def exec_sql(sql, params=None, many=False):
    conn = get_conn()
    cur = conn.cursor()
    if many:
        cur.executemany(sql, params)
    else:
        cur.execute(sql, params or [])
    conn.commit()
    rid = cur.lastrowid
    conn.close()
    return rid

def hash_password(text):
    return generate_password_hash(text, method='pbkdf2:sha256', salt_length=16)

def verify_password(stored, password):
    stored = stored or ''
    if len(stored) == 64 and all(c in '0123456789abcdef' for c in stored.lower()):
        return hashlib.sha256(password.encode()).hexdigest() == stored
    try: return check_password_hash(stored, password)
    except Exception: return False

def now_str():
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')

def today_str():
    return str(date.today())

def month_key(date_str):
    return str(date_str)[:7]

def gen_code(prefix):
    return f"{prefix}-{datetime.now().strftime('%Y%m%d%H%M%S')}"

def get_branch_id():
    """Return current user's branch_id or None (PUSAT)."""
    return session.get('branch_id')

def get_branch_name(branch_id=None):
    """Get branch name from id or return 'PUSAT'."""
    if not branch_id:
        return 'PUSAT'
    row = q_one('SELECT name FROM koperasi_branches WHERE id=?', [branch_id])
    return row['name'] if row else 'PUSAT'

def get_product_stock(product_id, branch_id=None):
    """Get stock for a product in a specific branch. If branch_id is None, return PUSAT stock."""
    if not branch_id:
        row = q_one('SELECT stock FROM products WHERE id=?', [product_id])
        return float(row['stock']) if row else 0
    row = q_one('SELECT stock FROM product_branch_stock WHERE product_id=? AND branch_id=?', [product_id, branch_id])
    return float(row['stock']) if row else 0

def update_branch_stock(product_id, qty, branch_id=None, operation='add'):
    """Add or subtract stock for a product in a branch.
    If branch_id is None, operate on PUSAT (products.stock).
    operation: 'add' or 'subtract'
    """
    if not branch_id:
        # PUSAT — use products table
        if operation == 'add':
            exec_sql('UPDATE products SET stock = COALESCE(stock,0) + ? WHERE id=?', [qty, product_id])
        else:
            current = get_product_stock(product_id, None)
            if current < qty:
                raise ValueError(f'Stok pusat tidak cukup. Tersedia {current}, diminta {qty}')
            exec_sql('UPDATE products SET stock = COALESCE(stock,0) - ? WHERE id=?', [qty, product_id])
    else:
        # Cabang — use product_branch_stock
        if operation == 'add':
            exec_sql('INSERT INTO product_branch_stock(product_id, branch_id, stock) VALUES (?, ?, ?) ON CONFLICT(product_id, branch_id) DO UPDATE SET stock = COALESCE(stock,0) + ?',
                     [product_id, branch_id, qty, qty])
        else:
            current = get_product_stock(product_id, branch_id)
            if current < qty:
                raise ValueError(f'Stok cabang tidak cukup. Tersedia {current}, diminta {qty}')
            exec_sql('UPDATE product_branch_stock SET stock = COALESCE(stock,0) - ? WHERE product_id=? AND branch_id=?',
                     [qty, product_id, branch_id])

def ensure_product_branch_stock(product_id, branch_id):
    """Ensure product_branch_stock row exists, create with 0 if not."""
    if not branch_id:
        return
    row = q_one('SELECT id FROM product_branch_stock WHERE product_id=? AND branch_id=?', [product_id, branch_id])
    if not row:
        exec_sql('INSERT OR IGNORE INTO product_branch_stock(product_id, branch_id, stock, min_stock) VALUES (?, ?, 0, 0)',
                 [product_id, branch_id])

def rupiah(x):
    try:
        return f"Rp {float(x):,.0f}".replace(',', '.')
    except Exception:
        return 'Rp 0'

def to_excel(data_dict, filename):
    """Export dict of lists to Excel file."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, rows in data_dict.items():
            df = pd.DataFrame(rows)
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name=filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

def add_timeline(loan_id, status, note='', created_by=None):
    exec_sql('INSERT INTO loan_timeline(loan_id, status, note, created_by) VALUES (?, ?, ?, ?)', [loan_id, status, note, created_by or session.get('user_id')])

def parse_float(v, default=0.0):
    try:
        return float(v)
    except Exception:
        return default

def current_user():
    if 'user_id' not in session:
        return None
    return q_one('SELECT * FROM users WHERE id=?', [session['user_id']])

def get_ip():
    return request.headers.get('X-Forwarded-For', request.remote_addr or '-')

def get_ua():
    return (request.headers.get('User-Agent') or '-')[:250]

def log_action(action, entity, entity_id='', detail=''):
    username = session.get('username', 'system')
    exec_sql(
        'INSERT INTO audit_logs(log_time, username, action, entity, entity_id, detail, ip_address, user_agent) VALUES (?, ?, ?, ?, ?, ?, ?, ?)',
        [now_str(), username, action, entity, str(entity_id), detail, get_ip(), get_ua()]
    )

def get_account_id(code):
    row = q_one('SELECT id FROM accounts WHERE account_code=?', [code])
    return int(row['id']) if row else None

def post_journal(entry_date, description, lines, ref_type=None, ref_id=None, created_by=None):
    entry_no = gen_code('JU')
    payload = []
    for line in lines:
        payload.append((entry_no, entry_date, description, line['account_id'], line.get('debit', 0), line.get('credit', 0), ref_type, ref_id, created_by))
    exec_sql('''
        INSERT INTO journal_entries(entry_no, entry_date, description, account_id, debit, credit, ref_type, ref_id, created_by)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', payload, many=True)
    return entry_no

def reverse_journal(entry_date, description, original_ref_type, original_ref_id, created_by=None):
    rows = q_all('SELECT account_id, debit, credit FROM journal_entries WHERE ref_type=? AND ref_id=?', [original_ref_type, original_ref_id])
    if not rows:
        return None
    lines = []
    for r in rows:
        lines.append({'account_id': r['account_id'], 'debit': float(r['credit']), 'credit': float(r['debit'])})
    return post_journal(entry_date, description, lines, ref_type=f'reversal_{original_ref_type}', ref_id=original_ref_id, created_by=created_by)

def is_period_locked(date_str):
    mk = month_key(date_str)
    row = q_one('SELECT 1 FROM period_locks WHERE period_month=? AND is_locked=1', [mk])
    return row is not None

def require_open_period(date_str):
    if is_period_locked(date_str):
        flash(f'Periode {month_key(date_str)} sudah ditutup / dikunci.', 'error')
        return False
    return True

def get_setting(key, default=None):
    row = q_one('SELECT value FROM settings WHERE key=?', [key])
    return row['value'] if row else default

def set_setting(key, value):
    exec_sql('INSERT INTO settings(key, value) VALUES(?, ?) ON CONFLICT(key) DO UPDATE SET value=excluded.value', [key, str(value)])

# =========================
# Menu Visibility System
# =========================
# All known menu items with unique IDs
ALL_MENU_ITEMS = {
    'dashboard': 'Dashboard',
    'quick_cashier': 'Kasir Cepat',
    'cashier': 'Kasir',
    'cashier_discount': 'Kasir + Diskon',
    'sales_history': 'Riwayat Penjualan',
    'sales_returns': 'Retur Penjualan',
    'daily_report': 'Laporan Harian',
    'ewallet_dashboard': 'E-Wallet Dashboard',
    'topup_approval': 'Approval Topup',
    'ewallet_accounts': 'Direktori Wallet',
    'ewallet_ledger': 'Ledger Wallet',
    'ewallet_security': 'Keamanan Wallet',
    'payment_accounts': 'Rekening Pembayaran',
    'admin_command_center': 'Command Center',
    'ewallet_adjustment': 'Penyesuaian Wallet',
    'members': 'Data Anggota',
    'member_import': 'Import Member',
    'products': 'Master Barang',
    'stock_movements': 'Mutasi Stok',
    'stock_transfer': 'Transfer Stok',
    'stock_history': 'Riwayat Stok',
    'stock_opname': 'Stock Opname',
    'maps': 'Maps Stok',
    'categories': 'Kategori',
    'product_import_excel': 'Import Excel Barang',
    'admin_branches': 'Kelola Cabang',
    'quick_cashier_queue': 'Verifikasi QC',
    'suppliers': 'Supplier',
    'purchase_orders': 'Pembelian (PO)',
    'purchase_order_history': 'History Order / Audit',
    'supplier_payments': 'Hutang Supplier',
    'supplier_portal': 'Portal Supplier',
    'savings': 'Simpanan',
    'loans': 'Kredit',
    'loan_types': 'Produk Kredit',
    'verify_loan_payments': 'Verifikasi Angsuran',
    'reports': 'Laporan',
    'financial_statements': 'Laba Rugi & Neraca',
    'shu_report': 'Laporan SHU',
    'approvals': 'Approval',
    'user_approval': 'Persetujuan Pengguna',
    'users': 'Manajemen Pengguna',
    'accounting': 'Akuntansi',
    'audit': 'Jejak Audit',
    'tutorial': 'Panduan Penggunaan',
    'innovation_center': 'Pusat Inovasi',
    'operations_plus': 'Pusat Tata Kelola',
    'credit_control': 'Kontrol Risiko Kredit',
    'settings': 'Pengaturan',
    # Member-specific
    'member_dashboard': 'Dashboard Member',
    'wallet': 'Dompet & Isi Saldo',
    'member_purchases': 'Riwayat Belanja',
    'member_digital_card': 'Kartu Digital',
    'member_card_pdf': 'Kartu PDF',
    'apply_loan': 'Ajukan Kredit',
    'my_payments': 'Riwayat Angsuran',
    'member_expense_history': 'Riwayat Transaksi',
    'shu_member': 'SHU Saya',
    'member_hub': 'Ruang Anggota',
}


# Canonical menu access. Sidebar, settings, and route permissions must follow this map.
MENU_ROLE_ACCESS = {
    'dashboard': {'admin','kasir','bendahara','supervisor','branch_admin','branch_cashier'},
    'quick_cashier': {'admin','kasir','branch_admin','branch_cashier'},
    'cashier': {'admin','kasir','branch_admin','branch_cashier'},
    'cashier_discount': {'admin','kasir','branch_admin','branch_cashier'},
    'sales_history': {'admin','kasir','branch_admin','branch_cashier'},
    'sales_returns': {'admin','kasir','branch_admin'},
    'daily_report': {'admin','kasir','bendahara','branch_admin'},
    'ewallet_dashboard': {'admin'}, 'topup_approval': {'admin'},
    'ewallet_accounts': {'admin'}, 'ewallet_ledger': {'admin'}, 'ewallet_security': {'admin'},
    'payment_accounts': {'admin','bendahara'}, 'admin_command_center': {'admin'}, 'ewallet_adjustment': {'admin'},
    'members': {'admin'}, 'member_import': {'admin'},
    'products': {'admin','kasir','branch_admin','branch_cashier'},
    'stock_movements': {'admin','kasir','branch_admin','branch_cashier'},
    'stock_transfer': {'admin','kasir','branch_admin','branch_cashier'},
    'stock_history': {'admin','kasir','branch_admin','branch_cashier'},
    'stock_opname': {'admin','supervisor','branch_admin'},
    'maps': {'admin','kasir','branch_admin','branch_cashier'},
    'categories': {'admin'}, 'product_import_excel': {'admin'},
    'admin_branches': {'admin','branch_admin'},
    'quick_cashier_queue': {'admin','branch_admin'},
    'suppliers': {'admin','kasir'}, 'purchase_orders': {'admin'}, 'supplier_payments': {'admin'},
    'supplier_portal': {'supplier'}, 'purchase_order_history': {'admin','supplier'},
    'savings': {'admin','bendahara','branch_admin'},
    'loans': {'admin','bendahara','user','branch_admin'},
    'loan_types': {'admin','bendahara'},
    'verify_loan_payments': {'admin','bendahara','branch_admin'},
    'reports': {'admin','bendahara','supervisor','branch_admin'},
    'financial_statements': {'admin','bendahara'}, 'shu_report': {'admin','bendahara'},
    'approvals': {'admin'}, 'user_approval': {'admin'}, 'users': {'admin'},
    'accounting': {'admin','bendahara'}, 'audit': {'admin','supervisor'},
    'tutorial': {'admin','kasir','bendahara','supervisor','user','branch_admin','branch_cashier','supplier'},
    'innovation_center': {'admin','bendahara','supervisor','branch_admin'},
    'operations_plus': {'admin','bendahara','supervisor','branch_admin'},
    'credit_control': {'admin','bendahara','supervisor','branch_admin'},
    'settings': {'admin','kasir','bendahara','supervisor','user','branch_admin','branch_cashier','supplier'},
    'member_dashboard': {'user'}, 'wallet': {'user'}, 'member_purchases': {'user'},
    'member_digital_card': {'user'}, 'member_card_pdf': {'user'}, 'apply_loan': {'user'},
    'my_payments': {'user'}, 'member_expense_history': {'user'}, 'shu_member': {'user'}, 'member_hub': {'user'},
}

def is_menu_allowed_for_role(menu_id, role):
    return role in MENU_ROLE_ACCESS.get(menu_id, set())

def get_menu_hidden(role):
    """Return set of menu IDs hidden for this role."""
    val = get_setting(f'menu_hidden_{role}', '')
    if not val:
        return set()
    try:
        return set(json.loads(val))
    except:
        return set()

def set_menu_hidden(role, hidden_set):
    """Save hidden menu IDs for a role as JSON array."""
    set_setting(f'menu_hidden_{role}', json.dumps(sorted(hidden_set)))

def is_menu_visible(menu_id, role):
    """Visible only when role is allowed and administrator enabled the menu."""
    if not is_menu_allowed_for_role(menu_id, role):
        return False
    if role == 'admin':
        return True
    return menu_id not in get_menu_hidden(role)

def date_range_from_request(prefix=''):
    start = request.args.get(f'{prefix}start', '') or request.form.get(f'{prefix}start', '')
    end = request.args.get(f'{prefix}end', '') or request.form.get(f'{prefix}end', '')
    return start, end

def add_date_filter(sql, date_field, start, end, params):
    if start:
        sql += f' AND {date_field} >= ?'
        params.append(start)
    if end:
        sql += f' AND {date_field} <= ?'
        params.append(end)
    return sql, params

def get_cart():
    if 'cart' not in session:
        session['cart'] = []
    return session['cart']

def save_cart(cart):
    session['cart'] = cart
    session.modified = True

DEFAULT_PER_PAGE = 10

def paginate_query(sql, params, page=None, per_page=None):
    """Generic pagination: count total, fetch page, return dict."""
    if per_page is None:
        per_page = 10
    page = max(1, int(page or 1))
    # count
    count_sql = f"SELECT COUNT(*) as n FROM ({sql})"
    total = q_one(count_sql, params)['n'] or 0
    total_pages = max(1, (total + per_page - 1) // per_page)
    if page > total_pages:
        page = total_pages
    offset = (page - 1) * per_page
    data_sql = f"{sql} LIMIT ? OFFSET ?"
    rows = q_all(data_sql, params + [per_page, offset])
    return {'rows': rows, 'page': page, 'per_page': per_page, 'total': total, 'total_pages': total_pages}

def render_pagination(p, endpoint, extra_params=None):
    """HTML pagination controls."""
    if p['total_pages'] <= 1:
        return ''
    if extra_params is None:
        extra_params = {}
    parts = ['<nav style="display:flex;gap:6px;align-items:center;justify-content:center;margin:16px 0;flex-wrap:wrap;">']
    def lnk(pg, label, disabled=False):
        cls = 'btn btn-sm btn-ghost' if not disabled else 'btn btn-sm btn-ghost" style="pointer-events:none;opacity:0.4'
        params = '&'.join(f'{k}={v}' for k, v in extra_params.items() if v)
        sep = '&' if params else ''
        return f'<a class="{cls}" href="{url_for(endpoint, page=pg)}{sep}{params}">{label}</a>'
    parts.append(lnk(1, '«', p['page'] == 1))
    parts.append(lnk(max(1, p['page']-1), '‹', p['page'] == 1))
    start = max(1, p['page'] - 2)
    end = min(p['total_pages'], p['page'] + 2)
    for pg in range(start, end + 1):
        if pg == p['page']:
            parts.append(f'<span class="btn btn-sm btn-ghost" style="background:var(--primary);color:white;border-color:var(--primary);pointer-events:none;">{pg}</span>')
        else:
            parts.append(lnk(pg, str(pg)))
    parts.append(lnk(min(p['total_pages'], p['page']+1), '›', p['page'] == p['total_pages']))
    parts.append(lnk(p['total_pages'], '»', p['page'] == p['total_pages']))
    parts.append(f'<span class="muted small" style="margin-left:8px;">Hal {p["page"]}/{p["total_pages"]} ({p["total"]} data)</span>')
    parts.append('</nav>')
    return ''.join(parts)

# =========================
# DB Init
# =========================
def init_db():
    conn = get_conn()
    cur = conn.cursor()

    cur.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE,
            full_name TEXT,
            phone TEXT,
            address TEXT,
            employee_number TEXT UNIQUE,
            password_hash TEXT,
            role TEXT DEFAULT 'user',
            status TEXT DEFAULT 'PENDING_APPROVAL',
            approved_by INTEGER,
            approved_at TEXT,
            reject_reason TEXT,
            active INTEGER DEFAULT 1,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(approved_by) REFERENCES users(id)
        )
    ''')

    def add_column_if_not_exists(table, column, definition):
        try:
            cur.execute(f'ALTER TABLE {table} ADD COLUMN {column} {definition}')
            conn.commit()
        except:
            pass
    
    add_column_if_not_exists('users', 'phone', 'TEXT')
    add_column_if_not_exists('users', 'address', 'TEXT')
    add_column_if_not_exists('users', 'status', 'TEXT DEFAULT "PENDING_APPROVAL"')
    add_column_if_not_exists('users', 'approved_by', 'INTEGER')
    add_column_if_not_exists('users', 'approved_at', 'TEXT')
    add_column_if_not_exists('users', 'employee_number', 'TEXT UNIQUE')
    add_column_if_not_exists('users', 'reject_reason', 'TEXT')
    add_column_if_not_exists('users', 'active', 'INTEGER DEFAULT 1')
    
    add_column_if_not_exists('users', 'branch_id', 'INTEGER')
    add_column_if_not_exists('users', 'supplier_id', 'INTEGER')
    
    add_column_if_not_exists('quick_cashier_items', 'stock_branch_id', 'INTEGER')
    
    add_column_if_not_exists('loans', 'loan_type_id', 'INTEGER')
    add_column_if_not_exists('loans', 'interest_rate', 'REAL DEFAULT 0')
    add_column_if_not_exists('loans', 'admin_note', 'TEXT')
    add_column_if_not_exists('loans', 'calculated_by', 'INTEGER')
    add_column_if_not_exists('loans', 'calculated_at', 'TEXT')
    add_column_if_not_exists('loans', 'total_paid', 'REAL DEFAULT 0')
    add_column_if_not_exists('loans', 'late_penalty', 'REAL DEFAULT 0')
    
    conn.commit()

    cur.execute('''
        CREATE TABLE IF NOT EXISTS members (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            member_code TEXT UNIQUE,
            name TEXT NOT NULL,
            phone TEXT,
            address TEXT,
            join_date TEXT,
            status TEXT DEFAULT 'Aktif',
            saldo REAL DEFAULT 0,
            shu_balance REAL DEFAULT 0,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    try:
        cur.execute('ALTER TABLE members ADD COLUMN saldo REAL DEFAULT 0')
    except:
        pass
    try:
        cur.execute('ALTER TABLE members ADD COLUMN shu_balance REAL DEFAULT 0')
    except:
        pass

    cur.execute('''
        CREATE TABLE IF NOT EXISTS topup_requests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            member_id INTEGER,
            nominal REAL DEFAULT 0,
            bukti_foto TEXT,
            status TEXT DEFAULT 'PENDING',
            note TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            approved_at TEXT,
            approved_by INTEGER,
            FOREIGN KEY(member_id) REFERENCES members(id),
            FOREIGN KEY(approved_by) REFERENCES users(id)
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS saldo_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            member_id INTEGER,
            tipe TEXT,
            nominal REAL DEFAULT 0,
            saldo_sebelum REAL DEFAULT 0,
            saldo_setelah REAL DEFAULT 0,
            keterangan TEXT,
            reference_id INTEGER,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            created_by INTEGER,
            FOREIGN KEY(member_id) REFERENCES members(id),
            FOREIGN KEY(created_by) REFERENCES users(id)
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            barcode TEXT UNIQUE,
            product_name TEXT NOT NULL,
            category TEXT,
            unit TEXT DEFAULT 'pcs',
            buy_price REAL DEFAULT 0,
            sell_price REAL DEFAULT 0,
            stock REAL DEFAULT 0,
            min_stock REAL DEFAULT 0,
            active INTEGER DEFAULT 1,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS sales (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            invoice_no TEXT UNIQUE,
            trx_date TEXT,
            member_id INTEGER,
            cashier_id INTEGER,
            customer_name TEXT,
            total REAL DEFAULT 0,
            paid REAL DEFAULT 0,
            change_amount REAL DEFAULT 0,
            note TEXT,
            status TEXT DEFAULT 'Posted',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(member_id) REFERENCES members(id),
            FOREIGN KEY(cashier_id) REFERENCES users(id)
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS sales_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sales_id INTEGER,
            product_id INTEGER,
            barcode TEXT,
            product_name TEXT,
            qty REAL DEFAULT 0,
            price REAL DEFAULT 0,
            subtotal REAL DEFAULT 0,
            FOREIGN KEY(sales_id) REFERENCES sales(id) ON DELETE CASCADE,
            FOREIGN KEY(product_id) REFERENCES products(id)
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS savings_transactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            trx_no TEXT,
            trx_date TEXT,
            member_id INTEGER,
            saving_type TEXT,
            direction TEXT,
            amount REAL DEFAULT 0,
            note TEXT,
            created_by INTEGER,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(member_id) REFERENCES members(id),
            FOREIGN KEY(created_by) REFERENCES users(id)
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS loan_types (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            description TEXT,
            interest_rate_monthly REAL DEFAULT 0,
            admin_fee_fixed REAL DEFAULT 0,
            admin_fee_percent REAL DEFAULT 0,
            min_tenor INTEGER DEFAULT 1,
            max_tenor INTEGER DEFAULT 36,
            max_amount REAL DEFAULT 25000000,
            metode_bunga TEXT DEFAULT 'FLAT',
            is_active INTEGER DEFAULT 1,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    add_column_if_not_exists('loan_types', 'metode_bunga', 'TEXT DEFAULT "FLAT"')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS loans (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            loan_no TEXT UNIQUE,
            member_id INTEGER,
            loan_type_id INTEGER,
            loan_date TEXT,
            principal REAL DEFAULT 0,
            service_fee REAL DEFAULT 0,
            interest_rate REAL DEFAULT 0,
            tenor_month INTEGER DEFAULT 1,
            total_receivable REAL DEFAULT 0,
            monthly_installment REAL DEFAULT 0,
            admin_note TEXT,
            status TEXT DEFAULT 'DRAFT',
            note TEXT,
            created_by INTEGER,
            calculated_by INTEGER,
            calculated_at TEXT,
            approved_by INTEGER,
            approved_at TEXT,
            total_paid REAL DEFAULT 0,
            late_penalty REAL DEFAULT 0,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(member_id) REFERENCES members(id),
            FOREIGN KEY(loan_type_id) REFERENCES loan_types(id),
            FOREIGN KEY(created_by) REFERENCES users(id),
            FOREIGN KEY(calculated_by) REFERENCES users(id),
            FOREIGN KEY(approved_by) REFERENCES users(id)
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS loan_schedules (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            loan_id INTEGER,
            installment_number INTEGER,
            due_date TEXT,
            amount REAL DEFAULT 0,
            principal_amount REAL DEFAULT 0,
            interest_amount REAL DEFAULT 0,
            paid_date TEXT,
            paid_amount REAL DEFAULT 0,
            penalty_amount REAL DEFAULT 0,
            status TEXT DEFAULT 'BELUM_BAYAR',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(loan_id) REFERENCES loans(id) ON DELETE CASCADE
        )
    ''')
    add_column_if_not_exists('loan_schedules', 'penalty_amount', 'REAL DEFAULT 0')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS loan_payments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            loan_id INTEGER,
            payment_date TEXT,
            amount REAL DEFAULT 0,
            penalty_amount REAL DEFAULT 0,
            payment_method TEXT DEFAULT 'tunai',
            transfer_bank TEXT,
            transfer_proof TEXT,
            status TEXT DEFAULT 'VERIFIED',
            verified_by INTEGER,
            verified_at TEXT,
            note TEXT,
            created_by INTEGER,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(loan_id) REFERENCES loans(id) ON DELETE CASCADE,
            FOREIGN KEY(created_by) REFERENCES users(id),
            FOREIGN KEY(verified_by) REFERENCES users(id)
        )
    ''')
    add_column_if_not_exists('loan_payments', 'penalty_amount', 'REAL DEFAULT 0')
    
    add_column_if_not_exists('loan_payments', 'payment_method', 'TEXT DEFAULT "tunai"')
    add_column_if_not_exists('loan_payments', 'transfer_bank', 'TEXT')
    add_column_if_not_exists('loan_payments', 'transfer_proof', 'TEXT')
    add_column_if_not_exists('loan_payments', 'status', 'TEXT DEFAULT "VERIFIED"')
    add_column_if_not_exists('loan_payments', 'verified_by', 'INTEGER')
    add_column_if_not_exists('loan_payments', 'verified_at', 'TEXT')
    add_column_if_not_exists('loan_schedules', 'payment_id', 'INTEGER')
    add_column_if_not_exists('loan_payments', 'verification_note', 'TEXT')
    add_column_if_not_exists('loan_payments', 'submitted_at', 'TEXT')
    add_column_if_not_exists('loan_payments', 'client_reference', 'TEXT')
    add_column_if_not_exists('loan_payments', 'destination_account_id', 'INTEGER')
    add_column_if_not_exists('loan_payments', 'destination_account_snapshot', 'TEXT')
    # FIX VERIFIKASI ANGSURAN: status pending lama/variasi ejaan tetap masuk antrian admin.
    try:
        cur.execute('''UPDATE loan_payments
                       SET status='PENDING_VERIFICATION'
                       WHERE UPPER(TRIM(COALESCE(status,''))) IN ('PENDING','PENDING VERIFICATION','PENDING_VERIFIKASI','MENUNGGU','MENUNGGU VERIFIKASI','MENUNGGU_VERIFIKASI')''')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_loan_payments_status_submitted ON loan_payments(status, submitted_at, id)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_loan_payments_recovery_queue ON loan_payments(loan_id,status,verified_by,submitted_at)')
        # Recovery data: jika user sudah upload bukti, tapi record keburu berstatus VERIFIED tanpa verifikator
        # dan belum pernah dipakai pada jadwal angsuran, tampilkan kembali di menu Verifikasi Angsuran.
        cur.execute('''UPDATE loan_payments
                       SET status='PENDING_VERIFICATION'
                       WHERE UPPER(TRIM(COALESCE(status,'')))='VERIFIED'
                         AND COALESCE(transfer_proof,'')<>''
                         AND submitted_at IS NOT NULL
                         AND verified_by IS NULL
                         AND verified_at IS NULL
                         AND NOT EXISTS (SELECT 1 FROM loan_schedules ls WHERE ls.payment_id=loan_payments.id)''')
    except Exception:
        pass
    cur.execute('''CREATE TABLE IF NOT EXISTS payment_accounts(id INTEGER PRIMARY KEY AUTOINCREMENT,bank_name TEXT NOT NULL,account_number TEXT NOT NULL,account_holder TEXT NOT NULL,account_type TEXT DEFAULT 'KREDIT',branch_id INTEGER,currency TEXT DEFAULT 'IDR',instructions TEXT,status TEXT DEFAULT 'ACTIVE',is_default INTEGER DEFAULT 0,version_no INTEGER DEFAULT 1,effective_from TEXT DEFAULT CURRENT_TIMESTAMP,effective_until TEXT,created_by INTEGER,created_at TEXT DEFAULT CURRENT_TIMESTAMP,updated_by INTEGER,updated_at TEXT,UNIQUE(bank_name,account_number))''')
    cur.execute('''CREATE TABLE IF NOT EXISTS payment_account_history(id INTEGER PRIMARY KEY AUTOINCREMENT,payment_account_id INTEGER,action TEXT,old_data TEXT,new_data TEXT,reason TEXT,changed_by INTEGER,changed_at TEXT DEFAULT CURRENT_TIMESTAMP,FOREIGN KEY(payment_account_id) REFERENCES payment_accounts(id))''')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_payment_accounts_type_status ON payment_accounts(account_type,status)')
    
    cur.execute('''
        CREATE TABLE IF NOT EXISTS loan_documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            loan_id INTEGER,
            name TEXT,
            description TEXT,
            required INTEGER DEFAULT 1,
            file_name TEXT,
            uploaded_at TEXT,
            uploaded_by INTEGER,
            status TEXT DEFAULT 'PENDING',
            admin_note TEXT,
            verified_by INTEGER,
            verified_at TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(loan_id) REFERENCES loans(id) ON DELETE CASCADE
        )
    ''')
    
    cur.execute('''
        CREATE TABLE IF NOT EXISTS loan_timeline (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            loan_id INTEGER,
            status TEXT,
            note TEXT,
            created_by INTEGER,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(loan_id) REFERENCES loans(id) ON DELETE CASCADE
        )
    ''')
    
    add_column_if_not_exists('loans', 'current_stage', 'TEXT DEFAULT "SUBMITTED"')
    add_column_if_not_exists('loans', 'progress_percent', 'INTEGER DEFAULT 10')
    add_column_if_not_exists('sales', 'payment_method', 'TEXT DEFAULT "tunai"')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS stock_movements (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            product_id INTEGER,
            movement_type TEXT,
            qty REAL DEFAULT 0,
            unit_cost REAL DEFAULT 0,
            note TEXT,
            created_by INTEGER,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(product_id) REFERENCES products(id)
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS suppliers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            supplier_code TEXT UNIQUE,
            name TEXT NOT NULL,
            phone TEXT,
            address TEXT,
            contact_person TEXT,
            is_active INTEGER DEFAULT 1,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS purchase_orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            po_no TEXT UNIQUE,
            po_date TEXT,
            supplier_id INTEGER,
            total REAL DEFAULT 0,
            status TEXT DEFAULT 'DRAFT',
            note TEXT,
            created_by INTEGER,
            received_at TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(supplier_id) REFERENCES suppliers(id),
            FOREIGN KEY(created_by) REFERENCES users(id)
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS purchase_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            po_id INTEGER,
            product_id INTEGER,
            qty REAL DEFAULT 0,
            price REAL DEFAULT 0,
            subtotal REAL DEFAULT 0,
            FOREIGN KEY(po_id) REFERENCES purchase_orders(id) ON DELETE CASCADE,
            FOREIGN KEY(product_id) REFERENCES products(id)
        )
    ''')

    add_column_if_not_exists('purchase_items', 'supplier_status', 'TEXT DEFAULT "PENDING"')
    add_column_if_not_exists('purchase_items', 'supplier_note', 'TEXT')
    add_column_if_not_exists('purchase_items', 'confirmed_qty', 'REAL DEFAULT 0')
    add_column_if_not_exists('purchase_items', 'supplier_price', 'REAL DEFAULT 0')
    add_column_if_not_exists('purchase_items', 'supplier_price_diff', 'REAL DEFAULT 0')
    add_column_if_not_exists('purchase_items', 'supplier_subtotal', 'REAL DEFAULT 0')
    add_column_if_not_exists('purchase_orders', 'supplier_feedback_at', 'TEXT')
    add_column_if_not_exists('purchase_orders', 'supplier_total', 'REAL DEFAULT 0')
    add_column_if_not_exists('purchase_orders', 'supplier_eta', 'TEXT')
    add_column_if_not_exists('purchase_orders', 'supplier_reference', 'TEXT')
    add_column_if_not_exists('purchase_orders', 'supplier_delivery_note', 'TEXT')
    add_column_if_not_exists('purchase_orders', 'has_price_change', 'INTEGER DEFAULT 0')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS supplier_po_feedback (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            po_id INTEGER,
            supplier_id INTEGER,
            message TEXT,
            created_by INTEGER,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS accounts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            account_code TEXT UNIQUE,
            account_name TEXT,
            category TEXT,
            normal_balance TEXT
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS journal_entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            entry_no TEXT,
            entry_date TEXT,
            description TEXT,
            account_id INTEGER,
            debit REAL DEFAULT 0,
            credit REAL DEFAULT 0,
            ref_type TEXT,
            ref_id INTEGER,
            created_by INTEGER,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(account_id) REFERENCES accounts(id),
            FOREIGN KEY(created_by) REFERENCES users(id)
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS approval_requests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            request_type TEXT,
            ref_table TEXT,
            ref_id INTEGER,
            status TEXT DEFAULT 'Pending',
            reason TEXT,
            created_by INTEGER,
            approved_by INTEGER,
            approved_at TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(created_by) REFERENCES users(id),
            FOREIGN KEY(approved_by) REFERENCES users(id)
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS period_locks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            period_month TEXT UNIQUE,
            is_locked INTEGER DEFAULT 1,
            note TEXT,
            locked_by INTEGER,
            locked_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(locked_by) REFERENCES users(id)
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS audit_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            log_time TEXT,
            username TEXT,
            action TEXT,
            entity TEXT,
            entity_id TEXT,
            detail TEXT,
            ip_address TEXT,
            user_agent TEXT
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS notifications (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            title TEXT,
            message TEXT,
            is_read INTEGER DEFAULT 0,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(user_id) REFERENCES users(id)
        )
    ''')

    default_users = [
        ('admin', 'Administrator', hash_password('admin123'), 'admin', 1, 'ACTIVE'),
        ('kasir', 'User Kasir', hash_password('kasir123'), 'kasir', 1, 'ACTIVE'),
        ('bendahara', 'User Bendahara', hash_password('bendahara123'), 'bendahara', 1, 'ACTIVE'),
        ('supervisor', 'User Supervisor', hash_password('supervisor123'), 'supervisor', 1, 'ACTIVE'),
    ]
    cur.executemany('INSERT OR IGNORE INTO users(username, full_name, password_hash, role, active, status) VALUES (?, ?, ?, ?, ?, ?)', default_users)

    default_accounts = [
        ('1001', 'Kas', 'Aset', 'Debit'),
        ('1101', 'Piutang Pinjaman', 'Aset', 'Debit'),
        ('1102', 'Piutang Denda', 'Aset', 'Debit'),
        ('1201', 'Persediaan Barang', 'Aset', 'Debit'),
        ('2001', 'Simpanan Anggota', 'Kewajiban', 'Kredit'),
        ('2101', 'SHU Belum Dibagi', 'Kewajiban', 'Kredit'),
        ('3001', 'Modal', 'Modal', 'Kredit'),
        ('4001', 'Penjualan', 'Pendapatan', 'Kredit'),
        ('4101', 'Pendapatan Jasa Pinjaman', 'Pendapatan', 'Kredit'),
        ('4201', 'Pendapatan Denda', 'Pendapatan', 'Kredit'),
        ('5001', 'Harga Pokok Penjualan', 'Beban', 'Debit'),
        ('6001', 'Beban Operasional', 'Beban', 'Debit')
    ]
    cur.executemany('INSERT OR IGNORE INTO accounts(account_code, account_name, category, normal_balance) VALUES (?, ?, ?, ?)', default_accounts)

    cur.execute("INSERT OR IGNORE INTO settings(key, value) VALUES('loan_auto_approve_limit', ?)", (str(LOAN_AUTO_APPROVE_LIMIT),))
    cur.execute("INSERT OR IGNORE INTO settings(key, value) VALUES('manual_journal_approve_limit', ?)", (str(MANUAL_JOURNAL_APPROVE_LIMIT),))
    cur.execute("INSERT OR IGNORE INTO settings(key, value) VALUES('late_penalty_percent', ?)", (str(LATE_PENALTY_PERCENT),))

    cur.execute('''
        CREATE TABLE IF NOT EXISTS categories (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            description TEXT,
            is_active INTEGER DEFAULT 1,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS quick_cashier_queue (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            invoice_no TEXT,
            trx_date TEXT,
            total REAL DEFAULT 0,
            payment_method TEXT DEFAULT 'tunai',
            member_id INTEGER,
            customer_name TEXT,
            note TEXT,
            status TEXT DEFAULT 'PENDING',
            created_by INTEGER,
            verified_by INTEGER,
            verified_at TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(member_id) REFERENCES members(id),
            FOREIGN KEY(created_by) REFERENCES users(id),
            FOREIGN KEY(verified_by) REFERENCES users(id)
        )
    ''')

    add_column_if_not_exists('quick_cashier_queue', 'stock_branch_id', 'INTEGER')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS quick_cashier_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            queue_id INTEGER,
            product_id INTEGER,
            product_name TEXT,
            qty REAL DEFAULT 0,
            price REAL DEFAULT 0,
            subtotal REAL DEFAULT 0,
            FOREIGN KEY(queue_id) REFERENCES quick_cashier_queue(id) ON DELETE CASCADE,
            FOREIGN KEY(product_id) REFERENCES products(id)
        )
    ''')
    add_column_if_not_exists('quick_cashier_items', 'stock_branch_id', 'INTEGER')
    add_column_if_not_exists('quick_cashier_items', 'stock_branch_id', 'INTEGER')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS sales_returns (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            return_no TEXT UNIQUE,
            sales_id INTEGER,
            member_id INTEGER,
            return_date TEXT,
            total REAL DEFAULT 0,
            reason TEXT,
            status TEXT DEFAULT 'Posted',
            created_by INTEGER,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(sales_id) REFERENCES sales(id),
            FOREIGN KEY(member_id) REFERENCES members(id),
            FOREIGN KEY(created_by) REFERENCES users(id)
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS sales_return_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            return_id INTEGER,
            product_id INTEGER,
            qty REAL DEFAULT 0,
            price REAL DEFAULT 0,
            subtotal REAL DEFAULT 0,
            FOREIGN KEY(return_id) REFERENCES sales_returns(id) ON DELETE CASCADE,
            FOREIGN KEY(product_id) REFERENCES products(id)
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS supplier_payments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            po_id INTEGER,
            supplier_id INTEGER,
            amount REAL DEFAULT 0,
            payment_date TEXT,
            payment_method TEXT DEFAULT 'tunai',
            note TEXT,
            created_by INTEGER,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(po_id) REFERENCES purchase_orders(id),
            FOREIGN KEY(supplier_id) REFERENCES suppliers(id),
            FOREIGN KEY(created_by) REFERENCES users(id)
        )
    ''')

    # =========================
    # Tabel Multi-Cabang (Branches & Stock per Branch)
    # =========================
    cur.execute('''
        CREATE TABLE IF NOT EXISTS koperasi_branches (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            branch_code TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL,
            address TEXT,
            lat REAL,
            lng REAL,
            phone TEXT,
            is_active INTEGER DEFAULT 1,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS product_branch_stock (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            product_id INTEGER NOT NULL,
            branch_id INTEGER NOT NULL,
            stock REAL DEFAULT 0,
            min_stock REAL DEFAULT 0,
            FOREIGN KEY(product_id) REFERENCES products(id) ON DELETE CASCADE,
            FOREIGN KEY(branch_id) REFERENCES koperasi_branches(id) ON DELETE CASCADE,
            UNIQUE(product_id, branch_id)
        )
    ''')

    # Seed cabang pusat jika belum ada
    cur.execute('SELECT COUNT(*) FROM koperasi_branches')
    if cur.fetchone()[0] == 0:
        cur.execute('INSERT INTO koperasi_branches(branch_code, name, address, lat, lng, phone, is_active) VALUES (?, ?, ?, ?, ?, ?, ?)',
            ('PUSAT', 'Koperasi Pusat', 'Jl. Merdeka No. 1, Jakarta Pusat', -6.2088, 106.8456, '021-12345678', 1))
        cur.execute('INSERT INTO koperasi_branches(branch_code, name, address, lat, lng, phone, is_active) VALUES (?, ?, ?, ?, ?, ?, ?)',
            ('KC-001', 'Koperasi Cabang Bandung', 'Jl. Asia Afrika No. 45, Bandung', -6.9175, 107.6191, '022-87654321', 1))
        cur.execute('INSERT INTO koperasi_branches(branch_code, name, address, lat, lng, phone, is_active) VALUES (?, ?, ?, ?, ?, ?, ?)',
            ('KC-002', 'Koperasi Cabang Surabaya', 'Jl. Tunjungan No. 12, Surabaya', -7.2504, 112.7688, '031-11223344', 1))
        # Copy existing product stock to pusat branch
        cur.execute('INSERT OR IGNORE INTO product_branch_stock(product_id, branch_id, stock, min_stock) SELECT p.id, b.id, p.stock, p.min_stock FROM products p, koperasi_branches b WHERE b.branch_code = "PUSAT"')
    
    cur.execute('SELECT COUNT(*) FROM members')
    if cur.fetchone()[0] == 0:
        cur.execute('INSERT INTO members(member_code, name, phone, address, join_date, status) VALUES (?, ?, ?, ?, ?, ?)', ('MBR-001', 'Member Demo', '08123456789', 'Alamat Demo', today_str(), 'Aktif'))

    cur.execute('SELECT COUNT(*) FROM products')
    if cur.fetchone()[0] == 0:
        demo_products = [
            ('899100100001', 'Beras 5 Kg', 'Sembako', 'sak', 60000, 67000, 25, 5, 1),
            ('899100100002', 'Minyak 1 L', 'Sembako', 'botol', 14500, 16500, 60, 10, 1),
            ('899100100003', 'Gula 1 Kg', 'Sembako', 'kg', 15500, 17500, 40, 8, 1),
        ]
        cur.executemany('INSERT INTO products(barcode, product_name, category, unit, buy_price, sell_price, stock, min_stock, active) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)', demo_products)

    # Tabel Riwayat Stok (baru)
    cur.execute('''
        CREATE TABLE IF NOT EXISTS stock_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            trx_no TEXT,
            trx_date TEXT,
            product_id INTEGER,
            product_name TEXT,
            barcode TEXT,
            qty REAL DEFAULT 0,
            movement_type TEXT,
            origin TEXT,
            destination TEXT,
            stock_before REAL DEFAULT 0,
            stock_after REAL DEFAULT 0,
            reference_type TEXT,
            reference_id INTEGER,
            note TEXT,
            created_by INTEGER,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(product_id) REFERENCES products(id),
            FOREIGN KEY(created_by) REFERENCES users(id)
        )
    ''')

    # Tabel Batch Stok (FIFO untuk HPP)
    cur.execute('''
        CREATE TABLE IF NOT EXISTS product_batches (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            product_id INTEGER NOT NULL,
            branch_id INTEGER,
            po_id INTEGER,
            batch_no TEXT,
            qty_remaining REAL DEFAULT 0,
            initial_qty REAL DEFAULT 0,
            unit_cost REAL DEFAULT 0,
            entry_date TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(product_id) REFERENCES products(id) ON DELETE CASCADE,
            FOREIGN KEY(po_id) REFERENCES purchase_orders(id),
            FOREIGN KEY(branch_id) REFERENCES koperasi_branches(id)
        )
    ''')
    
    # Tambah kolom unit_cost di sales_items jika belum ada
    add_column_if_not_exists('sales_items', 'unit_cost', 'REAL DEFAULT 0')

    conn.commit()
    conn.close()

    # =========================
    # Helper: Notifikasi
    # =========================
def add_notification(user_id, title, message):
    exec_sql('INSERT INTO notifications(user_id, title, message) VALUES (?, ?, ?)', [user_id, title, message])

def get_unread_notifications(user_id):
    return q_all('SELECT * FROM notifications WHERE user_id = ? AND is_read = 0 ORDER BY id DESC LIMIT 10', [user_id])

def _safe_count(sql, params=None):
    """Count helper untuk badge sidebar. Aman jika tabel/kolom fitur belum ada."""
    try:
        row = q_one(sql, params or [])
        return int((row['n'] if row and 'n' in row.keys() else 0) or 0)
    except Exception:
        return 0

def get_sidebar_badge_counts(user=None):
    """Return jumlah item yang butuh perhatian untuk bubble merah di sidebar."""
    user = user or current_user()
    if not user:
        return {}
    role = user['role']
    counts = {}

    def set_count(menu_id, value):
        try:
            value = int(value or 0)
        except Exception:
            value = 0
        if value > 0 and is_menu_visible(menu_id, role):
            counts[menu_id] = value

    # Antrean operasional yang perlu tindakan admin/petugas.
    set_count('quick_cashier_queue', _safe_count("SELECT COUNT(*) n FROM quick_cashier_queue WHERE status='PENDING'"))
    set_count('topup_approval', _safe_count("SELECT COUNT(*) n FROM topup_requests WHERE status='PENDING'"))
    set_count('verify_loan_payments', _safe_count("SELECT COUNT(*) n FROM loan_payments WHERE UPPER(TRIM(COALESCE(status,'')))='PENDING_VERIFICATION' OR (UPPER(TRIM(COALESCE(status,'')))='VERIFIED' AND COALESCE(transfer_proof,'')<>'' AND submitted_at IS NOT NULL AND verified_by IS NULL AND verified_at IS NULL AND NOT EXISTS (SELECT 1 FROM loan_schedules ls WHERE ls.payment_id=loan_payments.id))"))
    set_count('user_approval', _safe_count("SELECT COUNT(*) n FROM users WHERE status='PENDING_APPROVAL'"))
    set_count('approvals', _safe_count("SELECT COUNT(*) n FROM approval_requests WHERE status='Pending'"))
    set_count('stock_opname', _safe_count("SELECT COUNT(*) n FROM stock_opname_sessions WHERE status='SUBMITTED'"))

    # Indikator perhatian, bukan transaksi masuk, tapi tetap membantu admin.
    set_count('products', _safe_count('SELECT COUNT(*) n FROM products WHERE active=1 AND stock<=min_stock'))
    credit_alerts = (
        _safe_count("SELECT COUNT(*) n FROM member_kyc WHERE status='PENDING'") +
        _safe_count("SELECT COUNT(*) n FROM loan_guarantors WHERE status='PENDING'") +
        _safe_count("SELECT COUNT(*) n FROM loan_restructures WHERE status='PENDING'")
    )
    set_count('credit_control', credit_alerts)

    # Badge khusus anggota supaya anggota tahu ada status yang harus dilihat.
    if role == 'user':
        uid = user['id']
        set_count('wallet', _safe_count("""
            SELECT COUNT(*) n
            FROM topup_requests t
            JOIN members m ON m.id=t.member_id
            WHERE REPLACE(m.member_code,'EMP-','') = COALESCE((SELECT employee_number FROM users WHERE id=?),'')
              AND t.status='PENDING'
        """, [uid]))
        user_payment_alerts = _safe_count("SELECT COUNT(*) n FROM loan_payments WHERE created_by=? AND status IN ('RETURNED','REJECTED')", [uid])
        set_count('my_payments', user_payment_alerts)
        set_count('loans', user_payment_alerts)

    return counts

def notify_member(member_id, title, message):
    """Buat notifikasi untuk user yang terhubung ke member."""
    # member_id bisa integer (id dari tabel members) atau string (member_code)
    if isinstance(member_id, int):
        member = q_one('SELECT member_code FROM members WHERE id = ?', [member_id])
        if not member:
            return
        member_code = member['member_code']
    else:
        member_code = member_id
    
    # Extract employee_number dari member_code (format: EMP-123)
    emp_no = member_code.replace('EMP-', '') if member_code.startswith('EMP-') else member_code
    user = q_one('SELECT id FROM users WHERE employee_number = ?', [emp_no])
    if user:
        add_notification(user['id'], title, message)

def notify_loan_payment_verifiers(title, message):
    """Buat notifikasi untuk admin/bendahara/branch_admin saat angsuran baru masuk."""
    for r in q_all("SELECT id FROM users WHERE active=1 AND role IN ('admin','bendahara','branch_admin')"):
        add_notification(r['id'], title, message)

# =========================
# Helper: Riwayat Stok
# =========================
def log_stock_history(product_id, qty, movement_type, origin='', destination='', reference_type='', reference_id=0, note=''):
    """Log setiap perubahan stok ke tabel stock_history."""
    prod = q_one('SELECT barcode, product_name FROM products WHERE id=?', [product_id])
    if not prod:
        return
    # Hitung stok sebelum
    if not origin or origin == 'PUSAT':
        stock_before = float(q_one('SELECT stock FROM products WHERE id=?', [product_id])['stock'] or 0)
    else:
        try:
            bid = int(origin)
            stock_before = float(q_one('SELECT stock FROM product_branch_stock WHERE product_id=? AND branch_id=?', [product_id, bid])['stock'] or 0)
        except:
            stock_before = 0
    # Stok setelah
    if not destination or destination == 'PUSAT':
        stock_after = float(q_one('SELECT stock FROM products WHERE id=?', [product_id])['stock'] or 0)
    else:
        try:
            bid = int(destination)
            stock_after = float(q_one('SELECT stock FROM product_branch_stock WHERE product_id=? AND branch_id=?', [product_id, bid])['stock'] or 0)
        except:
            stock_after = 0
    exec_sql('''INSERT INTO stock_history(trx_no, trx_date, product_id, product_name, barcode, qty, movement_type, origin, destination, stock_before, stock_after, reference_type, reference_id, note, created_by)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
        [gen_code('SH'), today_str(), product_id, prod['product_name'], prod['barcode'], qty, movement_type, origin, destination, stock_before, stock_after, reference_type, reference_id, note, session.get('user_id')])

# =========================
# Helper: FIFO Batch Stok
# =========================
def add_to_batches(product_id, qty, unit_cost, po_id=None, branch_id=None, entry_date=None):
    """Add a new batch entry when stock comes in (e.g. PO receive)."""
    batch_no = gen_code('BATCH')
    exec_sql('''INSERT INTO product_batches(product_id, branch_id, po_id, batch_no, qty_remaining, initial_qty, unit_cost, entry_date)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)''',
        [product_id, branch_id, po_id, batch_no, qty, qty, unit_cost, entry_date or today_str()])

def add_stock_with_fifo(product_id, qty, unit_cost, po_id=None, branch_id=None, entry_date=None, movement_type='Stok Masuk', note=''):
    # Tambah stok fisik dan buat batch FIFO baru jika harga beli berbeda.
    qty = parse_float(qty)
    unit_cost = parse_float(unit_cost)
    if qty <= 0:
        return False
    update_branch_stock(product_id, qty, branch_id, 'add')
    exec_sql('UPDATE products SET buy_price=? WHERE id=?', [unit_cost, product_id])
    add_to_batches(product_id, qty, unit_cost, po_id=po_id, branch_id=branch_id, entry_date=entry_date or today_str())
    log_stock_history(product_id, qty, movement_type, destination=str(branch_id or 'PUSAT'), reference_type='purchase_orders' if po_id else 'stock_adjustment', reference_id=po_id or 0, note=note)
    return True

def remove_stock_with_fifo(product_id, qty, branch_id=None, movement_type='Stok Keluar', note=''):
    # Kurangi stok dan konsumsi batch FIFO agar qty_remaining tetap akurat.
    qty = parse_float(qty)
    if qty <= 0:
        return False
    consume_from_batches(product_id, qty, branch_id)
    update_branch_stock(product_id, qty, branch_id, 'subtract')
    log_stock_history(product_id, qty, movement_type, origin=str(branch_id or 'PUSAT'), reference_type='stock_adjustment', reference_id=0, note=note)
    return True

def ensure_fifo_batch_for_existing_stock(product_id, branch_id=None):
    """Buat batch awal untuk stok lama yang belum punya batch FIFO."""
    if branch_id:
        row = q_one('SELECT COALESCE(SUM(qty_remaining),0) n FROM product_batches WHERE product_id=? AND branch_id=?', [product_id, branch_id])
    else:
        row = q_one('SELECT COALESCE(SUM(qty_remaining),0) n FROM product_batches WHERE product_id=? AND branch_id IS NULL', [product_id])
    batch_sum = float(row['n'] or 0) if row else 0
    current = get_product_stock(product_id, branch_id)
    missing = float(current or 0) - batch_sum
    if missing > 0:
        prod = q_one('SELECT buy_price FROM products WHERE id=?', [product_id])
        cost = float(prod['buy_price'] or 0) if prod else 0
        add_to_batches(product_id, missing, cost, po_id=None, branch_id=branch_id, entry_date=today_str())

def consume_from_batches(product_id, qty, branch_id=None):
    """Consume qty from oldest batches (FIFO). Returns weighted average unit_cost used.
    Raises Exception if insufficient batch stock."""
    qty = parse_float(qty)
    ensure_fifo_batch_for_existing_stock(product_id, branch_id)
    remaining = qty
    total_cost = 0.0
    if branch_id:
        batches = q_all('SELECT * FROM product_batches WHERE product_id=? AND branch_id=? AND qty_remaining > 0 ORDER BY id ASC',
                        [product_id, branch_id])
    else:
        batches = q_all('SELECT * FROM product_batches WHERE product_id=? AND branch_id IS NULL AND qty_remaining > 0 ORDER BY id ASC',
                        [product_id])
    # Also check batches where branch_id matches current
    if not batches and branch_id:
        batches = q_all('SELECT * FROM product_batches WHERE product_id=? AND branch_id=? AND qty_remaining > 0 ORDER BY id ASC',
                        [product_id, branch_id])
    for b in batches:
        if remaining <= 0:
            break
        avail = float(b['qty_remaining'])
        take = min(avail, remaining)
        total_cost += take * float(b['unit_cost'])
        new_remaining = avail - take
        exec_sql('UPDATE product_batches SET qty_remaining=? WHERE id=?', [new_remaining, b['id']])
        remaining -= take
    if remaining > 0:
        raise ValueError(f'Batch FIFO tidak cukup untuk produk {product_id}. Kekurangan {remaining}')
    avg_cost = total_cost / qty if qty > 0 else 0
    return round(avg_cost, 2)

# =========================
# Helper: Hitung Denda
# =========================
def calculate_penalty(due_date_str, payment_date_str, amount, penalty_rate=None):
    if penalty_rate is None:
        penalty_rate = float(get_setting('late_penalty_percent', LATE_PENALTY_PERCENT))
    try:
        due = datetime.strptime(due_date_str, '%Y-%m-%d').date()
        pay = datetime.strptime(payment_date_str, '%Y-%m-%d').date()
        if pay <= due:
            return 0
        diff_days = (pay - due).days
        diff_months = max(1, diff_days // 30)
        penalty = amount * penalty_rate * diff_months
        return round(penalty, 2)
    except:
        return 0

# =========================
# Helper: SHU Advanced Periodic Allocation
# =========================
def ensure_shu_tables():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute('''CREATE TABLE IF NOT EXISTS shu_periods(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        year INTEGER UNIQUE,
        start_date TEXT,
        end_date TEXT,
        net_profit_snapshot REAL DEFAULT 0,
        distributable_profit REAL DEFAULT 0,
        reserve_percent REAL DEFAULT 40,
        savings_weight REAL DEFAULT 50,
        shopping_weight REAL DEFAULT 30,
        loan_weight REAL DEFAULT 20,
        status TEXT DEFAULT 'DRAFT',
        note TEXT,
        created_by INTEGER,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        updated_by INTEGER,
        updated_at TEXT,
        distributed_by INTEGER,
        distributed_at TEXT
    )''')
    cur.execute('''CREATE TABLE IF NOT EXISTS shu_member_adjustments(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        year INTEGER,
        member_id INTEGER,
        manual_amount REAL DEFAULT 0,
        note TEXT,
        created_by INTEGER,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(year, member_id)
    )''')
    cur.execute('''CREATE TABLE IF NOT EXISTS shu_distribution_details(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        year INTEGER,
        member_id INTEGER,
        savings_point REAL DEFAULT 0,
        shopping_point REAL DEFAULT 0,
        loan_point REAL DEFAULT 0,
        shu_savings REAL DEFAULT 0,
        shu_shopping REAL DEFAULT 0,
        shu_loan REAL DEFAULT 0,
        manual_amount REAL DEFAULT 0,
        total_shu REAL DEFAULT 0,
        distributed_at TEXT,
        UNIQUE(year, member_id)
    )''')
    cur.execute('''CREATE TABLE IF NOT EXISTS opening_balances(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        balance_date TEXT,
        account_id INTEGER,
        direction TEXT,
        amount REAL DEFAULT 0,
        note TEXT,
        journal_ref TEXT,
        created_by INTEGER,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )''')
    conn.commit(); conn.close()

def shu_period_bounds(year):
    ensure_shu_tables()
    row = q_one('SELECT * FROM shu_periods WHERE year=?', [year])
    if row:
        return row['start_date'], row['end_date']
    return f'{year}-01-01', f'{year}-12-31'

def shu_net_profit(start_date, end_date):
    rows = q_all('''SELECT a.category, COALESCE(SUM(j.debit),0) debit, COALESCE(SUM(j.credit),0) credit
                    FROM journal_entries j JOIN accounts a ON a.id=j.account_id
                    WHERE j.entry_date>=? AND j.entry_date<=?
                    GROUP BY a.category''', [start_date, end_date])
    revenue = expense = 0.0
    for r in rows:
        if r['category'] == 'Pendapatan':
            revenue += float(r['credit'] or 0) - float(r['debit'] or 0)
        elif r['category'] == 'Beban':
            expense += float(r['debit'] or 0) - float(r['credit'] or 0)
    return round(revenue - expense, 2), round(revenue, 2), round(expense, 2)

def ensure_shu_period(year):
    ensure_shu_tables()
    row = q_one('SELECT * FROM shu_periods WHERE year=?', [year])
    if row:
        return row
    start_date, end_date = f'{year}-01-01', f'{year}-12-31'
    net, revenue, expense = shu_net_profit(start_date, end_date)
    distributable = max(0, net * 0.60)
    exec_sql('''INSERT OR IGNORE INTO shu_periods(year,start_date,end_date,net_profit_snapshot,distributable_profit,reserve_percent,savings_weight,shopping_weight,loan_weight,status,created_by)
                VALUES(?,?,?,?,?,?,?,?,?,?,?)''', [year,start_date,end_date,net,distributable,40,50,30,20,'DRAFT',session.get('user_id')])
    return q_one('SELECT * FROM shu_periods WHERE year=?', [year])

def shu_member_points(member_id, start_date, end_date):
    savings = float(q_one('''SELECT COALESCE(SUM(CASE WHEN direction="Masuk" THEN amount ELSE -amount END),0) x
                             FROM savings_transactions WHERE member_id=? AND trx_date>=? AND trx_date<=?''', [member_id,start_date,end_date])['x'] or 0)
    shopping = float(q_one('''SELECT COALESCE(SUM(total),0) x FROM sales
                              WHERE member_id=? AND trx_date>=? AND trx_date<=? AND status="Posted"''', [member_id,start_date,end_date])['x'] or 0)
    loan_pay = float(q_one('''SELECT COALESCE(SUM(lp.amount),0) x FROM loan_payments lp
                              JOIN loans l ON l.id=lp.loan_id
                              WHERE l.member_id=? AND lp.payment_date>=? AND lp.payment_date<=? AND UPPER(TRIM(COALESCE(lp.status,'')))='VERIFIED' ''', [member_id,start_date,end_date])['x'] or 0)
    return savings, shopping, loan_pay

def calculate_shu_detail(member_id, year=None):
    if year is None:
        year = datetime.now().year
    year = int(year)
    period = ensure_shu_period(year)
    start_date, end_date = period['start_date'], period['end_date']
    net_profit, revenue, expense = shu_net_profit(start_date, end_date)
    # Jika admin isi distributable_profit, pakai angka itu. Jika belum, pakai laba bersih dikurangi cadangan.
    distributable = float(period['distributable_profit'] or 0)
    if distributable <= 0:
        distributable = max(0, net_profit * (100 - float(period['reserve_percent'] or 0)) / 100)
    sw, bw, lw = float(period['savings_weight'] or 0), float(period['shopping_weight'] or 0), float(period['loan_weight'] or 0)
    weight_total = sw + bw + lw
    if weight_total <= 0:
        sw, bw, lw, weight_total = 50, 30, 20, 100
    savings_pool = distributable * sw / weight_total
    shopping_pool = distributable * bw / weight_total
    loan_pool = distributable * lw / weight_total
    savings, shopping, loan_pay = shu_member_points(member_id, start_date, end_date)
    totals = q_one('''SELECT
        COALESCE((SELECT SUM(CASE WHEN direction="Masuk" THEN amount ELSE -amount END) FROM savings_transactions WHERE trx_date>=? AND trx_date<=?),0) savings_all,
        COALESCE((SELECT SUM(total) FROM sales WHERE trx_date>=? AND trx_date<=? AND status="Posted"),0) shopping_all,
        COALESCE((SELECT SUM(lp.amount) FROM loan_payments lp WHERE lp.payment_date>=? AND lp.payment_date<=? AND UPPER(TRIM(COALESCE(lp.status,'')))='VERIFIED'),0) loan_all''',
        [start_date,end_date,start_date,end_date,start_date,end_date])
    savings_all = float(totals['savings_all'] or 0)
    shopping_all = float(totals['shopping_all'] or 0)
    loan_all = float(totals['loan_all'] or 0)
    shu_savings = savings / savings_all * savings_pool if savings_all > 0 else 0
    shu_shopping = shopping / shopping_all * shopping_pool if shopping_all > 0 else 0
    shu_loan = loan_pay / loan_all * loan_pool if loan_all > 0 else 0
    adj = q_one('SELECT manual_amount,note FROM shu_member_adjustments WHERE year=? AND member_id=?', [year,member_id])
    manual = float(adj['manual_amount'] or 0) if adj else 0
    total = max(0, shu_savings + shu_shopping + shu_loan + manual)
    return {
        'year': year, 'period': period, 'start_date': start_date, 'end_date': end_date,
        'net_profit': net_profit, 'revenue': revenue, 'expense': expense, 'distributable': distributable,
        'savings': savings, 'shopping': shopping, 'loan_pay': loan_pay,
        'savings_all': savings_all, 'shopping_all': shopping_all, 'loan_all': loan_all,
        'shu_savings': round(shu_savings,2), 'shu_shopping': round(shu_shopping,2), 'shu_loan': round(shu_loan,2),
        'manual': round(manual,2), 'adjust_note': adj['note'] if adj else '', 'total': round(total,2)
    }

def calculate_shu(member_id, year=None):
    return calculate_shu_detail(member_id, year)['total']

# =========================
# Decorators
# =========================
def login_required(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return func(*args, **kwargs)
    return wrapper

def role_required(*roles):
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            user = current_user()
            if not user:
                return redirect(url_for('login'))
            if user['role'] not in roles:
                flash('Menu ini tidak aktif untuk role Anda. Hubungi admin jika akses diperlukan.', 'warning')
                return redirect(url_for('dashboard'))
            return func(*args, **kwargs)
        return wrapper
    return decorator

# =========================
# UI Helpers
# =========================
def bar_chart_svg(items, title='Chart', width=560, height=220, color='#22c55e'):
    if not items:
        return '<div class="muted">Belum ada data.</div>'
    values = [float(v) for _, v in items]
    maxv = max(values) if max(values) > 0 else 1
    left, top, bottom = 40, 20, 30
    plot_w = width - left - 10
    plot_h = height - top - bottom
    bw = max(20, plot_w // max(1, len(items) * 2))
    gap = bw
    x = left
    parts = [f'<svg width="100%" viewBox="0 0 {width} {height}" xmlns="http://www.w3.org/2000/svg">']
    parts.append(f'<text x="10" y="15" fill="#93c5fd" font-size="12">{title}</text>')
    parts.append(f'<line x1="{left}" y1="{top+plot_h}" x2="{left+plot_w}" y2="{top+plot_h}" stroke="#334155"/>')
    for label, val in items:
        h = (float(val) / maxv) * (plot_h - 10)
        y = top + plot_h - h
        parts.append(f'<rect x="{x}" y="{y}" width="{bw}" height="{h}" rx="6" fill="{color}"/>')
        parts.append(f'<text x="{x+bw/2}" y="{top+plot_h+14}" text-anchor="middle" fill="#94a3b8" font-size="10">{label}</text>')
        parts.append(f'<text x="{x+bw/2}" y="{max(top+12, y-4)}" text-anchor="middle" fill="#e5e7eb" font-size="10">{int(val)}</text>')
        x += bw + gap
    parts.append('</svg>')
    return ''.join(parts)

def grouped_bar_chart_svg(items_a, items_b, title='Chart', width=560, height=240, color_a='#2563eb', color_b='#10b981', label_a='Series A', label_b='Series B'):
    if not items_a and not items_b:
        return '<div class="muted">Belum ada data.</div>'
    labels = [l for l, _ in items_a]
    vals_a = [float(v) for _, v in items_a]
    b_map = {l: float(v) for l, v in items_b}
    vals_b = [b_map.get(l, 0) for l in labels]
    all_vals = vals_a + vals_b
    maxv = max(all_vals) if all_vals and max(all_vals) > 0 else 1
    left, top, bottom = 45, 20, 35
    plot_w = width - left - 10
    plot_h = height - top - bottom
    n = max(1, len(labels))
    group_w = plot_w / n
    bw = max(8, group_w / 4)
    x = left
    parts = [f'<svg width="100%" viewBox="0 0 {width} {height}" xmlns="http://www.w3.org/2000/svg">']
    parts.append(f'<text x="10" y="15" fill="#93c5fd" font-size="12">{title}</text>')
    parts.append(f'<line x1="{left}" y1="{top+plot_h}" x2="{left+plot_w}" y2="{top+plot_h}" stroke="#334155"/>')
    lx = left + 10
    parts.append(f'<rect x="{lx}" y="{height-14}" width="10" height="10" rx="2" fill="{color_a}"/>')
    parts.append(f'<text x="{lx+14}" y="{height-5}" fill="#94a3b8" font-size="10">{label_a}</text>')
    lx2 = lx + 80
    parts.append(f'<rect x="{lx2}" y="{height-14}" width="10" height="10" rx="2" fill="{color_b}"/>')
    parts.append(f'<text x="{lx2+14}" y="{height-5}" fill="#94a3b8" font-size="10">{label_b}</text>')
    for i, label in enumerate(labels):
        va = vals_a[i] if i < len(vals_a) else 0
        vb = vals_b[i] if i < len(vals_b) else 0
        ha = (va / maxv) * (plot_h - 10) if maxv > 0 else 0
        hb = (vb / maxv) * (plot_h - 10) if maxv > 0 else 0
        mid = x + group_w / 2
        ax = mid - bw - 1
        bx = mid + 1
        ya = top + plot_h - ha
        yb = top + plot_h - hb
        if ha > 0:
            parts.append(f'<rect x="{ax}" y="{ya}" width="{bw}" height="{ha}" rx="4" fill="{color_a}"/>')
            parts.append(f'<text x="{ax+bw/2}" y="{max(top+12, ya-3)}" text-anchor="middle" fill="#e5e7eb" font-size="9">{int(va):,}</text>')
        if hb > 0:
            parts.append(f'<rect x="{bx}" y="{yb}" width="{bw}" height="{hb}" rx="4" fill="{color_b}"/>')
            parts.append(f'<text x="{bx+bw/2}" y="{max(top+12, yb-3)}" text-anchor="middle" fill="#e5e7eb" font-size="9">{int(vb):,}</text>')
        parts.append(f'<text x="{mid}" y="{top+plot_h+14}" text-anchor="middle" fill="#94a3b8" font-size="10">{label}</text>')
        x += group_w
    parts.append('</svg>')
    return ''.join(parts)

def pie_chart_svg(items, title='Chart', width=400, height=280, colors=None):
    """SVG pie chart sederhana."""
    if not items:
        return '<div class="muted">Belum ada data.</div>'
    if colors is None:
        colors = ['#4f46e5','#10b981','#f59e0b','#ef4444','#8b5cf6','#ec4899','#14b8a6','#f97316','#6366f1','#84cc16']
    values = [max(0, float(v)) for _, v in items]
    total = sum(values)
    if total <= 0:
        return '<div class="muted">Belum ada data.</div>'
    cx, cy, r = 130, 140, 110
    parts = [f'<svg width="100%" viewBox="0 0 {width} {height}" xmlns="http://www.w3.org/2000/svg">']
    parts.append(f'<text x="10" y="15" fill="#93c5fd" font-size="12">{title}</text>')
    angle_start = -90
    cos = math.cos
    sin = math.sin
    for i, (label, val) in enumerate(items):
        val_f = max(0, float(val))
        if val_f <= 0: continue
        angle = (val_f / total) * 360
        angle_end = angle_start + angle
        a1_rad = angle_start * 3.14159 / 180
        a2_rad = angle_end * 3.14159 / 180
        x1 = cx + r * cos(a1_rad)
        y1 = cy + r * sin(a1_rad)
        x2 = cx + r * cos(a2_rad)
        y2 = cy + r * sin(a2_rad)
        large = 1 if angle > 180 else 0
        color = colors[i % len(colors)]
        parts.append(f'<path d="M{cx},{cy} L{x1:.1f},{y1:.1f} A{r},{r} 0 {large},1 {x2:.1f},{y2:.1f} Z" fill="{color}" stroke="white" stroke-width="2"/>')
        parts.append(f'<text x="260" y="{30 + i*22}" fill="#e5e7eb" font-size="12"><tspan fill="{color}" font-size="14">●</tspan> {label[:18]}</text>')
        parts.append(f'<text x="350" y="{30 + i*22}" text-anchor="end" fill="#94a3b8" font-size="12">{int(val_f):,}</text>')
        angle_start = angle_end
    parts.append('</svg>')
    return ''.join(parts)

def simple_qr_svg(data, size=140):
    """Generate a real QR code SVG from data string."""
    qr = qrcode.QRCode(version=None, error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=10, border=2)
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color='#1e293b', back_color='white').convert('RGB')
    # Convert to base64 PNG
    buf = _BytesIO()
    img.save(buf, format='PNG')
    b64 = base64.b64encode(buf.getvalue()).decode()
    return f'<img src="data:image/png;base64,{b64}" width="{size}" height="{size}" style="image-rendering:pixelated;" alt="QR Code"/>'

def horizontal_bar_chart_svg(items, title='Chart', width=560, height=200, colors=None):
    if not items:
        return '<div class="muted">Belum ada data.</div>'
    if colors is None:
        colors = ['#2563eb', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899']
    values = [float(v) for _, v in items]
    maxv = max(values) if max(values) > 0 else 1
    left = 110
    bar_h = max(14, min(24, (height - 20) // len(items)))
    gap = 4
    parts = [f'<svg width="100%" viewBox="0 0 {width} {height}" xmlns="http://www.w3.org/2000/svg">']
    parts.append(f'<text x="10" y="15" fill="#93c5fd" font-size="12">{title}</text>')
    y = 28
    for i, (label, val) in enumerate(items):
        color = colors[i % len(colors)]
        bw = (float(val) / maxv) * (width - left - 20)
        bw = max(2, bw)
        parts.append(f'<text x="{left-8}" y="{y+bar_h/2+4}" text-anchor="end" fill="#94a3b8" font-size="11">{label[:16]}</text>')
        parts.append(f'<rect x="{left}" y="{y}" width="{bw}" height="{bar_h}" rx="4" fill="{color}"/>')
        parts.append(f'<text x="{left+bw+6}" y="{y+bar_h/2+4}" fill="#e5e7eb" font-size="10">{int(val):,}</text>')
        y += bar_h + gap
    parts.append('</svg>')
    return ''.join(parts)

# =========================
# UI REDESIGN: render_page with modern, clean design
# =========================
CSS_DESIGN = '''
/* ==========================================================================
   KOPERASI ENTERPRISE — DESIGN SYSTEM V60 (total rebuild, mobile-first)
   Satu sumber kebenaran. Tidak ada lagi patch bertumpuk.
   ========================================================================== */

/* ---------- 1. RESET & TOKENS ---------- */
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
html{-webkit-text-size-adjust:100%;text-size-adjust:100%}
:root{
  --brand:#0f7a55;--brand-dark:#0b5c40;--brand-light:#eaf7f1;--brand-soft:#f3faf6;
  --ink:#10231a;--text:#1b2b23;--muted:#64746c;--muted-2:#8a978f;
  --border:#dfe8e2;--border-soft:#edf1ee;--bg:#f5f8f6;--surface:#ffffff;
  --danger:#c2372a;--danger-bg:#fdecea;--warning:#b45309;--warning-bg:#fff8e9;
  --info:#2563eb;--info-bg:#eaf1fd;--success:#0f7a55;--success-bg:#eaf7f1;
  --radius-sm:10px;--radius-md:16px;--radius-lg:22px;--radius-pill:999px;
  --shadow-sm:0 4px 14px rgba(16,45,37,.05);--shadow:0 8px 24px rgba(16,45,37,.07);--shadow-lg:0 18px 46px rgba(16,45,37,.12);
  --header-h:58px;--nav-h:64px;--sidebar-w:264px;--sidebar-w-compact:78px;
  --primary:var(--brand);--danger-color:var(--danger);
  --safe-b:env(safe-area-inset-bottom,0px);--safe-t:env(safe-area-inset-top,0px);
}
html,body{height:100%}
body{font-family:'Segoe UI',Inter,system-ui,-apple-system,Roboto,Arial,sans-serif;background:var(--bg);color:var(--text);font-size:14px;line-height:1.5;-webkit-font-smoothing:antialiased}
a{color:var(--brand);text-decoration:none}
a:hover{text-decoration:underline}
img{max-width:100%;display:block}
button{font-family:inherit;cursor:pointer}
input,select,textarea{font-family:inherit;color:inherit}
h1,h2,h3,h4{color:var(--ink);line-height:1.25;font-weight:800}
table{border-collapse:collapse;width:100%}
.sr-only{position:absolute;width:1px;height:1px;overflow:hidden;clip:rect(0,0,0,0);white-space:nowrap}
.skip-link{position:absolute;left:8px;top:-48px;background:var(--brand);color:#fff;padding:10px 16px;border-radius:var(--radius-sm);z-index:1000;transition:top .15s}
.skip-link:focus{top:8px}
::-webkit-scrollbar{width:10px;height:10px}
::-webkit-scrollbar-thumb{background:#c9d6cf;border-radius:99px}
::selection{background:var(--brand-light);color:var(--brand-dark)}

/* ---------- 2. LAYOUT SHELL ---------- */
.app-wrap{display:flex;min-height:100vh;background:var(--bg)}

/* Sidebar (desktop = fixed column, mobile = slide-in drawer) */
.sidebar{width:var(--sidebar-w);flex:0 0 var(--sidebar-w);background:var(--surface);border-right:1px solid var(--border);display:flex;flex-direction:column;position:sticky;top:0;height:100vh;overflow-y:auto;transition:width .18s ease,flex-basis .18s ease;z-index:40}
body.sidebar-compact .sidebar{width:var(--sidebar-w-compact);flex-basis:var(--sidebar-w-compact)}
.mobile-drawer-close{display:none}
.sidebar-brand{display:flex;align-items:center;gap:10px;padding:18px 16px;border-bottom:1px solid var(--border-soft)}
.brand-icon{width:38px;height:38px;border-radius:12px;background:var(--brand-light);display:grid;place-items:center;font-size:19px;flex:none}
.brand-text{font-weight:900;font-size:16px;color:var(--ink);letter-spacing:.02em;white-space:nowrap;overflow:hidden}
.sidebar-user{display:flex;align-items:center;gap:10px;padding:12px 16px;border-bottom:1px solid var(--border-soft)}
.avatar{width:36px;height:36px;border-radius:50%;background:var(--brand);color:#fff;display:grid;place-items:center;font-weight:800;flex:none}
.user-info{min-width:0;flex:1}
.uname{font-weight:800;font-size:12.5px;color:var(--ink);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.urole{font-size:10.5px;color:var(--muted);text-transform:capitalize}
.sidebar-collapse-btn{width:26px;height:26px;border-radius:8px;border:1px solid var(--border);background:var(--surface);color:var(--muted);flex:none;display:grid;place-items:center}
body.sidebar-compact .sidebar-collapse-btn span{transform:rotate(180deg)}
.sidebar-search{position:relative;margin:10px 14px;display:flex;align-items:center}
.sidebar-search input{width:100%;height:38px;border-radius:var(--radius-pill);border:1px solid var(--border);background:var(--bg);padding:0 34px;font-size:12.5px}
.sidebar-search-icon{position:absolute;left:12px;color:var(--muted-2);font-size:13px}
.sidebar-search-clear{position:absolute;right:30px;border:none;background:none;color:var(--muted-2);font-size:13px;display:none}
.sidebar-search-key{position:absolute;right:10px;font-size:10px;color:var(--muted-2);border:1px solid var(--border);border-radius:5px;padding:1px 5px;background:var(--surface)}
.sidebar-nav{flex:1;overflow-y:auto;padding:6px 10px 10px}
.sidebar-section-header{display:flex;align-items:center;gap:8px;width:100%;padding:9px 8px;border:none;background:none;color:var(--muted);font-weight:800;font-size:10.5px;text-transform:uppercase;letter-spacing:.07em;border-radius:8px}
.sidebar-section-header:hover{background:var(--brand-soft)}
.section-symbol{font-size:12px}
.section-name{flex:1;text-align:left;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.section-count{font-size:9.5px;background:var(--border-soft);border-radius:99px;padding:1px 6px}
.section-chevron{transition:transform .15s;font-size:10px}
.collapsed .section-chevron{transform:rotate(-90deg)}
.nav-icon{width:20px;text-align:center;font-size:14px;flex:none}
.nav-label{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.sidebar-nav a,.sidebar-menu-link{display:flex;align-items:center;gap:10px;padding:9px 10px;border-radius:10px;color:var(--text);font-size:12.5px;font-weight:600}
.sidebar-nav a:hover{background:var(--brand-soft);text-decoration:none}
.sidebar-nav a.active{background:var(--brand-light);color:var(--brand-dark);font-weight:800}
.sidebar-section-items{display:none;flex-direction:column;gap:1px;padding:2px 0 8px 4px}
.sidebar-section.open .sidebar-section-items{display:flex}
.sidebar-section.open>.sidebar-section-header .section-chevron{transform:rotate(180deg)}
.sidebar-notif-badge{margin-left:auto;min-width:18px;height:18px;padding:0 5px;border-radius:99px;display:none;align-items:center;justify-content:center;background:var(--danger);color:#fff;font-size:9.5px;font-weight:900;line-height:18px}
.sidebar-notif-badge.show{display:inline-flex}
.sidebar-fixed-footer{border-top:1px solid var(--border-soft);padding:10px;display:flex;flex-direction:column;gap:2px}
.sidebar-footer-link{display:flex;align-items:center;gap:10px;padding:9px 10px;border-radius:10px;border:none;background:none;width:100%;text-align:left;color:var(--text);font-size:12.5px;font-weight:700}
.sidebar-footer-link:hover{background:var(--brand-soft)}
.sidebar-logout{color:var(--danger)}
body.sidebar-compact .brand-text,body.sidebar-compact .user-info,body.sidebar-compact .sidebar-search,
body.sidebar-compact .nav-label,body.sidebar-compact .section-name,body.sidebar-compact .section-count,
body.sidebar-compact .section-chevron{display:none}
body.sidebar-compact .sidebar-nav a,body.sidebar-compact .sidebar-footer-link{justify-content:center}

.main-content{flex:1;min-width:0;display:flex;flex-direction:column}
.mobile-user-card,.mobile-grid-wrap,.mobile-header,.mobile-nav,.hamburger.desktop-menu-btn,
.mobile-notif-backdrop,.mobile-notif-sheet,.mobile-app-sheet,.mobile-app-backdrop,.mobile-drawer-backdrop{display:none}

.topbar{display:flex;align-items:flex-start;justify-content:space-between;gap:14px;padding:22px 28px 6px;flex-wrap:wrap}
.page-header h1{font-size:21px}
.page-header p{color:var(--muted);font-size:12.5px;margin-top:2px}
.topbar-actions{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
main.main-content>*:not(.topbar):not(.mobile-user-card):not(.mobile-grid-wrap){padding-left:28px;padding-right:28px}
main.main-content>.footer{padding-bottom:28px}
.footer{color:var(--muted-2);font-size:11px;text-align:center;padding:24px 0 10px}

/* ---------- 3. BUTTONS / BADGES / ALERTS ---------- */
.btn,.ds-btn{display:inline-flex;align-items:center;justify-content:center;gap:6px;height:40px;padding:0 16px;border-radius:12px;border:1px solid var(--border);background:var(--surface);color:var(--text);font-weight:700;font-size:12.5px;line-height:1;white-space:nowrap}
.btn:hover,.ds-btn:hover{border-color:var(--brand);color:var(--brand-dark);text-decoration:none}
.btn-success,.ds-btn:not(.ds-btn-secondary):not(.ds-btn-danger){background:var(--brand);border-color:var(--brand);color:#fff}
.btn-success:hover{background:var(--brand-dark);color:#fff}
.btn-danger,.ds-btn-danger{background:var(--danger);border-color:var(--danger);color:#fff}
.btn-danger:hover{background:#a12e22;color:#fff}
.btn-warn{background:var(--warning);border-color:var(--warning);color:#fff}
.btn-ghost,.ds-btn-secondary{background:transparent}
.btn-sm{height:32px;padding:0 12px;font-size:11.5px;border-radius:9px}
.btn[disabled],.btn.is-loading{opacity:.6;cursor:not-allowed}
.badge{display:inline-flex;align-items:center;gap:4px;padding:3px 10px;border-radius:var(--radius-pill);font-size:10.5px;font-weight:800;background:var(--border-soft);color:var(--muted)}
.badge-success{background:var(--success-bg);color:var(--brand-dark)}
.badge-danger{background:var(--danger-bg);color:var(--danger)}
.badge-warn{background:var(--warning-bg);color:var(--warning)}
.badge-info{background:var(--info-bg);color:var(--info)}
.badge-gray{background:var(--border-soft);color:var(--muted)}
.alert{padding:12px 16px;border-radius:var(--radius-sm);border:1px solid var(--border);background:var(--surface);font-size:12.5px;margin-bottom:14px}
.alert-info{background:var(--info-bg);border-color:#c7dcfb;color:#1d4ed8}
.flash{max-width:1180px;margin:14px auto 0;padding:12px 18px;border-radius:var(--radius-sm);font-size:12.5px;font-weight:700}
.flash.success,.flash-success{background:var(--success-bg);color:var(--brand-dark);border:1px solid #b7e4d5}
.flash.error,.flash-error{background:var(--danger-bg);color:var(--danger);border:1px solid #f3c3bd}
.flash-warning{background:var(--warning-bg);color:var(--warning);border:1px solid #f4d49b}
.spinner{width:16px;height:16px;border:2px solid var(--border);border-top-color:var(--brand);border-radius:50%;animation:spin .7s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
.done{color:var(--brand)}.edit{color:var(--info)}.danger{color:var(--danger)}.primary{color:var(--brand)}
.muted,.sub,.small{color:var(--muted);font-size:11.5px}
.green{color:var(--brand)}.blue{color:var(--info)}
.text-center{text-align:center}.text-right{text-align:right}

/* ---------- 4. LAYOUT PRIMITIVES: grid / card / table ---------- */
.grid{display:grid;grid-template-columns:repeat(12,1fr);gap:16px;margin:16px 0}
.col-4{grid-column:span 4}.col-5{grid-column:span 5}.col-6{grid-column:span 6}
.col-7{grid-column:span 7}.col-8{grid-column:span 8}.col-12{grid-column:span 12}
.card,.ds-card{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius-md);padding:18px;box-shadow:var(--shadow-sm)}
.card-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:14px;margin:16px 0}
.ds-card-head{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:12px}
.metrics{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;margin:16px 0}
.metric{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius-md);padding:16px;box-shadow:var(--shadow-sm)}
.metric .label,.metric label{display:block;font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.06em;color:var(--muted)}
.metric .value,.metric .val{display:block;margin-top:6px;font-size:21px;font-weight:900;color:var(--ink)}
.metric-compact{padding:12px}
.table-wrap{overflow-x:auto;background:var(--surface);border:1px solid var(--border);border-radius:var(--radius-md);box-shadow:var(--shadow-sm);-webkit-overflow-scrolling:touch}
table th{text-align:left;font-size:10px;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);padding:12px 14px;border-bottom:1px solid var(--border);background:var(--brand-soft);white-space:nowrap}
table td{padding:11px 14px;border-bottom:1px solid var(--border-soft);font-size:12.5px;vertical-align:middle}
table tr:last-child td{border-bottom:none}
table tr:hover td{background:var(--brand-soft)}
.cell-money{font-weight:800;white-space:nowrap;font-variant-numeric:tabular-nums}
.cell-status,.cell-actions{white-space:nowrap}
.cell-min{width:1%;white-space:nowrap}
.cell-barcode{font-family:monospace;font-size:11px}
.cell-product{min-width:160px}
.cell-stock{font-weight:800}
.timeline{display:flex;flex-direction:column;gap:0;position:relative}
.timeline-item{position:relative;padding:0 0 18px 26px;border-left:2px solid var(--border);margin-left:6px}
.timeline-item:last-child{border-color:transparent}
.timeline-dot{position:absolute;left:-7px;top:0;width:12px;height:12px;border-radius:50%;background:var(--brand);border:2px solid var(--surface)}
.activity-list,.activity-stream{display:flex;flex-direction:column;gap:2px}
.activity-item{display:flex;gap:10px;padding:10px 0;border-bottom:1px solid var(--border-soft);font-size:12px}

/* ---------- 5. FORMS ---------- */
.form-group{margin-bottom:14px}
.form-group label,.label,.lbl,.lb{display:block;font-size:11.5px;font-weight:800;color:var(--ink);margin-bottom:5px}
input,select,textarea{width:100%;min-height:42px;padding:0 13px;border:1px solid var(--border);border-radius:11px;background:var(--surface);font-size:14px}
textarea{min-height:90px;padding:10px 13px}
input:focus,select:focus,textarea:focus{outline:none;border-color:var(--brand);box-shadow:0 0 0 3px var(--brand-light)}
input[type=file]{min-height:auto;padding:8px;font-size:12px}
input[type=checkbox],input[type=radio]{width:16px;min-height:16px;height:16px;flex:none}
.field-hint,.field-note{font-size:10.5px;color:var(--muted);margin-top:4px}
.form-section{border-top:1px solid var(--border-soft);padding-top:16px;margin-top:16px}
.form-section:first-child{border-top:none;padding-top:0;margin-top:0}
.form-step{padding:16px;border:1px solid var(--border);border-radius:var(--radius-md);margin-bottom:12px}
.ds-field{margin-bottom:14px}
.ds-form-grid,.settings-form-grid,.role-menu-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:14px}
.ds-form-error{color:var(--danger);font-size:11px;margin-top:4px}
.input-suffix{position:relative}
.money-input{font-variant-numeric:tabular-nums}
.range-scale{display:flex;justify-content:space-between;font-size:10px;color:var(--muted)}
.progress-bar{width:100%;height:8px;border-radius:99px;background:var(--border-soft);overflow:hidden}
.progress-fill{height:100%;background:var(--brand);border-radius:99px}

/* ---------- 6. HERO PATTERN (dipakai banyak halaman: bank/gov/loan/mx/qc/verification/dsb) ---------- */
.bank-hero,.gov-hero,.loan-hero,.mx-hero,.qc-verify-hero,.verification-hero,.verification-v2-hero,
.po-audit-hero,.digital-card-hero,.ccs-hero,.command-hero,.neo-command-hero,.mh-insight-hero,
.wallet-risk-hero,.command-center-v13{
  background:linear-gradient(135deg,var(--brand-dark),var(--brand));color:#fff;border-radius:var(--radius-lg);
  padding:24px;display:flex;align-items:center;justify-content:space-between;gap:16px;flex-wrap:wrap;
  box-shadow:var(--shadow);margin-bottom:16px;position:relative;overflow:hidden
}
.bank-hero *,.gov-hero *,.loan-hero *,.mx-hero *,.qc-verify-hero *,.verification-hero *,.verification-v2-hero *,
.po-audit-hero *,.digital-card-hero *,.ccs-hero *,.command-hero *,.neo-command-hero *,.mh-insight-hero *,
.wallet-risk-hero *,.command-center-v13 *{color:#fff}
.bank-hero h1,.bank-hero h2,.gov-hero h1,.gov-hero h2,.loan-hero h1,.loan-hero h2,.mx-hero h1,.mx-hero h2,
.qc-verify-hero h1,.qc-verify-hero h2,.verification-hero h1,.verification-hero h2,.po-audit-hero h1,.po-audit-hero h2,
.digital-card-hero h1,.digital-card-hero h2,.ccs-hero h1,.ccs-hero h2,.command-hero h1,.command-hero h2{font-size:19px}
.bank-hero p,.gov-hero p,.loan-hero p,.mx-hero p,.qc-verify-hero p,.verification-hero p,.po-audit-hero p{opacity:.88;font-size:12px;margin-top:4px}
.loan-eyebrow,.mx-overline,.command-kicker,.coop-kicker,.settings-kicker,.inventory-kicker{font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.09em;opacity:.85;margin-bottom:6px}
.bank-hero-stats,.mh-insight-hero,.po-audit-kpis,.qc-kpis,.mx-glance{display:flex;gap:10px;flex-wrap:wrap}

/* ---------- 7. KPI / STAT GRIDS ---------- */
.qc-kpis,.po-audit-kpis,.analytics-card,.mx-glance,.bank-hero-stats,.gov-cards,.neo-metrics,.shu-config-grid,.mx-loan-grid{
  display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin:14px 0
}
.qc-kpis article,.po-audit-kpis article,.analytics-card,.neo-metrics article,.gov-cards article{
  padding:16px;border:1px solid var(--border);border-radius:16px;background:var(--surface);box-shadow:var(--shadow-sm)
}
.qc-kpis span,.po-audit-kpis span{display:block;color:var(--muted);font-size:9px;font-weight:900;text-transform:uppercase;letter-spacing:.08em}
.qc-kpis b,.po-audit-kpis b{display:block;margin-top:6px;font-size:20px}
.qc-kpis .ok b,.po-audit-kpis .ok b{color:var(--brand)}
.qc-kpis .warn b,.po-audit-kpis .warn b{color:var(--danger)}

/* ---------- 8. GENERIC SUB-PAGE WRAPPERS (grid gap-y layout, kept consistent) ---------- */
.credit-page,.credit-detail-page,.credit-analysis-page,.credit-payment-page,.credit-products-page,.credit-verification-page,
.digital-card-page,.shu-adv-page,.qc-verify-page,.ccs-layout,.gov-layout,.bank-layout,.loan-apply-layout,
.po-audit-layout,.po-audit-page,.wallet-admin-suite,.inventory-workspace,.settings-shell,.analytics-studio,
.analytics-layout,.role-menu-editor,.verification-workspace,.command-actions,.neo-panel{
  display:grid;gap:16px
}
.bank-two,.mh-two,.ds-split,.gov-layout,.ccs-layout,.po-detail-top,.credit-account-options{
  display:grid;grid-template-columns:2fr 1fr;gap:16px
}
.bank-grid,.loan-product-grid,.card-grid,.role-menu-grid,.backup-grid,.wallet-account-grid,
.mx-service-grid,.qc-location-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px}

/* row-style lists (loan-item, po-row, vt-row, wallet-history-row, backup-row, restore-row, ccs-list rows) */
.loan-item,.po-row,.vt-row,.wallet-history-row,.backup-row,.restore-row,.rpt-row,.bank-check,
.mx-activity-row,.gov-record,.inventory-toolbar,.qc-product-row,.sim-row{
  display:flex;align-items:center;justify-content:space-between;gap:12px;padding:12px 14px;
  border:1px solid var(--border);border-radius:14px;background:var(--surface);margin-bottom:8px;flex-wrap:wrap
}
.loan-item-main,.po-order-main,.vt-member{min-width:0;flex:1}
.loan-number,.po-number,.vt-ref{font-family:monospace;font-size:11px;color:var(--muted)}
.loan-amount,.po-order-money,.vt-amount{font-weight:900;color:var(--ink);white-space:nowrap}
.loan-actions,.po-filter-actions,.top-actions,.bottom-actions,.vt-action{display:flex;gap:8px;flex-wrap:wrap}

/* empty states everywhere */
.loan-empty,.po-empty,.po-detail-empty,.bank-empty,.bank-empty-small,.gov-empty,.analytics-empty,
.ds-empty,.wallet-empty,.ccs-empty,.mh-empty,.mx-empty,.qc-location-empty,.account-unavailable{
  text-align:center;padding:32px 16px;border:1px dashed var(--border);border-radius:var(--radius-md);color:var(--muted);background:var(--brand-soft)
}

/* ---------- 9. FORMS: page-specific form cards ---------- */
.bank-form,.loan-form-card,.coop-form,.ds-critical-form,.qc-search-form,.qc-add-form,.wallet-limit-form,
.role-menu-form,.settings-form,.ccs-form{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius-md);padding:18px;box-shadow:var(--shadow-sm)}
.qc-search-form,.qc-add-form{display:flex;gap:8px;flex-wrap:wrap;padding:0;border:none;box-shadow:none;background:none}
.qc-search-form input,.qc-add-form input{flex:1;min-width:140px}
.bank-form-note{font-size:11px;color:var(--muted);margin-top:6px}

/* ---------- 10. LOGIN / COOP AUTH ---------- */
.login-wrap,.coop-login{min-height:100vh;display:grid;grid-template-columns:1.1fr 1fr;background:var(--bg)}
.login-left,.coop-visual{background:linear-gradient(150deg,var(--brand-dark),var(--brand));color:#fff;padding:48px;display:flex;flex-direction:column;justify-content:center;gap:20px}
.login-left *,.coop-visual *{color:#fff}
.login-right,.coop-auth{display:flex;align-items:center;justify-content:center;padding:24px}
.login-card,.coop-auth-card{width:100%;max-width:400px;background:var(--surface);border:1px solid var(--border);border-radius:var(--radius-lg);padding:32px;box-shadow:var(--shadow-lg)}
.login-brand,.coop-brand{display:flex;align-items:center;gap:10px;margin-bottom:18px}
.login-brand-icon,.coop-auth-mark{width:42px;height:42px;border-radius:12px;background:var(--brand-light);display:grid;place-items:center;font-size:20px}
.login-title,.coop-auth-kicker{font-size:20px;font-weight:900;color:var(--ink)}
.login-sub,.coop-auth-sub{color:var(--muted);font-size:12px;margin-bottom:18px}
.login-field,.coop-field{margin-bottom:14px}
.login-input-wrap{position:relative}
.login-input-icon{position:absolute;left:13px;top:50%;transform:translateY(-50%);color:var(--muted-2)}
.login-input-wrap input{padding-left:36px}
.login-btn,.coop-submit{width:100%;height:46px;border-radius:12px;background:var(--brand);color:#fff;border:none;font-weight:800;font-size:13px}
.login-btn:hover,.coop-submit:hover{background:var(--brand-dark)}
.login-footer,.coop-auth-footer{text-align:center;margin-top:16px;font-size:11.5px;color:var(--muted)}
.login-features,.coop-principles{display:flex;flex-direction:column;gap:14px}
.login-feat{display:flex;gap:10px;align-items:flex-start}
.login-feat-icon{font-size:20px}
.login-feat-title{font-weight:800;font-size:13px}
.login-feat-desc{font-size:11.5px;opacity:.85}
.coop-demo,.coop-access-help{font-size:11px;background:rgba(255,255,255,.12);border-radius:10px;padding:10px 12px;margin-top:10px}

/* ---------- 11. MEMBER CARD / DIGITAL CARD / QR ---------- */
.member-card,.digital-card-page .card,.kartu{max-width:380px;margin:0 auto;background:linear-gradient(135deg,var(--brand-dark),var(--brand));color:#fff;border-radius:20px;padding:22px;box-shadow:var(--shadow-lg)}
.member-card *{color:#fff}
.member-card-chip{width:38px;height:28px;border-radius:6px;background:rgba(255,255,255,.35);margin-bottom:16px}
.member-card-qr,.qr-inner{background:#fff;border-radius:12px;padding:10px;display:inline-block}
.member-card-detail{margin-top:12px;font-size:11px;opacity:.9}
.member-card-stage{display:flex;justify-content:center;padding:20px 0}
.digital-card-actions{display:flex;gap:8px;justify-content:center;margin-top:14px;flex-wrap:wrap}

/* ---------- 12. TABS / PILL NAV ---------- */
.mh-tabs,.gov-tabs,.ccs-tabs{display:flex;gap:6px;overflow-x:auto;padding-bottom:6px;-webkit-overflow-scrolling:touch}
.mh-tabs a,.mh-tabs button,.gov-tabs a,.gov-tabs button,.ccs-tabs a,.ccs-tabs button{white-space:nowrap;padding:8px 14px;border-radius:99px;border:1px solid var(--border);background:var(--surface);font-size:11.5px;font-weight:700;color:var(--text)}
.mh-tabs .active,.gov-tabs .active,.ccs-tabs .active{background:var(--brand);border-color:var(--brand);color:#fff}
.settings-anchor-nav{display:flex;gap:6px;flex-wrap:wrap}

/* ---------- 13. MISC SMALL COMPONENTS ---------- */
.zero-pay-warning{border:1px solid #f4d49b;background:var(--warning-bg);border-radius:var(--radius-md);padding:14px}
.zero-pay-warning h3{color:var(--warning)}
.payment-schedules>div{display:flex;justify-content:space-between;align-items:center;gap:10px;padding:11px 0;border-bottom:1px solid var(--border-soft)}
.payment-schedules i{font-style:normal;border-radius:99px;padding:5px 10px;font-size:10px;font-weight:800;background:var(--warning-bg);color:var(--warning)}
.payment-schedules i.paid{background:var(--success-bg);color:var(--brand)}
.payment-submit-shell{display:grid;grid-template-columns:1.5fr 1fr;gap:16px}
.payment-submit-result{padding:14px 18px;border:1px solid #b7e4d5;border-radius:16px;background:var(--success-bg);color:var(--brand-dark)}
.notif-bell-wrap{position:relative}
.notif-badge{background:var(--danger);color:#fff;border-radius:50%;min-width:18px;height:18px;font-size:10px;display:flex;align-items:center;justify-content:center}
.section-caption{font-size:10px;text-transform:uppercase;letter-spacing:.08em;color:var(--muted);font-weight:800;margin:14px 0 6px}
.card-orb,.orb-a,.orb-b{display:none}
.card-shine{display:none}
.flow-node,.node-a,.node-b,.node-c,.node-d,.node-center{padding:12px;border:1px solid var(--border);border-radius:14px;background:var(--surface);text-align:center;font-size:11px}
.flow-track,.neo-flow-track,.network-line{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.flow-arrow{color:var(--muted-2)}
.thread-field,.thread-core{padding:12px;border:1px solid var(--border);border-radius:12px;background:var(--surface)}

/* ---------- 13b. EXTRA DYNAMIC COMPONENTS ---------- */
.po-order-card{display:flex;flex-wrap:wrap;align-items:center;gap:10px;padding:14px;border:1px solid var(--border);border-radius:14px;background:var(--surface);margin-bottom:8px;transition:.15s}
.po-order-card:hover,.po-order-card.is-active{border-color:var(--brand);box-shadow:var(--shadow-sm)}
.qc-visible-message{position:sticky;top:10px;z-index:50;font-weight:800}
.analytics-delta{display:inline-block;padding:4px 9px;border-radius:99px;font-size:10px;font-weight:800;background:var(--border-soft)}
.analytics-delta.up{background:var(--success-bg);color:var(--brand-dark)}
.analytics-delta.down{background:var(--danger-bg);color:var(--danger)}
[class*=" status-"],[class^="status-"]{font-style:normal;font-size:10.5px;font-weight:800;color:var(--muted)}
.neg{color:var(--danger)!important}.pos{color:var(--brand)!important}

/* ---------- 14. TOAST / DIALOG ---------- */
.ds-toast-region{position:fixed;top:16px;right:16px;z-index:1200;display:flex;flex-direction:column;gap:8px}
.ds-dialog-layer{position:fixed;inset:0;z-index:1300;display:grid;place-items:center;padding:16px}
.ds-dialog-backdrop{position:absolute;inset:0;background:rgba(16,35,26,.5)}
.ds-dialog{position:relative;background:var(--surface);border-radius:var(--radius-md);padding:22px;max-width:360px;width:100%;box-shadow:var(--shadow-lg);text-align:center}
.ds-dialog-icon{width:42px;height:42px;border-radius:50%;background:var(--warning-bg);color:var(--warning);display:grid;place-items:center;margin:0 auto 12px;font-weight:900}
.ds-dialog-actions{display:flex;gap:8px;justify-content:center;margin-top:16px}

/* ==========================================================================
   15. MOBILE (<=900px) — konsolidasi total, satu blok, mobile-first shell
   ========================================================================== */
@media(max-width:900px){
  .app-wrap{display:block}
  .sidebar,body.sidebar-compact .sidebar{position:fixed;top:0;left:0;bottom:0;width:82vw!important;max-width:320px!important;
    flex-basis:auto!important;transform:translateX(-100%);transition:transform .22s ease;z-index:200;box-shadow:var(--shadow-lg);padding-top:var(--safe-t)}
  .sidebar.open{transform:translateX(0)}
  .sidebar-collapse-btn{display:none}
  .sidebar-section-header,.sidebar-nav a,.sidebar-footer-link{justify-content:flex-start!important}
  .brand-text,.user-info,.sidebar-search,.nav-label,.section-name,.section-count,.section-chevron{display:initial!important}
  .mobile-drawer-close{display:flex;align-items:center;justify-content:center;position:absolute;right:12px;top:12px;
    width:32px;height:32px;border-radius:50%;border:none;background:var(--border-soft);font-size:18px;z-index:2}
  .mobile-drawer-backdrop{position:fixed;inset:0;background:rgba(16,35,26,.5);z-index:190}
  .mobile-drawer-backdrop.open,body.mobile-drawer-open .mobile-drawer-backdrop{display:block}
  body.mobile-drawer-open{overflow:hidden}
  body.mobile-drawer-open .sidebar{transform:translateX(0)}

  .hamburger.desktop-menu-btn{display:none!important}

  /* Top mobile header */
  .mobile-header{display:flex;align-items:center;justify-content:space-between;gap:8px;
    position:fixed;top:0;left:0;right:0;height:calc(var(--header-h) + var(--safe-t));padding:var(--safe-t) 12px 0;
    background:var(--surface);border-bottom:1px solid var(--border);z-index:150;box-shadow:var(--shadow-sm)}
  .mh-hamburger{width:38px;height:38px;flex:none;border:none;background:var(--brand-soft);border-radius:12px;
    display:flex;flex-direction:column;align-items:center;justify-content:center;gap:4px}
  .mh-hamburger span{width:16px;height:2px;background:var(--brand-dark);border-radius:2px}
  .mh-center{flex:1;min-width:0;text-align:center}
  .mh-kicker{font-size:9px;font-weight:800;color:var(--muted);letter-spacing:.08em}
  .mh-title{font-size:13.5px;font-weight:800;color:var(--ink);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  .mh-notif{position:relative;width:38px;height:38px;flex:none;border:none;background:var(--brand-soft);border-radius:12px;display:grid;place-items:center}
  .mh-bell{width:16px;height:16px;display:block}
  .mh-bell::before{content:'\\1F514';font-size:15px}
  .mh-badge{position:absolute;top:2px;right:2px;background:var(--danger);color:#fff;border-radius:50%;
    min-width:16px;height:16px;font-size:9px;display:flex;align-items:center;justify-content:center;font-weight:800}

  /* Bottom mobile nav */
  .mobile-nav{display:flex;align-items:stretch;justify-content:space-around;
    position:fixed;left:0;right:0;bottom:0;height:calc(var(--nav-h) + var(--safe-b));padding-bottom:var(--safe-b);
    background:var(--surface);border-top:1px solid var(--border);z-index:150;box-shadow:0 -6px 20px rgba(16,45,37,.08)}
  .mn-item{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:2px;
    border:none;background:none;color:var(--muted);font-size:9.5px;font-weight:700;position:relative;min-width:0;padding:6px 2px}
  .mn-item.active{color:var(--brand)}
  .mn-icon{font-size:18px;line-height:1}
  .mobile-nav-badge{position:absolute;top:2px;right:14%;background:var(--danger);color:#fff;border-radius:50%;
    min-width:15px;height:15px;font-size:8.5px;display:flex;align-items:center;justify-content:center}

  /* Mobile app-drawer / notif sheets */
  .mobile-app-backdrop,.mobile-notif-backdrop{position:fixed;inset:0;background:rgba(16,35,26,.5);z-index:210}
  .mobile-app-backdrop.open,.mobile-notif-backdrop.open{display:block}
  .mobile-app-sheet,.mobile-notif-sheet{position:fixed;left:0;right:0;bottom:0;max-height:82vh;background:var(--surface);
    border-radius:22px 22px 0 0;z-index:220;transform:translateY(100%);transition:transform .22s ease;
    display:flex;flex-direction:column;padding-bottom:var(--safe-b);box-shadow:0 -10px 40px rgba(16,45,37,.2)}
  .mobile-app-sheet.open,.mobile-notif-sheet.open{transform:translateY(0)}
  .mas-handle,.mns-handle{width:40px;height:4px;background:var(--border);border-radius:99px;margin:10px auto}
  .mobile-app-sheet header,.mobile-notif-sheet header{display:flex;align-items:center;justify-content:space-between;padding:4px 18px 10px;border-bottom:1px solid var(--border-soft)}
  .mobile-app-sheet header span,.mobile-notif-sheet header span{font-size:9px;font-weight:800;color:var(--muted);text-transform:uppercase;letter-spacing:.08em}
  .mobile-app-sheet header h3,.mobile-notif-sheet header h3{font-size:15px}
  .mobile-app-sheet header button,.mobile-notif-sheet header button{width:30px;height:30px;border-radius:50%;border:none;background:var(--border-soft);font-size:16px}
  .mobile-app-search{padding:12px 16px 6px}
  .mobile-app-search input{border-radius:99px;background:var(--bg)}
  .mobile-app-grid-wrap{overflow-y:auto;padding:4px 14px calc(16px + var(--safe-b))}
  .mobile-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}
  .mg-item{display:flex;flex-direction:column;gap:8px;padding:14px;border:1px solid var(--border);border-radius:18px;background:var(--brand-soft);text-align:left}
  .mg-icon{width:38px;height:38px;border-radius:12px;background:var(--surface);display:grid;place-items:center;font-size:18px}
  .mg-label{font-size:10.5px;font-weight:800;color:var(--ink);line-height:1.25}
  .mobile-notif-list{overflow-y:auto;padding:6px 16px;flex:1}
  .mobile-notif-empty{text-align:center;color:var(--muted);padding:24px;font-size:12px}
  .mobile-app-sheet footer,.mobile-notif-sheet footer{display:flex;gap:8px;padding:12px 16px;border-top:1px solid var(--border-soft)}
  .mobile-app-sheet footer .btn,.mobile-notif-sheet footer .btn{flex:1}

  .mobile-user-card{display:flex!important;align-items:center;justify-content:space-between;
    background:var(--surface);border:1px solid var(--border);border-radius:16px;padding:12px 14px;margin:calc(var(--header-h) + var(--safe-t) + 12px) 14px 0}
  .muc-name{font-weight:800;font-size:13px}
  .muc-role{font-size:10.5px;color:var(--muted);text-transform:capitalize}
  .muc-avatar{font-size:22px}

  /* Main content spacing to clear fixed header/nav */
  .main-content{padding-top:0}
  main.main-content>.topbar{margin-top:calc(var(--header-h) + var(--safe-t) + 10px)}
  main.main-content>.mobile-user-card ~ .topbar,.mobile-user-card + .mobile-grid-wrap + .topbar{margin-top:12px}
  .topbar{padding:0 14px 4px}
  main.main-content>*:not(.topbar):not(.mobile-user-card):not(.mobile-grid-wrap){padding-left:14px;padding-right:14px}
  main.main-content{padding-bottom:calc(var(--nav-h) + var(--safe-b) + 20px)}
  .page-header h1{font-size:17px}
  .page-header p{font-size:11px}

  /* Grids collapse to single column */
  .grid,.card-grid,.metrics,.ds-form-grid,.settings-form-grid,.role-menu-grid,.bank-grid,.loan-product-grid,
  .backup-grid,.wallet-account-grid,.mx-service-grid,.qc-location-grid,.bank-two,.mh-two,.ds-split,
  .gov-layout,.ccs-layout,.po-detail-top,.credit-account-options,.payment-submit-shell,.qc-kpis,
  .po-audit-kpis,.analytics-card,.mx-glance,.bank-hero-stats,.gov-cards,.neo-metrics,.shu-config-grid,.mx-loan-grid{
    grid-template-columns:1fr!important
  }
  .col-4,.col-5,.col-6,.col-7,.col-8,.col-12{grid-column:1/-1}

  /* Hero blocks stack and shrink */
  .bank-hero,.gov-hero,.loan-hero,.mx-hero,.qc-verify-hero,.verification-hero,.verification-v2-hero,
  .po-audit-hero,.digital-card-hero,.ccs-hero,.command-hero,.neo-command-hero,.mh-insight-hero,
  .wallet-risk-hero,.command-center-v13{flex-direction:column;align-items:flex-start;padding:18px;border-radius:18px}

  /* Row-style lists stack their contents */
  .loan-item,.po-row,.vt-row,.wallet-history-row,.backup-row,.restore-row,.rpt-row,.bank-check,
  .mx-activity-row,.gov-record,.inventory-toolbar,.qc-product-row,.sim-row{flex-direction:column;align-items:stretch}
  .loan-actions,.po-filter-actions,.top-actions,.bottom-actions,.vt-action{width:100%}
  .loan-actions .btn,.top-actions .btn,.bottom-actions .btn{flex:1}

  /* Login/auth stacks */
  .login-wrap,.coop-login{grid-template-columns:1fr}
  .login-left,.coop-visual{display:none}
  .login-card,.coop-auth-card{box-shadow:none;border:none;padding:20px 16px}

  /* Tables: allow horizontal scroll, never overflow viewport */
  .table-wrap{max-width:100vw;border-radius:14px}
  .card,.ds-card,.metric{border-radius:16px}

  /* Forms: 16px font stops iOS auto-zoom, full width everything */
  input,select,textarea{font-size:16px!important;min-height:44px}
  .btn,.ds-btn{min-height:44px;width:100%}
  .topbar-actions .btn,.top-actions .btn{width:auto}
  .notif-bell-wrap{display:none} /* replaced by mobile header bell */
  .desktop-menu-btn{display:none}
  input[type=file]{font-size:12px!important;min-height:auto}

  .flash{margin:12px 14px 0}
}

@media(max-width:420px){
  .mn-item{font-size:8.5px}
  .mg-label{font-size:9.5px}
  .mobile-grid{grid-template-columns:repeat(2,1fr)}
  .mh-title{font-size:12.5px}
}

@media(min-width:901px){
  .mobile-drawer-backdrop,.mobile-app-backdrop,.mobile-notif-backdrop{display:none!important}
}


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
