import os
import json
import hashlib
import uuid
from datetime import datetime, timedelta
from flask import Flask, request, jsonify, send_from_directory, g
from flask_cors import CORS
from functools import wraps
import openpyxl
import psycopg2
import psycopg2.extras

app = Flask(__name__, static_folder='static')
CORS(app)

DATABASE_URL = os.environ.get('DATABASE_URL', '')
CLOUDINARY_CLOUD_NAME = os.environ.get('CLOUDINARY_CLOUD_NAME', '')
CLOUDINARY_API_KEY = os.environ.get('CLOUDINARY_API_KEY', '')
CLOUDINARY_API_SECRET = os.environ.get('CLOUDINARY_API_SECRET', '')

# ERP database (erp_sourcing) — synced by BI/Airflow every hour
ERP_DATABASE_URL = os.environ.get('ERP_DATABASE_URL', '')
if not ERP_DATABASE_URL and DATABASE_URL:
    ERP_DATABASE_URL = DATABASE_URL.rsplit('/', 1)[0] + '/erp_sourcing'

TEAMS = ['CX', 'Sales/KAM', 'Sales Co', 'Merchandise', 'Inbound', 'Outbound']
FAULT_TEAMS = TEAMS + ['Customer']
CASE_TYPES = ['Complain', 'Claim', 'Update Invoice', 'วางบิล']
CLAIM_SUBTYPES = ['ด่วน (ภายในวัน)', 'รอรอบถัดไป (ไม่รู้วัน)', 'รอรอบถัดไป (รู้วันแล้ว)']
ROOT_CAUSES = ['สินค้าตกหล่น', 'คุณภาพไม่ผ่าน/ไม่ได้ spec', 'น้ำหนักไม่ครบ', 'ส่งผิด SKU', 'Master SKU ผิด', 'เอกสารผิดพลาด', 'อื่นๆ']
PRIORITIES = ['High', 'Medium']

# ---------------------------------------------------------------------------
# DB helpers
# ---------------------------------------------------------------------------

def get_db():
    conn = psycopg2.connect(DATABASE_URL, options="-c timezone=Asia/Bangkok")
    return conn

def query(sql, params=(), one=False):
    sql = sql.replace('?', '%s')
    conn = get_db()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    cur.execute(sql, params)
    rows = cur.fetchall()
    conn.close()
    result = [dict(r) for r in rows]
    return result[0] if one and result else (None if one else result)

def erp_query(sql, params=(), one=False):
    """Query the ERP database (erp_sourcing.sourcing_erp_order_items)."""
    sql = sql.replace('?', '%s')
    conn = psycopg2.connect(ERP_DATABASE_URL, options="-c timezone=Asia/Bangkok")
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    cur.execute(sql, params)
    rows = cur.fetchall()
    conn.close()
    result = [dict(r) for r in rows]
    return result[0] if one and result else (None if one else result)

def mutate(sql, params=()):
    sql = sql.replace('?', '%s')
    conn = get_db()
    cur = conn.cursor()
    cur.execute(sql, params)
    last_id = None
    if cur.description:
        row = cur.fetchone()
        if row:
            last_id = row[0]
    conn.commit()
    conn.close()
    return last_id

# ---------------------------------------------------------------------------
# One-time team name migration (old → new naming convention)
# ---------------------------------------------------------------------------

def migrate_team_names():
    """Rename old team names to new ones across all relevant tables/columns.
    Each statement runs in its own transaction so a missing table won't abort the rest."""
    renames = [
        ('Sales',             'Sales/KAM'),
        ('KAM',               'Sales/KAM'),
        ('Inbound/QC',        'Inbound'),
        ('Outbound/Logistics', 'Outbound'),
        # Management ถูกลบออกจากระบบ → ใช้ CX เป็น fallback
        ('Management',        'CX'),
    ]
    # แก้ค่า status ที่หลุดเข้าไปใน current_team โดยผิดพลาด
    status_cleanups = [
        "UPDATE tickets SET current_team = opener_team WHERE current_team IN ('pending_ack','pending_fault','open','in_progress','closed')",
    ]
    for sql in status_cleanups:
        try:
            conn = get_db()
            cur = conn.cursor()
            cur.execute(sql)
            rows = cur.rowcount
            conn.commit()
            conn.close()
            if rows:
                print(f"[migrate] cleanup status-in-team: {rows} rows fixed")
        except Exception as e:
            print(f"[migrate] cleanup error: {e}")
    stmts = [
        ("tickets",                  "fault_team"),
        ("tickets",                  "current_team"),
        ("tickets",                  "opener_team"),
        ("ticket_fault_attribution", "fault_team"),
        ("ticket_workflow_log",      "from_team"),
        ("ticket_workflow_log",      "to_team"),
        ("ticket_assignments",       "team"),
        ("users",                    "team"),
        ("employees",                "team"),
    ]
    total = 0
    for old, new in renames:
        for table, col in stmts:
            try:
                conn = get_db()
                cur = conn.cursor()
                cur.execute(f"UPDATE {table} SET {col}=%s WHERE {col}=%s", (new, old))
                rows = cur.rowcount
                conn.commit()
                conn.close()
                if rows:
                    print(f"[migrate] {table}.{col}: '{old}'→'{new}' ({rows} rows)")
                    total += rows
            except Exception as e:
                print(f"[migrate] skip {table}.{col}: {e}")
    print(f"[migrate_team_names] done — {total} rows updated")

migrate_team_names()

def migrate_orders_delivery_date():
    """Add delivery_date column to orders if it doesn't exist yet."""
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("""
            ALTER TABLE orders ADD COLUMN IF NOT EXISTS delivery_date TEXT
        """)
        conn.commit()
        conn.close()
        print("[migrate] orders.delivery_date column ensured")
    except Exception as e:
        print(f"[migrate] orders.delivery_date: {e}")

migrate_orders_delivery_date()

def migrate_orders_reset_and_fix():
    """One-time: clear orders + add UNIQUE index on erp_item_id.
    Skips if the unique index already exists (already ran)."""
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("SELECT 1 FROM pg_indexes WHERE indexname = 'orders_erp_item_id_uidx'")
        if cur.fetchone():
            conn.close()
            return  # already done
        # Index missing → old table without constraint → clear + fix
        cur.execute("TRUNCATE TABLE orders RESTART IDENTITY")
        cur.execute("CREATE UNIQUE INDEX orders_erp_item_id_uidx ON orders (erp_item_id)")
        conn.commit()
        conn.close()
        print("[migrate] orders cleared + unique index created")
    except Exception as e:
        print(f"[migrate] orders reset: {e}")

migrate_orders_reset_and_fix()

def migrate_accounts_outlets_unique():
    """Add missing UNIQUE indexes on accounts.name, outlets.erp_outlet_id, outlets(account_id,name)."""
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("""
            CREATE UNIQUE INDEX IF NOT EXISTS accounts_name_uidx ON accounts (name)
        """)
        cur.execute("""
            CREATE UNIQUE INDEX IF NOT EXISTS outlets_erp_outlet_id_uidx ON outlets (erp_outlet_id)
            WHERE erp_outlet_id IS NOT NULL
        """)
        cur.execute("""
            CREATE UNIQUE INDEX IF NOT EXISTS outlets_account_name_uidx ON outlets (account_id, name)
        """)
        conn.commit()
        conn.close()
        print("[migrate] accounts/outlets unique indexes ensured")
    except Exception as e:
        print(f"[migrate] accounts/outlets unique: {e}")

migrate_accounts_outlets_unique()

# ---------------------------------------------------------------------------
# Auth helpers
# ---------------------------------------------------------------------------

def hash_password(p):
    return hashlib.sha256(f"smm-crm-salt-2024{p}".encode()).hexdigest()

def _get_user_from_token():
    auth = request.headers.get('Authorization', '')
    token = auth.replace('Bearer ', '').strip()
    if not token:
        return None
    row = query("""
        SELECT u.id AS user_id, u.username, u.display_name, u.team, u.role
        FROM user_sessions s JOIN users u ON u.id = s.user_id
        WHERE s.token = %s AND u.status = 'active'
    """, (token,), one=True)
    return row

def require_auth(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        user = _get_user_from_token()
        if not user:
            return jsonify({'error': 'Unauthorized'}), 401
        g.user = user
        return f(*args, **kwargs)
    return decorated

def require_admin(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        user = _get_user_from_token()
        if not user:
            return jsonify({'error': 'Unauthorized'}), 401
        g.user = user
        if g.user['role'] != 'admin':
            return jsonify({'error': 'Admin only'}), 403
        return f(*args, **kwargs)
    return decorated

# ---------------------------------------------------------------------------
# DB init
# ---------------------------------------------------------------------------

SCHEMA_BASE = """
CREATE TABLE IF NOT EXISTS accounts (
    id SERIAL PRIMARY KEY,
    name TEXT NOT NULL UNIQUE,
    owner TEXT,
    status TEXT DEFAULT 'active',
    created_at TEXT DEFAULT (to_char(now(), 'YYYY-MM-DD"T"HH24:MI:SS'))
);

CREATE TABLE IF NOT EXISTS outlets (
    id SERIAL PRIMARY KEY,
    account_id INTEGER NOT NULL,
    name TEXT NOT NULL,
    erp_customer_id TEXT,
    erp_outlet_id TEXT UNIQUE,
    csc_code TEXT,
    status TEXT DEFAULT 'active',
    FOREIGN KEY (account_id) REFERENCES accounts(id)
);

CREATE TABLE IF NOT EXISTS account_notes (
    id SERIAL PRIMARY KEY,
    account_id INTEGER NOT NULL,
    note TEXT,
    created_by TEXT,
    created_at TEXT DEFAULT (to_char(now(), 'YYYY-MM-DD"T"HH24:MI:SS')),
    FOREIGN KEY (account_id) REFERENCES accounts(id)
);

CREATE TABLE IF NOT EXISTS suppliers (
    id SERIAL PRIMARY KEY,
    name TEXT NOT NULL,
    supplier_type TEXT,
    contact_person TEXT,
    phone TEXT,
    line_id TEXT,
    email TEXT,
    payment_method TEXT,
    credit_days INTEGER,
    rating INTEGER DEFAULT 0,
    status TEXT DEFAULT 'active',
    created_at TEXT DEFAULT (to_char(now(), 'YYYY-MM-DD"T"HH24:MI:SS'))
);

CREATE TABLE IF NOT EXISTS supplier_notes (
    id SERIAL PRIMARY KEY,
    supplier_id INTEGER NOT NULL,
    note TEXT,
    note_type TEXT,
    created_by TEXT,
    created_at TEXT DEFAULT (to_char(now(), 'YYYY-MM-DD"T"HH24:MI:SS')),
    FOREIGN KEY (supplier_id) REFERENCES suppliers(id)
);

CREATE TABLE IF NOT EXISTS employees (
    id SERIAL PRIMARY KEY,
    name TEXT NOT NULL,
    employee_code TEXT,
    department TEXT,
    team TEXT,
    role TEXT,
    phone TEXT,
    email TEXT,
    status TEXT DEFAULT 'active',
    created_at TEXT DEFAULT (to_char(now(), 'YYYY-MM-DD"T"HH24:MI:SS'))
);

CREATE TABLE IF NOT EXISTS orders (
    id SERIAL PRIMARY KEY,
    erp_item_id TEXT UNIQUE,
    order_id TEXT,
    invoice_number TEXT,
    doc_no TEXT,
    doc_date TEXT,
    outlet_id INTEGER,
    sku_code TEXT,
    product_name TEXT,
    qty REAL,
    unit TEXT,
    total_sales REAL,
    vat_price REAL,
    is_vat INTEGER,
    sku_group TEXT,
    sku_category TEXT,
    sku_type TEXT,
    delivery_date TEXT,
    delivery_started_at TEXT,
    delivery_finished_at TEXT,
    loaded_at TEXT,
    FOREIGN KEY (outlet_id) REFERENCES outlets(id)
);

CREATE TABLE IF NOT EXISTS users (
    id SERIAL PRIMARY KEY,
    username TEXT NOT NULL UNIQUE,
    password_hash TEXT NOT NULL,
    display_name TEXT NOT NULL,
    team TEXT NOT NULL,
    role TEXT DEFAULT 'staff',
    status TEXT DEFAULT 'active',
    created_at TEXT DEFAULT (to_char(now(), 'YYYY-MM-DD"T"HH24:MI:SS'))
);

CREATE TABLE IF NOT EXISTS user_sessions (
    id SERIAL PRIMARY KEY,
    user_id INTEGER NOT NULL,
    token TEXT NOT NULL UNIQUE,
    created_at TEXT DEFAULT (to_char(now(), 'YYYY-MM-DD"T"HH24:MI:SS')),
    FOREIGN KEY (user_id) REFERENCES users(id)
);

CREATE TABLE IF NOT EXISTS tickets (
    id SERIAL PRIMARY KEY,
    ticket_no TEXT UNIQUE,
    outlet_id INTEGER,
    invoice_number TEXT,
    sku_code TEXT,
    product_name TEXT,
    case_type TEXT,
    case_subtype TEXT,
    root_cause TEXT,
    priority TEXT DEFAULT 'Medium',
    status TEXT DEFAULT 'open',
    current_team TEXT DEFAULT 'CX',
    opener_team TEXT DEFAULT 'CX',
    opener_user_id INTEGER,
    fault_team TEXT,
    description TEXT,
    created_by TEXT,
    created_at TEXT DEFAULT (to_char(now(), 'YYYY-MM-DD"T"HH24:MI:SS')),
    closed_at TEXT,
    fault_attributed_at TEXT,
    FOREIGN KEY (outlet_id) REFERENCES outlets(id)
);

CREATE TABLE IF NOT EXISTS ticket_assignments (
    id SERIAL PRIMARY KEY,
    ticket_id INTEGER NOT NULL,
    team TEXT NOT NULL,
    note TEXT,
    employee_id INTEGER,
    acknowledged_by TEXT,
    acknowledged_user_id INTEGER,
    acknowledged_at TEXT,
    created_at TEXT DEFAULT (to_char(now(), 'YYYY-MM-DD"T"HH24:MI:SS')),
    FOREIGN KEY (ticket_id) REFERENCES tickets(id),
    FOREIGN KEY (employee_id) REFERENCES employees(id)
);

CREATE TABLE IF NOT EXISTS ticket_workflow_log (
    id SERIAL PRIMARY KEY,
    ticket_id INTEGER NOT NULL,
    from_team TEXT,
    to_team TEXT,
    action TEXT,
    note TEXT,
    image_url TEXT,
    user_id INTEGER,
    created_by TEXT,
    created_at TEXT DEFAULT (to_char(now(), 'YYYY-MM-DD"T"HH24:MI:SS')),
    FOREIGN KEY (ticket_id) REFERENCES tickets(id)
);

CREATE TABLE IF NOT EXISTS case_reads (
    user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    ticket_id INTEGER NOT NULL REFERENCES tickets(id) ON DELETE CASCADE,
    read_at TIMESTAMP NOT NULL DEFAULT NOW(),
    PRIMARY KEY (user_id, ticket_id)
);

CREATE TABLE IF NOT EXISTS ticket_fault_attribution (
    id SERIAL PRIMARY KEY,
    ticket_id INTEGER NOT NULL,
    fault_team TEXT NOT NULL,
    employee_id INTEGER,
    note TEXT,
    attributed_by TEXT,
    attributed_user_id INTEGER,
    created_at TEXT DEFAULT (to_char(now(), 'YYYY-MM-DD"T"HH24:MI:SS')),
    FOREIGN KEY (ticket_id) REFERENCES tickets(id),
    FOREIGN KEY (employee_id) REFERENCES employees(id)
);

CREATE TABLE IF NOT EXISTS ticket_comments (
    id SERIAL PRIMARY KEY,
    ticket_id INTEGER NOT NULL,
    comment TEXT,
    created_by TEXT,
    created_at TEXT DEFAULT (to_char(now(), 'YYYY-MM-DD"T"HH24:MI:SS')),
    FOREIGN KEY (ticket_id) REFERENCES tickets(id)
);

CREATE TABLE IF NOT EXISTS leads (
    id SERIAL PRIMARY KEY,
    lead_no TEXT UNIQUE,
    company_name TEXT NOT NULL,
    contact_name TEXT,
    contact_phone TEXT,
    contact_email TEXT,
    stage TEXT DEFAULT 'Cold Call/Email',
    owner_user_id INTEGER,
    description TEXT,
    created_by TEXT,
    created_at TEXT DEFAULT (to_char(now(), 'YYYY-MM-DD"T"HH24:MI:SS')),
    updated_at TEXT DEFAULT (to_char(now(), 'YYYY-MM-DD"T"HH24:MI:SS')),
    converted_at TEXT,
    converted_account_id INTEGER,
    FOREIGN KEY (owner_user_id) REFERENCES users(id),
    FOREIGN KEY (converted_account_id) REFERENCES accounts(id)
);

CREATE TABLE IF NOT EXISTS lead_activities (
    id SERIAL PRIMARY KEY,
    lead_id INTEGER NOT NULL,
    activity_type TEXT NOT NULL,
    subject TEXT,
    description TEXT,
    created_by TEXT,
    created_at TEXT DEFAULT (to_char(now(), 'YYYY-MM-DD"T"HH24:MI:SS')),
    FOREIGN KEY (lead_id) REFERENCES leads(id)
);
"""

def init_db():
    conn = get_db()
    cur = conn.cursor()

    # Create all tables — commit immediately so nothing below can roll these back
    for stmt in SCHEMA_BASE.strip().split(';'):
        stmt = stmt.strip()
        if stmt:
            cur.execute(stmt)

    conn.commit()  # commit table creation before anything that might rollback

    # Outlets unique constraint (in its own transaction)
    try:
        cur.execute("""
            DO $$
            BEGIN
                IF NOT EXISTS (
                    SELECT 1 FROM pg_constraint WHERE conname = 'outlets_account_id_name_key'
                ) THEN
                    ALTER TABLE outlets ADD CONSTRAINT outlets_account_id_name_key UNIQUE (account_id, name);
                END IF;
            END$$;
        """)
        conn.commit()
    except Exception as e:
        conn.rollback()
        print(f"[init_db] outlets constraint: {e}")

    # Migrations — each runs in its own transaction
    migrations = [
        "ALTER TABLE tickets ADD COLUMN IF NOT EXISTS case_subtype TEXT",
        "ALTER TABLE tickets ADD COLUMN IF NOT EXISTS root_cause TEXT",
        "ALTER TABLE tickets ADD COLUMN IF NOT EXISTS opener_team TEXT",
        "ALTER TABLE tickets ADD COLUMN IF NOT EXISTS opener_user_id INTEGER",
        "ALTER TABLE tickets ADD COLUMN IF NOT EXISTS closed_at TEXT",
        "ALTER TABLE tickets ADD COLUMN IF NOT EXISTS fault_attributed_at TEXT",
        "ALTER TABLE ticket_workflow_log ADD COLUMN IF NOT EXISTS image_url TEXT",
        "ALTER TABLE ticket_workflow_log ADD COLUMN IF NOT EXISTS user_id INTEGER",
        "CREATE UNIQUE INDEX IF NOT EXISTS outlets_erp_outlet_id_idx ON outlets(erp_outlet_id) WHERE erp_outlet_id IS NOT NULL AND erp_outlet_id != ''",
        "UPDATE tickets SET status='pending_ack' WHERE case_type='Complain' AND current_team='pending_ack' AND status='open'",
        "ALTER TABLE tickets ADD COLUMN IF NOT EXISTS claim_items TEXT",
        "ALTER TABLE tickets ADD COLUMN IF NOT EXISTS resolution_type TEXT",
        "ALTER TABLE ticket_comments ADD COLUMN IF NOT EXISTS image_url TEXT",
        "CREATE SEQUENCE IF NOT EXISTS leads_lead_no_seq START 1",
        "ALTER TABLE leads ADD COLUMN IF NOT EXISTS lost_from_stage TEXT",
        "ALTER TABLE leads ADD COLUMN IF NOT EXISTS contact_position TEXT",
        "ALTER TABLE leads ADD COLUMN IF NOT EXISTS branch TEXT",
    ]
    for m in migrations:
        try:
            cur.execute(m)
            conn.commit()
        except Exception:
            conn.rollback()

    # Default admin user
    cur.execute("SELECT COUNT(*) FROM users")
    cnt = cur.fetchone()[0]
    if cnt == 0:
        cur.execute(
            "INSERT INTO users (username, password_hash, display_name, team, role) VALUES (%s,%s,%s,%s,%s)",
            ('admin', hash_password('admin123'), 'Admin', 'CX', 'admin')
        )
        print("Created default admin user: admin / admin123")

    conn.commit()
    conn.close()
    print("Database initialized.")

# ---------------------------------------------------------------------------
# Static
# ---------------------------------------------------------------------------

@app.route('/')
def index():
    return send_from_directory('static', 'index.html')

# ---------------------------------------------------------------------------
# Config (public)
# ---------------------------------------------------------------------------

@app.route('/api/config', methods=['GET'])
def get_config():
    return jsonify({
        'teams': TEAMS,
        'fault_teams': FAULT_TEAMS,
        'case_types': CASE_TYPES,
        'claim_subtypes': CLAIM_SUBTYPES,
        'root_causes': ROOT_CAUSES,
        'priorities': PRIORITIES,
    })

@app.route('/api/workflows', methods=['GET'])
def get_workflows():
    return jsonify({})

# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------

@app.route('/api/auth/login', methods=['POST'])
def auth_login():
    d = request.json
    user = query(
        "SELECT * FROM users WHERE username=%s AND status='active'",
        (d.get('username', ''),), one=True
    )
    if not user or user['password_hash'] != hash_password(d.get('password', '')):
        return jsonify({'error': 'Username หรือ Password ไม่ถูกต้อง'}), 401
    token = str(uuid.uuid4())
    mutate("INSERT INTO user_sessions (user_id, token) VALUES (%s,%s) RETURNING id",
           (user['id'], token))
    return jsonify({
        'token': token,
        'user': {
            'id': user['id'],
            'username': user['username'],
            'display_name': user['display_name'],
            'team': user['team'],
            'role': user['role'],
        }
    })

@app.route('/api/auth/logout', methods=['POST'])
@require_auth
def auth_logout():
    auth = request.headers.get('Authorization', '')
    token = auth.replace('Bearer ', '').strip()
    mutate("DELETE FROM user_sessions WHERE token=%s", (token,))
    return jsonify({'ok': True})

@app.route('/api/auth/me', methods=['GET'])
@require_auth
def auth_me():
    return jsonify(g.user)

# ---------------------------------------------------------------------------
# User management (admin)
# ---------------------------------------------------------------------------

@app.route('/api/users', methods=['GET'])
@require_admin
def get_users():
    has_leads = request.args.get('has_leads', '0') == '1'
    if has_leads:
        rows = query("""
            SELECT DISTINCT u.id, u.username, u.display_name, u.team, u.role, u.status
            FROM users u
            INNER JOIN leads l ON l.owner_user_id = u.id
            ORDER BY u.display_name
        """)
    else:
        rows = query("SELECT id, username, display_name, team, role, status, created_at FROM users ORDER BY team, display_name")
    return jsonify(rows)

@app.route('/api/users', methods=['POST'])
@require_admin
def create_user():
    d = request.json
    if not d.get('username') or not d.get('password') or not d.get('display_name'):
        return jsonify({'error': 'username, password และ display_name จำเป็น'}), 400
    existing = query("SELECT id FROM users WHERE username=%s", (d['username'],), one=True)
    if existing:
        return jsonify({'error': 'Username นี้มีอยู่แล้ว'}), 400
    id_ = mutate(
        "INSERT INTO users (username, password_hash, display_name, team, role, status) VALUES (%s,%s,%s,%s,%s,%s) RETURNING id",
        (d['username'], hash_password(d['password']), d['display_name'],
         d.get('team', 'CX'), d.get('role', 'staff'), d.get('status', 'active'))
    )
    return jsonify({'id': id_}), 201

@app.route('/api/users/<int:uid>', methods=['PUT'])
@require_admin
def update_user(uid):
    d = request.json
    mutate("UPDATE users SET display_name=%s, team=%s, role=%s, status=%s WHERE id=%s",
           (d['display_name'], d['team'], d.get('role', 'staff'), d.get('status', 'active'), uid))
    return jsonify({'ok': True})

@app.route('/api/users/<int:uid>', methods=['DELETE'])
@require_admin
def delete_user(uid):
    if uid == g.user.get('id') or uid == g.user.get('user_id'):
        return jsonify({'error': 'ไม่สามารถลบตัวเองได้'}), 400
    conn = get_db()
    cur = conn.cursor()
    try:
        # Clear FK references before deleting
        cur.execute("DELETE FROM user_sessions WHERE user_id=%s", (uid,))
        cur.execute("UPDATE tickets SET opener_user_id=NULL WHERE opener_user_id=%s", (uid,))
        cur.execute("UPDATE ticket_workflow_log SET user_id=NULL WHERE user_id=%s", (uid,))
        cur.execute("UPDATE ticket_assignments SET acknowledged_user_id=NULL WHERE acknowledged_user_id=%s", (uid,))
        cur.execute("UPDATE ticket_fault_attribution SET attributed_user_id=NULL WHERE attributed_user_id=%s", (uid,))
        cur.execute("DELETE FROM users WHERE id=%s", (uid,))
        conn.commit()
        conn.close()
        return jsonify({'ok': True})
    except Exception as e:
        conn.rollback(); conn.close()
        return jsonify({'error': str(e)}), 500

@app.route('/api/users/<int:uid>/password', methods=['PUT'])
@require_admin
def reset_password(uid):
    d = request.json
    if not d.get('new_password'):
        return jsonify({'error': 'กรุณาใส่ password ใหม่'}), 400
    mutate("UPDATE users SET password_hash=%s WHERE id=%s",
           (hash_password(d['new_password']), uid))
    return jsonify({'ok': True})

@app.route('/api/admin/clear-tickets', methods=['POST'])
@require_admin
def clear_all_tickets():
    """Delete all ticket data (admin only). Irreversible."""
    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute("DELETE FROM ticket_fault_attribution")
        cur.execute("DELETE FROM ticket_comments")
        cur.execute("DELETE FROM ticket_workflow_log")
        cur.execute("DELETE FROM ticket_assignments")
        cur.execute("DELETE FROM tickets")
        conn.commit()
        conn.close()
        return jsonify({'ok': True, 'message': 'ลบข้อมูลเคสทั้งหมดเรียบร้อยแล้ว'})
    except Exception as e:
        conn.rollback(); conn.close()
        return jsonify({'error': str(e)}), 500

@app.route('/api/admin/seed-tickets', methods=['POST'])
@require_admin
def seed_tickets():
    """Create realistic sample tickets for dashboard testing (admin only)."""
    import random
    from datetime import datetime, timedelta

    n = int(request.json.get('count', 60))
    days_back = int(request.json.get('days', 30))
    now = datetime.now()

    teams      = ['CX','Sales/KAM','Sales Co','Merchandise','Inbound','Outbound']
    fault_teams= ['CX','Sales/KAM','Sales Co','Merchandise','Inbound','Outbound','Customer']
    case_types = ['Complain','Claim','Complain','Complain','Claim','Update Invoice']
    priorities = ['Urgent','High','High','Medium','Medium','Medium','Low']
    statuses   = ['open','in_progress','pending_ack','pending_fault','closed']
    root_causes= ['สินค้าตกหล่น','คุณภาพไม่ผ่าน/ไม่ได้ spec','น้ำหนักไม่ครบ','ส่งผิด SKU','เอกสารผิดพลาด','อื่นๆ']
    skus       = ['SKU-001','SKU-002','SKU-003','SKU-004','SKU-005','SKU-006','SKU-007']
    products   = ['สินค้า A','สินค้า B','สินค้า C','สินค้า D','สินค้า E']
    accounts   = ['ร้านตัวอย่าง A','ร้านตัวอย่าง B','ร้านตัวอย่าง C','ร้านตัวอย่าง D']

    conn = get_db()
    cur  = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    try:
        # get opener user id
        cur.execute("SELECT id, username FROM users LIMIT 1")
        u = cur.fetchone()
        opener_id   = u['id']   if u else None
        opener_name = u['username'] if u else 'admin'

        # pull real invoice numbers from ERP so Case/Invoice % works
        erp_inv_rows = erp_query("""
            SELECT DISTINCT invoice_number FROM sourcing_erp_order_items
            WHERE invoice_number IS NOT NULL AND invoice_number != ''
            ORDER BY RANDOM() LIMIT 500
        """)
        real_invoices = [r['invoice_number'] for r in erp_inv_rows]

        created = 0
        for i in range(n):
            dt = now - timedelta(days=random.randint(0, days_back),
                                 hours=random.randint(0,23), minutes=random.randint(0,59))
            dt_str  = dt.strftime('%Y-%m-%dT%H:%M:%S')
            opener  = random.choice(teams)
            ctype   = random.choice(case_types)
            prio    = random.choice(priorities)
            # ใช้ invoice จริงจาก orders (ถ้ามี) เพื่อให้ Case/Invoice % ทำงานได้
            inv_no  = random.choice(real_invoices) if real_invoices else f'INV-SEED-{random.randint(10000,99999)}'
            sku     = random.choice(skus)
            product = random.choice(products)
            account = random.choice(accounts)

            # determine status weighted toward realistic distribution
            status = random.choices(statuses, weights=[15,20,10,10,45])[0]
            cur_team = opener if status in ('open','in_progress') else random.choice(teams)
            closed_at = None
            if status == 'closed':
                close_dt = dt + timedelta(hours=random.randint(2, 72))
                closed_at = close_dt.strftime('%Y-%m-%dT%H:%M:%S')

            ticket_no = f'TK-SEED-{dt.strftime("%Y%m")}-{i+1:04d}'
            cur.execute("""
                INSERT INTO tickets
                (ticket_no, invoice_number, sku_code, product_name, case_type, priority,
                 status, current_team, opener_team, opener_user_id,
                 description, created_by, created_at, closed_at, root_cause)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                RETURNING id
            """, (ticket_no, inv_no, sku, product, ctype, prio,
                  status, cur_team, opener, opener_id,
                  f'[Seed] ปัญหาเกี่ยวกับ {product} ร้าน {account}',
                  opener_name, dt_str, closed_at,
                  random.choice(root_causes) if ctype in ('Claim','Complain') else None))
            tid = cur.fetchone()['id']

            # workflow log
            cur.execute("""
                INSERT INTO ticket_workflow_log (ticket_id, from_team, to_team, action, note, created_by, created_at)
                VALUES (%s,%s,%s,'created',%s,%s,%s)
            """, (tid, None, opener, f'เปิดเคส {ctype}', opener_name, dt_str))

            if status in ('in_progress','pending_ack','pending_fault','closed'):
                fwd_dt = (dt + timedelta(hours=random.randint(1,8))).strftime('%Y-%m-%dT%H:%M:%S')
                fwd_team = random.choice([t for t in teams if t != opener])
                cur.execute("""
                    INSERT INTO ticket_workflow_log (ticket_id, from_team, to_team, action, note, created_by, created_at)
                    VALUES (%s,%s,%s,'forwarded','ส่งต่อเพื่อตรวจสอบ',%s,%s)
                """, (tid, opener, fwd_team, opener_name, fwd_dt))

            # fault attribution for closed/pending_fault Claim/Complain tickets
            if status in ('closed','pending_fault') and ctype in ('Claim','Complain'):
                if random.random() > 0.2:  # 80% have attribution
                    fat = random.choice(fault_teams)
                    fa_dt = (dt + timedelta(hours=random.randint(4,24))).strftime('%Y-%m-%dT%H:%M:%S')
                    # อัปเดต tickets.fault_team ด้วย (ให้ Fault Team card ใน Overview มีข้อมูล)
                    cur.execute("UPDATE tickets SET fault_team=%s WHERE id=%s", (fat, tid))
                    cur.execute("""
                        INSERT INTO ticket_fault_attribution
                        (ticket_id, fault_team, note, attributed_by, created_at)
                        VALUES (%s,%s,'[Seed] ระบุ fault อัตโนมัติ',%s,%s)
                    """, (tid, fat, opener_name, fa_dt))
            created += 1

        conn.commit()
        conn.close()
        return jsonify({'ok': True, 'message': f'สร้างข้อมูลตัวอย่างเรียบร้อย {created} เคส'})
    except Exception as e:
        conn.rollback(); conn.close()
        return jsonify({'error': str(e)}), 500

# ---------------------------------------------------------------------------
# Accounts
# ---------------------------------------------------------------------------

@app.route('/api/accounts', methods=['GET'])
@require_auth
def get_accounts():
    rows = query("""
        SELECT a.*,
               COUNT(DISTINCT o.id) AS outlet_count,
               COUNT(DISTINCT CASE WHEN t.status NOT IN ('closed') THEN t.id END) AS open_tickets
        FROM accounts a
        LEFT JOIN outlets o ON o.account_id = a.id
        LEFT JOIN tickets t ON t.outlet_id = o.id
        GROUP BY a.id ORDER BY a.name
    """)
    return jsonify(rows)

@app.route('/api/accounts', methods=['POST'])
@require_auth
def create_account():
    d = request.json
    id_ = mutate("INSERT INTO accounts (name, owner, status) VALUES (%s,%s,%s) RETURNING id",
                 (d['name'], d.get('owner', ''), d.get('status', 'active')))
    return jsonify({'id': id_}), 201

@app.route('/api/accounts/<int:aid>', methods=['GET'])
@require_auth
def get_account(aid):
    row = query("SELECT * FROM accounts WHERE id=%s", (aid,), one=True)
    if not row:
        return jsonify({'error': 'Not found'}), 404
    return jsonify(row)

@app.route('/api/accounts/<int:aid>', methods=['PUT'])
@require_auth
def update_account(aid):
    d = request.json
    mutate("UPDATE accounts SET name=%s, owner=%s, status=%s WHERE id=%s",
           (d['name'], d.get('owner', ''), d.get('status', 'active'), aid))
    return jsonify({'ok': True})

@app.route('/api/accounts/<int:aid>/outlets', methods=['GET'])
@require_auth
def get_account_outlets(aid):
    rows = query("""
        SELECT o.*,
               COUNT(DISTINCT CASE WHEN t.status NOT IN ('closed') THEN t.id END) AS open_tickets
        FROM outlets o
        LEFT JOIN tickets t ON t.outlet_id = o.id
        WHERE o.account_id=%s GROUP BY o.id ORDER BY o.name
    """, (aid,))
    return jsonify(rows)

@app.route('/api/accounts/<int:aid>/notes', methods=['GET'])
@require_auth
def get_account_notes(aid):
    return jsonify(query("SELECT * FROM account_notes WHERE account_id=%s ORDER BY created_at DESC", (aid,)))

@app.route('/api/accounts/<int:aid>/notes', methods=['POST'])
@require_auth
def add_account_note(aid):
    d = request.json
    id_ = mutate("INSERT INTO account_notes (account_id, note, created_by) VALUES (%s,%s,%s) RETURNING id",
                 (aid, d['note'], d.get('created_by', g.user['display_name'])))
    return jsonify({'id': id_}), 201

# ---------------------------------------------------------------------------
# Outlets
# ---------------------------------------------------------------------------

@app.route('/api/outlets', methods=['GET'])
@require_auth
def get_outlets():
    search = request.args.get('q', '')
    if search:
        rows = query("""
            SELECT o.*, a.name AS account_name, a.owner
            FROM outlets o JOIN accounts a ON a.id=o.account_id
            WHERE o.name ILIKE %s OR a.name ILIKE %s
            ORDER BY o.name LIMIT 100
        """, (f'%{search}%', f'%{search}%'))
    else:
        rows = query("""
            SELECT o.*, a.name AS account_name, a.owner
            FROM outlets o JOIN accounts a ON a.id=o.account_id
            ORDER BY o.name
        """)
    return jsonify(rows)

@app.route('/api/outlets/<int:oid>', methods=['GET'])
@require_auth
def get_outlet(oid):
    row = query("""
        SELECT o.*, a.name AS account_name, a.owner
        FROM outlets o JOIN accounts a ON a.id=o.account_id
        WHERE o.id=%s
    """, (oid,), one=True)
    if not row:
        return jsonify({'error': 'Not found'}), 404
    return jsonify(row)

@app.route('/api/outlets/<int:oid>/invoices', methods=['GET'])
@require_auth
def get_outlet_invoices(oid):
    outlet = query("SELECT erp_outlet_id FROM outlets WHERE id=%s", (oid,), one=True)
    if not outlet or not outlet.get('erp_outlet_id'):
        return jsonify([])
    rows = erp_query("""
        SELECT invoice_number, doc_date, COUNT(*) AS sku_count, SUM(total_sales) AS total_sales
        FROM sourcing_erp_order_items WHERE outlet_id=%s
        GROUP BY invoice_number, doc_date ORDER BY doc_date DESC
    """, (outlet['erp_outlet_id'],))
    return jsonify(rows)

@app.route('/api/outlets/<int:oid>/tickets', methods=['GET'])
@require_auth
def get_outlet_tickets(oid):
    rows = query("SELECT * FROM tickets WHERE outlet_id=%s ORDER BY created_at DESC", (oid,))
    return jsonify(rows)

@app.route('/api/outlets/<int:oid>/notes', methods=['GET'])
@require_auth
def get_outlet_notes(oid):
    outlet = query("SELECT account_id FROM outlets WHERE id=%s", (oid,), one=True)
    if not outlet:
        return jsonify([])
    return jsonify(query("SELECT * FROM account_notes WHERE account_id=%s ORDER BY created_at DESC",
                         (outlet['account_id'],)))

# ---------------------------------------------------------------------------
# Invoices
# ---------------------------------------------------------------------------

@app.route('/api/invoices/<invoice_number>/skus', methods=['GET'])
@require_auth
def get_invoice_skus(invoice_number):
    rows = erp_query("""
        SELECT DISTINCT sku AS sku_code, product_name, qty, unit
        FROM sourcing_erp_order_items WHERE invoice_number=%s ORDER BY product_name
    """, (invoice_number,))
    return jsonify(rows)

# ---------------------------------------------------------------------------
# Suppliers
# ---------------------------------------------------------------------------

@app.route('/api/suppliers', methods=['GET'])
@require_auth
def get_suppliers():
    return jsonify(query("SELECT * FROM suppliers ORDER BY name"))

@app.route('/api/suppliers', methods=['POST'])
@require_auth
def create_supplier():
    d = request.json
    id_ = mutate("""INSERT INTO suppliers
        (name, supplier_type, contact_person, phone, line_id, email,
         payment_method, credit_days, rating, status)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) RETURNING id""",
        (d['name'], d.get('supplier_type',''), d.get('contact_person',''),
         d.get('phone',''), d.get('line_id',''), d.get('email',''),
         d.get('payment_method',''), d.get('credit_days',0),
         d.get('rating',0), d.get('status','active')))
    return jsonify({'id': id_}), 201

@app.route('/api/suppliers/<int:sid>', methods=['GET'])
@require_auth
def get_supplier(sid):
    row = query("SELECT * FROM suppliers WHERE id=%s", (sid,), one=True)
    if not row:
        return jsonify({'error': 'Not found'}), 404
    return jsonify(row)

@app.route('/api/suppliers/<int:sid>', methods=['PUT'])
@require_auth
def update_supplier(sid):
    d = request.json
    mutate("""UPDATE suppliers SET name=%s, supplier_type=%s, contact_person=%s,
        phone=%s, line_id=%s, email=%s, payment_method=%s, credit_days=%s, rating=%s, status=%s
        WHERE id=%s""",
        (d['name'], d.get('supplier_type',''), d.get('contact_person',''),
         d.get('phone',''), d.get('line_id',''), d.get('email',''),
         d.get('payment_method',''), d.get('credit_days',0),
         d.get('rating',0), d.get('status','active'), sid))
    return jsonify({'ok': True})

@app.route('/api/suppliers/<int:sid>/notes', methods=['GET'])
@require_auth
def get_supplier_notes(sid):
    return jsonify(query("SELECT * FROM supplier_notes WHERE supplier_id=%s ORDER BY created_at DESC", (sid,)))

@app.route('/api/suppliers/<int:sid>/notes', methods=['POST'])
@require_auth
def add_supplier_note(sid):
    d = request.json
    id_ = mutate("INSERT INTO supplier_notes (supplier_id, note, note_type, created_by) VALUES (%s,%s,%s,%s) RETURNING id",
                 (sid, d['note'], d.get('note_type','general'), d.get('created_by', g.user['display_name'])))
    return jsonify({'id': id_}), 201

# ---------------------------------------------------------------------------
# Employees
# ---------------------------------------------------------------------------

@app.route('/api/employees', methods=['GET'])
@require_auth
def get_employees():
    team = request.args.get('team')
    if team:
        return jsonify(query("SELECT * FROM employees WHERE team=%s AND status='active' ORDER BY name", (team,)))
    return jsonify(query("SELECT * FROM employees ORDER BY team, name"))

@app.route('/api/employees', methods=['POST'])
@require_auth
def create_employee():
    d = request.json
    id_ = mutate("""INSERT INTO employees (name, employee_code, department, team, role, phone, email, status)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s) RETURNING id""",
        (d['name'], d.get('employee_code',''), d.get('department',''),
         d.get('team',''), d.get('role',''), d.get('phone',''),
         d.get('email',''), d.get('status','active')))
    return jsonify({'id': id_}), 201

@app.route('/api/employees/<int:eid>', methods=['GET'])
@require_auth
def get_employee(eid):
    row = query("SELECT * FROM employees WHERE id=%s", (eid,), one=True)
    if not row:
        return jsonify({'error': 'Not found'}), 404
    return jsonify(row)

@app.route('/api/employees/<int:eid>', methods=['PUT'])
@require_auth
def update_employee(eid):
    d = request.json
    mutate("""UPDATE employees SET name=%s, employee_code=%s, department=%s, team=%s,
        role=%s, phone=%s, email=%s, status=%s WHERE id=%s""",
        (d['name'], d.get('employee_code',''), d.get('department',''),
         d.get('team',''), d.get('role',''), d.get('phone',''),
         d.get('email',''), d.get('status','active'), eid))
    return jsonify({'ok': True})

@app.route('/api/employees/<int:eid>', methods=['DELETE'])
@require_auth
def delete_employee(eid):
    emp = query("SELECT id, name FROM employees WHERE id=%s", (eid,), one=True)
    if not emp:
        return jsonify({'error': 'ไม่พบพนักงานนี้'}), 404
    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute("UPDATE ticket_fault_attribution SET employee_id=NULL WHERE employee_id=%s", (eid,))
        cur.execute("UPDATE ticket_assignments SET employee_id=NULL WHERE employee_id=%s", (eid,))
        cur.execute("DELETE FROM employees WHERE id=%s", (eid,))
        conn.commit()
    except Exception as e:
        conn.rollback()
        return jsonify({'error': str(e)}), 500
    finally:
        conn.close()
    return jsonify({'ok': True})

# ---------------------------------------------------------------------------
# Tickets
# ---------------------------------------------------------------------------

def next_ticket_no():
    row = query("SELECT COUNT(*) AS cnt FROM tickets", one=True)
    return f"TK-{(row['cnt'] or 0) + 1:04d}"

@app.route('/api/tickets', methods=['GET'])
@require_auth
def get_tickets():
    team = request.args.get('team')
    status = request.args.get('status')
    outlet_id = request.args.get('outlet_id')
    wheres = []

    current_user_id = g.user.get('user_id') or g.user.get('id')
    base = """
        SELECT DISTINCT t.*, o.name AS outlet_name, a.name AS account_name,
               (SELECT image_url FROM ticket_workflow_log
                WHERE ticket_id = t.id AND image_url IS NOT NULL AND image_url != ''
                ORDER BY created_at DESC LIMIT 1) AS latest_image_url,
               EXISTS(SELECT 1 FROM case_reads cr WHERE cr.user_id=%s AND cr.ticket_id=t.id) AS is_read
        FROM tickets t
        LEFT JOIN outlets o ON o.id = t.outlet_id
        LEFT JOIN accounts a ON a.id = o.account_id
        LEFT JOIN ticket_assignments ta ON ta.ticket_id = t.id AND ta.acknowledged_at IS NULL
    """
    params = [current_user_id]

    if team and team != 'Management':
        wheres.append("(t.current_team=%s OR (ta.team=%s AND t.case_type='Complain'))")
        params.extend([team, team])
    if status:
        wheres.append("t.status=%s")
        params.append(status)
    if outlet_id:
        wheres.append("t.outlet_id=%s")
        params.append(outlet_id)

    where_clause = ("WHERE " + " AND ".join(wheres)) if wheres else ""
    rows = query(f"{base} {where_clause} ORDER BY t.created_at DESC", params)
    return jsonify(rows)

@app.route('/api/tickets', methods=['POST'])
@require_auth
def create_ticket():
    d = request.json
    ticket_no = next_ticket_no()
    case_type = d.get('case_type', '')
    opener_team = g.user['team']
    opener_user_id = g.user['user_id']
    initial_teams = d.get('initial_teams', [])
    if not initial_teams:
        return jsonify({'error': 'กรุณาเลือกทีมที่จะส่งงานให้'}), 400

    if case_type == 'Complain':
        current_team = 'pending_ack'
        status = 'pending_ack'
    else:
        current_team = initial_teams[0]
        status = 'open'

    claim_items = d.get('claim_items')  # JSON string of [{sku_code,product_name,qty,unit,claimed_qty}]
    # Use first claim item as primary sku for dashboard
    primary_sku, primary_name = d.get('sku_code'), d.get('product_name')
    if claim_items:
        import json as _json
        try:
            items = _json.loads(claim_items) if isinstance(claim_items, str) else claim_items
            if items:
                primary_sku = items[0].get('sku_code', primary_sku)
                primary_name = items[0].get('product_name', primary_name)
            claim_items = _json.dumps(items) if not isinstance(claim_items, str) else claim_items
        except Exception:
            pass

    id_ = mutate("""INSERT INTO tickets
        (ticket_no, outlet_id, invoice_number, sku_code, product_name,
         case_type, case_subtype, root_cause, priority, status,
         current_team, opener_team, opener_user_id, description, created_by,
         claim_items, resolution_type)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) RETURNING id""",
        (ticket_no, d.get('outlet_id'), d.get('invoice_number'),
         primary_sku, primary_name,
         case_type, d.get('case_subtype',''), d.get('root_cause',''),
         d.get('priority','Medium'), status,
         current_team, opener_team, opener_user_id,
         d.get('description',''), g.user['display_name'],
         claim_items, d.get('resolution_type','')))

    # For Complain: create assignments for each team
    if case_type == 'Complain':
        for team in initial_teams:
            mutate("""INSERT INTO ticket_assignments (ticket_id, team) VALUES (%s,%s) RETURNING id""",
                   (id_, team))

    # Log creation
    mutate("""INSERT INTO ticket_workflow_log
        (ticket_id, from_team, to_team, action, note, image_url, user_id, created_by)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s)""",
        (id_, opener_team, ', '.join(initial_teams), 'สร้างเคส',
         d.get('description',''), d.get('image_url',''),
         opener_user_id, g.user['display_name']))

    return jsonify({'id': id_, 'ticket_no': ticket_no}), 201

@app.route('/api/tickets/<int:tid>', methods=['GET'])
@require_auth
def get_ticket(tid):
    row = query("""
        SELECT t.*, o.name AS outlet_name, a.name AS account_name,
               u.display_name AS opener_name
        FROM tickets t
        LEFT JOIN outlets o ON o.id = t.outlet_id
        LEFT JOIN accounts a ON a.id = o.account_id
        LEFT JOIN users u ON u.id = t.opener_user_id
        WHERE t.id=%s
    """, (tid,), one=True)
    if not row:
        return jsonify({'error': 'Not found'}), 404
    return jsonify(row)

@app.route('/api/tickets/<int:tid>/assignments', methods=['GET'])
@require_auth
def get_ticket_assignments(tid):
    rows = query("""
        SELECT ta.*, e.name AS employee_name
        FROM ticket_assignments ta
        LEFT JOIN employees e ON e.id = ta.employee_id
        WHERE ta.ticket_id=%s ORDER BY ta.team
    """, (tid,))
    return jsonify(rows)

@app.route('/api/tickets/<int:tid>/forward', methods=['POST'])
@require_auth
def forward_ticket(tid):
    d = request.json
    to_team = (d.get('to_team') or '').strip()
    note = (d.get('note') or '').strip()
    image_url = d.get('image_url')
    if not to_team:
        return jsonify({'error': 'กรุณาเลือกทีมที่จะส่งต่อ'}), 400
    ticket = query("SELECT * FROM tickets WHERE id=%s", (tid,), one=True)
    if not ticket:
        return jsonify({'error': 'Not found'}), 404
    if ticket['status'] == 'closed':
        return jsonify({'error': 'เคสนี้ปิดแล้ว'}), 400

    from_team = ticket['current_team']
    mutate("UPDATE tickets SET current_team=%s, status='in_progress' WHERE id=%s", (to_team, tid))
    mutate("""INSERT INTO ticket_workflow_log
        (ticket_id, from_team, to_team, action, note, image_url, user_id, created_by)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s)""",
        (tid, from_team, to_team, 'ส่งต่อ', note, image_url,
         g.user['user_id'], g.user['display_name']))
    return jsonify({'ok': True, 'next_team': to_team})

@app.route('/api/tickets/<int:tid>/complete', methods=['POST'])
@require_auth
def complete_ticket(tid):
    d = request.json
    note = (d.get('note') or '').strip()
    image_url = d.get('image_url')
    ticket = query("SELECT * FROM tickets WHERE id=%s", (tid,), one=True)
    if not ticket:
        return jsonify({'error': 'Not found'}), 404
    if ticket['status'] == 'closed':
        return jsonify({'error': 'เคสนี้ปิดแล้ว'}), 400

    from_team = ticket['current_team']
    back_team = ticket['opener_team'] or 'CX'
    mutate("UPDATE tickets SET current_team=%s, status='in_progress' WHERE id=%s", (back_team, tid))
    mutate("""INSERT INTO ticket_workflow_log
        (ticket_id, from_team, to_team, action, note, image_url, user_id, created_by)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s)""",
        (tid, from_team, back_team, f'เสร็จแล้ว → กลับ {back_team}', note, image_url,
         g.user['user_id'], g.user['display_name']))
    return jsonify({'ok': True, 'back_to': back_team})

@app.route('/api/tickets/<int:tid>/acknowledge', methods=['POST'])
@require_auth
def acknowledge_ticket(tid):
    d = request.json
    note = (d.get('note') or '').strip()
    employee_id = d.get('employee_id')
    if not note:
        return jsonify({'error': 'กรุณาใส่ note ก่อน Acknowledge'}), 400
    if not employee_id:
        return jsonify({'error': 'กรุณาเลือกพนักงานที่รับผิดชอบ'}), 400

    ticket = query("SELECT * FROM tickets WHERE id=%s", (tid,), one=True)
    if not ticket:
        return jsonify({'error': 'Not found'}), 404

    my_team = g.user['team']
    assignment = query(
        "SELECT * FROM ticket_assignments WHERE ticket_id=%s AND team=%s AND acknowledged_at IS NULL",
        (tid, my_team), one=True
    )
    if not assignment:
        return jsonify({'error': 'ไม่มี assignment สำหรับทีมนี้ หรือ acknowledge แล้ว'}), 400

    now = datetime.now().isoformat()
    mutate("""UPDATE ticket_assignments
        SET note=%s, employee_id=%s, acknowledged_by=%s, acknowledged_user_id=%s, acknowledged_at=%s
        WHERE id=%s""",
        (note, employee_id, g.user['display_name'], g.user['user_id'], now, assignment['id']))

    mutate("""INSERT INTO ticket_workflow_log
        (ticket_id, from_team, to_team, action, note, user_id, created_by)
        VALUES (%s,%s,%s,%s,%s,%s,%s)""",
        (tid, my_team, my_team, 'Acknowledge', note, g.user['user_id'], g.user['display_name']))

    # Check if all assignments acknowledged
    pending = query(
        "SELECT COUNT(*) AS cnt FROM ticket_assignments WHERE ticket_id=%s AND acknowledged_at IS NULL",
        (tid,), one=True
    )
    if pending['cnt'] == 0:
        # Auto-close: record fault attribution for each acknowledged team/employee
        acknowledged = query(
            "SELECT * FROM ticket_assignments WHERE ticket_id=%s AND acknowledged_at IS NOT NULL",
            (tid,)
        )
        for ack in acknowledged:
            if ack['employee_id']:
                mutate("""INSERT INTO ticket_fault_attribution
                    (ticket_id, fault_team, employee_id, note, attributed_by, attributed_user_id)
                    VALUES (%s,%s,%s,%s,%s,%s)""",
                    (tid, ack['team'], ack['employee_id'], ack.get('note') or '',
                     'system (auto-close)', g.user['user_id']))
        mutate("""UPDATE tickets SET status='closed', closed_at=%s, fault_attributed_at=%s
            WHERE id=%s""", (now, now, tid))
        mutate("""INSERT INTO ticket_workflow_log
            (ticket_id, from_team, to_team, action, note, created_by)
            VALUES (%s,%s,%s,%s,%s,%s)""",
            (tid, 'system', '', 'ปิดเคสอัตโนมัติ — ทุกทีม Acknowledge ครบแล้ว', '', 'system'))

    return jsonify({'ok': True})

@app.route('/api/tickets/<int:tid>/close', methods=['POST'])
@require_auth
def close_ticket(tid):
    d = request.json
    fault_team = (d.get('fault_team') or '').strip()
    note = (d.get('note') or '').strip()

    ticket = query("SELECT * FROM tickets WHERE id=%s", (tid,), one=True)
    if not ticket:
        return jsonify({'error': 'Not found'}), 404

    is_wang_bil = ticket.get('case_type') == 'วางบิล'
    if not fault_team and not is_wang_bil:
        return jsonify({'error': 'กรุณาเลือกทีมรับผิดชอบก่อนปิดเคส'}), 400

    my_team = g.user['team']
    is_admin = g.user['role'] == 'admin'
    if ticket['opener_team'] != my_team and not is_admin:
        return jsonify({'error': 'เฉพาะทีมที่เปิดเคสเท่านั้นที่ปิดได้'}), 403

    now = datetime.now().isoformat()
    if is_wang_bil:
        new_status = 'closed'
        mutate("""UPDATE tickets SET status='closed', closed_at=%s, fault_attributed_at=%s WHERE id=%s""",
               (now, now, tid))
    elif fault_team == 'Customer':
        new_status = 'closed'
        mutate("""UPDATE tickets SET status='closed', fault_team=%s, closed_at=%s, fault_attributed_at=%s WHERE id=%s""",
               (fault_team, now, now, tid))
    else:
        new_status = 'pending_fault'
        # ตั้ง current_team=fault_team ด้วย เพื่อให้ fault team มองเห็นเคสใน My Team
        mutate("""UPDATE tickets SET status='pending_fault', fault_team=%s, current_team=%s, closed_at=%s WHERE id=%s""",
               (fault_team, fault_team, now, tid))

    mutate("""INSERT INTO ticket_workflow_log
        (ticket_id, from_team, to_team, action, note, user_id, created_by)
        VALUES (%s,%s,%s,%s,%s,%s,%s)""",
        (tid, my_team, '', 'ปิดเคส', note, g.user['user_id'], g.user['display_name']))

    return jsonify({'ok': True, 'status': new_status})

@app.route('/api/tickets/<int:tid>/attribute-fault', methods=['POST'])
@require_auth
def attribute_fault(tid):
    d = request.json
    employee_id = d.get('employee_id')
    note = (d.get('note') or '').strip()
    if not employee_id:
        return jsonify({'error': 'กรุณาเลือกพนักงานที่รับผิดชอบ'}), 400

    ticket = query("SELECT * FROM tickets WHERE id=%s", (tid,), one=True)
    if not ticket:
        return jsonify({'error': 'Not found'}), 404

    my_team = g.user['team']
    is_admin = g.user['role'] == 'admin'
    if ticket['fault_team'] != my_team and not is_admin:
        return jsonify({'error': 'เฉพาะทีมรับผิดชอบเท่านั้นที่ระบุพนักงานได้'}), 403

    now = datetime.now().isoformat()
    mutate("""INSERT INTO ticket_fault_attribution
        (ticket_id, fault_team, employee_id, note, attributed_by, attributed_user_id)
        VALUES (%s,%s,%s,%s,%s,%s) RETURNING id""",
        (tid, ticket['fault_team'], employee_id, note,
         g.user['display_name'], g.user['user_id']))

    mutate("UPDATE tickets SET status='closed', fault_attributed_at=%s WHERE id=%s", (now, tid))

    mutate("""INSERT INTO ticket_workflow_log
        (ticket_id, from_team, to_team, action, note, user_id, created_by)
        VALUES (%s,%s,%s,%s,%s,%s,%s)""",
        (tid, my_team, '', 'ระบุผู้รับผิดชอบ', note, g.user['user_id'], g.user['display_name']))

    return jsonify({'ok': True})

@app.route('/api/tickets/<int:tid>/dispute', methods=['POST'])
@require_auth
def dispute_fault(tid):
    d = request.json
    note = (d.get('note') or '').strip()
    if not note:
        return jsonify({'error': 'กรุณาระบุเหตุผลในการโต้แย้ง'}), 400

    ticket = query("SELECT * FROM tickets WHERE id=%s", (tid,), one=True)
    if not ticket:
        return jsonify({'error': 'Not found'}), 404

    my_team = g.user['team']
    is_admin = g.user['role'] == 'admin'

    if ticket['status'] != 'pending_fault':
        return jsonify({'error': 'ไม่สามารถโต้แย้งได้ในสถานะนี้'}), 400
    if ticket['fault_team'] != my_team and not is_admin:
        return jsonify({'error': 'เฉพาะทีมที่ถูกระบุเท่านั้นที่โต้แย้งได้'}), 403

    # Send case back to CX, keep status = pending_fault
    mutate("UPDATE tickets SET current_team='CX' WHERE id=%s", (tid,))

    # Mark as unread for all CX users so badge + dot light up
    mutate("""
        DELETE FROM case_reads
        WHERE ticket_id = %s
          AND user_id IN (SELECT id FROM users WHERE team = 'CX')
    """, (tid,))

    mutate("""INSERT INTO ticket_workflow_log
        (ticket_id, from_team, to_team, action, note, user_id, created_by)
        VALUES (%s,%s,%s,%s,%s,%s,%s)""",
        (tid, my_team, 'CX', 'Dispute', note, g.user['user_id'], g.user['display_name']))

    return jsonify({'ok': True})

@app.route('/api/tickets/<int:tid>/note', methods=['POST'])
@require_auth
def add_ticket_note(tid):
    d = request.json
    note = (d.get('note') or '').strip()
    image_url = d.get('image_url', '')
    if not note and not image_url:
        return jsonify({'error': 'กรุณาใส่ข้อความหรือแนบไฟล์ก่อน'}), 400
    ticket = query("SELECT id FROM tickets WHERE id=%s", (tid,), one=True)
    if not ticket:
        return jsonify({'error': 'Not found'}), 404
    mutate("""INSERT INTO ticket_workflow_log
        (ticket_id, from_team, to_team, action, note, image_url, user_id, created_by)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s)""",
        (tid, g.user['team'], '', 'note', note, image_url,
         g.user['user_id'], g.user['display_name']))
    return jsonify({'ok': True})

@app.route('/api/tickets/<int:tid>/log', methods=['GET'])
@require_auth
def get_ticket_log(tid):
    rows = query("""
        SELECT l.*, u.display_name AS user_display_name, u.team AS user_team
        FROM ticket_workflow_log l
        LEFT JOIN users u ON u.id = l.user_id
        WHERE l.ticket_id=%s ORDER BY l.created_at
    """, (tid,))
    for i, row in enumerate(rows):
        if i == 0:
            row['duration_seconds'] = None
        else:
            try:
                prev_t = datetime.fromisoformat(rows[i-1]['created_at'].replace('Z',''))
                curr_t = datetime.fromisoformat(row['created_at'].replace('Z',''))
                row['duration_seconds'] = int((curr_t - prev_t).total_seconds())
            except Exception:
                row['duration_seconds'] = None
    return jsonify(rows)

@app.route('/api/tickets/<int:tid>/comments', methods=['GET'])
@require_auth
def get_ticket_comments(tid):
    return jsonify(query("SELECT * FROM ticket_comments WHERE ticket_id=%s ORDER BY created_at", (tid,)))

@app.route('/api/tickets/<int:tid>/comments', methods=['POST'])
@require_auth
def add_ticket_comment(tid):
    d = request.json
    text = d.get('note') or d.get('comment', '')
    image_url = d.get('image_url', '')
    id_ = mutate("INSERT INTO ticket_comments (ticket_id, comment, image_url, created_by) VALUES (%s,%s,%s,%s) RETURNING id",
                 (tid, text, image_url, g.user['display_name']))
    return jsonify({'ok': True, 'id': id_})

@app.route('/api/tickets/<int:tid>/fault-attribution', methods=['GET'])
@require_auth
def get_fault_attribution(tid):
    row = query("""
        SELECT fa.*, e.name AS employee_name
        FROM ticket_fault_attribution fa
        LEFT JOIN employees e ON e.id = fa.employee_id
        WHERE fa.ticket_id=%s ORDER BY fa.created_at DESC LIMIT 1
    """, (tid,), one=True)
    return jsonify(row)

# ---------------------------------------------------------------------------
# Image upload
# ---------------------------------------------------------------------------

@app.route('/api/upload/image', methods=['POST'])
@require_auth
def upload_image():
    if not CLOUDINARY_CLOUD_NAME:
        return jsonify({'error': 'Cloudinary ยังไม่ได้ตั้งค่า กรุณาเพิ่ม CLOUDINARY_CLOUD_NAME ใน environment variables'}), 400
    if 'file' not in request.files:
        return jsonify({'error': 'No file'}), 400
    try:
        import cloudinary
        import cloudinary.uploader
        cloudinary.config(
            cloud_name=CLOUDINARY_CLOUD_NAME,
            api_key=CLOUDINARY_API_KEY,
            api_secret=CLOUDINARY_API_SECRET
        )
        f = request.files['file']
        fname = (f.filename or '').lower()
        resource_type = 'raw' if fname.endswith('.pdf') else 'image'
        result = cloudinary.uploader.upload(f, folder='smm-crm', resource_type=resource_type)
        return jsonify({'url': result['secure_url']})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ---------------------------------------------------------------------------
# Dashboard
# ---------------------------------------------------------------------------

def date_filter_sql(alias='t'):
    start = request.args.get('start')
    end = request.args.get('end')
    parts = []
    params = []
    if start:
        parts.append(f"{alias}.created_at >= %s")
        params.append(start)
    if end:
        parts.append(f"{alias}.created_at <= %s")
        params.append(end + 'T23:59:59')
    return (' AND ' + ' AND '.join(parts)) if parts else '', params

@app.route('/api/dashboard/overview', methods=['GET'])
@require_auth
def dashboard_overview():
    extra, params = date_filter_sql()
    total = query(f"SELECT COUNT(*) AS cnt FROM tickets t WHERE 1=1 {extra}", params, one=True)
    by_status = query(f"SELECT status, COUNT(*) AS cnt FROM tickets t WHERE 1=1 {extra} GROUP BY status", params)
    by_case = query(f"SELECT case_type, COUNT(*) AS cnt FROM tickets t WHERE 1=1 {extra} GROUP BY case_type ORDER BY cnt DESC", params)
    by_root = query(f"SELECT root_cause, COUNT(*) AS cnt FROM tickets t WHERE root_cause IS NOT NULL AND root_cause!='' {extra} GROUP BY root_cause ORDER BY cnt DESC", params)
    by_priority = query(f"SELECT priority, COUNT(*) AS cnt FROM tickets t WHERE 1=1 {extra} GROUP BY priority", params)
    by_opener = query(f"SELECT opener_team, COUNT(*) AS count FROM tickets t WHERE opener_team IS NOT NULL {extra} GROUP BY opener_team ORDER BY count DESC", params)
    # เคสที่ยังค้างอยู่กับแต่ละทีม (ไม่นับ closed)
    by_backlog = query(f"""
        SELECT t.current_team AS team, COUNT(*) AS cnt
        FROM tickets t
        WHERE t.status NOT IN ('closed')
          AND t.current_team IS NOT NULL
          AND t.current_team != ''
          AND t.current_team != 'pending_ack'
          {extra}
        GROUP BY t.current_team
        ORDER BY cnt DESC
    """, params)
    status_map = {r['status']: r['cnt'] for r in by_status}
    return jsonify({
        'total': total['cnt'] if total else 0,
        'open': status_map.get('open', 0),
        'in_progress': status_map.get('in_progress', 0),
        'pending_ack': status_map.get('pending_ack', 0),
        'pending_fault': status_map.get('pending_fault', 0),
        'closed': status_map.get('closed', 0),
        'by_status': by_status,
        'by_case_type': by_case,
        'by_root_cause': by_root,
        'by_priority': by_priority,
        'by_opener_team': by_opener,
        'by_team_backlog': by_backlog,
    })

@app.route('/api/dashboard/sku-problems', methods=['GET'])
@require_auth
def dashboard_sku():
    extra, params = date_filter_sql()
    rows = query(f"""
        SELECT sku_code, product_name, COUNT(*) AS cnt,
               STRING_AGG(DISTINCT case_type, ', ') AS case_types,
               STRING_AGG(DISTINCT root_cause, ', ') AS root_causes
        FROM tickets t WHERE sku_code IS NOT NULL AND sku_code != '' {extra}
        GROUP BY sku_code, product_name ORDER BY cnt DESC LIMIT 20
    """, params)
    return jsonify(rows)

@app.route('/api/dashboard/by-employee', methods=['GET'])
@require_auth
def dashboard_employee():
    extra, params = date_filter_sql()
    total_row = query(f"SELECT COUNT(*) AS cnt FROM tickets t WHERE 1=1 {extra}", params, one=True)
    total = total_row['cnt'] if total_row and total_row['cnt'] else 1
    rows = query(f"""
        SELECT e.id, e.name, e.team, COUNT(DISTINCT tfa.ticket_id) AS cnt,
               ROUND(COUNT(DISTINCT tfa.ticket_id)*100.0/%s, 1) AS pct
        FROM ticket_fault_attribution tfa
        JOIN employees e ON e.id = tfa.employee_id
        JOIN tickets t ON t.id = tfa.ticket_id
        WHERE 1=1 {extra}
        GROUP BY e.id, e.name, e.team ORDER BY cnt DESC
    """, [total] + params)
    return jsonify(rows)

@app.route('/api/dashboard/by-customer', methods=['GET'])
@require_auth
def dashboard_customer():
    extra, params = date_filter_sql()
    rows = query(f"""
        SELECT o.id AS outlet_id, o.name AS outlet_name, a.name AS account_name,
               COUNT(t.id) AS cnt,
               STRING_AGG(DISTINCT t.case_type, ', ') AS case_types
        FROM outlets o
        JOIN accounts a ON a.id = o.account_id
        LEFT JOIN tickets t ON t.outlet_id = o.id {('AND 1=1' + extra) if extra else ''}
        GROUP BY o.id, o.name, a.name ORDER BY cnt DESC LIMIT 20
    """, params)
    return jsonify(rows)

@app.route('/api/dashboard/bottleneck', methods=['GET'])
@require_auth
def dashboard_bottleneck():
    extra, params = date_filter_sql()
    ticket_ids_row = query(f"SELECT id FROM tickets t WHERE 1=1 {extra}", params)
    ticket_ids = [r['id'] for r in ticket_ids_row] if ticket_ids_row else []
    if not ticket_ids:
        return jsonify([])
    placeholders = ','.join(['%s'] * len(ticket_ids))
    logs = query(f"""
        SELECT ticket_id, from_team, created_at
        FROM ticket_workflow_log
        WHERE ticket_id IN ({placeholders})
        ORDER BY ticket_id, created_at
    """, ticket_ids)
    team_totals = {}
    team_counts = {}
    by_ticket = {}
    for row in logs:
        tid = row['ticket_id']
        if tid not in by_ticket:
            by_ticket[tid] = []
        by_ticket[tid].append(row)
    for tid, entries in by_ticket.items():
        for i in range(1, len(entries)):
            try:
                prev_t = datetime.fromisoformat(entries[i-1]['created_at'].replace('Z',''))
                curr_t = datetime.fromisoformat(entries[i]['created_at'].replace('Z',''))
                secs = (curr_t - prev_t).total_seconds()
                team = entries[i-1]['from_team'] or 'unknown'
                if team in ('system', 'unknown', ''):
                    continue
                team_totals[team] = team_totals.get(team, 0) + secs
                team_counts[team] = team_counts.get(team, 0) + 1
            except Exception:
                pass
    result = []
    for team in TEAMS:
        cnt = team_counts.get(team, 0)
        total = team_totals.get(team, 0)
        result.append({
            'team': team,
            'avg_seconds': int(total / cnt) if cnt > 0 else 0,
            'ticket_count': cnt
        })
    result.sort(key=lambda x: x['avg_seconds'], reverse=True)
    return jsonify(result)

@app.route('/api/dashboard/fault-by-team', methods=['GET'])
@require_auth
def dashboard_fault_team():
    rows = query("""
        SELECT fault_team, COUNT(*) AS cnt FROM tickets
        WHERE fault_team IS NOT NULL AND fault_team != ''
        GROUP BY fault_team ORDER BY cnt DESC
    """)
    return jsonify(rows)

@app.route('/api/dashboard/trend', methods=['GET'])
@require_auth
def dashboard_trend():
    extra, params = date_filter_sql()
    rows = query(f"""
        SELECT DATE(t.created_at) AS day, COUNT(*) AS cnt
        FROM tickets t WHERE 1=1 {extra}
        GROUP BY DATE(t.created_at) ORDER BY day
    """, params)
    return jsonify(rows or [])

@app.route('/api/dashboard/aging', methods=['GET'])
@require_auth
def dashboard_aging():
    rows = query("""
        SELECT
            CASE
                WHEN EXTRACT(EPOCH FROM (NOW() - created_at::timestamp))/3600 < 24 THEN '< 1 วัน'
                WHEN EXTRACT(EPOCH FROM (NOW() - created_at::timestamp))/3600 < 72 THEN '1-3 วัน'
                WHEN EXTRACT(EPOCH FROM (NOW() - created_at::timestamp))/3600 < 168 THEN '3-7 วัน'
                ELSE '> 7 วัน'
            END AS bucket,
            COUNT(*) AS cnt,
            MIN(EXTRACT(EPOCH FROM (NOW() - created_at::timestamp))) AS min_age
        FROM tickets
        WHERE status NOT IN ('closed')
        GROUP BY bucket ORDER BY min_age
    """)
    return jsonify(rows or [])

@app.route('/api/dashboard/case-invoice-ratio', methods=['GET'])
@require_auth
def dashboard_case_invoice_ratio():
    # กรองด้วย delivery_date (วันที่ส่งสินค้าจริงจาก ERP)
    # ถ้า delivery_date เป็น NULL จะ fallback ไป doc_date
    # ตัวหาร = จำนวน invoice ที่ส่งในช่วงนั้น
    # ตัวตั้ง = Claim+Complain ที่ invoice นั้นอยู่ในช่วงเดียวกัน
    start = request.args.get('start', '')
    end = request.args.get('end', '')
    date_extra = ''
    date_params = []
    if start:
        date_extra += " AND COALESCE(o.delivery_started_at::date, o.doc_date) >= %s"
        date_params.append(start)
    if end:
        date_extra += " AND COALESCE(o.delivery_started_at::date, o.doc_date) <= %s"
        date_params.append(end)

    # ตัวหาร: invoice ที่ส่งในช่วง delivery_date (จาก ERP)
    inv_row = erp_query(f"""
        SELECT COUNT(DISTINCT o.invoice_number) AS total_invoices
        FROM sourcing_erp_order_items o WHERE 1=1 {date_extra}
    """, date_params, one=True)
    total_invoices = int(inv_row['total_invoices']) if inv_row and inv_row['total_invoices'] else 0
    if total_invoices == 0:
        return jsonify({'total_invoices': 0, 'teams': []})

    # ตัวตั้ง: Claim+Complain ที่ invoice ส่งในช่วงนั้น แยกตาม fault_team
    # ดึง invoice list จาก ERP ก่อน แล้วค่อย query CRM
    erp_invoices = erp_query(f"""
        SELECT DISTINCT invoice_number FROM sourcing_erp_order_items o WHERE 1=1 {date_extra}
    """, date_params)
    inv_list = [r['invoice_number'] for r in erp_invoices if r.get('invoice_number')]

    if not inv_list:
        return jsonify({'total_invoices': total_invoices, 'teams': []})

    rows = query("""
        SELECT COALESCE(tfa.fault_team, 'ยังไม่ระบุ') AS fault_team, COUNT(*) AS cnt
        FROM tickets t
        LEFT JOIN ticket_fault_attribution tfa ON tfa.ticket_id = t.id
        WHERE t.case_type IN ('Claim', 'Complain', 'Update Invoice')
          AND t.invoice_number = ANY(%s)
        GROUP BY COALESCE(tfa.fault_team, 'ยังไม่ระบุ') ORDER BY cnt DESC
    """, (inv_list,))
    teams = [{'team': r['fault_team'], 'cnt': int(r['cnt']),
               'pct': round(int(r['cnt']) / total_invoices * 100, 2)} for r in (rows or [])]
    return jsonify({'total_invoices': total_invoices, 'teams': teams})

@app.route('/api/dashboard/case-invoice-debug', methods=['GET'])
@require_auth
def dashboard_case_invoice_debug():
    """Show why tickets are excluded from Case/Invoice % calculation."""
    extra, params = date_filter_sql()
    start = request.args.get('start', '')
    end   = request.args.get('end', '')
    date_extra = ''
    date_params = []
    if start:
        date_extra += " AND COALESCE(o.delivery_date, o.doc_date) >= %s"
        date_params.append(start)
    if end:
        date_extra += " AND COALESCE(o.delivery_date, o.doc_date) <= %s"
        date_params.append(end)

    # ดึง tickets ทั้งหมดในช่วงเวลา (จาก created_at เหมือน Overview)
    tickets_all = query(f"""
        SELECT t.id, t.ticket_no, t.case_type, t.invoice_number,
               t.created_at, t.status
        FROM tickets t WHERE 1=1 {extra}
        ORDER BY t.created_at DESC
    """, params)

    result = []
    for t in (tickets_all or []):
        reason = None
        included = False
        inv = t['invoice_number']

        if t['case_type'] not in ('Claim', 'Complain', 'Update Invoice'):
            reason = f"ประเภท '{t['case_type']}' ไม่ใช่ Claim/Complain/Update Invoice"
        elif not inv or inv.strip() == '':
            reason = 'ไม่มี Invoice Number'
        else:
            # check invoice exists in ERP
            in_erp = erp_query(
                "SELECT 1 FROM sourcing_erp_order_items WHERE invoice_number=%s LIMIT 1", (inv,), one=True)
            if not in_erp:
                reason = f'Invoice {inv} ไม่มีในระบบ ERP'
            else:
                # check delivery_date in range
                if date_extra:
                    in_range = erp_query(
                        f"SELECT 1 FROM sourcing_erp_order_items o WHERE o.invoice_number=%s {date_extra} LIMIT 1",
                        [inv] + date_params, one=True)
                    if not in_range:
                        reason = f'Invoice {inv} อยู่นอกช่วงวันที่ที่เลือก'
                    else:
                        included = True
                        reason = '✓ นับใน Case/Invoice %'
                else:
                    included = True
                    reason = '✓ นับใน Case/Invoice %'

        result.append({
            'ticket_no': t['ticket_no'],
            'case_type': t['case_type'],
            'invoice_number': inv or '-',
            'created_at': t['created_at'],
            'status': t['status'],
            'included': included,
            'reason': reason,
        })

    total = len(result)
    included_count = sum(1 for r in result if r['included'])
    return jsonify({'total': total, 'included': included_count, 'rows': result})

@app.route('/api/dashboard/fault-rate-trend', methods=['GET'])
@require_auth
def dashboard_fault_rate_trend():
    """Fault rate trend by date or week — for Case/Invoice % chart."""
    start    = request.args.get('start', '')
    end      = request.args.get('end', '')
    group_by = request.args.get('group_by', 'day')  # 'day' | 'week'

    date_extra = ''
    date_params = []
    if start:
        date_extra += " AND COALESCE(o.delivery_started_at::date, o.doc_date) >= %s"
        date_params.append(start)
    if end:
        date_extra += " AND COALESCE(o.delivery_started_at::date, o.doc_date) <= %s"
        date_params.append(end)

    # For weekly: group by Sun-Sat week
    if group_by == 'week':
        dt_expr = "(COALESCE(o.delivery_started_at::date, o.doc_date) - CAST(EXTRACT(DOW FROM COALESCE(o.delivery_started_at::date, o.doc_date)) AS INT))"
    else:
        dt_expr = "COALESCE(o.delivery_started_at::date, o.doc_date)"

    # Invoices per period — from ERP
    inv_rows = erp_query(f"""
        SELECT {dt_expr} AS dt,
               COUNT(DISTINCT o.invoice_number) AS inv_cnt
        FROM sourcing_erp_order_items o WHERE 1=1 {date_extra}
        GROUP BY dt ORDER BY dt
    """, date_params)

    if not inv_rows:
        return jsonify([])

    # Get invoice list per period for CRM join
    erp_inv_by_dt = erp_query(f"""
        SELECT {dt_expr} AS dt, invoice_number
        FROM sourcing_erp_order_items o WHERE 1=1 {date_extra}
    """, date_params)
    dt_to_invoices = {}
    for r in (erp_inv_by_dt or []):
        dt_to_invoices.setdefault(r['dt'], set()).add(r['invoice_number'])
    all_inv_list = list({inv for s in dt_to_invoices.values() for inv in s})

    # Fault cases per period — overall
    fault_rows = query("""
        SELECT t.invoice_number, COUNT(DISTINCT t.id) AS fault_cnt
        FROM tickets t
        WHERE t.case_type IN ('Claim','Complain','Update Invoice')
          AND t.invoice_number = ANY(%s)
        GROUP BY t.invoice_number
    """, (all_inv_list,)) if all_inv_list else []

    # Fault cases per period per team
    team_rows = query("""
        SELECT t.invoice_number,
               COALESCE(t.fault_team, 'ยังไม่ระบุ') AS team,
               COUNT(DISTINCT t.id) AS fault_cnt
        FROM tickets t
        WHERE t.case_type IN ('Claim','Complain','Update Invoice')
          AND t.invoice_number = ANY(%s)
        GROUP BY t.invoice_number, team
    """, (all_inv_list,)) if all_inv_list else []

    # Build per-dt fault maps
    inv_fault_map = {}
    for r in (fault_rows or []):
        inv_fault_map[r['invoice_number']] = int(r['fault_cnt'])
    inv_team_map = {}
    for r in (team_rows or []):
        inv_team_map.setdefault(r['invoice_number'], {})[r['team']] = int(r['fault_cnt'])

    fault_map = {}
    team_map = {}
    for dt, inv_set in dt_to_invoices.items():
        fault_map[dt] = sum(inv_fault_map.get(inv, 0) for inv in inv_set)
        for inv in inv_set:
            for team, cnt in inv_team_map.get(inv, {}).items():
                team_map.setdefault(team, {})
                team_map[team][dt] = team_map[team].get(dt, 0) + cnt

    # build maps
    fault_map = {r['dt']: int(r['fault_cnt']) for r in (fault_rows or [])}
    team_map = {}  # team -> {dt -> cnt}
    for r in (team_rows or []):
        team = r['team']
        if team not in team_map:
            team_map[team] = {}
        team_map[team][r['dt']] = int(r['fault_cnt'])

    result = []
    for r in inv_rows:
        dt = r['dt']
        inv = int(r['inv_cnt'])
        fault = fault_map.get(dt, 0)
        row = {
            'date': str(dt),
            'invoices': inv,
            'faults': fault,
            'rate': round(fault / inv * 100, 2) if inv else 0,
            'by_team': {}
        }
        for team, dmap in team_map.items():
            cnt = dmap.get(dt, 0)
            row['by_team'][team] = {
                'faults': cnt,
                'rate': round(cnt / inv * 100, 2) if inv else 0
            }
        result.append(row)
    return jsonify(result)

@app.route('/api/my-team/active', methods=['GET'])
@require_auth
def my_team_active():
    """Lightweight endpoint for notification polling — returns active cases for current user's team."""
    team = g.user.get('team') or g.user.get('team', '')
    uid  = g.user.get('user_id') or g.user.get('id')
    if not team:
        return jsonify({'count': 0, 'unread_count': 0, 'cases': []})
    rows = query("""
        SELECT t.id, t.ticket_no, t.case_type, t.created_at,
               COALESCE(o.name, '') AS outlet_name,
               COALESCE(a.name, '') AS account_name,
               NOT EXISTS(SELECT 1 FROM case_reads cr WHERE cr.user_id=%s AND cr.ticket_id=t.id) AS is_unread
        FROM tickets t
        LEFT JOIN outlets o ON o.id = t.outlet_id
        LEFT JOIN accounts a ON a.id = o.account_id
        WHERE t.current_team = %s AND t.status NOT IN ('closed')
        ORDER BY t.created_at DESC
    """, (uid, team))
    unread_count = sum(1 for r in rows if r.get('is_unread'))

    # For CX: also return active dispute cases so frontend can notify
    disputed_cases = []
    if team == 'CX':
        disputed_cases = query("""
            SELECT DISTINCT ON (wl.ticket_id)
                   wl.id AS log_id, wl.ticket_id,
                   wl.from_team, wl.note AS dispute_note, wl.created_at AS disputed_at,
                   t.ticket_no, t.case_type,
                   COALESCE(o.name,'') AS outlet_name
            FROM ticket_workflow_log wl
            JOIN tickets t ON t.id = wl.ticket_id
            LEFT JOIN outlets o ON o.id = t.outlet_id
            WHERE wl.action = 'Dispute'
              AND t.current_team = 'CX'
              AND t.status = 'pending_fault'
            ORDER BY wl.ticket_id, wl.created_at DESC
        """)

    return jsonify({'count': len(rows), 'unread_count': unread_count,
                    'cases': rows, 'disputed_cases': disputed_cases})

@app.route('/api/tickets/<int:tid>/read', methods=['POST'])
@require_auth
def mark_ticket_read(tid):
    uid = g.user.get('user_id') or g.user.get('id')
    mutate("""
        INSERT INTO case_reads (user_id, ticket_id)
        VALUES (%s, %s)
        ON CONFLICT (user_id, ticket_id) DO NOTHING
    """, (uid, tid))
    return jsonify({'ok': True})

@app.route('/api/dashboard/fault-by-employee', methods=['GET'])
@require_auth
def dashboard_fault_employee():
    rows = query("""
        SELECT e.name, e.team, COUNT(*) AS cnt
        FROM ticket_fault_attribution fa
        JOIN employees e ON e.id = fa.employee_id
        GROUP BY e.id, e.name, e.team ORDER BY cnt DESC LIMIT 20
    """)
    return jsonify(rows)

# ---------------------------------------------------------------------------
# ERP Import
# ---------------------------------------------------------------------------

def safe_str(v):
    if v is None:
        return None
    s = str(v).strip()
    return None if s.lower() in ('nan', 'none', '') else s

def safe_float(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None

@app.route('/api/import/erp', methods=['POST'])
@require_auth
def import_erp():
    if 'file' not in request.files:
        return jsonify({'error': 'No file'}), 400
    f = request.files['file']
    wb = openpyxl.load_workbook(f, data_only=True)

    sheet = None
    for name in wb.sheetnames:
        if 'invoice' in name.lower() or 'archive' in name.lower():
            sheet = wb[name]
            break
    if sheet is None:
        sheet = wb.active

    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return jsonify({'error': 'Empty file'}), 400

    # Normalize header: lowercase + replace spaces/dashes with underscore
    # ทำให้ "Order Item Id", "order_item_id", "OrderItemId" → "order_item_id" หมด
    def norm_col(h):
        return str(h).strip().lower().replace(' ', '_').replace('-', '_') if h is not None else ''

    headers = [str(h).strip() if h is not None else '' for h in rows[0]]
    col_idx = {norm_col(h): i for i, h in enumerate(headers)}  # normalized → column index

    # COL_ALIASES: key = canonical name (normalized), values = alternate normalized names
    COL_ALIASES = {
        'order_item_id':      ['order_item_id', 'order_iter', 'orderitemid', 'order_item'],
        'total_sales':        ['total_sales', 'total_sale'],
        'outlet_id':          ['outlet_id', 'customer_outlet_id', 'outletid'],
        'customer_id':        ['customer_id', 'customerid', 'erp_customer_id'],
        'account_name':       ['account_name', 'accountname', 'account'],
        'customer_name':      ['customer_name', 'customername'],
        'product_name':       ['product_name', 'productname'],
        'sku_category':       ['sku_category', 'sku_categ'],
        'delivery_date':      ['delivery_date', 'delivery_dt'],
        'delivery_started_at': ['delivery_started_at', 'delivery_s', 'delivery_start'],
        'delivery_finished_at':['delivery_finished_at', 'delivery_f', 'delivery_finish'],
        'invoice_number':     ['invoice_number', 'invoice_no', 'invoicenumber'],
        'doc_no':             ['doc_no', 'docno'],
        'doc_date':           ['doc_date'],
        'sku':                ['sku'],
        'qty':                ['qty'],
        'csc_code':           ['csc_code', 'csccode'],
        'order_id':           ['order_id', 'orderid'],
        'owner':              ['owner'],
        'unit':               ['unit'],
        'vat_price':          ['vat_price', 'vatprice'],
        'is_vat':             ['is_vat', 'isvat'],
        'sku_group':          ['sku_group', 'skupgroup'],
        'sku_type':           ['sku_type', 'skutype'],
        'loaded_at':          ['loaded_at', 'loadedat'],
    }

    def get(row, erp_col):
        key = norm_col(erp_col)
        candidates = COL_ALIASES.get(key, [key])
        for name in candidates:
            idx = col_idx.get(norm_col(name))
            if idx is not None and idx < len(row):
                return row[idx]
        return None

    def parse_date_str(v):
        """แปลง string วันที่หลายรูปแบบ → YYYY-MM-DD หรือ None
        ตรรกะ:
          - มี AM/PM → ใช้ US format M/D/YYYY H:MM:SS AM/PM  (เช่น delivery_started_at)
          - ไม่มี AM/PM → ใช้ Thai format DD/MM/YYYY ก่อน   (เช่น doc_date)
        """
        if not v:
            return None
        s = str(v).strip()
        if not s or s.lower() in ('none', 'nan', ''):
            return None
        # Already ISO date YYYY-MM-DD
        if len(s) == 10 and s[4] == '-':
            return s
        # Excel serial number (numeric)
        if s.replace('.', '', 1).isdigit():
            try:
                from openpyxl.utils.datetime import from_excel
                return str(from_excel(float(s)).date())
            except Exception:
                pass
        from datetime import datetime as _dt
        su = s.upper()
        if 'AM' in su or 'PM' in su:
            # US-style datetime: M/D/YYYY H:MM:SS AM/PM  (เช่น "4/1/2026 1:00:00 AM")
            for fmt in ('%m/%d/%Y %I:%M:%S %p', '%m/%d/%Y %I:%M %p'):
                try:
                    return str(_dt.strptime(s, fmt).date())
                except Exception:
                    pass
        else:
            # Thai-style date: DD/MM/YYYY [HH:MM:SS]  (เช่น "01/04/2026")
            for fmt in ('%d/%m/%Y %H:%M:%S', '%d/%m/%Y',
                        '%Y-%m-%dT%H:%M:%S', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d',
                        '%m/%d/%Y %H:%M:%S', '%m/%d/%Y'):
                try:
                    return str(_dt.strptime(s, fmt).date())
                except Exception:
                    pass
        return s  # return as-is if nothing matched

    parsed = []
    for row in rows[1:]:
        erp_item_id = safe_str(get(row, 'order_item_id'))
        if not erp_item_id:
            continue

        doc_date = parse_date_str(safe_str(get(row, 'doc_date')))

        # delivery_date: ใช้คอลัม delivery_date ถ้ามี, ถ้าไม่มีให้ดึงวันจาก delivery_started_at
        delivery_date_raw = safe_str(get(row, 'delivery_date'))
        if delivery_date_raw:
            delivery_date = parse_date_str(delivery_date_raw)
        else:
            # extract date portion from delivery_started_at (เช่น "4/1/2026 1:00:00 AM")
            delivery_date = parse_date_str(safe_str(get(row, 'delivery_started_at')))

        is_vat_val = get(row, 'is_vat')
        if isinstance(is_vat_val, bool):
            is_vat = 1 if is_vat_val else 0
        elif isinstance(is_vat_val, str):
            is_vat = 1 if is_vat_val.strip().lower() in ('true', '1', 'yes') else 0
        else:
            is_vat = 1 if is_vat_val else 0

        parsed.append({
            'erp_item_id':      erp_item_id,
            'account_name':     safe_str(get(row, 'account_name')) or 'Unknown',
            'outlet_name':      safe_str(get(row, 'customer_name')) or 'Unknown',
            'owner':            safe_str(get(row, 'owner')) or '',
            'erp_customer_id':  safe_str(get(row, 'customer_id')),
            'erp_outlet_id':    safe_str(get(row, 'outlet_id')),
            'csc_code':         safe_str(get(row, 'csc_code')),
            'order_id':         safe_str(get(row, 'order_id')),
            'invoice_number':   safe_str(get(row, 'invoice_number')),
            'doc_no':           safe_str(get(row, 'doc_no')),
            'doc_date':         doc_date,
            'delivery_date':    delivery_date,
            'sku_code':         safe_str(get(row, 'sku')),
            'product_name':     safe_str(get(row, 'product_name')),
            'qty':              safe_float(get(row, 'qty')),
            'unit':             safe_str(get(row, 'unit')),
            'total_sales':      safe_float(get(row, 'total_sales')),
            'vat_price':        safe_float(get(row, 'vat_price')),
            'is_vat':           is_vat,
            'sku_group':        safe_str(get(row, 'sku_group')),
            'sku_category':     safe_str(get(row, 'sku_category')),
            'sku_type':         safe_str(get(row, 'sku_type')),
            'delivery_started_at':  safe_str(get(row, 'delivery_started_at')),
            'delivery_finished_at': safe_str(get(row, 'delivery_finished_at')),
            'loaded_at':        safe_str(get(row, 'loaded_at')),
        })

    inserted = 0
    skipped = 0
    errors = []

    conn = get_db()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    try:
        all_ids = [p['erp_item_id'] for p in parsed]
        cur.execute("SELECT erp_item_id FROM orders WHERE erp_item_id = ANY(%s)", (all_ids,))
        existing_ids = {r['erp_item_id'] for r in cur.fetchall()}
        new_rows = [p for p in parsed if p['erp_item_id'] not in existing_ids]
        skipped = len(parsed) - len(new_rows)

        account_map = {}
        unique_accounts = {(p['account_name'], p['owner']) for p in new_rows}
        for acc_name, owner in unique_accounts:
            cur.execute("SELECT id FROM accounts WHERE name=%s", (acc_name,))
            r = cur.fetchone()
            if r:
                account_map[acc_name] = r['id']
            else:
                cur.execute("INSERT INTO accounts (name, owner) VALUES (%s,%s) RETURNING id", (acc_name, owner))
                account_map[acc_name] = cur.fetchone()['id']

        outlet_map = {}
        # Dedupe outlets by (acc_name, out_name) — same outlet may appear many times with diff erp_oid/csc
        seen_outlet_keys = {}
        for p in new_rows:
            key = (p['account_name'], p['outlet_name'])
            if key not in seen_outlet_keys:
                seen_outlet_keys[key] = p
        for (acc_name, out_name), p in seen_outlet_keys.items():
            acc_id = account_map.get(acc_name)
            if acc_id is None:
                continue
            erp_cid = p['erp_customer_id']
            erp_oid = p['erp_outlet_id']
            csc = p['csc_code']
            # Always look up by (account_id, name) — most reliable key
            cur.execute("SELECT id, name FROM outlets WHERE account_id=%s AND name=%s", (acc_id, out_name))
            r = cur.fetchone()
            if r:
                outlet_map[(acc_id, r['name'])] = r['id']
            else:
                cur.execute("""
                    INSERT INTO outlets (account_id, name, erp_customer_id, erp_outlet_id, csc_code)
                    VALUES (%s,%s,%s,%s,%s) RETURNING id, name
                """, (acc_id, out_name, erp_cid, erp_oid, csc))
                r = cur.fetchone()
                outlet_map[(acc_id, r['name'])] = r['id']

        # Fallback: look up outlets from DB if not found in outlet_map (e.g. existed from prev import)
        missing = {(p['account_name'], p['outlet_name']) for p in new_rows
                   if account_map.get(p['account_name']) and
                   (account_map[p['account_name']], p['outlet_name']) not in outlet_map}
        for acc_name, out_name in missing:
            acc_id = account_map[acc_name]
            cur.execute("SELECT id FROM outlets WHERE account_id=%s AND name=%s", (acc_id, out_name))
            r = cur.fetchone()
            if r:
                outlet_map[(acc_id, out_name)] = r['id']

        order_tuples = []
        for p in new_rows:
            acc_id = account_map.get(p['account_name'])
            out_id = outlet_map.get((acc_id, p['outlet_name'])) if acc_id else None
            order_tuples.append((
                p['erp_item_id'], p['order_id'], p['invoice_number'], p['doc_no'],
                p['doc_date'], p['delivery_date'], out_id, p['sku_code'], p['product_name'],
                p['qty'], p['unit'], p['total_sales'], p['vat_price'], p['is_vat'],
                p['sku_group'], p['sku_category'], p['sku_type'],
                p['delivery_started_at'], p['delivery_finished_at'], p['loaded_at'],
            ))

        if order_tuples:
            psycopg2.extras.execute_values(cur, """
                INSERT INTO orders
                (erp_item_id, order_id, invoice_number, doc_no, doc_date, delivery_date,
                 outlet_id, sku_code, product_name, qty, unit,
                 total_sales, vat_price, is_vat, sku_group, sku_category,
                 sku_type, delivery_started_at, delivery_finished_at, loaded_at)
                VALUES %s
            """, order_tuples, page_size=500)
            inserted = len(order_tuples)

        conn.commit()
    except Exception as e:
        conn.rollback()
        return jsonify({'error': str(e)}), 500
    finally:
        cur.close()
        conn.close()

    return jsonify({'inserted': inserted, 'skipped': skipped, 'errors': errors})

# ---------------------------------------------------------------------------
# Run
# ---------------------------------------------------------------------------

init_db()

@app.route('/api/debug/all-tickets-fault', methods=['GET'])
def debug_all_tickets_fault():
    """Show all tickets with fault attribution and invoice match status."""
    rows = query("""
        SELECT
            t.id, t.case_type, t.status,
            t.invoice_number,
            t.fault_team AS tickets_fault_team,
            tfa.fault_team AS tfa_fault_team,
            CASE WHEN t.invoice_number IS NOT NULL AND t.invoice_number != ''
                 THEN 1 ELSE 0
            END AS invoice_in_orders,
            t.opener_team, t.description
        FROM tickets t
        LEFT JOIN ticket_fault_attribution tfa ON tfa.ticket_id = t.id
        ORDER BY t.id
    """)
    return jsonify(rows)

@app.route('/api/debug/complain-no-invoice', methods=['GET'])
def debug_complain_no_invoice():
    rows = query("""
        SELECT t.id, t.ticket_no, t.status, t.priority,
               t.description, t.opener_team, t.current_team,
               t.fault_team, t.created_at,
               o.name AS outlet_name, a.name AS account_name
        FROM tickets t
        LEFT JOIN outlets o ON o.id = t.outlet_id
        LEFT JOIN accounts a ON a.id = o.account_id
        WHERE t.case_type = 'Complain'
          AND (t.invoice_number IS NULL OR t.invoice_number = '')
        ORDER BY t.created_at DESC
    """)
    return jsonify(rows)

# ---------------------------------------------------------------------------
# Pipeline / Leads
# ---------------------------------------------------------------------------

LEAD_STAGES = ['Cold Call/Email', 'Meeting', 'Follow Up', 'Closed Win', 'Closed Lost']

def next_lead_no(cur):
    cur.execute("SELECT lead_no FROM leads ORDER BY id DESC LIMIT 1")
    row = cur.fetchone()
    lead_no_val = row['lead_no'] if row else None
    if not lead_no_val:
        return 'LD-0001'
    try:
        num = int(lead_no_val.split('-')[1]) + 1
    except Exception:
        num = 1
    return f'LD-{num:04d}'

@app.route('/api/leads', methods=['GET'])
@require_auth
def get_leads():
    stage  = request.args.get('stage')
    owner  = request.args.get('owner_user_id')
    wheres = []
    params = []
    if stage:
        wheres.append("l.stage = %s"); params.append(stage)
    # Sales rep: enforce own-leads-only filter regardless of query params
    if g.user['role'] == 'sales':
        wheres.append("l.owner_user_id = %s"); params.append(g.user['user_id'])
    elif owner:
        wheres.append("l.owner_user_id = %s"); params.append(owner)
    where_sql = ('WHERE ' + ' AND '.join(wheres)) if wheres else ''
    rows = query(f"""
        SELECT l.*, u.display_name AS owner_name,
               a.name AS converted_account_name,
               (SELECT MAX(la.created_at) FROM lead_activities la
                WHERE la.lead_id = l.id) AS last_activity_date
        FROM leads l
        LEFT JOIN users u ON u.id = l.owner_user_id
        LEFT JOIN accounts a ON a.id = l.converted_account_id
        {where_sql}
        ORDER BY l.created_at DESC
    """, params)
    return jsonify(rows)

@app.route('/api/leads', methods=['POST'])
@require_auth
def create_lead():
    d = request.json
    conn = get_db()
    try:
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        lead_no = next_lead_no(cur)
        cur.execute("""
            INSERT INTO leads (lead_no, company_name, contact_name, contact_phone,
                               contact_email, contact_position, branch,
                               stage, owner_user_id, description, created_by)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) RETURNING id
        """, (lead_no, d['company_name'], d.get('contact_name'), d.get('contact_phone'),
              d.get('contact_email'), d.get('contact_position'), d.get('branch'),
              d.get('stage', 'Cold Call/Email'),
              d.get('owner_user_id') or None, d.get('description'), g.user['display_name']))
        new_id = cur.fetchone()['id']
        conn.commit()
        conn.close()
        return jsonify({'id': new_id, 'lead_no': lead_no}), 201
    except Exception as e:
        conn.rollback(); conn.close()
        print(f"[create_lead] error: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/leads/<int:lid>', methods=['GET'])
@require_auth
def get_lead(lid):
    row = query("""
        SELECT l.*, u.display_name AS owner_name,
               a.name AS converted_account_name
        FROM leads l
        LEFT JOIN users u ON u.id = l.owner_user_id
        LEFT JOIN accounts a ON a.id = l.converted_account_id
        WHERE l.id = %s
    """, (lid,), one=True)
    if not row:
        return jsonify({'error': 'Not found'}), 404
    return jsonify(row)

@app.route('/api/leads/<int:lid>', methods=['PUT'])
@require_auth
def update_lead(lid):
    d = request.json
    new_stage = d.get('stage')
    conn = get_db()
    try:
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        # Capture lost_from_stage when transitioning to Closed Lost
        lost_from = None
        if new_stage == 'Closed Lost':
            cur.execute("SELECT stage FROM leads WHERE id=%s", (lid,))
            row = cur.fetchone()
            if row and row['stage'] not in ('Closed Lost', 'Closed Win'):
                lost_from = row['stage']
        now = "to_char(now(),'YYYY-MM-DD\"T\"HH24:MI:SS')"
        if lost_from is not None:
            cur.execute(f"""
                UPDATE leads SET company_name=%s, contact_name=%s, contact_phone=%s,
                    contact_email=%s, contact_position=%s, branch=%s,
                    stage=%s, owner_user_id=%s, description=%s,
                    lost_from_stage=%s, updated_at=({now})
                WHERE id=%s
            """, (d['company_name'], d.get('contact_name'), d.get('contact_phone'),
                  d.get('contact_email'), d.get('contact_position'), d.get('branch'),
                  new_stage, d.get('owner_user_id'),
                  d.get('description'), lost_from, lid))
        else:
            cur.execute(f"""
                UPDATE leads SET company_name=%s, contact_name=%s, contact_phone=%s,
                    contact_email=%s, contact_position=%s, branch=%s,
                    stage=%s, owner_user_id=%s, description=%s,
                    updated_at=({now})
                WHERE id=%s
            """, (d['company_name'], d.get('contact_name'), d.get('contact_phone'),
                  d.get('contact_email'), d.get('contact_position'), d.get('branch'),
                  new_stage, d.get('owner_user_id'),
                  d.get('description'), lid))
        conn.commit()
        conn.close()
        return jsonify({'ok': True})
    except Exception as e:
        conn.rollback()
        conn.close()
        return jsonify({'error': str(e)}), 500

@app.route('/api/leads/<int:lid>', methods=['DELETE'])
@require_auth
def delete_lead(lid):
    mutate("DELETE FROM lead_activities WHERE lead_id=%s", (lid,))
    mutate("DELETE FROM leads WHERE id=%s", (lid,))
    return jsonify({'ok': True})

@app.route('/api/leads/<int:lid>/convert', methods=['POST'])
@require_auth
def convert_lead(lid):
    d = request.json  # { account_id: int|null }
    account_id = d.get('account_id') or None
    now_expr = "to_char(now(), 'YYYY-MM-DD\"T\"HH24:MI:SS')"
    mutate(f"""
        UPDATE leads SET stage='Closed Win', converted_at=({now_expr}),
            converted_account_id=%s, updated_at=({now_expr})
        WHERE id=%s
    """, (account_id, lid))
    return jsonify({'ok': True})

@app.route('/api/leads/<int:lid>/activities', methods=['GET'])
@require_auth
def get_lead_activities(lid):
    rows = query("""
        SELECT * FROM lead_activities WHERE lead_id=%s ORDER BY created_at DESC
    """, (lid,))
    return jsonify(rows)

@app.route('/api/leads/<int:lid>/activities', methods=['POST'])
@require_auth
def add_lead_activity(lid):
    d = request.json
    mutate("""
        INSERT INTO lead_activities (lead_id, activity_type, subject, description, created_by)
        VALUES (%s,%s,%s,%s,%s)
    """, (lid, d['activity_type'], d.get('subject'), d.get('description'),
          g.user['display_name']))
    # bump updated_at on lead
    mutate("UPDATE leads SET updated_at=to_char(now(),'YYYY-MM-DD\"T\"HH24:MI:SS') WHERE id=%s", (lid,))
    return jsonify({'ok': True}), 201

@app.route('/api/sales-dashboard', methods=['GET'])
@require_auth
def sales_dashboard():
    period     = request.args.get('period', '30d')
    start_date = request.args.get('start_date', '')
    end_date   = request.args.get('end_date', '')

    now = datetime.now()
    if period == 'today':
        since = now.strftime('%Y-%m-%dT00:00:00')
        until = now.strftime('%Y-%m-%dT23:59:59')
    elif period == 'yesterday':
        y = (now - timedelta(days=1))
        since = y.strftime('%Y-%m-%dT00:00:00')
        until = y.strftime('%Y-%m-%dT23:59:59')
    elif period == 'this_month':
        since = now.replace(day=1).strftime('%Y-%m-%dT00:00:00')
        until = now.strftime('%Y-%m-%dT23:59:59')
    elif period == 'last_month':
        first_this = now.replace(day=1)
        last_prev  = first_this - timedelta(days=1)
        since = last_prev.replace(day=1).strftime('%Y-%m-%dT00:00:00')
        until = last_prev.strftime('%Y-%m-%dT23:59:59')
    elif period == 'custom' and start_date and end_date:
        since = start_date + 'T00:00:00'
        until = end_date + 'T23:59:59'
    else:
        since = now.replace(day=1).strftime('%Y-%m-%dT00:00:00')
        until = now.strftime('%Y-%m-%dT23:59:59')

    week_ago = (now - timedelta(days=7)).strftime('%Y-%m-%dT00:00:00')
    is_sales = g.user['role'] == 'sales'
    uid      = g.user['user_id']

    own = " AND owner_user_id=%s"

    total = query(
        "SELECT COUNT(*) AS cnt FROM leads" + (" WHERE owner_user_id=%s" if is_sales else ""),
        ([uid] if is_sales else []), one=True)

    by_stage = query(
        "SELECT stage, COUNT(*) AS cnt FROM leads" +
        (" WHERE owner_user_id=%s" if is_sales else "") +
        " GROUP BY stage",
        ([uid] if is_sales else []))

    wins = query(
        "SELECT COUNT(*) AS cnt FROM leads WHERE stage='Closed Win' AND updated_at>=%s AND updated_at<=%s" +
        (own if is_sales else ""),
        ([since, until, uid] if is_sales else [since, until]), one=True)

    if is_sales:
        acts = query(
            "SELECT COUNT(*) AS cnt FROM lead_activities la"
            " JOIN leads l ON l.id=la.lead_id"
            " WHERE la.created_at>=%s AND l.owner_user_id=%s",
            [week_ago, uid], one=True)
    else:
        acts = query(
            "SELECT COUNT(*) AS cnt FROM lead_activities WHERE created_at>=%s",
            [week_ago], one=True)

    return jsonify({
        'total_leads':    total['cnt']   if total else 0,
        'by_stage':       by_stage,
        'closed_wins':    wins['cnt']    if wins  else 0,
        'activities_week': acts['cnt']   if acts  else 0,
        'period': period, 'since': since, 'until': until,
    })


@app.route('/api/revenue/owners', methods=['GET'])
@require_auth
def revenue_owners():
  try:
    now   = datetime.now()
    start = request.args.get('start') or now.replace(day=1).strftime('%Y-%m-%d')
    end   = request.args.get('end')   or now.strftime('%Y-%m-%d')

    crm_rows = query("""
        SELECT a.owner, a.id AS account_id, o.erp_outlet_id
        FROM accounts a
        JOIN outlets o ON o.account_id = a.id
        WHERE a.owner IS NOT NULL AND trim(a.owner) != ''
          AND o.erp_outlet_id IS NOT NULL AND trim(o.erp_outlet_id) != ''
    """)
    if not crm_rows:
        acct_cnt  = query("SELECT COUNT(*) AS c FROM accounts WHERE owner IS NOT NULL AND trim(owner)!=''", one=True)
        outl_cnt  = query("SELECT COUNT(*) AS c FROM outlets WHERE erp_outlet_id IS NOT NULL AND trim(erp_outlet_id)!=''", one=True)
        return jsonify({'rows': [], 'total_revenue': 0, 'start': start, 'end': end})

    all_outlet_ids = list({r['erp_outlet_id'] for r in crm_rows})
    erp_rows = []
    erp_error = None
    try:
        erp_rows = erp_query("""
            SELECT outlet_id::text AS outlet_id,
                   COALESCE(SUM(total_sales), 0) AS revenue,
                   COALESCE(SUM(qty), 0)          AS volume,
                   COUNT(DISTINCT invoice_number) AS orders
            FROM sourcing_erp_order_items
            WHERE doc_date >= %s AND doc_date <= %s
              AND outlet_id::text = ANY(%s)
            GROUP BY outlet_id
        """, (start, end, [str(x) for x in all_outlet_ids]))
    except Exception as ex:
        erp_error = str(ex)
    erp_map = {r['outlet_id']: r for r in (erp_rows or [])}

    owner_data = {}
    for row in crm_rows:
        owner = row['owner']
        erp   = erp_map.get(row['erp_outlet_id'], {})
        if owner not in owner_data:
            owner_data[owner] = {'revenue': 0, 'volume': 0, 'orders': 0, 'account_ids': set()}
        owner_data[owner]['revenue']     += float(erp.get('revenue') or 0)
        owner_data[owner]['volume']      += float(erp.get('volume')  or 0)
        owner_data[owner]['orders']      += int(erp.get('orders')    or 0)
        owner_data[owner]['account_ids'].add(row['account_id'])

    total_revenue = sum(d['revenue'] for d in owner_data.values())
    result = []
    for owner, d in owner_data.items():
        rev  = d['revenue']
        ords = d['orders']
        result.append({
            'owner':     owner,
            'revenue':   round(rev, 2),
            'volume':    round(d['volume'], 2),
            'orders':    ords,
            'accounts':  len(d['account_ids']),
            'avg_order': round(rev / ords, 2) if ords > 0 else 0,
            'share':     round(rev / total_revenue * 100, 1) if total_revenue > 0 else 0,
        })
    result.sort(key=lambda x: -x['revenue'])
    return jsonify({'rows': result, 'total_revenue': round(total_revenue, 2), 'start': start, 'end': end})
  except Exception as e:
    return jsonify({'rows': [], 'total_revenue': 0}), 200

@app.route('/api/revenue/owner-accounts', methods=['GET'])
@require_auth
def revenue_owner_accounts():
    owner = request.args.get('owner', '')
    now   = datetime.now()
    start = request.args.get('start') or (now - timedelta(days=30)).strftime('%Y-%m-%d')
    end   = request.args.get('end')   or now.strftime('%Y-%m-%d')
    if not owner:
        return jsonify([])

    crm_rows = query("""
        SELECT a.id AS account_id, a.name AS account_name, o.erp_outlet_id
        FROM accounts a
        JOIN outlets o ON o.account_id = a.id
        WHERE a.owner = %s
          AND o.erp_outlet_id IS NOT NULL AND trim(o.erp_outlet_id) != ''
    """, (owner,))
    if not crm_rows:
        return jsonify([])

    outlet_ids = [str(r['erp_outlet_id']) for r in crm_rows]
    erp_rows = erp_query("""
        SELECT outlet_id::text AS outlet_id,
               COALESCE(SUM(total_sales), 0) AS revenue,
               COALESCE(SUM(qty), 0)          AS volume,
               COUNT(DISTINCT invoice_number) AS orders
        FROM sourcing_erp_order_items
        WHERE delivery_started_at::date >= %s AND delivery_started_at::date <= %s
          AND outlet_id::text = ANY(%s)
        GROUP BY outlet_id
    """, (start, end, outlet_ids))
    erp_map = {r['outlet_id']: r for r in (erp_rows or [])}

    acc_data = {}
    for row in crm_rows:
        aid = row['account_id']
        if aid not in acc_data:
            acc_data[aid] = {'account_name': row['account_name'], 'revenue': 0, 'volume': 0, 'orders': 0, 'outlet_count': 0}
        acc_data[aid]['outlet_count'] += 1
        erp = erp_map.get(row['erp_outlet_id'], {})
        acc_data[aid]['revenue'] += float(erp.get('revenue') or 0)
        acc_data[aid]['volume']  += float(erp.get('volume')  or 0)
        acc_data[aid]['orders']  += int(erp.get('orders')    or 0)

    result = []
    for aid, d in acc_data.items():
        rev  = d['revenue']
        ords = d['orders']
        result.append({
            'account_id': aid, 'account_name': d['account_name'],
            'outlet_count': d['outlet_count'],
            'revenue':  round(rev, 2),
            'volume':   round(d['volume'], 2),
            'orders':   ords,
            'avg_order': round(rev / ords, 2) if ords > 0 else 0,
        })
    result.sort(key=lambda x: -x['revenue'])
    return jsonify(result)

@app.route('/api/sales-dashboard/reps', methods=['GET'])
@require_auth
def sales_dashboard_reps():
    if g.user['role'] not in ('admin',):
        return jsonify({'error': 'Forbidden'}), 403
    now      = datetime.now()
    week_ago = (now - timedelta(days=7)).strftime('%Y-%m-%dT00:00:00')
    rows = query("""
        SELECT
            u.id   AS user_id,
            u.display_name,
            COUNT(DISTINCT l.id)                                           AS total,
            COUNT(DISTINCT CASE WHEN l.stage='Cold Call/Email' THEN l.id END) AS cold_call,
            COUNT(DISTINCT CASE WHEN l.stage='Meeting'         THEN l.id END) AS meeting,
            COUNT(DISTINCT CASE WHEN l.stage='Follow Up'       THEN l.id END) AS follow_up,
            COUNT(DISTINCT CASE WHEN l.stage='Closed Win'      THEN l.id END) AS won,
            COUNT(DISTINCT CASE WHEN l.stage='Closed Lost'     THEN l.id END) AS lost,
            COUNT(CASE WHEN la.created_at >= %s              THEN la.id END) AS activities_week,
            MAX(la.created_at)                                             AS last_active
        FROM users u
        LEFT JOIN leads l  ON l.owner_user_id = u.id
        LEFT JOIN lead_activities la ON la.lead_id = l.id
        WHERE u.role IN ('sales','admin') AND u.status = 'active'
        GROUP BY u.id, u.display_name
        ORDER BY u.display_name
    """, (week_ago,))
    return jsonify(rows)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
