import os
import json
from datetime import datetime, date
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import openpyxl
import psycopg2
import psycopg2.extras

app = Flask(__name__, static_folder='static')
CORS(app)

DATABASE_URL = os.environ.get('DATABASE_URL', '')

# ---------------------------------------------------------------------------
# DB helpers
# ---------------------------------------------------------------------------

def get_db():
    conn = psycopg2.connect(DATABASE_URL)
    return conn


def query(sql, params=(), one=False):
    # Convert ? placeholders to %s for psycopg2
    sql = sql.replace('?', '%s')
    conn = get_db()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    cur.execute(sql, params)
    rows = cur.fetchall()
    conn.close()
    result = [dict(r) for r in rows]
    return result[0] if one and result else (None if one else result)


def mutate(sql, params=()):
    sql = sql.replace('?', '%s')
    # For INSERT ... RETURNING id
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
# DB init
# ---------------------------------------------------------------------------

SCHEMA = """
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
    erp_outlet_id TEXT,
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
    delivery_started_at TEXT,
    delivery_finished_at TEXT,
    loaded_at TEXT,
    FOREIGN KEY (outlet_id) REFERENCES outlets(id)
);

CREATE TABLE IF NOT EXISTS tickets (
    id SERIAL PRIMARY KEY,
    ticket_no TEXT UNIQUE,
    outlet_id INTEGER,
    invoice_number TEXT,
    sku_code TEXT,
    product_name TEXT,
    case_type TEXT,
    fault_category TEXT,
    fault_team TEXT,
    priority TEXT DEFAULT 'Medium',
    status TEXT DEFAULT 'open',
    current_team TEXT DEFAULT 'CX',
    workflow_step INTEGER DEFAULT 0,
    workflow_branch TEXT,
    assigned_employee_id INTEGER,
    description TEXT,
    created_by TEXT,
    created_at TEXT DEFAULT (to_char(now(), 'YYYY-MM-DD"T"HH24:MI:SS')),
    resolved_at TEXT,
    FOREIGN KEY (outlet_id) REFERENCES outlets(id),
    FOREIGN KEY (assigned_employee_id) REFERENCES employees(id)
);

CREATE TABLE IF NOT EXISTS ticket_workflow_log (
    id SERIAL PRIMARY KEY,
    ticket_id INTEGER NOT NULL,
    from_team TEXT,
    to_team TEXT,
    action TEXT,
    note TEXT,
    employee_id INTEGER,
    created_by TEXT,
    created_at TEXT DEFAULT (to_char(now(), 'YYYY-MM-DD"T"HH24:MI:SS')),
    FOREIGN KEY (ticket_id) REFERENCES tickets(id)
);

CREATE TABLE IF NOT EXISTS ticket_comments (
    id SERIAL PRIMARY KEY,
    ticket_id INTEGER NOT NULL,
    comment TEXT,
    created_by TEXT,
    created_at TEXT DEFAULT (to_char(now(), 'YYYY-MM-DD"T"HH24:MI:SS')),
    FOREIGN KEY (ticket_id) REFERENCES tickets(id)
);
"""

WORKFLOWS = {
    'Claim': [
        {'team': 'CX', 'label': 'รับเรื่อง + ถามลูกค้า', 'branches': ['ส่งใหม่', 'ตัดออก']},
        {'team': 'BRANCH', 'label': 'เลือกแนวทาง', 'is_branch': True},
        {'team': 'Merchandise', 'label': 'สั่งสินค้าใหม่', 'branch': 'ส่งใหม่'},
        {'team': 'Inbound/QC', 'label': 'ตรวจรับสินค้า', 'branch': 'ส่งใหม่'},
        {'team': 'Outbound/Logistics', 'label': 'เรียกรถ + จัดส่ง', 'branch': 'ส่งใหม่'},
        {'team': 'Sales Co', 'label': 'ตัด SKU ออก', 'branch': 'ตัดออก'},
        {'team': 'CX', 'label': 'ปิดเคส'},
    ],
    'น้ำหนักไม่ครบ': [
        {'team': 'CX', 'label': 'รับเรื่อง + ถามลูกค้า', 'branches': ['ส่งเพิ่ม', 'ตัด Invoice', 'รอรอบถัดไป']},
        {'team': 'Outbound/Logistics', 'label': 'จัดส่งเพิ่ม', 'branch': 'ส่งเพิ่ม'},
        {'team': 'Sales Co', 'label': 'ตัด Invoice', 'branch': 'ตัด Invoice'},
        {'team': 'CX', 'label': 'Note + รอรอบถัดไป', 'branch': 'รอรอบถัดไป'},
        {'team': 'CX', 'label': 'ปิดเคส'},
    ],
    'ตกหล่น': [
        {'team': 'CX', 'label': 'รับเรื่อง + ถามลูกค้า', 'branches': ['ส่งทันที', 'รอรอบถัดไป']},
        {'team': 'Outbound/Logistics', 'label': 'จัดส่งทันที', 'branch': 'ส่งทันที'},
        {'team': 'CX', 'label': 'Note + รอรอบถัดไป', 'branch': 'รอรอบถัดไป'},
        {'team': 'CX', 'label': 'ปิดเคส'},
    ],
    'คุณภาพไม่ผ่าน/ผิดสเปค': [
        {'team': 'CX', 'label': 'รับเรื่อง + ถามลูกค้า'},
        {'team': 'CX', 'label': 'ตรวจ master spec', 'branches': ['master ถูก', 'master ผิด']},
        {'team': 'Merchandise', 'label': 'หาสินค้าใหม่', 'branch': 'master ถูก'},
        {'team': 'Inbound/QC', 'label': 'ตรวจรับสินค้า', 'branch': 'master ถูก'},
        {'team': 'Merchandise', 'label': 'แก้ไข master spec', 'branch': 'master ผิด'},
        {'team': 'CX', 'label': 'ปิดเคส'},
    ],
    'Sales Co Error': [
        {'team': 'CX', 'label': 'รับเรื่อง'},
        {'team': 'Sales Co', 'label': 'ดำเนินการ (ตัด SKU / แนบรูป / แก้ order)'},
        {'team': 'CX', 'label': 'ปิดเคส'},
    ],
}


def init_db():
    conn = get_db()
    cur = conn.cursor()
    for stmt in SCHEMA.strip().split(';'):
        stmt = stmt.strip()
        if stmt:
            cur.execute(stmt)
    conn.commit()
    conn.close()
    print("Database initialized.")


# ---------------------------------------------------------------------------
# Static files
# ---------------------------------------------------------------------------

@app.route('/')
def index():
    return send_from_directory('static', 'index.html')


# ---------------------------------------------------------------------------
# Accounts
# ---------------------------------------------------------------------------

@app.route('/api/accounts', methods=['GET'])
def get_accounts():
    rows = query("""
        SELECT a.*,
               COUNT(DISTINCT o.id) AS outlet_count,
               COUNT(DISTINCT CASE WHEN t.status NOT IN ('resolved','closed') THEN t.id END) AS open_tickets
        FROM accounts a
        LEFT JOIN outlets o ON o.account_id = a.id
        LEFT JOIN tickets t ON t.outlet_id = o.id
        GROUP BY a.id ORDER BY a.name
    """)
    return jsonify(rows)


@app.route('/api/accounts', methods=['POST'])
def create_account():
    d = request.json
    id_ = mutate("INSERT INTO accounts (name, owner, status) VALUES (%s,%s,%s) RETURNING id",
                 (d['name'], d.get('owner', ''), d.get('status', 'active')))
    return jsonify({'id': id_}), 201


@app.route('/api/accounts/<int:aid>', methods=['GET'])
def get_account(aid):
    row = query("SELECT * FROM accounts WHERE id=%s", (aid,), one=True)
    if not row:
        return jsonify({'error': 'Not found'}), 404
    return jsonify(row)


@app.route('/api/accounts/<int:aid>', methods=['PUT'])
def update_account(aid):
    d = request.json
    mutate("UPDATE accounts SET name=%s, owner=%s, status=%s WHERE id=%s",
           (d['name'], d.get('owner', ''), d.get('status', 'active'), aid))
    return jsonify({'ok': True})


@app.route('/api/accounts/<int:aid>/outlets', methods=['GET'])
def get_account_outlets(aid):
    rows = query("""
        SELECT o.*,
               COUNT(DISTINCT CASE WHEN t.status NOT IN ('resolved','closed') THEN t.id END) AS open_tickets
        FROM outlets o
        LEFT JOIN tickets t ON t.outlet_id = o.id
        WHERE o.account_id=%s GROUP BY o.id ORDER BY o.name
    """, (aid,))
    return jsonify(rows)


@app.route('/api/accounts/<int:aid>/notes', methods=['GET'])
def get_account_notes(aid):
    return jsonify(query("SELECT * FROM account_notes WHERE account_id=%s ORDER BY created_at DESC", (aid,)))


@app.route('/api/accounts/<int:aid>/notes', methods=['POST'])
def add_account_note(aid):
    d = request.json
    id_ = mutate("INSERT INTO account_notes (account_id, note, created_by) VALUES (%s,%s,%s) RETURNING id",
                 (aid, d['note'], d.get('created_by', 'CX')))
    return jsonify({'id': id_}), 201


# ---------------------------------------------------------------------------
# Outlets
# ---------------------------------------------------------------------------

@app.route('/api/outlets', methods=['GET'])
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
def get_outlet_invoices(oid):
    rows = query("""
        SELECT invoice_number, doc_date, COUNT(*) AS sku_count, SUM(total_sales) AS total_sales
        FROM orders WHERE outlet_id=%s
        GROUP BY invoice_number, doc_date ORDER BY doc_date DESC
    """, (oid,))
    return jsonify(rows)


@app.route('/api/outlets/<int:oid>/tickets', methods=['GET'])
def get_outlet_tickets(oid):
    rows = query("SELECT * FROM tickets WHERE outlet_id=%s ORDER BY created_at DESC", (oid,))
    return jsonify(rows)


@app.route('/api/outlets/<int:oid>/notes', methods=['GET'])
def get_outlet_notes(oid):
    outlet = query("SELECT account_id FROM outlets WHERE id=%s", (oid,), one=True)
    if not outlet:
        return jsonify([])
    return jsonify(query("SELECT * FROM account_notes WHERE account_id=%s ORDER BY created_at DESC",
                         (outlet['account_id'],)))


# ---------------------------------------------------------------------------
# Invoice / SKU lookup
# ---------------------------------------------------------------------------

@app.route('/api/invoices/<invoice_number>/skus', methods=['GET'])
def get_invoice_skus(invoice_number):
    rows = query("""
        SELECT DISTINCT sku_code, product_name, qty, unit
        FROM orders WHERE invoice_number=%s
        ORDER BY product_name
    """, (invoice_number,))
    return jsonify(rows)


# ---------------------------------------------------------------------------
# Suppliers
# ---------------------------------------------------------------------------

@app.route('/api/suppliers', methods=['GET'])
def get_suppliers():
    return jsonify(query("SELECT * FROM suppliers ORDER BY name"))


@app.route('/api/suppliers', methods=['POST'])
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
def get_supplier(sid):
    row = query("SELECT * FROM suppliers WHERE id=%s", (sid,), one=True)
    if not row:
        return jsonify({'error': 'Not found'}), 404
    return jsonify(row)


@app.route('/api/suppliers/<int:sid>', methods=['PUT'])
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
def get_supplier_notes(sid):
    return jsonify(query("SELECT * FROM supplier_notes WHERE supplier_id=%s ORDER BY created_at DESC", (sid,)))


@app.route('/api/suppliers/<int:sid>/notes', methods=['POST'])
def add_supplier_note(sid):
    d = request.json
    id_ = mutate("INSERT INTO supplier_notes (supplier_id, note, note_type, created_by) VALUES (%s,%s,%s,%s) RETURNING id",
                 (sid, d['note'], d.get('note_type','general'), d.get('created_by','CX')))
    return jsonify({'id': id_}), 201


# ---------------------------------------------------------------------------
# Employees
# ---------------------------------------------------------------------------

@app.route('/api/employees', methods=['GET'])
def get_employees():
    team = request.args.get('team')
    if team:
        return jsonify(query("SELECT * FROM employees WHERE team=%s AND status='active' ORDER BY name", (team,)))
    return jsonify(query("SELECT * FROM employees ORDER BY team, name"))


@app.route('/api/employees', methods=['POST'])
def create_employee():
    d = request.json
    id_ = mutate("""INSERT INTO employees (name, employee_code, department, team, role, phone, email, status)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s) RETURNING id""",
        (d['name'], d.get('employee_code',''), d.get('department',''),
         d.get('team',''), d.get('role',''), d.get('phone',''),
         d.get('email',''), d.get('status','active')))
    return jsonify({'id': id_}), 201


@app.route('/api/employees/<int:eid>', methods=['GET'])
def get_employee(eid):
    row = query("SELECT * FROM employees WHERE id=%s", (eid,), one=True)
    if not row:
        return jsonify({'error': 'Not found'}), 404
    return jsonify(row)


@app.route('/api/employees/<int:eid>', methods=['PUT'])
def update_employee(eid):
    d = request.json
    mutate("""UPDATE employees SET name=%s, employee_code=%s, department=%s, team=%s,
        role=%s, phone=%s, email=%s, status=%s WHERE id=%s""",
        (d['name'], d.get('employee_code',''), d.get('department',''),
         d.get('team',''), d.get('role',''), d.get('phone',''),
         d.get('email',''), d.get('status','active'), eid))
    return jsonify({'ok': True})


# ---------------------------------------------------------------------------
# Tickets
# ---------------------------------------------------------------------------

def next_ticket_no():
    row = query("SELECT COUNT(*) AS cnt FROM tickets", one=True)
    return f"TK-{(row['cnt'] or 0) + 1:04d}"


def get_workflow_steps(case_type, branch=None):
    steps = WORKFLOWS.get(case_type, [])
    if branch:
        filtered = [s for s in steps if 'branch' not in s or s.get('branch') == branch or s.get('is_branch')]
        return filtered
    return steps


@app.route('/api/tickets', methods=['GET'])
def get_tickets():
    team = request.args.get('team')
    status = request.args.get('status')
    outlet_id = request.args.get('outlet_id')
    params = []
    wheres = []
    if team and team != 'Management':
        wheres.append("t.current_team=%s")
        params.append(team)
    if status:
        wheres.append("t.status=%s")
        params.append(status)
    if outlet_id:
        wheres.append("t.outlet_id=%s")
        params.append(outlet_id)
    where_clause = ("WHERE " + " AND ".join(wheres)) if wheres else ""
    rows = query(f"""
        SELECT t.*, o.name AS outlet_name, a.name AS account_name,
               e.name AS assigned_employee_name
        FROM tickets t
        LEFT JOIN outlets o ON o.id=t.outlet_id
        LEFT JOIN accounts a ON a.id=o.account_id
        LEFT JOIN employees e ON e.id=t.assigned_employee_id
        {where_clause}
        ORDER BY t.created_at DESC
    """, params)
    return jsonify(rows)


@app.route('/api/tickets', methods=['POST'])
def create_ticket():
    d = request.json
    ticket_no = next_ticket_no()
    case_type = d.get('case_type', '')
    workflow = WORKFLOWS.get(case_type, [])
    first_team = workflow[0]['team'] if workflow else 'CX'

    id_ = mutate("""INSERT INTO tickets
        (ticket_no, outlet_id, invoice_number, sku_code, product_name,
         case_type, fault_category, fault_team, priority, status,
         current_team, workflow_step, description, created_by)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) RETURNING id""",
        (ticket_no, d.get('outlet_id'), d.get('invoice_number'),
         d.get('sku_code'), d.get('product_name'),
         case_type, d.get('fault_category',''), d.get('fault_team',''),
         d.get('priority','Medium'), 'open',
         first_team, 0,
         d.get('description',''), d.get('created_by','CX')))

    mutate("""INSERT INTO ticket_workflow_log
        (ticket_id, from_team, to_team, action, note, created_by)
        VALUES (%s,%s,%s,%s,%s,%s)""",
        (id_, '', first_team, 'สร้างเคส', d.get('description',''), d.get('created_by','CX')))

    return jsonify({'id': id_, 'ticket_no': ticket_no}), 201


@app.route('/api/tickets/<int:tid>', methods=['GET'])
def get_ticket(tid):
    row = query("""
        SELECT t.*, o.name AS outlet_name, a.name AS account_name,
               e.name AS assigned_employee_name
        FROM tickets t
        LEFT JOIN outlets o ON o.id=t.outlet_id
        LEFT JOIN accounts a ON a.id=o.account_id
        LEFT JOIN employees e ON e.id=t.assigned_employee_id
        WHERE t.id=%s
    """, (tid,), one=True)
    if not row:
        return jsonify({'error': 'Not found'}), 404
    return jsonify(row)


@app.route('/api/tickets/<int:tid>', methods=['PUT'])
def update_ticket(tid):
    d = request.json
    mutate("""UPDATE tickets SET priority=%s, status=%s, fault_category=%s,
        fault_team=%s, assigned_employee_id=%s, description=%s WHERE id=%s""",
        (d.get('priority','Medium'), d.get('status','open'),
         d.get('fault_category',''), d.get('fault_team',''),
         d.get('assigned_employee_id'), d.get('description',''), tid))
    return jsonify({'ok': True})


@app.route('/api/tickets/<int:tid>/advance', methods=['POST'])
def advance_ticket(tid):
    d = request.json
    ticket = query("SELECT * FROM tickets WHERE id=%s", (tid,), one=True)
    if not ticket:
        return jsonify({'error': 'Not found'}), 404

    action = d.get('action', '')
    note = d.get('note', '')
    created_by = d.get('created_by', 'CX')
    branch = d.get('branch', ticket.get('workflow_branch'))
    employee_id = d.get('employee_id')

    case_type = ticket['case_type']
    workflow = WORKFLOWS.get(case_type, [])
    current_step = ticket['workflow_step']
    new_step = current_step + 1

    next_team = ticket['current_team']
    new_status = ticket['status']

    effective_steps = [s for s in workflow if 'branch' not in s or s.get('branch') == branch or s.get('is_branch')]
    if new_step < len(effective_steps):
        next_step_def = effective_steps[new_step]
        next_team = next_step_def['team']
        if next_team == 'BRANCH':
            branch = d.get('branch', '')
            next_team = ticket['current_team']
    else:
        new_status = 'resolved'
        next_team = 'CX'

    resolved_at = datetime.now().isoformat() if new_status == 'resolved' else None

    mutate("""UPDATE tickets SET current_team=%s, workflow_step=%s, workflow_branch=%s,
        status=%s, assigned_employee_id=%s, resolved_at=%s WHERE id=%s""",
        (next_team, new_step, branch, new_status, employee_id, resolved_at, tid))

    mutate("""INSERT INTO ticket_workflow_log
        (ticket_id, from_team, to_team, action, note, employee_id, created_by)
        VALUES (%s,%s,%s,%s,%s,%s,%s)""",
        (tid, ticket['current_team'], next_team, action, note, employee_id, created_by))

    return jsonify({'ok': True, 'next_team': next_team, 'status': new_status})


@app.route('/api/tickets/<int:tid>/acknowledge', methods=['POST'])
def acknowledge_ticket(tid):
    d = request.json
    employee_id = d.get('employee_id')
    created_by = d.get('created_by', '')
    ticket = query("SELECT * FROM tickets WHERE id=%s", (tid,), one=True)
    if not ticket:
        return jsonify({'error': 'Not found'}), 404
    mutate("UPDATE tickets SET assigned_employee_id=%s, status='in_progress' WHERE id=%s",
           (employee_id, tid))
    mutate("""INSERT INTO ticket_workflow_log
        (ticket_id, from_team, to_team, action, note, employee_id, created_by)
        VALUES (%s,%s,%s,%s,%s,%s,%s)""",
        (tid, ticket['current_team'], ticket['current_team'], 'Acknowledge', '', employee_id, created_by))
    return jsonify({'ok': True})


@app.route('/api/tickets/<int:tid>/close', methods=['POST'])
def close_ticket(tid):
    d = request.json
    mutate("UPDATE tickets SET status='closed', resolved_at=%s WHERE id=%s",
           (datetime.now().isoformat(), tid))
    mutate("""INSERT INTO ticket_workflow_log
        (ticket_id, from_team, to_team, action, note, created_by)
        VALUES (%s,%s,%s,%s,%s,%s)""",
        (tid, 'CX', '', 'ปิดเคส', d.get('note',''), d.get('created_by','CX')))
    return jsonify({'ok': True})


@app.route('/api/tickets/<int:tid>/comments', methods=['GET'])
def get_ticket_comments(tid):
    return jsonify(query("SELECT * FROM ticket_comments WHERE ticket_id=%s ORDER BY created_at", (tid,)))


@app.route('/api/tickets/<int:tid>/comments', methods=['POST'])
def add_ticket_comment(tid):
    d = request.json
    id_ = mutate("INSERT INTO ticket_comments (ticket_id, comment, created_by) VALUES (%s,%s,%s) RETURNING id",
                 (tid, d['comment'], d.get('created_by','CX')))
    return jsonify({'id': id_}), 201


@app.route('/api/tickets/<int:tid>/log', methods=['GET'])
def get_ticket_log(tid):
    rows = query("""
        SELECT l.*, e.name AS employee_name
        FROM ticket_workflow_log l
        LEFT JOIN employees e ON e.id=l.employee_id
        WHERE l.ticket_id=%s ORDER BY l.created_at
    """, (tid,))
    return jsonify(rows)


@app.route('/api/workflows', methods=['GET'])
def get_workflows():
    return jsonify(WORKFLOWS)


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
def dashboard_overview():
    extra, params = date_filter_sql()
    total = query(f"SELECT COUNT(*) AS cnt FROM tickets t WHERE 1=1 {extra}", params, one=True)
    by_status = query(f"SELECT status, COUNT(*) AS cnt FROM tickets t WHERE 1=1 {extra} GROUP BY status", params)
    by_case = query(f"SELECT case_type, COUNT(*) AS cnt FROM tickets t WHERE 1=1 {extra} GROUP BY case_type ORDER BY cnt DESC", params)
    by_fault = query(f"SELECT fault_team, COUNT(*) AS cnt FROM tickets t WHERE 1=1 {extra} GROUP BY fault_team ORDER BY cnt DESC", params)
    by_priority = query(f"SELECT priority, COUNT(*) AS cnt FROM tickets t WHERE 1=1 {extra} GROUP BY priority", params)
    return jsonify({
        'total': total['cnt'] if total else 0,
        'by_status': by_status,
        'by_case_type': by_case,
        'by_fault_team': by_fault,
        'by_priority': by_priority,
    })


@app.route('/api/dashboard/sku-problems', methods=['GET'])
def dashboard_sku():
    extra, params = date_filter_sql()
    rows = query(f"""
        SELECT sku_code, product_name, COUNT(*) AS cnt,
               STRING_AGG(DISTINCT case_type, ',') AS case_types,
               STRING_AGG(DISTINCT fault_team, ',') AS fault_teams
        FROM tickets t WHERE sku_code IS NOT NULL AND sku_code != '' {extra}
        GROUP BY sku_code, product_name ORDER BY cnt DESC LIMIT 20
    """, params)
    return jsonify(rows)


@app.route('/api/dashboard/by-employee', methods=['GET'])
def dashboard_employee():
    extra, params = date_filter_sql()
    total_row = query(f"SELECT COUNT(*) AS cnt FROM tickets t WHERE 1=1 {extra}", params, one=True)
    total = total_row['cnt'] if total_row else 1
    rows = query(f"""
        SELECT e.id, e.name, e.team, COUNT(t.id) AS cnt,
               ROUND(COUNT(t.id)*100.0/%s, 1) AS pct,
               STRING_AGG(DISTINCT t.case_type, ',') AS case_types
        FROM employees e
        LEFT JOIN tickets t ON t.assigned_employee_id=e.id {('AND 1=1' + extra) if extra else ''}
        GROUP BY e.id, e.name, e.team ORDER BY cnt DESC
    """, [total] + params)
    return jsonify(rows)


@app.route('/api/dashboard/by-customer', methods=['GET'])
def dashboard_customer():
    extra, params = date_filter_sql()
    rows = query(f"""
        SELECT o.id AS outlet_id, o.name AS outlet_name, a.name AS account_name,
               COUNT(t.id) AS cnt,
               STRING_AGG(DISTINCT t.case_type, ',') AS case_types
        FROM outlets o
        JOIN accounts a ON a.id=o.account_id
        LEFT JOIN tickets t ON t.outlet_id=o.id {('AND 1=1' + extra) if extra else ''}
        GROUP BY o.id, o.name, a.name ORDER BY cnt DESC LIMIT 20
    """, params)
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

    headers = [str(h).strip() if h is not None else '' for h in rows[0]]
    col_idx = {h: i for i, h in enumerate(headers)}

    COL_ALIASES = {
        'order_item_id': ['order_item_id', 'order_iter', 'OrderItemId', 'order_item'],
        'Total_Sales':   ['Total_Sales', 'Total_Sale', 'total_sales', 'total_sale'],
        'outlet_id':     ['outlet_id', 'customer_outlet_id', 'OutletId'],
        'customer_id':   ['customer_id', 'CustomerId', 'erp_customer_id'],
        'account_name':  ['account_name', 'AccountName', 'account'],
        'Customer_Name': ['Customer_Name', 'Customer_name', 'customer_name', 'CustomerName'],
        'Product_Name':  ['Product_Name', 'Product_name', 'product_name', 'ProductName'],
        'sku_category':  ['sku_category', 'sku_categ', 'SKU_Category'],
        'delivery_started_at':  ['delivery_started_at', 'delivery_s', 'delivery_start'],
        'delivery_finished_at': ['delivery_finished_at', 'delivery_f', 'delivery_finish'],
        'invoice_number': ['invoice_number', 'invoice_no', 'InvoiceNumber'],
        'doc_no':        ['doc_no', 'DocNo', 'Doc_No'],
    }

    def get(row, erp_col):
        candidates = COL_ALIASES.get(erp_col, [erp_col])
        for name in candidates:
            idx = col_idx.get(name)
            if idx is not None and idx < len(row):
                return row[idx]
        return None

    inserted = 0
    skipped = 0
    errors = []

    conn = get_db()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    try:
        for row_num, row in enumerate(rows[1:], start=2):
            erp_item_id = safe_str(get(row, 'order_item_id'))
            if not erp_item_id:
                continue

            cur.execute("SELECT id FROM orders WHERE erp_item_id=%s", (erp_item_id,))
            if cur.fetchone():
                skipped += 1
                continue

            account_name = safe_str(get(row, 'account_name')) or 'Unknown'
            outlet_name = safe_str(get(row, 'Customer_Name')) or 'Unknown'
            owner = safe_str(get(row, 'Owner')) or ''
            erp_customer_id = safe_str(get(row, 'customer_id'))
            erp_outlet_id = safe_str(get(row, 'outlet_id'))
            csc_code = safe_str(get(row, 'csc_code'))

            cur.execute("SELECT id FROM accounts WHERE name=%s", (account_name,))
            acc = cur.fetchone()
            if acc:
                account_id = acc['id']
                cur.execute("UPDATE accounts SET owner=%s WHERE id=%s", (owner, account_id))
            else:
                cur.execute("INSERT INTO accounts (name, owner) VALUES (%s,%s) RETURNING id", (account_name, owner))
                account_id = cur.fetchone()['id']

            cur.execute("SELECT id FROM outlets WHERE account_id=%s AND name=%s", (account_id, outlet_name))
            out = cur.fetchone()
            if out:
                outlet_id = out['id']
                cur.execute("UPDATE outlets SET erp_customer_id=%s, erp_outlet_id=%s, csc_code=%s WHERE id=%s",
                            (erp_customer_id, erp_outlet_id, csc_code, outlet_id))
            else:
                cur.execute("""INSERT INTO outlets (account_id, name, erp_customer_id, erp_outlet_id, csc_code)
                    VALUES (%s,%s,%s,%s,%s) RETURNING id""",
                    (account_id, outlet_name, erp_customer_id, erp_outlet_id, csc_code))
                outlet_id = cur.fetchone()['id']

            doc_date = safe_str(get(row, 'Doc_date'))
            if doc_date and 'T' not in doc_date and '-' not in doc_date:
                try:
                    from openpyxl.utils.datetime import from_excel
                    doc_date = str(from_excel(float(doc_date)).date())
                except Exception:
                    pass

            cur.execute("""INSERT INTO orders
                (erp_item_id, order_id, invoice_number, doc_no, doc_date,
                 outlet_id, sku_code, product_name, qty, unit,
                 total_sales, vat_price, is_vat, sku_group, sku_category,
                 sku_type, delivery_started_at, delivery_finished_at, loaded_at)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                ON CONFLICT (erp_item_id) DO NOTHING""",
                (erp_item_id,
                 safe_str(get(row, 'order_id')),
                 safe_str(get(row, 'invoice_number')),
                 safe_str(get(row, 'doc_no')),
                 doc_date,
                 outlet_id,
                 safe_str(get(row, 'SKU')),
                 safe_str(get(row, 'Product_Name')),
                 safe_float(get(row, 'QTY')),
                 safe_str(get(row, 'unit')),
                 safe_float(get(row, 'Total_Sales')),
                 safe_float(get(row, 'vat_price')),
                 1 if get(row, 'is_vat') else 0,
                 safe_str(get(row, 'sku_group')),
                 safe_str(get(row, 'sku_category')),
                 safe_str(get(row, 'sku_type')),
                 safe_str(get(row, 'delivery_started_at')),
                 safe_str(get(row, 'delivery_finished_at')),
                 safe_str(get(row, 'loaded_at'))))
            inserted += 1

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

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
