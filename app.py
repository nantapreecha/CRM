from flask import Flask, jsonify, request, send_from_directory
from flask_cors import CORS
import sqlite3
import os
from datetime import datetime
import openpyxl

app = Flask(__name__, static_folder='static')
CORS(app)

DB_PATH = os.path.join(os.path.dirname(__file__), 'crm.db')

# ─────────────────────────────────────────
# DATABASE SETUP
# ─────────────────────────────────────────

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    conn.executescript('''
        CREATE TABLE IF NOT EXISTS customers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            customer_type TEXT,
            address TEXT,
            phone TEXT,
            email TEXT,
            contact_person TEXT,
            credit_days INTEGER DEFAULT 0,
            status TEXT DEFAULT 'active',
            created_at TEXT DEFAULT (datetime('now','localtime')),
            updated_at TEXT DEFAULT (datetime('now','localtime'))
        );

        CREATE TABLE IF NOT EXISTS customer_notes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            customer_id INTEGER NOT NULL,
            note TEXT NOT NULL,
            created_by TEXT DEFAULT '',
            created_at TEXT DEFAULT (datetime('now','localtime')),
            FOREIGN KEY (customer_id) REFERENCES customers(id)
        );

        CREATE TABLE IF NOT EXISTS suppliers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            supplier_type TEXT,
            contact_person TEXT,
            phone TEXT,
            line_id TEXT,
            email TEXT,
            payment_method TEXT,
            credit_days INTEGER DEFAULT 0,
            rating INTEGER DEFAULT 3,
            status TEXT DEFAULT 'active',
            created_at TEXT DEFAULT (datetime('now','localtime')),
            updated_at TEXT DEFAULT (datetime('now','localtime'))
        );

        CREATE TABLE IF NOT EXISTS supplier_notes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            supplier_id INTEGER NOT NULL,
            note TEXT NOT NULL,
            note_type TEXT DEFAULT 'general',
            created_by TEXT DEFAULT '',
            created_at TEXT DEFAULT (datetime('now','localtime')),
            FOREIGN KEY (supplier_id) REFERENCES suppliers(id)
        );

        CREATE TABLE IF NOT EXISTS orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            erp_id TEXT,
            document_no TEXT,
            doc_date TEXT,
            customer_name TEXT,
            customer_id INTEGER,
            product_name TEXT,
            product_code TEXT,
            price REAL,
            qty REAL,
            discount TEXT,
            total_sales REAL,
            product_cost REAL,
            assign_name TEXT,
            vat TEXT,
            claim_cut_check INTEGER DEFAULT 0,
            claim_cut_qty REAL,
            claim_delivery_check INTEGER DEFAULT 0,
            claim_delivery_qty REAL,
            delivery_cost REAL,
            month_source TEXT
        );

        CREATE TABLE IF NOT EXISTS employees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            employee_code TEXT,
            department TEXT,
            team TEXT,
            role TEXT,
            phone TEXT,
            email TEXT,
            status TEXT DEFAULT 'active',
            created_at TEXT DEFAULT (datetime('now','localtime'))
        );

        CREATE TABLE IF NOT EXISTS tickets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ticket_no TEXT UNIQUE,
            title TEXT NOT NULL,
            description TEXT,
            ticket_type TEXT DEFAULT 'general',
            priority TEXT DEFAULT 'medium',
            status TEXT DEFAULT 'open',
            related_type TEXT,
            related_id INTEGER,
            related_name TEXT,
            assigned_department TEXT,
            assigned_employee_id INTEGER,
            assigned_employee_name TEXT,
            created_by TEXT DEFAULT '',
            created_at TEXT DEFAULT (datetime('now','localtime')),
            updated_at TEXT DEFAULT (datetime('now','localtime')),
            resolved_at TEXT
        );

        CREATE TABLE IF NOT EXISTS ticket_comments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ticket_id INTEGER NOT NULL,
            comment TEXT NOT NULL,
            created_by TEXT DEFAULT '',
            created_at TEXT DEFAULT (datetime('now','localtime')),
            FOREIGN KEY (ticket_id) REFERENCES tickets(id)
        );
    ''')
    conn.commit()
    conn.close()

def row_to_dict(row):
    return dict(row) if row else {}

# ─────────────────────────────────────────
# CUSTOMERS
# ─────────────────────────────────────────

@app.route('/api/customers', methods=['GET'])
def get_customers():
    search = request.args.get('search', '')
    status = request.args.get('status', '')
    conn = get_db()
    q = "SELECT c.*, (SELECT MAX(o.doc_date) FROM orders o WHERE o.customer_id=c.id) as last_order FROM customers c WHERE 1=1"
    params = []
    if search:
        q += " AND c.name LIKE ?"
        params.append(f'%{search}%')
    if status:
        q += " AND c.status=?"
        params.append(status)
    q += " ORDER BY c.name"
    rows = conn.execute(q, params).fetchall()
    conn.close()
    return jsonify([row_to_dict(r) for r in rows])

@app.route('/api/customers', methods=['POST'])
def create_customer():
    d = request.json
    conn = get_db()
    try:
        conn.execute(
            'INSERT INTO customers (name, customer_type, address, phone, email, contact_person, credit_days, status) VALUES (?,?,?,?,?,?,?,?)',
            (d.get('name'), d.get('customer_type',''), d.get('address',''), d.get('phone',''),
             d.get('email',''), d.get('contact_person',''), d.get('credit_days',0), d.get('status','active'))
        )
        conn.commit()
        new_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        row = conn.execute("SELECT * FROM customers WHERE id=?", (new_id,)).fetchone()
        conn.close()
        return jsonify(row_to_dict(row)), 201
    except Exception as e:
        conn.close()
        return jsonify({'error': str(e)}), 400

@app.route('/api/customers/<int:cid>', methods=['GET'])
def get_customer(cid):
    conn = get_db()
    row = conn.execute("SELECT * FROM customers WHERE id=?", (cid,)).fetchone()
    if not row:
        conn.close()
        return jsonify({'error': 'ไม่พบข้อมูล'}), 404
    c = row_to_dict(row)
    c['notes'] = [row_to_dict(n) for n in conn.execute(
        "SELECT * FROM customer_notes WHERE customer_id=? ORDER BY created_at DESC", (cid,)
    ).fetchall()]
    c['orders'] = [row_to_dict(o) for o in conn.execute(
        '''SELECT document_no, doc_date, product_name, price, qty, total_sales, product_cost,
                  assign_name, claim_cut_check, claim_cut_qty, claim_delivery_check, claim_delivery_qty, month_source
           FROM orders WHERE customer_id=? ORDER BY doc_date DESC LIMIT 200''', (cid,)
    ).fetchall()]
    stats = conn.execute(
        '''SELECT COUNT(DISTINCT document_no) as total_orders,
                  COALESCE(SUM(total_sales),0) as total_revenue,
                  COALESCE(SUM(product_cost),0) as total_cost,
                  SUM(CASE WHEN claim_cut_check=1 THEN 1 ELSE 0 END) as claim_cut_count,
                  SUM(CASE WHEN claim_delivery_check=1 THEN 1 ELSE 0 END) as claim_delivery_count
           FROM orders WHERE customer_id=?''', (cid,)
    ).fetchone()
    c['stats'] = row_to_dict(stats)
    conn.close()
    return jsonify(c)

@app.route('/api/customers/<int:cid>', methods=['PUT'])
def update_customer(cid):
    d = request.json
    conn = get_db()
    conn.execute(
        '''UPDATE customers SET name=?, customer_type=?, address=?, phone=?, email=?,
           contact_person=?, credit_days=?, status=?, updated_at=datetime('now','localtime') WHERE id=?''',
        (d.get('name'), d.get('customer_type',''), d.get('address',''), d.get('phone',''),
         d.get('email',''), d.get('contact_person',''), d.get('credit_days',0), d.get('status','active'), cid)
    )
    conn.commit()
    row = conn.execute("SELECT * FROM customers WHERE id=?", (cid,)).fetchone()
    conn.close()
    return jsonify(row_to_dict(row))

@app.route('/api/customers/<int:cid>/notes', methods=['POST'])
def add_customer_note(cid):
    d = request.json
    conn = get_db()
    conn.execute("INSERT INTO customer_notes (customer_id, note, created_by) VALUES (?,?,?)",
                 (cid, d.get('note',''), d.get('created_by','')))
    conn.commit()
    conn.close()
    return jsonify({'success': True}), 201

# ─────────────────────────────────────────
# SUPPLIERS
# ─────────────────────────────────────────

@app.route('/api/suppliers', methods=['GET'])
def get_suppliers():
    search = request.args.get('search', '')
    conn = get_db()
    q = "SELECT * FROM suppliers WHERE 1=1"
    params = []
    if search:
        q += " AND name LIKE ?"
        params.append(f'%{search}%')
    q += " ORDER BY name"
    rows = conn.execute(q, params).fetchall()
    conn.close()
    return jsonify([row_to_dict(r) for r in rows])

@app.route('/api/suppliers', methods=['POST'])
def create_supplier():
    d = request.json
    conn = get_db()
    conn.execute(
        'INSERT INTO suppliers (name, supplier_type, contact_person, phone, line_id, email, payment_method, credit_days, rating, status) VALUES (?,?,?,?,?,?,?,?,?,?)',
        (d.get('name'), d.get('supplier_type',''), d.get('contact_person',''), d.get('phone',''),
         d.get('line_id',''), d.get('email',''), d.get('payment_method',''),
         d.get('credit_days',0), d.get('rating',3), d.get('status','active'))
    )
    conn.commit()
    new_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    row = conn.execute("SELECT * FROM suppliers WHERE id=?", (new_id,)).fetchone()
    conn.close()
    return jsonify(row_to_dict(row)), 201

@app.route('/api/suppliers/<int:sid>', methods=['GET'])
def get_supplier(sid):
    conn = get_db()
    row = conn.execute("SELECT * FROM suppliers WHERE id=?", (sid,)).fetchone()
    if not row:
        conn.close()
        return jsonify({'error': 'ไม่พบข้อมูล'}), 404
    s = row_to_dict(row)
    s['notes'] = [row_to_dict(n) for n in conn.execute(
        "SELECT * FROM supplier_notes WHERE supplier_id=? ORDER BY created_at DESC", (sid,)
    ).fetchall()]
    s['tickets'] = [row_to_dict(t) for t in conn.execute(
        "SELECT * FROM tickets WHERE related_type='supplier' AND related_id=? ORDER BY created_at DESC LIMIT 30", (sid,)
    ).fetchall()]
    conn.close()
    return jsonify(s)

@app.route('/api/suppliers/<int:sid>', methods=['PUT'])
def update_supplier(sid):
    d = request.json
    conn = get_db()
    conn.execute(
        '''UPDATE suppliers SET name=?, supplier_type=?, contact_person=?, phone=?, line_id=?,
           email=?, payment_method=?, credit_days=?, rating=?, status=?,
           updated_at=datetime('now','localtime') WHERE id=?''',
        (d.get('name'), d.get('supplier_type',''), d.get('contact_person',''), d.get('phone',''),
         d.get('line_id',''), d.get('email',''), d.get('payment_method',''),
         d.get('credit_days',0), d.get('rating',3), d.get('status','active'), sid)
    )
    conn.commit()
    row = conn.execute("SELECT * FROM suppliers WHERE id=?", (sid,)).fetchone()
    conn.close()
    return jsonify(row_to_dict(row))

@app.route('/api/suppliers/<int:sid>/notes', methods=['POST'])
def add_supplier_note(sid):
    d = request.json
    conn = get_db()
    conn.execute("INSERT INTO supplier_notes (supplier_id, note, note_type, created_by) VALUES (?,?,?,?)",
                 (sid, d.get('note',''), d.get('note_type','general'), d.get('created_by','')))
    conn.commit()
    conn.close()
    return jsonify({'success': True}), 201

# ─────────────────────────────────────────
# TICKETS
# ─────────────────────────────────────────

def next_ticket_no():
    conn = get_db()
    count = conn.execute("SELECT COUNT(*) FROM tickets").fetchone()[0]
    conn.close()
    return f"TK-{datetime.now().year}-{str(count+1).zfill(4)}"

@app.route('/api/tickets', methods=['GET'])
def get_tickets():
    status = request.args.get('status', '')
    dept = request.args.get('department', '')
    search = request.args.get('search', '')
    conn = get_db()
    q = "SELECT * FROM tickets WHERE 1=1"
    params = []
    if status:
        q += " AND status=?"
        params.append(status)
    if dept:
        q += " AND assigned_department=?"
        params.append(dept)
    if search:
        q += " AND (title LIKE ? OR related_name LIKE ? OR ticket_no LIKE ?)"
        params += [f'%{search}%', f'%{search}%', f'%{search}%']
    q += " ORDER BY created_at DESC"
    rows = conn.execute(q, params).fetchall()
    conn.close()
    return jsonify([row_to_dict(r) for r in rows])

@app.route('/api/tickets', methods=['POST'])
def create_ticket():
    d = request.json
    tno = next_ticket_no()
    conn = get_db()
    conn.execute(
        '''INSERT INTO tickets (ticket_no, title, description, ticket_type, priority, status,
           related_type, related_id, related_name, assigned_department,
           assigned_employee_id, assigned_employee_name, created_by) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
        (tno, d.get('title'), d.get('description',''), d.get('ticket_type','general'),
         d.get('priority','medium'), 'open', d.get('related_type',''),
         d.get('related_id'), d.get('related_name',''), d.get('assigned_department',''),
         d.get('assigned_employee_id'), d.get('assigned_employee_name',''), d.get('created_by',''))
    )
    conn.commit()
    new_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    row = conn.execute("SELECT * FROM tickets WHERE id=?", (new_id,)).fetchone()
    conn.close()
    return jsonify(row_to_dict(row)), 201

@app.route('/api/tickets/<int:tid>', methods=['GET'])
def get_ticket(tid):
    conn = get_db()
    row = conn.execute("SELECT * FROM tickets WHERE id=?", (tid,)).fetchone()
    if not row:
        conn.close()
        return jsonify({'error': 'ไม่พบข้อมูล'}), 404
    t = row_to_dict(row)
    t['comments'] = [row_to_dict(c) for c in conn.execute(
        "SELECT * FROM ticket_comments WHERE ticket_id=? ORDER BY created_at ASC", (tid,)
    ).fetchall()]
    conn.close()
    return jsonify(t)

@app.route('/api/tickets/<int:tid>', methods=['PUT'])
def update_ticket(tid):
    d = request.json
    conn = get_db()
    old = conn.execute("SELECT resolved_at FROM tickets WHERE id=?", (tid,)).fetchone()
    resolved_at = old['resolved_at'] if old else None
    if d.get('status') in ('resolved','closed') and not resolved_at:
        resolved_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    conn.execute(
        '''UPDATE tickets SET title=?, description=?, ticket_type=?, priority=?, status=?,
           assigned_department=?, assigned_employee_id=?, assigned_employee_name=?,
           updated_at=datetime('now','localtime'), resolved_at=? WHERE id=?''',
        (d.get('title'), d.get('description',''), d.get('ticket_type','general'),
         d.get('priority','medium'), d.get('status','open'), d.get('assigned_department',''),
         d.get('assigned_employee_id'), d.get('assigned_employee_name',''), resolved_at, tid)
    )
    conn.commit()
    row = conn.execute("SELECT * FROM tickets WHERE id=?", (tid,)).fetchone()
    conn.close()
    return jsonify(row_to_dict(row))

@app.route('/api/tickets/<int:tid>/comments', methods=['POST'])
def add_ticket_comment(tid):
    d = request.json
    conn = get_db()
    conn.execute("INSERT INTO ticket_comments (ticket_id, comment, created_by) VALUES (?,?,?)",
                 (tid, d.get('comment',''), d.get('created_by','')))
    conn.execute("UPDATE tickets SET updated_at=datetime('now','localtime') WHERE id=?", (tid,))
    conn.commit()
    conn.close()
    return jsonify({'success': True}), 201

# ─────────────────────────────────────────
# EMPLOYEES
# ─────────────────────────────────────────

@app.route('/api/employees', methods=['GET'])
def get_employees():
    dept = request.args.get('department', '')
    conn = get_db()
    q = "SELECT * FROM employees WHERE status='active'"
    params = []
    if dept:
        q += " AND department=?"
        params.append(dept)
    q += " ORDER BY department, name"
    rows = conn.execute(q, params).fetchall()
    conn.close()
    return jsonify([row_to_dict(r) for r in rows])

@app.route('/api/employees', methods=['POST'])
def create_employee():
    d = request.json
    conn = get_db()
    conn.execute(
        'INSERT INTO employees (name, employee_code, department, team, role, phone, email, status) VALUES (?,?,?,?,?,?,?,?)',
        (d.get('name'), d.get('employee_code',''), d.get('department',''), d.get('team',''),
         d.get('role',''), d.get('phone',''), d.get('email',''), d.get('status','active'))
    )
    conn.commit()
    new_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    row = conn.execute("SELECT * FROM employees WHERE id=?", (new_id,)).fetchone()
    conn.close()
    return jsonify(row_to_dict(row)), 201

@app.route('/api/employees/<int:eid>', methods=['PUT'])
def update_employee(eid):
    d = request.json
    conn = get_db()
    conn.execute(
        'UPDATE employees SET name=?, department=?, team=?, role=?, phone=?, email=?, status=? WHERE id=?',
        (d.get('name'), d.get('department',''), d.get('team',''), d.get('role',''),
         d.get('phone',''), d.get('email',''), d.get('status','active'), eid)
    )
    conn.commit()
    row = conn.execute("SELECT * FROM employees WHERE id=?", (eid,)).fetchone()
    conn.close()
    return jsonify(row_to_dict(row))

# ─────────────────────────────────────────
# IMPORT ERP
# ─────────────────────────────────────────

@app.route('/api/import/erp', methods=['POST'])
def import_erp():
    if 'file' not in request.files:
        return jsonify({'error': 'ไม่พบไฟล์'}), 400
    file = request.files['file']
    month_source = request.form.get('month', 'UNKNOWN').upper()
    try:
        wb = openpyxl.load_workbook(file, read_only=True, data_only=True)
        sheet_name = wb.sheetnames[0]
        for name in wb.sheetnames:
            if any(k in name.lower() for k in ['invoice','archive','table']):
                sheet_name = name
                break
        ws = wb[sheet_name]
        all_rows = list(ws.iter_rows(values_only=True))
        if len(all_rows) < 2:
            return jsonify({'error': 'ไฟล์ไม่มีข้อมูล'}), 400

        headers = [str(h).strip() if h else '' for h in all_rows[0]]
        def col(row_data, name):
            try:
                return row_data[headers.index(name)]
            except (ValueError, IndexError):
                return None

        conn = get_db()
        imported = skipped = 0

        for row in all_rows[1:]:
            if not any(v for v in row if v is not None):
                continue
            erp_id = str(col(row,'ID') or '').strip()
            if not erp_id or erp_id in ('None',''):
                continue
            if conn.execute("SELECT 1 FROM orders WHERE erp_id=? AND month_source=?", (erp_id, month_source)).fetchone():
                skipped += 1
                continue

            cname = str(col(row,'Customer_Name') or '').strip()
            cid = None
            if cname:
                cr = conn.execute("SELECT id FROM customers WHERE name=?", (cname,)).fetchone()
                if cr:
                    cid = cr['id']
                else:
                    conn.execute("INSERT INTO customers (name, status) VALUES (?,?)", (cname,'active'))
                    conn.commit()
                    cid = conn.execute("SELECT last_insert_rowid()").fetchone()[0]

            dd = col(row,'Doc_date')
            if dd and hasattr(dd,'isoformat'):
                dd = dd.strftime('%Y-%m-%d')
            elif dd:
                dd = str(dd)[:10]

            conn.execute(
                '''INSERT INTO orders (erp_id,document_no,doc_date,customer_name,customer_id,
                   product_name,product_code,price,qty,discount,total_sales,product_cost,
                   assign_name,vat,claim_cut_check,claim_cut_qty,
                   claim_delivery_check,claim_delivery_qty,delivery_cost,month_source)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                (erp_id, str(col(row,'Document_No') or ''), dd, cname, cid,
                 str(col(row,'Product_Name') or ''), str(col(row,'Product_Code') or ''),
                 col(row,'Price'), col(row,'QTY'), str(col(row,'Discount') or ''),
                 col(row,'Total_Sales'), col(row,'Product_Cost'), str(col(row,'Assign_Name') or ''),
                 str(col(row,'Vat') or ''),
                 1 if col(row,'Claim_CUT_Check') else 0, col(row,'Claim_CUT_QTY'),
                 1 if col(row,'Claim_Delivery_Check') else 0, col(row,'Claim_Delivery_QTY'),
                 col(row,'Delivery_Cost'), month_source)
            )
            imported += 1

        conn.commit()
        conn.close()
        return jsonify({'success': True, 'imported': imported, 'skipped': skipped,
                        'message': f'นำเข้าสำเร็จ {imported} รายการ (ข้าม {skipped} ซ้ำ)'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ─────────────────────────────────────────
# DASHBOARD
# ─────────────────────────────────────────

@app.route('/api/dashboard', methods=['GET'])
def dashboard():
    conn = get_db()

    monthly = conn.execute(
        '''SELECT month_source,
                  COALESCE(SUM(total_sales),0) as total_sales,
                  COALESCE(SUM(product_cost),0) as total_cost,
                  COUNT(DISTINCT customer_name) as unique_customers,
                  COUNT(DISTINCT document_no) as total_orders
           FROM orders GROUP BY month_source ORDER BY month_source'''
    ).fetchall()

    claims = conn.execute(
        '''SELECT SUM(CASE WHEN claim_cut_check=1 THEN 1 ELSE 0 END) as cut_count,
                  SUM(CASE WHEN claim_delivery_check=1 THEN 1 ELSE 0 END) as delivery_count,
                  COUNT(*) as total_lines FROM orders'''
    ).fetchone()

    ticket_status = conn.execute(
        "SELECT status, COUNT(*) as cnt FROM tickets GROUP BY status"
    ).fetchall()

    ticket_dept = conn.execute(
        "SELECT assigned_department, COUNT(*) as cnt FROM tickets WHERE status NOT IN ('closed') GROUP BY assigned_department"
    ).fetchall()

    top_customers = conn.execute(
        '''SELECT customer_name, COALESCE(SUM(total_sales),0) as revenue, COUNT(DISTINCT document_no) as orders
           FROM orders GROUP BY customer_name ORDER BY revenue DESC LIMIT 10'''
    ).fetchall()

    summary = {
        'customers': conn.execute("SELECT COUNT(*) FROM customers WHERE status='active'").fetchone()[0],
        'suppliers': conn.execute("SELECT COUNT(*) FROM suppliers WHERE status='active'").fetchone()[0],
        'open_tickets': conn.execute("SELECT COUNT(*) FROM tickets WHERE status='open'").fetchone()[0],
        'inprogress_tickets': conn.execute("SELECT COUNT(*) FROM tickets WHERE status='in_progress'").fetchone()[0],
        'total_revenue': conn.execute("SELECT COALESCE(SUM(total_sales),0) FROM orders").fetchone()[0],
        'employees': conn.execute("SELECT COUNT(*) FROM employees WHERE status='active'").fetchone()[0],
    }

    conn.close()
    return jsonify({
        'monthly': [row_to_dict(r) for r in monthly],
        'claims': row_to_dict(claims),
        'ticket_status': [row_to_dict(r) for r in ticket_status],
        'ticket_dept': [row_to_dict(r) for r in ticket_dept],
        'top_customers': [row_to_dict(r) for r in top_customers],
        'summary': summary
    })

# ─────────────────────────────────────────
# SERVE FRONTEND
# ─────────────────────────────────────────

@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def serve(path):
    if path and os.path.exists(os.path.join('static', path)):
        return send_from_directory('static', path)
    return send_from_directory('static', 'index.html')

# ─────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────

if __name__ == '__main__':
    os.makedirs('static', exist_ok=True)
    init_db()
    print("=" * 50)
    print("  Sourcing BU CRM - กำลังเริ่มทำงาน")
    print("=" * 50)
    print(f"  เปิด browser ที่: http://localhost:5000")
    print(f"  หรือเข้าจาก PC อื่น: http://[IP ของเครื่อง]:5000")
    print("=" * 50)
    app.run(debug=False, host='0.0.0.0', port=5000)
