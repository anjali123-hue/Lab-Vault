import hashlib
import os
import secrets
import smtplib
import sqlite3
import threading
from datetime import datetime, timedelta, timezone
from email.message import EmailMessage
from functools import wraps
from pathlib import Path

from dotenv import load_dotenv
from flask import (
    Flask,
    flash,
    g,
    redirect,
    render_template,
    request,
    session,
    url_for,
)
from itsdangerous import BadSignature, SignatureExpired, URLSafeTimedSerializer
from werkzeug.security import check_password_hash, generate_password_hash

load_dotenv()

ROOT = Path(__file__).resolve().parent
DB_PATH = Path(os.getenv("SQLITE_PATH", ROOT / "data" / "labvault.db"))
DB_PATH.parent.mkdir(parents=True, exist_ok=True)
SECRET_KEY = os.getenv("SECRET_KEY", secrets.token_hex(32))
MAIL_SENDER = os.getenv("MAIL_USERNAME", "labvault.lab@gmail.com")
MAIL_PASSWORD = os.getenv("MAIL_PASSWORD", "")
ADMIN_USERNAME = os.getenv("ADMIN_USERNAME", "admin")
ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD", "change-this-admin-password")
HOD_EMAIL = os.getenv("HOD_EMAIL", "")
APP_URL = os.getenv("APP_URL", "")
serializer = URLSafeTimedSerializer(SECRET_KEY)

app = Flask(__name__)
app.config.update(
    SECRET_KEY=SECRET_KEY,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.getenv("SESSION_COOKIE_SECURE", "0") == "1",
    PERMANENT_SESSION_LIFETIME=timedelta(hours=8),
)
DB_LOCK = threading.Lock()


def utcnow():
    return datetime.now(timezone.utc).replace(microsecond=0)


def db():
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
    return g.db


@app.teardown_appcontext
def close_db(_error=None):
    conn = g.pop("db", None)
    if conn:
        conn.close()


def query(sql, args=(), one=False):
    cur = db().execute(sql, args)
    rows = cur.fetchone() if one else cur.fetchall()
    cur.close()
    return rows


def execute(sql, args=()):
    with DB_LOCK:
        cur = db().execute(sql, args)
        db().commit()
        last = cur.lastrowid
        cur.close()
    return last


def init_db():
    with app.app_context():
        db().executescript(
            """
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                role TEXT NOT NULL CHECK(role IN ('student','admin')),
                full_name TEXT NOT NULL,
                uid TEXT UNIQUE,
                email TEXT UNIQUE,
                phone TEXT,
                branch TEXT,
                division TEXT,
                year TEXT,
                password_hash TEXT NOT NULL,
                created_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS inventory (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                source_file TEXT NOT NULL,
                source_key TEXT NOT NULL UNIQUE,
                stock_number TEXT,
                name TEXT NOT NULL,
                category TEXT NOT NULL,
                description TEXT,
                location TEXT,
                total_qty INTEGER NOT NULL DEFAULT 0,
                available_qty INTEGER NOT NULL DEFAULT 0,
                minimum_stock INTEGER NOT NULL DEFAULT 1,
                condition TEXT NOT NULL DEFAULT 'Good',
                maintenance_status TEXT NOT NULL DEFAULT 'Clear',
                active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS requests (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                request_code TEXT NOT NULL UNIQUE,
                student_id INTEGER NOT NULL REFERENCES users(id),
                project_title TEXT NOT NULL,
                purpose TEXT NOT NULL,
                mentor_name TEXT NOT NULL,
                mentor_email TEXT NOT NULL,
                mentor_phone TEXT,
                hod_name TEXT NOT NULL,
                hod_email TEXT NOT NULL,
                hod_phone TEXT,
                due_date TEXT NOT NULL,
                mentor_status TEXT NOT NULL DEFAULT 'PENDING',
                hod_status TEXT NOT NULL DEFAULT 'PENDING',
                overall_status TEXT NOT NULL DEFAULT 'PENDING_MENTOR',
                rejection_reason TEXT,
                collection_otp_ready INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS request_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                request_id INTEGER NOT NULL REFERENCES requests(id) ON DELETE CASCADE,
                inventory_id INTEGER NOT NULL REFERENCES inventory(id),
                quantity INTEGER NOT NULL,
                UNIQUE(request_id, inventory_id)
            );
            CREATE TABLE IF NOT EXISTS approval_tokens (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                request_id INTEGER NOT NULL REFERENCES requests(id) ON DELETE CASCADE,
                stage TEXT NOT NULL,
                token_hash TEXT NOT NULL UNIQUE,
                status TEXT NOT NULL DEFAULT 'PENDING',
                created_at TEXT NOT NULL,
                used_at TEXT
            );
            CREATE TABLE IF NOT EXISTS otp_tokens (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                request_id INTEGER NOT NULL REFERENCES requests(id) ON DELETE CASCADE,
                kind TEXT NOT NULL,
                otp_hash TEXT NOT NULL,
                expires_at TEXT NOT NULL,
                attempts INTEGER NOT NULL DEFAULT 0,
                used_at TEXT,
                created_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS issues (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                request_id INTEGER NOT NULL REFERENCES requests(id),
                student_id INTEGER NOT NULL REFERENCES users(id),
                issued_at TEXT NOT NULL,
                due_date TEXT NOT NULL,
                returned_at TEXT,
                return_condition TEXT,
                return_remarks TEXT
            );
            CREATE TABLE IF NOT EXISTS notifications (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL REFERENCES users(id),
                title TEXT NOT NULL,
                body TEXT NOT NULL,
                kind TEXT NOT NULL,
                is_read INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS audit_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER,
                role TEXT,
                action TEXT NOT NULL,
                description TEXT,
                created_at TEXT NOT NULL
            );
            """
        )
        db().commit()
        admin = query("SELECT id FROM users WHERE role='admin' LIMIT 1", one=True)
        if not admin:
            execute(
                """INSERT INTO users
                (role, full_name, email, password_hash, created_at)
                VALUES ('admin', 'Lab Assistant', ?, ?, ?)""",
                (MAIL_SENDER, generate_password_hash(ADMIN_PASSWORD), utcnow().isoformat()),
            )
        import_inventory()


def category_for(name):
    n = name.lower()
    groups = [
        ("Development Boards", ["board", "arduino", "esp", "raspberry", "raspi", "microcontroller"]),
        ("Sensors", ["sensor", "dht", "ds18", "probe", "ldr"]),
        ("Power Supplies", ["power", "adapter", "battery", "charger", "supply"]),
        ("Communication Modules", ["wifi", "bluetooth", "gsm", "zigbee", "router", "ethernet"]),
        ("Displays", ["display", "lcd", "oled", "projector", "monitor"]),
        ("Motors", ["motor", "servo", "stepper", "wheel"]),
        ("Cables", ["cable", "hdmi", "usb", "wire", "connector"]),
        ("Storage", ["pendrive", "ssd", "memory", "disk"]),
        ("Tools", ["printer", "camera", "headset", "multimeter", "solder", "tool"]),
        ("ICs & Components", ["ic", "resistor", "capacitor", "relay", "transistor"]),
    ]
    for label, words in groups:
        if any(word in n for word in words):
            return label
    return "Consumables"


def import_inventory():
    """Idempotently import only the two uploaded workbooks."""
    try:
        from openpyxl import load_workbook
    except ImportError:
        return
    files = [
        (ROOT / "attached_assets" / "master_stock_1786273496625.xlsx", "master_stock"),
        (ROOT / "attached_assets" / "consumables_1786273504248.xlsx", "consumables"),
    ]
    for path, source in files:
        if not path.exists():
            continue
        try:
            workbook = load_workbook(path, read_only=True, data_only=True)
            sheet = workbook.active
            rows = list(sheet.iter_rows(values_only=True))
            header_idx = next(
                (i for i, row in enumerate(rows) if any(str(v or "").strip().lower() in {"name of the article", "qty"} for v in row)),
                None,
            )
            if header_idx is None:
                continue
            headers = [str(v or "").strip().lower() for v in rows[header_idx]]
            def col(*names):
                return next((headers.index(n) for n in names if n in headers), None)
            name_i = col("name of the article")
            qty_i = col("qty", "quantity")
            if name_i is None:
                continue
            stock_i = col("stock register no")
            sr_i = col("sr no", "sr.no")
            loc_i = col("allocated to", "location")
            for row_number, row in enumerate(rows[header_idx + 1 :], 1):
                name = str(row[name_i] or "").strip() if name_i < len(row) else ""
                if not name:
                    continue
                stock = str(row[stock_i] or "").strip() if stock_i is not None and stock_i < len(row) else ""
                sr = str(row[sr_i] or "").strip() if sr_i is not None and sr_i < len(row) else ""
                key = f"{source}:{stock or sr or row_number}:{name.lower()}"
                raw_qty = row[qty_i] if qty_i is not None and qty_i < len(row) else 0
                try:
                    qty = max(0, int(float(raw_qty or 0)))
                except (TypeError, ValueError):
                    qty = 0
                location = str(row[loc_i] or "").strip() if loc_i is not None and loc_i < len(row) else ""
                existing = query("SELECT id FROM inventory WHERE source_key=?", (key,), one=True)
                now = utcnow().isoformat()
                if existing:
                    execute(
                        """UPDATE inventory SET stock_number=?, name=?, category=?, location=?,
                        total_qty=?, updated_at=? WHERE source_key=?""",
                        (stock or sr or None, name, category_for(name), location or "Main Lab", qty, now, key),
                    )
                else:
                    execute(
                        """INSERT INTO inventory
                        (source_file, source_key, stock_number, name, category, description, location,
                         total_qty, available_qty, minimum_stock, created_at, updated_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                        (source, key, stock or sr or None, name, category_for(name),
                         "Imported from the LabVault inventory workbook.",
                         location or "Main Lab", qty, qty, max(1, min(3, qty // 4 or 1)), now, now),
                    )
            workbook.close()
        except Exception as exc:
            app.logger.warning("Inventory import skipped for %s: %s", path.name, exc)


def csrf_token():
    if "csrf" not in session:
        session["csrf"] = secrets.token_urlsafe(24)
    return session["csrf"]


@app.context_processor
def inject_globals():
    user = current_user()
    unread = query("SELECT COUNT(*) AS count FROM notifications WHERE user_id=? AND is_read=0", (user["id"],), one=True)["count"] if user else 0
    return {"current_user": user, "csrf_token": csrf_token(), "unread_count": unread, "now": utcnow()}


def current_user():
    user_id = session.get("user_id")
    return query("SELECT * FROM users WHERE id=?", (user_id,), one=True) if user_id else None


def login_required(role=None):
    def decorator(fn):
        @wraps(fn)
        def wrapped(*args, **kwargs):
            user = current_user()
            if not user:
                flash("Please sign in to continue.", "info")
                return redirect(url_for("login", role=role or "student"))
            if role and user["role"] != role:
                flash("You do not have access to that workspace.", "error")
                return redirect(url_for("dashboard"))
            return fn(*args, **kwargs)
        return wrapped
    return decorator


def valid_csrf():
    return secrets.compare_digest(session.get("csrf", ""), request.form.get("csrf", ""))


def audit(action, description="", user=None):
    user = user or current_user()
    execute(
        "INSERT INTO audit_logs (user_id, role, action, description, created_at) VALUES (?, ?, ?, ?, ?)",
        (user["id"] if user else None, user["role"] if user else None, action, description, utcnow().isoformat()),
    )


def notify(user_id, title, body, kind="info"):
    execute(
        "INSERT INTO notifications (user_id, title, body, kind, created_at) VALUES (?, ?, ?, ?, ?)",
        (user_id, title, body, kind, utcnow().isoformat()),
    )


def signed_token(request_id, stage):
    raw = serializer.dumps({"request_id": request_id, "stage": stage})
    execute(
        "INSERT INTO approval_tokens (request_id, stage, token_hash, created_at) VALUES (?, ?, ?, ?)",
        (request_id, stage, hashlib.sha256(raw.encode()).hexdigest(), utcnow().isoformat()),
    )
    return raw


def token_payload(raw, stage):
    try:
        payload = serializer.loads(raw, max_age=60 * 60 * 24 * 7)
    except (BadSignature, SignatureExpired):
        return None, "This approval link has expired or is invalid."
    if payload.get("stage") != stage:
        return None, "This approval link is not valid for this approval stage."
    token = query(
        "SELECT * FROM approval_tokens WHERE token_hash=? AND stage=?",
        (hashlib.sha256(raw.encode()).hexdigest(), stage),
        one=True,
    )
    if not token or token["status"] != "PENDING":
        return None, "This request has already been reviewed."
    req = query("SELECT * FROM requests WHERE id=?", (payload.get("request_id"),), one=True)
    if not req:
        return None, "The requested approval could not be found."
    expected = "PENDING_MENTOR" if stage == "mentor" else "PENDING_HOD"
    if req["overall_status"] != expected:
        return None, "This request has already been reviewed."
    return (payload, token), None


def send_mail(recipient, subject, body, action_url=None):
    if not recipient:
        return False
    html = f"""<!doctype html><html><body style="font-family:Arial,sans-serif;color:#18243a">
    <div style="max-width:620px;margin:auto;padding:32px;background:#f6f8fc">
    <div style="background:#17213a;color:white;border-radius:18px;padding:24px">
    <div style="font-size:24px;font-weight:800">Lab<span style="color:#8c6cff">Vault</span></div>
    <p style="color:#c4cee4">{body.replace(chr(10), '<br>')}</p>
    {f'<a href="{action_url}" style="display:inline-block;background:#5277f7;color:white;padding:13px 20px;border-radius:10px;text-decoration:none;font-weight:700">Review in LabVault</a>' if action_url else ''}
    </div><p style="font-size:12px;color:#7d889b">Sent by LabVault · {MAIL_SENDER}</p></div></body></html>"""
    try:
        if not MAIL_PASSWORD:
            app.logger.warning("MAIL_PASSWORD is not configured; email preview not sent to %s", recipient)
            return False
        message = EmailMessage()
        message["Subject"] = subject
        message["From"] = f"LabVault <{MAIL_SENDER}>"
        message["To"] = recipient
        message.set_content(body)
        message.add_alternative(html, subtype="html")
        with smtplib.SMTP(os.getenv("MAIL_SERVER", "smtp.gmail.com"), int(os.getenv("MAIL_PORT", "587"))) as smtp:
            smtp.starttls()
            smtp.login(MAIL_SENDER, MAIL_PASSWORD)
            smtp.send_message(message)
        return True
    except Exception as exc:
        app.logger.warning("Email delivery failed: %s", exc)
        return False


def request_items(request_id):
    return query(
        """SELECT ri.quantity, i.name, i.category, i.location
        FROM request_items ri JOIN inventory i ON i.id=ri.inventory_id
        WHERE ri.request_id=? ORDER BY i.name""",
        (request_id,),
    )


@app.route("/")
def landing():
    stats = {
        "items": query("SELECT COUNT(*) AS n FROM inventory WHERE active=1", one=True)["n"],
        "units": query("SELECT COALESCE(SUM(total_qty),0) AS n FROM inventory WHERE active=1", one=True)["n"],
        "categories": query("SELECT COUNT(DISTINCT category) AS n FROM inventory WHERE active=1", one=True)["n"],
    }
    return render_template("landing.html", stats=stats)


@app.route("/login", methods=["GET", "POST"])
def login():
    role = request.args.get("role", request.form.get("role", "student"))
    if request.method == "POST":
        if not valid_csrf():
            flash("Your session expired. Please try again.", "error")
        else:
            identity = request.form.get("identity", "").strip()
            password = request.form.get("password", "")
            if role == "admin":
                user = query(
                    "SELECT * FROM users WHERE role='admin' AND (email=? OR full_name=? OR ?=?)",
                    (identity, identity, identity, ADMIN_USERNAME),
                    one=True,
                )
            else:
                user = query("SELECT * FROM users WHERE role='student' AND uid=?", (identity.upper(),), one=True)
            if user and check_password_hash(user["password_hash"], password):
                session.clear()
                session.permanent = True
                session["user_id"] = user["id"]
                session["csrf"] = secrets.token_urlsafe(24)
                audit("LOGIN", "Successful sign in.", user)
                return redirect(url_for("dashboard"))
            flash("Invalid credentials. Please check your details and try again.", "error")
    return render_template("login.html", role=role)


@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        if not valid_csrf():
            flash("Your session expired. Please try again.", "error")
            return render_template("register.html")
        values = {key: request.form.get(key, "").strip() for key in ["full_name", "phone", "uid", "branch", "division", "year", "email"]}
        password = request.form.get("password", "")
        if len(values["full_name"]) < 2 or len(values["uid"]) < 5 or "@" not in values["email"] or len(password) < 8:
            flash("Please complete all fields and use a password of at least 8 characters.", "error")
            return render_template("register.html")
        try:
            execute(
                """INSERT INTO users (role, full_name, phone, uid, branch, division, year, email, password_hash, created_at)
                VALUES ('student', ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (*values.values(), generate_password_hash(password), utcnow().isoformat()),
            )
            flash("Account created. Sign in with your UID to continue.", "success")
            return redirect(url_for("login", role="student"))
        except sqlite3.IntegrityError:
            flash("That UID or email is already registered.", "error")
    return render_template("register.html")


@app.route("/logout")
def logout():
    user = current_user()
    if user:
        audit("LOGOUT", "Signed out.", user)
    session.clear()
    return redirect(url_for("landing"))


@app.route("/dashboard")
@login_required()
def dashboard():
    return redirect(url_for("admin_dashboard" if current_user()["role"] == "admin" else "student_dashboard"))


@app.route("/student")
@login_required("student")
def student_dashboard():
    user = current_user()
    term = request.args.get("q", "").strip()
    category = request.args.get("category", "")
    availability = request.args.get("availability", "")
    sql = "SELECT * FROM inventory WHERE active=1"
    args = []
    if term:
        sql += " AND (name LIKE ? OR category LIKE ? OR stock_number LIKE ?)"
        args += [f"%{term}%", f"%{term}%", f"%{term}%"]
    if category:
        sql += " AND category=?"
        args.append(category)
    if availability == "available":
        sql += " AND available_qty > 0"
    elif availability == "low":
        sql += " AND available_qty <= minimum_stock AND available_qty > 0"
    sql += " ORDER BY name LIMIT 80"
    items = query(sql, args)
    categories = query("SELECT DISTINCT category FROM inventory WHERE active=1 ORDER BY category")
    requests = query("SELECT * FROM requests WHERE student_id=? ORDER BY created_at DESC LIMIT 6", (user["id"],))
    issues = query(
        """SELECT issues.*, requests.request_code FROM issues JOIN requests ON requests.id=issues.request_id
        WHERE issues.student_id=? AND issues.returned_at IS NULL ORDER BY issues.due_date""",
        (user["id"],),
    )
    cart = session.get("cart", {})
    return render_template(
        "student.html", items=items, categories=categories, requests=requests, issues=issues,
        cart=cart, term=term, selected_category=category, availability=availability,
    )


@app.post("/student/cart/add/<int:item_id>")
@login_required("student")
def cart_add(item_id):
    if not valid_csrf():
        flash("Your session expired. Please try again.", "error")
        return redirect(url_for("student_dashboard"))
    item = query("SELECT * FROM inventory WHERE id=? AND active=1", (item_id,), one=True)
    if not item or item["available_qty"] < 1:
        flash("That item is currently unavailable.", "error")
        return redirect(request.referrer or url_for("student_dashboard"))
    cart = session.setdefault("cart", {})
    key = str(item_id)
    cart[key] = min(int(cart.get(key, 0)) + 1, item["available_qty"])
    session.modified = True
    flash(f"{item['name']} added to your request list.", "success")
    return redirect(request.referrer or url_for("student_dashboard"))


@app.post("/student/cart/remove/<int:item_id>")
@login_required("student")
def cart_remove(item_id):
    if valid_csrf():
        cart = session.setdefault("cart", {})
        cart.pop(str(item_id), None)
        session.modified = True
    return redirect(request.referrer or url_for("student_dashboard"))


@app.route("/student/request", methods=["GET", "POST"])
@login_required("student")
def create_request():
    cart = session.get("cart", {})
    ids = [int(key) for key in cart.keys() if str(key).isdigit()]
    items = query(f"SELECT * FROM inventory WHERE id IN ({','.join('?' for _ in ids)})", ids) if ids else []
    if request.method == "POST":
        if not valid_csrf():
            flash("Your session expired. Please try again.", "error")
        elif not items:
            flash("Add at least one available component first.", "error")
        elif not request.form.get("responsibility"):
            flash("Please accept the responsibility statement before submitting.", "error")
        else:
            try:
                days = min(7, max(1, int(request.form.get("days", "3"))))
            except ValueError:
                days = 3
            due = (utcnow() + timedelta(days=days)).date().isoformat()
            year = utcnow().year
            count = query("SELECT COUNT(*) AS n FROM requests", one=True)["n"] + 1
            code = f"REQ-{year}-{count:06d}"
            user = current_user()
            req_id = execute(
                """INSERT INTO requests
                (request_code, student_id, project_title, purpose, mentor_name, mentor_email, mentor_phone,
                 hod_name, hod_email, hod_phone, due_date, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (code, user["id"], request.form.get("project_title", "").strip(),
                 request.form.get("purpose", "").strip(), request.form.get("mentor_name", "").strip(),
                 request.form.get("mentor_email", "").strip(), request.form.get("mentor_phone", "").strip(),
                 request.form.get("hod_name", "").strip(), request.form.get("hod_email", "").strip(),
                 request.form.get("hod_phone", "").strip(), due, utcnow().isoformat(), utcnow().isoformat()),
            )
            for item in items:
                execute(
                    "INSERT INTO request_items (request_id, inventory_id, quantity) VALUES (?, ?, ?)",
                    (req_id, item["id"], min(int(cart[str(item["id"])]), item["available_qty"])),
                )
            token = signed_token(req_id, "mentor")
            link = f"{APP_URL.rstrip('/') if APP_URL else request.host_url.rstrip('/')}{url_for('approval', stage='mentor', token=token)}"
            body = f"""A component request requires your review.

Student: {user['full_name']} ({user['uid']})
Project: {request.form.get('project_title', '').strip()}
Return date: {due}
Request ID: {code}"""
            sent = send_mail(request.form.get("mentor_email", "").strip(), "Component Request Requires Your Approval", body, link)
            notify(user["id"], "Request submitted", f"{code} is waiting for mentor approval.", "request")
            audit("REQUEST_CREATED", code)
            session["cart"] = {}
            flash(f"{code} submitted successfully. {'Mentor notification sent.' if sent else 'The request is saved; email delivery is not configured.'}", "success")
            return redirect(url_for("student_dashboard"))
    return render_template("request.html", items=items, cart=cart)


@app.route("/approval/<stage>/<token>", methods=["GET", "POST"])
def approval(stage, token):
    if stage not in {"mentor", "hod"}:
        return render_template("approval.html", error="That approval page could not be found.", stage=stage)
    result, error = token_payload(token, stage)
    if error:
        return render_template("approval.html", error=error, stage=stage)
    payload, token_row = result
    req = query(
        """SELECT requests.*, users.full_name AS student_name, users.uid, users.branch, users.division, users.year
        FROM requests JOIN users ON users.id=requests.student_id WHERE requests.id=?""",
        (payload["request_id"],), one=True,
    )
    items = request_items(req["id"])
    if request.method == "POST":
        if not valid_csrf():
            flash("Your approval session expired. Please try again.", "error")
        else:
            decision = request.form.get("decision")
            if decision not in {"approve", "reject"}:
                flash("Choose approve or reject.", "error")
            elif decision == "reject" and not request.form.get("reason", "").strip():
                flash("A rejection reason is required.", "error")
            else:
                now = utcnow().isoformat()
                if stage == "mentor":
                    if decision == "approve":
                        execute("UPDATE requests SET mentor_status='APPROVED', overall_status='PENDING_HOD', updated_at=? WHERE id=?", (now, req["id"]))
                        hod_token = signed_token(req["id"], "hod")
                        link = f"{APP_URL.rstrip('/') if APP_URL else request.host_url.rstrip('/')}{url_for('approval', stage='hod', token=hod_token)}"
                        send_mail(req["hod_email"], "HOD Approval Required — Mentor Has Approved", f"The Mentor has approved {req['request_code']}. Your approval is now required.", link)
                        notify(req["student_id"], "Mentor approved your request", f"{req['request_code']} is now waiting for HOD approval.", "approval")
                        audit("MENTOR_APPROVED", req["request_code"])
                        message = "MENTOR APPROVAL COMPLETED"
                    else:
                        reason = request.form["reason"].strip()
                        execute("UPDATE requests SET mentor_status='REJECTED', overall_status='REJECTED', rejection_reason=?, updated_at=? WHERE id=?", (reason, now, req["id"]))
                        notify(req["student_id"], "Request rejected by mentor", reason, "rejection")
                        audit("MENTOR_REJECTED", req["request_code"])
                        message = "MENTOR REJECTION RECORDED"
                else:
                    if decision == "approve":
                        execute("UPDATE requests SET hod_status='APPROVED', overall_status='FULLY_APPROVED', collection_otp_ready=1, updated_at=? WHERE id=?", (now, req["id"]))
                        notify(req["student_id"], "Ready for collection", f"{req['request_code']} is approved by Mentor and HOD. Generate an OTP when you are ready to collect.", "success")
                        audit("HOD_APPROVED", req["request_code"])
                        message = "HOD APPROVAL COMPLETED"
                    else:
                        reason = request.form["reason"].strip()
                        execute("UPDATE requests SET hod_status='REJECTED', overall_status='REJECTED', rejection_reason=?, updated_at=? WHERE id=?", (reason, now, req["id"]))
                        notify(req["student_id"], "Request rejected by HOD", reason, "rejection")
                        audit("HOD_REJECTED", req["request_code"])
                        message = "HOD REJECTION RECORDED"
                execute("UPDATE approval_tokens SET status='USED', used_at=? WHERE id=?", (now, token_row["id"]))
                return render_template("approval.html", complete=message, stage=stage, req=req, items=items)
    return render_template("approval.html", req=req, items=items, stage=stage)


def make_otp(request_id, kind):
    raw = f"{secrets.randbelow(1_000_000):06d}"
    now = utcnow()
    execute("UPDATE otp_tokens SET used_at=? WHERE request_id=? AND kind=? AND used_at IS NULL", (now.isoformat(), request_id, kind))
    execute(
        "INSERT INTO otp_tokens (request_id, kind, otp_hash, expires_at, created_at) VALUES (?, ?, ?, ?, ?)",
        (request_id, kind, hashlib.sha256(raw.encode()).hexdigest(), (now + timedelta(minutes=10)).isoformat(), now.isoformat()),
    )
    return raw


@app.post("/student/otp/<int:request_id>")
@login_required("student")
def generate_collection_otp(request_id):
    req = query("SELECT * FROM requests WHERE id=? AND student_id=? AND overall_status='FULLY_APPROVED'", (request_id, current_user()["id"]), one=True)
    if not req or not valid_csrf():
        flash("That request is not ready for collection.", "error")
    else:
        code = make_otp(request_id, "collection")
        notify(current_user()["id"], "Collection OTP generated", f"Your one-time collection code is {code}. It expires in 10 minutes.", "otp")
        flash(f"Collection OTP: {code} · valid for 10 minutes", "success")
    return redirect(url_for("student_dashboard"))


@app.route("/admin", methods=["GET", "POST"])
@login_required("admin")
def admin_dashboard():
    if request.method == "POST" and valid_csrf():
        action = request.form.get("action")
        if action == "issue":
            req = query("SELECT * FROM requests WHERE request_code=? AND overall_status='FULLY_APPROVED'", (request.form.get("request_code", "").strip(),), one=True)
            otp = query("SELECT * FROM otp_tokens WHERE request_id=? AND kind='collection' AND used_at IS NULL ORDER BY id DESC LIMIT 1", (req["id"],), one=True) if req else None
            if not req or not otp or datetime.fromisoformat(otp["expires_at"]) < utcnow() or otp["attempts"] >= 5 or not secrets.compare_digest(otp["otp_hash"], hashlib.sha256(request.form.get("otp", "").strip().encode()).hexdigest()):
                if otp:
                    execute("UPDATE otp_tokens SET attempts=attempts+1 WHERE id=?", (otp["id"],))
                flash("Invalid or expired collection OTP.", "error")
            else:
                items = request_items(req["id"])
                for item in items:
                    execute("UPDATE inventory SET available_qty=MAX(0, available_qty-?), updated_at=? WHERE name=?", (item["quantity"], utcnow().isoformat(), item["name"]))
                execute("UPDATE otp_tokens SET used_at=? WHERE id=?", (utcnow().isoformat(), otp["id"]))
                issue_id = execute("INSERT INTO issues (request_id, student_id, issued_at, due_date) VALUES (?, ?, ?, ?)", (req["id"], req["student_id"], utcnow().isoformat(), req["due_date"]))
                notify(req["student_id"], "Components issued", f"{req['request_code']} was collected. Return by {req['due_date']}.", "issue")
                audit("COMPONENTS_ISSUED", req["request_code"])
                flash(f"Collection completed for {req['request_code']}. Issue record #{issue_id} created.", "success")
        elif action == "add_inventory":
            name = request.form.get("name", "").strip()
            try:
                qty = max(0, int(request.form.get("quantity", "0")))
            except ValueError:
                qty = 0
            if name:
                execute(
                    """INSERT INTO inventory (source_file, source_key, name, category, description, location,
                    total_qty, available_qty, minimum_stock, created_at, updated_at) VALUES ('admin', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (f"admin:{secrets.token_hex(8)}", name, request.form.get("category", "Other"), "Added by Lab Assistant.", request.form.get("location", "Main Lab"), qty, qty, max(1, int(request.form.get("minimum", "1") or 1)), utcnow().isoformat(), utcnow().isoformat()),
                )
                audit("INVENTORY_CREATED", name)
                flash("Inventory item added.", "success")
    metrics = {
        "items": query("SELECT COUNT(*) AS n FROM inventory WHERE active=1", one=True)["n"],
        "units": query("SELECT COALESCE(SUM(available_qty),0) AS n FROM inventory WHERE active=1", one=True)["n"],
        "issued": query("SELECT COUNT(*) AS n FROM issues WHERE returned_at IS NULL", one=True)["n"],
        "low": query("SELECT COUNT(*) AS n FROM inventory WHERE active=1 AND available_qty <= minimum_stock AND available_qty > 0", one=True)["n"],
        "pending": query("SELECT COUNT(*) AS n FROM requests WHERE overall_status LIKE 'PENDING%'", one=True)["n"],
        "overdue": query("SELECT COUNT(*) AS n FROM issues WHERE returned_at IS NULL AND due_date < ?", (utcnow().date().isoformat(),), one=True)["n"],
    }
    inventory = query("SELECT * FROM inventory WHERE active=1 ORDER BY name LIMIT 60")
    recent = query(
        """SELECT requests.*, users.full_name AS student_name FROM requests JOIN users ON users.id=requests.student_id
        ORDER BY requests.created_at DESC LIMIT 8"""
    )
    return render_template("admin.html", metrics=metrics, inventory=inventory, recent=recent)


@app.post("/admin/inventory/<int:item_id>/toggle")
@login_required("admin")
def toggle_inventory(item_id):
    if valid_csrf():
        execute("UPDATE inventory SET active=CASE active WHEN 1 THEN 0 ELSE 1 END, updated_at=? WHERE id=?", (utcnow().isoformat(), item_id))
        flash("Inventory status updated.", "success")
    return redirect(url_for("admin_dashboard"))


@app.post("/student/return/<int:issue_id>")
@login_required("student")
def generate_return_otp(issue_id):
    issue = query("SELECT * FROM issues WHERE id=? AND student_id=? AND returned_at IS NULL", (issue_id, current_user()["id"]), one=True)
    if issue and valid_csrf():
        code = make_otp(issue["request_id"], "return")
        notify(current_user()["id"], "Return OTP generated", f"Your return code is {code}. Share it with the Lab Assistant within 10 minutes.", "otp")
        flash(f"Return OTP: {code} · valid for 10 minutes", "success")
    else:
        flash("That issue is no longer active.", "error")
    return redirect(url_for("student_dashboard"))


@app.route("/admin/returns", methods=["GET", "POST"])
@login_required("admin")
def returns():
    if request.method == "POST" and valid_csrf():
        req = query("SELECT * FROM requests WHERE request_code=?", (request.form.get("request_code", "").strip(),), one=True)
        otp = query("SELECT * FROM otp_tokens WHERE request_id=? AND kind='return' AND used_at IS NULL ORDER BY id DESC LIMIT 1", (req["id"],), one=True) if req else None
        if not req or not otp or datetime.fromisoformat(otp["expires_at"]) < utcnow() or not secrets.compare_digest(otp["otp_hash"], hashlib.sha256(request.form.get("otp", "").strip().encode()).hexdigest()):
            flash("Invalid or expired return OTP.", "error")
        else:
            issue = query("SELECT * FROM issues WHERE request_id=? AND returned_at IS NULL", (req["id"],), one=True)
            if issue:
                condition = request.form.get("condition", "GOOD")
                execute("UPDATE issues SET returned_at=?, return_condition=?, return_remarks=? WHERE id=?", (utcnow().isoformat(), condition, request.form.get("remarks", "").strip(), issue["id"]))
                for item in request_items(req["id"]):
                    execute("UPDATE inventory SET available_qty=MIN(total_qty, available_qty+?), updated_at=? WHERE name=?", (item["quantity"], utcnow().isoformat(), item["name"]))
                execute("UPDATE otp_tokens SET used_at=? WHERE id=?", (utcnow().isoformat(), otp["id"]))
                notify(req["student_id"], "Components returned", f"Return completed for {req['request_code']} with condition {condition}.", "return")
                audit("COMPONENTS_RETURNED", req["request_code"])
                flash("Return completed and inventory updated.", "success")
    active = query(
        """SELECT issues.*, requests.request_code, users.full_name FROM issues
        JOIN requests ON requests.id=issues.request_id JOIN users ON users.id=issues.student_id
        WHERE issues.returned_at IS NULL ORDER BY issues.due_date"""
    )
    return render_template("returns.html", active=active)


@app.route("/notifications/read-all", methods=["POST"])
@login_required()
def read_notifications():
    if valid_csrf():
        execute("UPDATE notifications SET is_read=1 WHERE user_id=?", (current_user()["id"],))
    return redirect(request.referrer or url_for("dashboard"))


@app.route("/student/assistant", methods=["POST"])
@login_required("student")
def assistant():
    prompt = f"{request.form.get('title', '')} {request.form.get('purpose', '')}".lower()
    words = [word for word in prompt.replace(",", " ").split() if len(word) > 3]
    found = []
    for word in words[:12]:
        found.extend(query("SELECT * FROM inventory WHERE active=1 AND name LIKE ? LIMIT 3", (f"%{word}%",)))
    if not found:
        found = query("SELECT * FROM inventory WHERE active=1 AND available_qty > 0 ORDER BY available_qty DESC LIMIT 4")
    unique = {row["id"]: row for row in found}
    return render_template("assistant.html", recommendations=list(unique.values())[:8], title=request.form.get("title", "Your project"))


@app.errorhandler(404)
def not_found(_error):
    return render_template("error.html", code=404, title="Page not found", message="The page you requested is not available."), 404


@app.errorhandler(500)
def server_error(_error):
    return render_template("error.html", code=500, title="Something went wrong", message="We could not complete that action. Your data has not been removed."), 500


with app.app_context():
    init_db()


if __name__ == "__main__":
    port = int(os.getenv("PORT", "5000"))
    print("LabVault AI Lab Inventory Management")
    print(f"Server running at http://127.0.0.1:{port}")
    app.run(host="0.0.0.0", port=port, debug=False)