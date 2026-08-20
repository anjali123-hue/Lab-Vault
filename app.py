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
from zoneinfo import ZoneInfo
import csv
from io import StringIO
import re
import time
from flask import has_request_context
from openpyxl import load_workbook
import os
import sqlite3
import psycopg
from psycopg.rows import dict_row
import smtplib
import shutil
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


IST = ZoneInfo("Asia/Kolkata")


def format_ist(value):
    if not value:
        return ""

    try:
        dt = datetime.fromisoformat(value)

        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)

        return dt.astimezone(IST).strftime("%Y-%m-%d %H:%M")

    except (ValueError, TypeError):
        return value

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
    jsonify,
    Response,
)
from itsdangerous import BadSignature, SignatureExpired, URLSafeTimedSerializer
from werkzeug.security import check_password_hash, generate_password_hash

load_dotenv()

ROOT = Path(__file__).resolve().parent

DATABASE_URL = os.getenv("DATABASE_URL", "").strip()

DB_PATH = Path(
    os.getenv(
        "SQLITE_PATH",
        ROOT / "data" / "labvault.db"
    )
)

print("==========================================")
print("LABVAULT DATABASE:", DB_PATH.resolve())
print("==========================================")

DB_PATH.parent.mkdir(parents=True, exist_ok=True)
DB_PATH.parent.mkdir(parents=True, exist_ok=True)
SECRET_KEY = os.getenv("SECRET_KEY", secrets.token_hex(32))
MAIL_SENDER = os.getenv("MAIL_USERNAME", "labvault.lab@gmail.com")
MAIL_PASSWORD = os.getenv("MAIL_PASSWORD", "")
ADMIN_USERNAME = os.getenv("ADMIN_USERNAME", "admin")
ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD", "change-this-admin-password")
HOD_EMAIL = os.getenv("HOD_EMAIL", "")
HOD_NAME = os.getenv("HOD_NAME", "Head of Department").strip()
HOD_PASSWORD = os.getenv("HOD_PASSWORD","")
APP_URL = os.getenv("APP_URL", "")
serializer = URLSafeTimedSerializer(SECRET_KEY)

app = Flask(__name__)

@app.template_filter("ist")
def ist_filter(value):
    return format_ist(value)

app.config.update(
    SECRET_KEY=SECRET_KEY,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.getenv("SESSION_COOKIE_SECURE", "0") == "1",
    PERMANENT_SESSION_LIFETIME=timedelta(hours=8),
)
DB_LOCK = threading.Lock()
PG_BOOTSTRAP = False


def utcnow():
    return datetime.now(timezone.utc).replace(microsecond=0)


def adapt_sql(sql):
    """
    Convert the application's SQLite-style '?' placeholders
    to PostgreSQL '%s' placeholders when PostgreSQL is active.

    SQLite remains unchanged.
    """
    if DATABASE_URL:
        return sql.replace("?", "%s")

    return sql


def db():
    if "db" not in g:

        if DATABASE_URL:

            g.db = psycopg.connect(
                DATABASE_URL,
                row_factory=dict_row,
            )

        else:

            g.db = sqlite3.connect(
                DB_PATH
            )

            g.db.row_factory = sqlite3.Row

            g.db.execute(
                "PRAGMA foreign_keys = ON"
            )

    return g.db


@app.teardown_appcontext
def close_db(_error=None):

    conn = g.pop("db", None)

    if conn:
        conn.close()


def query(sql, args=(), one=False):

    sql = adapt_sql(sql)

    cur = db().execute(
        sql,
        args
    )

    rows = (
        cur.fetchone()
        if one
        else cur.fetchall()
    )

    cur.close()

    return rows


def execute(sql, args=()):
    global PG_BOOTSTRAP

    sql = adapt_sql(sql)

    with DB_LOCK:

        # -----------------------------------------------------
        # PostgreSQL
        # -----------------------------------------------------

        if DATABASE_URL:

            clean_sql = sql.strip().rstrip(";")

            is_insert = (
                clean_sql
                .lstrip()
                .lower()
                .startswith("insert")
            )

            if is_insert:

                if " returning " not in clean_sql.lower():
                    clean_sql += " RETURNING id"

                cur = db().execute(
                    clean_sql,
                    args
                )

                row = cur.fetchone()

                # During fresh PostgreSQL initialization,
                # all Excel imports are committed together.
                if not PG_BOOTSTRAP:
                    db().commit()

                cur.close()

                if row:
                    return row["id"]

                return None

            cur = db().execute(
                sql,
                args
            )

            if not PG_BOOTSTRAP:
                db().commit()

            cur.close()

            return None

        # -----------------------------------------------------
        # SQLite
        # -----------------------------------------------------

        cur = db().execute(
            sql,
            args
        )

        db().commit()

        last = cur.lastrowid

        cur.close()

        return last

def sync_faculty_from_excel():
    """
    Import/update IoT faculty from the Excel file.

    Excel structure:
        Row 1 -> IoT Department Stafflist
        Row 2 -> Sr. No. | Name of Faculty | Designation | Email-Id
        Row 3+ -> Faculty records

    The Excel file is used as the initial/official source.
    The database remains the live source for the application.

    Faculty records are matched by email.
    Existing records are updated.
    New records are inserted.
    Existing active/inactive status is preserved.
    No faculty records are deleted automatically.
    """

    faculty_file = os.path.join(
        BASE_DIR,
        "attached_assets",
        "iot_faculty.xlsx",
    )

    if not os.path.isfile(faculty_file):
        app.logger.warning(
            "IoT faculty Excel file not found: %s",
            faculty_file,
        )
        return

    try:
        workbook = load_workbook(
            filename=faculty_file,
            data_only=True,
            read_only=True,
        )

        sheet = workbook.active

        rows = list(
            sheet.iter_rows(
                values_only=True
            )
        )

        workbook.close()

        # -----------------------------------------------------
        # Minimum structure check
        # -----------------------------------------------------

        if len(rows) < 3:

            app.logger.warning(
                "IoT faculty Excel file does not contain "
                "the expected title, header and data rows."
            )

            return

        # -----------------------------------------------------
        # Row 1 = title
        # Row 2 = actual headers
        # -----------------------------------------------------

        headers = [
            str(value).strip()
            if value is not None
            else ""
            for value in rows[1]
        ]

        expected_headers = [
            "Sr. No.",
            "Name of Faculty",
            "Designation",
            "Email-Id",
        ]

        # Compare only the first four columns because
        # the sheet may contain unused columns after Email-Id.
        if headers[:4] != expected_headers:

            app.logger.error(
                "Faculty Excel headers do not match expected "
                "headers. Expected=%s Found=%s",
                expected_headers,
                headers,
            )

            return

        # Exact column positions from the Excel sheet.
        serial_col = 0
        name_col = 1
        designation_col = 2
        email_col = 3

        imported = 0
        updated = 0
        skipped = 0

        # -----------------------------------------------------
        # Row 3 onward contains faculty records
        # -----------------------------------------------------

        for row in rows[2:]:

            if not row:
                skipped += 1
                continue

            def cell(index):
                if index >= len(row):
                    return ""

                value = row[index]

                if value is None:
                    return ""

                return str(value).strip()

            serial_no = cell(serial_col)
            full_name = cell(name_col)
            designation = cell(designation_col)
            email = cell(email_col).lower()

            # -------------------------------------------------
            # Ignore completely blank rows
            # -------------------------------------------------

            if not full_name and not designation and not email:
                skipped += 1
                continue

            # -------------------------------------------------
            # Name and email are required
            # -------------------------------------------------

            if not full_name or not email:

                app.logger.warning(
                    "Skipping faculty row with missing "
                    "name/email. Sr. No.=%s Name=%s Email=%s",
                    serial_no,
                    full_name,
                    email,
                )

                skipped += 1
                continue

            # -------------------------------------------------
            # Basic email validation
            # -------------------------------------------------

            if "@" not in email:

                app.logger.warning(
                    "Skipping faculty row with invalid email: %s",
                    email,
                )

                skipped += 1
                continue

            # -------------------------------------------------
            # Match existing faculty by email
            # -------------------------------------------------

            existing = query(
                """
                SELECT
                    id,
                    active
                FROM faculty
                WHERE lower(email)=?
                LIMIT 1
                """,
                (email,),
                one=True,
            )

            now = utcnow().isoformat()

            # -------------------------------------------------
            # Existing faculty -> update details
            # -------------------------------------------------

            if existing:

                execute(
                    """
                    UPDATE faculty
                    SET
                        full_name=?,
                        designation=?,
                        department='IoT',
                        updated_at=?
                    WHERE id=?
                    """,
                    (
                        full_name,
                        designation,
                        now,
                        existing["id"],
                    ),
                )

                updated += 1

            # -------------------------------------------------
            # New faculty -> insert
            # -------------------------------------------------

            else:

                execute(
                    """
                    INSERT INTO faculty
                    (
                        full_name,
                        designation,
                        department,
                        email,
                        active,
                        created_at,
                        updated_at
                    )
                    VALUES
                    (
                        ?, ?, 'IoT', ?, 1, ?, ?
                    )
                    """,
                    (
                        full_name,
                        designation,
                        email,
                        now,
                        now,
                    ),
                )

                imported += 1

        db().commit()

        app.logger.info(
            "IoT faculty sync completed. "
            "Inserted=%s Updated=%s Skipped=%s",
            imported,
            updated,
            skipped,
        )

    except PermissionError:

        app.logger.exception(
            "Permission denied while reading the IoT faculty Excel file: %s",
            faculty_file,
        )

    except Exception:

        app.logger.exception(
            "IoT faculty Excel sync failed."
        )

def init_db_sqlite():
    with app.app_context():
        db().executescript(
            """
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                role TEXT NOT NULL CHECK(role IN ('student','admin')),
                full_name TEXT NOT NULL,
                erp_id TEXT UNIQUE,
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
                return_days INTEGER NOT NULL DEFAULT 3,
                admin_status TEXT NOT NULL DEFAULT 'PENDING',
                mentor_status TEXT NOT NULL DEFAULT 'PENDING',
                hod_status TEXT NOT NULL DEFAULT 'PENDING',
                overall_status TEXT NOT NULL DEFAULT 'PENDING_ADMIN',
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
            CREATE TABLE IF NOT EXISTS password_reset_tokens (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
                token_hash TEXT NOT NULL UNIQUE,
                status TEXT NOT NULL DEFAULT 'PENDING',
                created_at TEXT NOT NULL,
                used_at TEXT
            );
            CREATE TABLE IF NOT EXISTS admin_password_requests (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                admin_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
                new_password_hash TEXT NOT NULL,
                token_hash TEXT NOT NULL UNIQUE,
                status TEXT NOT NULL DEFAULT 'PENDING',
                created_at TEXT NOT NULL,
                used_at TEXT
            );
            CREATE TABLE IF NOT EXISTS admin_password_change_requests (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                admin_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
                reason TEXT NOT NULL,
                token_hash TEXT NOT NULL UNIQUE,
                status TEXT NOT NULL DEFAULT 'PENDING',
                created_at TEXT NOT NULL,
                authorized_at TEXT
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
            CREATE TABLE IF NOT EXISTS past_component_usage (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                source_key TEXT UNIQUE NOT NULL,
                student_name TEXT,
                student_email TEXT,
                year TEXT,
                department TEXT,
                division TEXT,
                project_name TEXT,
                component_name TEXT NOT NULL,
                quantity INTEGER NOT NULL DEFAULT 1,
                issue_date TEXT,
                returned_date TEXT,
                condition TEXT,
                phone TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );
            """
        )
        db().commit()

        # =========================================================
        # HOD / FACULTY MANAGEMENT
        # =========================================================

        db().execute(
            """
            CREATE TABLE IF NOT EXISTS faculty (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                full_name TEXT NOT NULL,
                designation TEXT,
                department TEXT NOT NULL DEFAULT 'IoT',
                email TEXT NOT NULL UNIQUE,
                phone TEXT,
                active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )

        db().commit()
        sync_faculty_from_excel()


        # =========================================================
        # HOD FLAG FOR USERS
        # =========================================================

        user_columns = {
            row["name"]
            for row in db().execute(
                "PRAGMA table_info(users)"
            ).fetchall()
        }

        if "is_hod" not in user_columns:
            db().execute(
                """
                ALTER TABLE users
                ADD COLUMN is_hod INTEGER NOT NULL DEFAULT 0
                """
            )
            db().commit()


        # =========================================================
        # REQUEST MENTOR VERIFICATION
        # =========================================================

        request_columns = {
            row["name"]
            for row in db().execute(
                "PRAGMA table_info(requests)"
            ).fetchall()
        }

        if "faculty_id" not in request_columns:
            db().execute(
                """
                ALTER TABLE requests
                ADD COLUMN faculty_id INTEGER
                """
            )
            db().commit()

        if "mentor_type" not in request_columns:
            db().execute(
                """
                ALTER TABLE requests
                ADD COLUMN mentor_type TEXT
                DEFAULT 'faculty'
                """
            )
            db().commit()

        if "mentor_verified" not in request_columns:
            db().execute(
                """
                ALTER TABLE requests
                ADD COLUMN mentor_verified INTEGER
                NOT NULL DEFAULT 1
                """
            )
            db().commit()

        if "mentor_verified_at" not in request_columns:
            db().execute(
                """
                ALTER TABLE requests
                ADD COLUMN mentor_verified_at TEXT
                """
            )
            db().commit()

        if "mentor_verified_by" not in request_columns:
            db().execute(
                """
                ALTER TABLE requests
                ADD COLUMN mentor_verified_by INTEGER
                """
            )
            db().commit()

        if "mentor_verification_reason" not in request_columns:
            db().execute(
                """
                ALTER TABLE requests
                ADD COLUMN mentor_verification_reason TEXT
                """
            )
            db().commit()

                # --------------------------------------------------
        # SAFE MIGRATION: ADD ERP ID
        # --------------------------------------------------

        user_columns = {
            row["name"]
            for row in db().execute(
                "PRAGMA table_info(users)"
            ).fetchall()
        }

        if "erp_id" not in user_columns:
            db().execute(
                "ALTER TABLE users ADD COLUMN erp_id TEXT"
            )
            db().commit()

        # Unique ERP ID for students
        db().execute(
            """
            CREATE UNIQUE INDEX IF NOT EXISTS
            idx_users_erp_id
            ON users(erp_id)
            WHERE erp_id IS NOT NULL
            """
        )
        db().commit()

        password_change_columns = {
            row["name"]
            for row in db().execute(
                "PRAGMA table_info(admin_password_change_requests)"
            ).fetchall()
        }

        if "reason" not in password_change_columns:
            db().execute(
                "ALTER TABLE admin_password_change_requests ADD COLUMN reason TEXT"
            )
            db().commit()

        if "status" not in password_change_columns:
            db().execute(
                "ALTER TABLE admin_password_change_requests "
                "ADD COLUMN status TEXT NOT NULL DEFAULT 'PENDING'"
            )
            db().commit()

        if "authorized_at" not in password_change_columns:
            db().execute(
                "ALTER TABLE admin_password_change_requests "
                "ADD COLUMN authorized_at TEXT"
            )
            db().commit()

        request_columns = {row["name"] for row in db().execute("PRAGMA table_info(requests)").fetchall()}
        if "return_days" not in request_columns:
            db().execute("ALTER TABLE requests ADD COLUMN return_days INTEGER NOT NULL DEFAULT 3")
            db().commit()
        if "admin_status" not in request_columns:
            db().execute("ALTER TABLE requests ADD COLUMN admin_status TEXT NOT NULL DEFAULT 'PENDING'")
            db().commit()
            db().execute(
                "UPDATE requests SET admin_status='APPROVED' WHERE overall_status NOT IN ('PENDING_ADMIN','REJECTED')"
            )
            db().commit()

        issue_columns = {row["name"] for row in db().execute("PRAGMA table_info(issues)").fetchall()}
        if "extensions_used" not in issue_columns:
            db().execute("ALTER TABLE issues ADD COLUMN extensions_used INTEGER NOT NULL DEFAULT 0")
            db().commit()
            
        if "overdue_notified_at" not in issue_columns:
            db().execute("ALTER TABLE issues ADD COLUMN overdue_notified_at TEXT")
            db().commit()

        if "due_notified_at" not in issue_columns:
            db().execute("ALTER TABLE issues ADD COLUMN due_notified_at TEXT")
            db().commit()

        if "extension_reason" not in issue_columns:
            db().execute("ALTER TABLE issues ADD COLUMN extension_reason TEXT")
            db().commit()
        if "extension_requested_days" not in issue_columns:
            db().execute("ALTER TABLE issues ADD COLUMN extension_requested_days INTEGER")
            db().commit()
        if "extension_status" not in issue_columns:
            db().execute("ALTER TABLE issues ADD COLUMN extension_status TEXT NOT NULL DEFAULT 'NONE'")
            db().commit()
        if "extension_admin_status" not in issue_columns:
            db().execute("ALTER TABLE issues ADD COLUMN extension_admin_status TEXT NOT NULL DEFAULT 'PENDING'")
            db().commit()
        if "extension_mentor_status" not in issue_columns:
            db().execute("ALTER TABLE issues ADD COLUMN extension_mentor_status TEXT NOT NULL DEFAULT 'PENDING'")
            db().commit()
        if "extension_hod_status" not in issue_columns:
            db().execute("ALTER TABLE issues ADD COLUMN extension_hod_status TEXT NOT NULL DEFAULT 'PENDING'")
            db().commit()

        otp_columns = {
                        row["name"]
                        for row in db().execute("PRAGMA table_info(otp_tokens)").fetchall()
}

        if "student_id" not in otp_columns:
            db().execute(
                "ALTER TABLE otp_tokens ADD COLUMN student_id INTEGER"
            )
            db().commit()

            # Populate the new column for existing OTP records
            db().execute(
                """UPDATE otp_tokens
                SET student_id = (
                    SELECT student_id
                    FROM requests
                    WHERE requests.id = otp_tokens.request_id
                )
                WHERE student_id IS NULL"""
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
        elif ADMIN_PASSWORD and ADMIN_PASSWORD != "change-this-admin-password":
            existing_admin = query("SELECT id, password_hash FROM users WHERE role='admin' LIMIT 1", one=True)
            if existing_admin and check_password_hash(existing_admin["password_hash"], "change-this-admin-password"):
                execute(
                    "UPDATE users SET password_hash=?, updated_at=? WHERE id=?"
                    if "updated_at" in {row["name"] for row in db().execute("PRAGMA table_info(users)").fetchall()}
                    else "UPDATE users SET password_hash=? WHERE id=?",
                    (generate_password_hash(ADMIN_PASSWORD), utcnow().isoformat(), existing_admin["id"])
                    if "updated_at" in {row["name"] for row in db().execute("PRAGMA table_info(users)").fetchall()}
                    else (generate_password_hash(ADMIN_PASSWORD), existing_admin["id"]),
                )
        import_inventory()
        import_past_component_usage()
        sync_kits_from_excel()

        # =========================================================
        # HOD ACCOUNT
        # =========================================================

        if HOD_EMAIL:

            # First, remove the HOD flag from any previous HOD account.
            execute(
                """
                UPDATE users
                SET is_hod=0
                WHERE is_hod=1
                """,
            )

            # Find the account for the currently configured HOD email.
            hod = query(
                """
                SELECT
                    id,
                    email,
                    password_hash
                FROM users
                WHERE lower(email)=?
                LIMIT 1
                """,
                (HOD_EMAIL.lower(),),
                one=True,
            )

            if hod:

                # Make this account the current HOD account.
                execute(
                    """
                    UPDATE users
                    SET
                        full_name=?,
                        is_hod=1,
                        role='admin'
                    WHERE id=?
                    """,
                    (
                        HOD_NAME,
                        hod["id"],
                    ),
                )

                # Set the configured HOD password.
                if HOD_PASSWORD:

                    execute(
                        """
                        UPDATE users
                        SET password_hash=?
                        WHERE id=?
                        """,
                        (
                            generate_password_hash(
                                HOD_PASSWORD
                            ),
                            hod["id"],
                        ),
                    )

            elif HOD_PASSWORD:

                # No account exists yet for the new HOD email.
                execute(
                    """
                    INSERT INTO users
                    (
                        role,
                        full_name,
                        email,
                        password_hash,
                        is_hod,
                        created_at
                    )
                    VALUES (?, ?, ?, ?, 1, ?)
                    """,
                    (
                        "admin",
                        HOD_NAME,
                        HOD_EMAIL.lower(),
                        generate_password_hash(
                            HOD_PASSWORD
                        ),
                        utcnow().isoformat(),
                    ),
                )

def init_db_postgres():
    """
    Initialize a completely fresh PostgreSQL LabVault database.

    No data is copied from local labvault.db.

    Excel data is imported:
        - IoT faculty
        - inventory
        - past component usage

    Transactional application data starts fresh.
    """

    global PG_BOOTSTRAP

    schema = """
    CREATE TABLE IF NOT EXISTS users (
        id BIGSERIAL PRIMARY KEY,
        role TEXT NOT NULL CHECK(role IN ('student','admin')),
        full_name TEXT NOT NULL,
        erp_id TEXT UNIQUE,
        uid TEXT UNIQUE,
        email TEXT UNIQUE,
        phone TEXT,
        branch TEXT,
        division TEXT,
        year TEXT,
        password_hash TEXT NOT NULL,
        created_at TEXT NOT NULL,
        is_hod INTEGER NOT NULL DEFAULT 0
    );

    CREATE TABLE IF NOT EXISTS inventory (
        id BIGSERIAL PRIMARY KEY,
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
        id BIGSERIAL PRIMARY KEY,
        request_code TEXT NOT NULL UNIQUE,
        student_id BIGINT NOT NULL REFERENCES users(id),
        project_title TEXT NOT NULL,
        purpose TEXT NOT NULL,
        mentor_name TEXT NOT NULL,
        mentor_email TEXT NOT NULL,
        mentor_phone TEXT,
        hod_name TEXT NOT NULL,
        hod_email TEXT NOT NULL,
        hod_phone TEXT,
        due_date TEXT NOT NULL,
        return_days INTEGER NOT NULL DEFAULT 3,
        admin_status TEXT NOT NULL DEFAULT 'PENDING',
        mentor_status TEXT NOT NULL DEFAULT 'PENDING',
        hod_status TEXT NOT NULL DEFAULT 'PENDING',
        overall_status TEXT NOT NULL DEFAULT 'PENDING_ADMIN',
        rejection_reason TEXT,
        collection_otp_ready INTEGER NOT NULL DEFAULT 0,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL,
        faculty_id BIGINT,
        mentor_type TEXT DEFAULT 'faculty',
        mentor_verified INTEGER NOT NULL DEFAULT 1,
        mentor_verified_at TEXT,
        mentor_verified_by BIGINT,
        mentor_verification_reason TEXT
    );

    CREATE TABLE IF NOT EXISTS request_items (
        id BIGSERIAL PRIMARY KEY,
        request_id BIGINT NOT NULL REFERENCES requests(id) ON DELETE CASCADE,
        inventory_id BIGINT NOT NULL REFERENCES inventory(id),
        quantity INTEGER NOT NULL,
        UNIQUE(request_id, inventory_id)
    );

    CREATE TABLE IF NOT EXISTS approval_tokens (
        id BIGSERIAL PRIMARY KEY,
        request_id BIGINT NOT NULL REFERENCES requests(id) ON DELETE CASCADE,
        stage TEXT NOT NULL,
        token_hash TEXT NOT NULL UNIQUE,
        status TEXT NOT NULL DEFAULT 'PENDING',
        created_at TEXT NOT NULL,
        used_at TEXT
    );

    CREATE TABLE IF NOT EXISTS password_reset_tokens (
        id BIGSERIAL PRIMARY KEY,
        user_id BIGINT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
        token_hash TEXT NOT NULL UNIQUE,
        status TEXT NOT NULL DEFAULT 'PENDING',
        created_at TEXT NOT NULL,
        used_at TEXT
    );

    CREATE TABLE IF NOT EXISTS admin_password_requests (
        id BIGSERIAL PRIMARY KEY,
        admin_id BIGINT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
        new_password_hash TEXT NOT NULL,
        token_hash TEXT NOT NULL UNIQUE,
        status TEXT NOT NULL DEFAULT 'PENDING',
        created_at TEXT NOT NULL,
        used_at TEXT
    );

    CREATE TABLE IF NOT EXISTS admin_password_change_requests (
        id BIGSERIAL PRIMARY KEY,
        admin_id BIGINT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
        reason TEXT NOT NULL,
        token_hash TEXT NOT NULL UNIQUE,
        status TEXT NOT NULL DEFAULT 'PENDING',
        created_at TEXT NOT NULL,
        authorized_at TEXT
    );

    CREATE TABLE IF NOT EXISTS otp_tokens (
        id BIGSERIAL PRIMARY KEY,
        request_id BIGINT NOT NULL REFERENCES requests(id) ON DELETE CASCADE,
        kind TEXT NOT NULL,
        otp_hash TEXT NOT NULL,
        expires_at TEXT NOT NULL,
        attempts INTEGER NOT NULL DEFAULT 0,
        used_at TEXT,
        created_at TEXT NOT NULL,
        student_id BIGINT
    );

    CREATE TABLE IF NOT EXISTS issues (
        id BIGSERIAL PRIMARY KEY,
        request_id BIGINT NOT NULL REFERENCES requests(id),
        student_id BIGINT NOT NULL REFERENCES users(id),
        issued_at TEXT NOT NULL,
        due_date TEXT NOT NULL,
        returned_at TEXT,
        return_condition TEXT,
        return_remarks TEXT,
        extensions_used INTEGER NOT NULL DEFAULT 0,
        overdue_notified_at TEXT,
        due_notified_at TEXT,
        extension_reason TEXT,
        extension_requested_days INTEGER,
        extension_status TEXT NOT NULL DEFAULT 'NONE',
        extension_admin_status TEXT NOT NULL DEFAULT 'PENDING',
        extension_mentor_status TEXT NOT NULL DEFAULT 'PENDING',
        extension_hod_status TEXT NOT NULL DEFAULT 'PENDING'
    );

    CREATE TABLE IF NOT EXISTS notifications (
        id BIGSERIAL PRIMARY KEY,
        user_id BIGINT NOT NULL REFERENCES users(id),
        title TEXT NOT NULL,
        body TEXT NOT NULL,
        kind TEXT NOT NULL,
        is_read INTEGER NOT NULL DEFAULT 0,
        created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS audit_logs (
        id BIGSERIAL PRIMARY KEY,
        user_id BIGINT,
        role TEXT,
        action TEXT NOT NULL,
        description TEXT,
        created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS past_component_usage (
        id BIGSERIAL PRIMARY KEY,
        source_key TEXT UNIQUE NOT NULL,
        student_name TEXT,
        student_email TEXT,
        year TEXT,
        department TEXT,
        division TEXT,
        project_name TEXT,
        component_name TEXT NOT NULL,
        quantity INTEGER NOT NULL DEFAULT 1,
        issue_date TEXT,
        returned_date TEXT,
        condition TEXT,
        phone TEXT,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS faculty (
        id BIGSERIAL PRIMARY KEY,
        full_name TEXT NOT NULL,
        designation TEXT,
        department TEXT NOT NULL DEFAULT 'IoT',
        email TEXT NOT NULL UNIQUE,
        phone TEXT,
        active INTEGER NOT NULL DEFAULT 1,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL
    );
    """

    try:

        # -----------------------------------------------------
        # CREATE SCHEMA
        # -----------------------------------------------------

        db().execute(schema)

        # -----------------------------------------------------
        # START ONE TRANSACTION FOR INITIAL DATA
        # -----------------------------------------------------

        PG_BOOTSTRAP = True

        # -----------------------------------------------------
        # FRESH LAB ASSISTANT
        # -----------------------------------------------------

        admin = query(
            """
            SELECT id
            FROM users
            WHERE role='admin'
            AND COALESCE(is_hod, 0)=0
            LIMIT 1
            """,
            one=True,
        )

        if not admin:

            execute(
                """
                INSERT INTO users
                (
                    role,
                    full_name,
                    email,
                    password_hash,
                    is_hod,
                    created_at
                )
                VALUES (?, ?, ?, ?, 0, ?)
                """,
                (
                    "admin",
                    "Lab Assistant",
                    MAIL_SENDER.lower(),
                    generate_password_hash(
                        ADMIN_PASSWORD
                    ),
                    utcnow().isoformat(),
                ),
            )

        # -----------------------------------------------------
        # FRESH HOD
        # -----------------------------------------------------

        if HOD_EMAIL and HOD_PASSWORD:

            hod = query(
                """
                SELECT id
                FROM users
                WHERE lower(email)=?
                LIMIT 1
                """,
                (
                    HOD_EMAIL.lower(),
                ),
                one=True,
            )

            if not hod:

                execute(
                    """
                    INSERT INTO users
                    (
                        role,
                        full_name,
                        email,
                        password_hash,
                        is_hod,
                        created_at
                    )
                    VALUES (?, ?, ?, ?, 1, ?)
                    """,
                    (
                        "admin",
                        HOD_NAME,
                        HOD_EMAIL.lower(),
                        generate_password_hash(
                            HOD_PASSWORD
                        ),
                        utcnow().isoformat(),
                    ),
                )

        # -----------------------------------------------------
        # EXCEL DATA
        # -----------------------------------------------------

        sync_faculty_from_excel()
        import_inventory()
        import_past_component_usage()

        # -----------------------------------------------------
        # COMMIT EVERYTHING TOGETHER
        # -----------------------------------------------------

        db().commit()

        PG_BOOTSTRAP = False

        app.logger.info(
            "Fresh PostgreSQL database initialized successfully."
        )

    except Exception:

        db().rollback()

        PG_BOOTSTRAP = False

        app.logger.exception(
            "Fresh PostgreSQL initialization failed."
        )

        raise

def init_db():

    if DATABASE_URL:

        init_db_postgres()

    else:

        init_db_sqlite()

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

def sync_kits_from_excel():
    """
    Import new kits from attached_assets/kits.xlsx.

    Expected columns:
        Name
        Quantity

    Kits are stored in the normal inventory table so students
    can search for and request them like other inventory items.

    Existing kit records are NOT overwritten on application restart.
    This prevents the Excel quantity from resetting live inventory
    after students/admins have already used or adjusted stock.

    New kits found in the Excel file are inserted automatically.
    """

    kits_file = os.path.join(
        BASE_DIR,
        "attached_assets",
        "kits.xlsx",
    )

    if not os.path.isfile(kits_file):
        app.logger.warning(
            "Kits Excel file not found: %s",
            kits_file,
        )
        return

    try:
        workbook = load_workbook(
            filename=kits_file,
            data_only=True,
            read_only=True,
        )

        sheet = workbook.active

        rows = list(
            sheet.iter_rows(
                values_only=True
            )
        )

        workbook.close()

        # -----------------------------------------------------
        # Basic file check
        # -----------------------------------------------------

        if len(rows) < 2:

            app.logger.warning(
                "kits.xlsx does not contain any kit records."
            )

            return

        # -----------------------------------------------------
        # Expected headers:
        # Name | Quantity
        # -----------------------------------------------------

        headers = [
            str(value).strip()
            if value is not None
            else ""
            for value in rows[0]
        ]

        if (
            "Name" not in headers
            or "Quantity" not in headers
        ):

            app.logger.error(
                "kits.xlsx must contain 'Name' and 'Quantity' columns. "
                "Found: %s",
                headers,
            )

            return

        name_col = headers.index("Name")
        quantity_col = headers.index("Quantity")

        inserted = 0
        skipped = 0

        now = utcnow().isoformat()

        # -----------------------------------------------------
        # Read kit rows
        # -----------------------------------------------------

        for row in rows[1:]:

            if not row:
                continue

            if name_col >= len(row):
                skipped += 1
                continue

            name_value = row[name_col]

            if name_value is None:
                skipped += 1
                continue

            name = str(
                name_value
            ).strip()

            if not name:
                skipped += 1
                continue

            # -------------------------------------------------
            # Read quantity
            # -------------------------------------------------

            if quantity_col >= len(row):
                skipped += 1
                continue

            raw_quantity = row[quantity_col]

            try:

                quantity = int(
                    float(
                        raw_quantity
                        if raw_quantity is not None
                        else 0
                    )
                )

            except (
                TypeError,
                ValueError,
            ):

                app.logger.warning(
                    "Skipping kit with invalid quantity. "
                    "Name=%s Quantity=%s",
                    name,
                    raw_quantity,
                )

                skipped += 1
                continue

            if quantity <= 0:

                skipped += 1
                continue

            # -------------------------------------------------
            # Stable inventory key
            # -------------------------------------------------

            source_key = (
                "kit::"
                + name.lower()
                .strip()
                .replace(" ", "_")
            )

            existing = query(
                """
                SELECT id
                FROM inventory
                WHERE source_key=?
                LIMIT 1
                """,
                (source_key,),
                one=True,
            )

            # -------------------------------------------------
            # Existing kit:
            # DO NOT reset quantity from Excel.
            # -------------------------------------------------

            if existing:

                continue

            # -------------------------------------------------
            # New kit:
            # Insert it into inventory.
            # -------------------------------------------------

            execute(
                """
                INSERT INTO inventory
                (
                    source_file,
                    source_key,
                    stock_number,
                    name,
                    category,
                    description,
                    location,
                    total_qty,
                    available_qty,
                    minimum_stock,
                    condition,
                    maintenance_status,
                    active,
                    created_at,
                    updated_at
                )
                VALUES
                (
                    ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
                )
                """,
                (
                    "kits.xlsx",
                    source_key,
                    "",
                    name,
                    "Kits",
                    "Laboratory kit",
                    "Main Lab",
                    quantity,
                    quantity,
                    1,
                    "Good",
                    "Clear",
                    1,
                    now,
                    now,
                ),
            )

            inserted += 1

        db().commit()

        app.logger.info(
            "Kit import completed. Inserted=%s Skipped=%s",
            inserted,
            skipped,
        )

    except Exception:

        app.logger.exception(
            "Kit Excel import failed."
        )

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

def parse_past_date(value):
    """
    Convert Excel dates / strings into YYYY-MM-DD.
    Returns empty string when no valid date is available.
    """
    if value is None or value == "":
        return ""

    # Already a datetime
    if isinstance(value, datetime):
        return value.date().isoformat()

    # Excel serial date
    if isinstance(value, (int, float)):
        try:
            excel_origin = datetime(1899, 12, 30)
            converted = excel_origin + timedelta(days=float(value))
            return converted.date().isoformat()
        except (ValueError, TypeError, OverflowError):
            return ""

    text = str(value).strip()

    if not text:
        return ""

    # Try common date formats
    formats = [
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%m/%d/%Y",
        "%d-%b-%Y",
        "%d %b %Y",
        "%d %B %Y",
    ]

    for fmt in formats:
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            continue

    return text


def import_past_component_usage():
    """
    Import historical component usage from the Excel workbook.

    The workbook contains student details only on the first row
    of each student's/project's record. Component-only rows inherit
    those details from the previous populated row.

    This data is historical only and does NOT create current users,
    requests, approvals, or issues.
    """

    try:
        from openpyxl import load_workbook
    except ImportError:
        app.logger.warning(
            "openpyxl is not installed. Past data import skipped."
        )
        return

    path = (
        ROOT
        / "attached_assets"
        / "pastdata_labvault.xlsx"
    )

    if not path.exists():
        app.logger.info(
            "Past-data workbook not found: %s",
            path,
        )
        return

    try:
        workbook = load_workbook(
            path,
            read_only=True,
            data_only=True,
        )

        sheet = workbook.active
        rows = list(
            sheet.iter_rows(
                values_only=True
            )
        )

        if not rows:
            workbook.close()
            return

        headers = [
            str(value or "").strip().lower()
            for value in rows[0]
        ]

        def col(name):
            try:
                return headers.index(name)
            except ValueError:
                return None

        name_i = col("name")
        email_i = col("email")
        year_i = col("year")
        department_i = col("department")
        division_i = col("div")
        project_i = col("project")
        component_i = col("component name")
        quantity_i = col("quantity")
        issue_i = col("issue date")
        returned_i = col("returned date")
        condition_i = col("condition")
        phone_i = col("phone no.")

        required_columns = [
            name_i,
            project_i,
            component_i,
            quantity_i,
            issue_i,
            returned_i,
        ]

        if any(index is None for index in required_columns):
            app.logger.warning(
                "Past-data workbook is missing one or more required columns."
            )
            workbook.close()
            return

        # These values are carried forward for component-only rows.
        current_name = ""
        current_email = ""
        current_year = ""
        current_department = ""
        current_division = ""
        current_project = ""
        current_issue_date = ""
        current_returned_date = ""
        current_condition = ""
        current_phone = ""

        imported = 0
        updated = 0

        for excel_row_number, row in enumerate(
            rows[1:],
            start=2
        ):

            def get_value(index):
                if index is None:
                    return ""
                if index >= len(row):
                    return ""
                return str(
                    row[index] or ""
                ).strip()

            # Update current student/project context
            # whenever a value is present.
            value = get_value(name_i)
            if value:
                current_name = value

            value = get_value(email_i)
            if value:
                current_email = value

            value = get_value(year_i)
            if value:
                current_year = value

            value = get_value(department_i)
            if value:
                current_department = value

            value = get_value(division_i)
            if value:
                current_division = value

            value = get_value(project_i)
            if value:
                current_project = value

            if issue_i is not None and row[issue_i] not in (None, ""):
                current_issue_date = parse_past_date(
                    row[issue_i]
                )

            if returned_i is not None and row[returned_i] not in (None, ""):
                current_returned_date = parse_past_date(
                    row[returned_i]
                )

            value = get_value(condition_i)
            if value:
                current_condition = value

            value = get_value(phone_i)
            if value:
                current_phone = value

            component_name = get_value(
                component_i
            )

            # Ignore blank component rows.
            if not component_name:
                continue

            # Quantity
            quantity = 1

            try:
                raw_quantity = row[quantity_i]

                if raw_quantity not in (None, ""):
                    quantity = max(
                        1,
                        int(float(raw_quantity))
                    )
            except (
                TypeError,
                ValueError,
            ):
                quantity = 1

            source_key = (
                f"pastdata:"
                f"{excel_row_number}"
            )

            now = utcnow().isoformat()

            existing = query(
                """
                SELECT id
                FROM past_component_usage
                WHERE source_key=?
                """,
                (source_key,),
                one=True,
            )

            if existing:

                execute(
                    """
                    UPDATE past_component_usage
                    SET
                        student_name=?,
                        student_email=?,
                        year=?,
                        department=?,
                        division=?,
                        project_name=?,
                        component_name=?,
                        quantity=?,
                        issue_date=?,
                        returned_date=?,
                        condition=?,
                        phone=?,
                        updated_at=?
                    WHERE source_key=?
                    """,
                    (
                        current_name,
                        current_email,
                        current_year,
                        current_department,
                        current_division,
                        current_project,
                        component_name,
                        quantity,
                        current_issue_date,
                        current_returned_date,
                        current_condition,
                        current_phone,
                        now,
                        source_key,
                    ),
                )

                updated += 1

            else:

                execute(
                    """
                    INSERT INTO past_component_usage
                    (
                        source_key,
                        student_name,
                        student_email,
                        year,
                        department,
                        division,
                        project_name,
                        component_name,
                        quantity,
                        issue_date,
                        returned_date,
                        condition,
                        phone,
                        created_at,
                        updated_at
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        source_key,
                        current_name,
                        current_email,
                        current_year,
                        current_department,
                        current_division,
                        current_project,
                        component_name,
                        quantity,
                        current_issue_date,
                        current_returned_date,
                        current_condition,
                        current_phone,
                        now,
                        now,
                    ),
                )

                imported += 1

        workbook.close()

        print(
            f"Past component usage import completed. "
            f"Inserted={imported} Updated={updated}"
        )

    except Exception as exc:
        app.logger.exception(
            "Past-data import failed: %s",
            exc,
        )

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

def hod_required(fn):

    @wraps(fn)
    def wrapped(*args, **kwargs):

        user = current_user()

        if not user:

            flash(
                "Please sign in to continue.",
                "info"
            )

            return redirect(
                url_for(
                    "login",
                    role="hod"
                )
            )

        if not user["is_hod"]:

            flash(
                "HOD access is required.",
                "error"
            )

            return redirect(
                url_for("dashboard")
            )

        return fn(
            *args,
            **kwargs
        )

    return wrapped


def valid_csrf():
    return secrets.compare_digest(session.get("csrf", ""), request.form.get("csrf", ""))


def audit(action, description="", user=None):
    """
    Write an audit entry.

    HTTP request:
        use the logged-in user.

    Background worker:
        record as a system-generated event.
    """
    if user is None and has_request_context():
        user = current_user()

    execute(
        """
        INSERT INTO audit_logs
        (user_id, role, action, description, created_at)
        VALUES (?, ?, ?, ?, ?)
        """,
        (
            user["id"] if user else None,
            user["role"] if user else "system",
            action,
            description,
            utcnow().isoformat(),
        ),
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
        "item_count": query("SELECT COUNT(*) AS n FROM inventory WHERE active=1", one=True)["n"],
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
                    """
                    SELECT *
                    FROM users
                    WHERE role='admin'
                    AND is_hod=0
                    AND (
                        email=?
                        OR full_name=?
                        OR ?=?
                    )
                    """,
                    (
                        identity,
                        identity,
                        identity,
                        ADMIN_USERNAME,
                    ),
                    one=True,
                )

            elif role == "hod":

                user = query(
                    """
                    SELECT *
                    FROM users
                    WHERE is_hod=1
                    AND lower(email)=?
                    """,
                    (
                        identity.lower(),
                    ),
                    one=True,
                )

            else:

                user = query(
                    """
                    SELECT *
                    FROM users
                    WHERE role='student'
                    AND upper(erp_id)=?
                    """,
                    (
                        identity.upper(),
                    ),
                    one=True,
                )
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
            flash(
                "Your session expired. Please try again.",
                "error"
            )
            return render_template("register.html")

        values = {
            key: request.form.get(key, "").strip()
            for key in [
                "full_name",
                "erp_id",
                "phone",
                "uid",
                "branch",
                "division",
                "year",
                "email",
            ]
        }

        values["full_name"] = values["full_name"].strip()
        values["erp_id"] = values["erp_id"].upper().strip()
        values["uid"] = values["uid"].upper().strip()
        values["email"] = values["email"].lower().strip()

        # --------------------------------------------------
        # PHONE NUMBER NORMALIZATION
        # --------------------------------------------------

        phone = re.sub(r"\D", "", values["phone"])

        # Allow user to enter 91XXXXXXXXXX as well
        if phone.startswith("91") and len(phone) == 12:
            phone = phone[2:]

        # Store normalized phone with +91
        values["phone"] = phone

        password = request.form.get("password", "")
        confirm_password = request.form.get("confirm_password", "")

        # --------------------------------------------------
        # VALIDATION
        # --------------------------------------------------

        if len(values["full_name"]) < 2:
            flash(
                "Please enter your full name.",
                "error"
            )
            return render_template("register.html")

        # ERP ID example: S1032230450
        if not re.fullmatch(r"S\d{10}", values["erp_id"]):
            flash(
                "Enter a valid ERP ID. Example: S1032230450",
                "error"
            )
            return render_template("register.html")

        # Indian mobile number:
        # exactly 10 digits, starting with 6, 7, 8 or 9
        if not re.fullmatch(r"[6-9]\d{9}", values["phone"]):
            flash(
                "Please enter a valid Phone No.",
                "error"
            )
            return render_template("register.html")

        if len(values["uid"]) < 5:
            flash(
                "Please enter a valid UID.",
                "error"
            )
            return render_template("register.html")

        if "@" not in values["email"]:
            flash(
                "Please enter a valid email address.",
                "error"
            )
            return render_template("register.html")

        if len(password) < 8:
            flash(
                "Password must contain at least 8 characters.",
                "error"
            )
            return render_template("register.html")

        if password != confirm_password:
            flash(
                "Passwords do not match.",
                "error"
            )
            return render_template("register.html")

        # --------------------------------------------------
        # CHECK ERP ID
        # --------------------------------------------------

        existing_erp = query(
            """
            SELECT id
            FROM users
            WHERE upper(erp_id)=?
            """,
            (values["erp_id"],),
            one=True,
        )

        if existing_erp:
            flash(
                "That ERP ID is already registered.",
                "error"
            )
            return render_template("register.html")

        # --------------------------------------------------
        # CHECK UID
        # --------------------------------------------------

        existing_uid = query(
            """
            SELECT id
            FROM users
            WHERE upper(uid)=?
            """,
            (values["uid"],),
            one=True,
        )

        if existing_uid:
            flash(
                "That UID is already registered.",
                "error"
            )
            return render_template("register.html")

        # --------------------------------------------------
        # CHECK EMAIL
        # --------------------------------------------------

        existing_email = query(
            """
            SELECT id
            FROM users
            WHERE lower(email)=?
            """,
            (values["email"],),
            one=True,
        )

        if existing_email:
            flash(
                "That email address is already registered.",
                "error"
            )
            return render_template("register.html")

        # --------------------------------------------------
        # STORE PHONE WITH INDIA COUNTRY CODE
        # --------------------------------------------------

        values["phone"] = "+91" + values["phone"]

        # --------------------------------------------------
        # CREATE STUDENT
        # --------------------------------------------------

        try:

            execute(
                """
                INSERT INTO users
                (
                    role,
                    full_name,
                    erp_id,
                    phone,
                    uid,
                    branch,
                    division,
                    year,
                    email,
                    password_hash,
                    created_at
                )
                VALUES
                (
                    'student',
                    ?,
                    ?,
                    ?,
                    ?,
                    ?,
                    ?,
                    ?,
                    ?,
                    ?,
                    ?
                )
                """,
                (
                    values["full_name"],
                    values["erp_id"],
                    values["phone"],
                    values["uid"],
                    values["branch"],
                    values["division"],
                    values["year"],
                    values["email"],
                    generate_password_hash(password),
                    utcnow().isoformat(),
                ),
            )

            flash(
                "Account created successfully. Sign in using your ERP ID.",
                "success"
            )

            return redirect(
                url_for("login", role="student")
            )

        except (sqlite3.IntegrityError, psycopg.IntegrityError):
            flash(
                "That ERP ID, UID, or email is already registered.",
                "error"
            )

    return render_template("register.html")
        
@app.route("/forgot-password", methods=["GET", "POST"])
def forgot_password():
    if request.method == "POST":
        if not valid_csrf():
            flash("Your session expired. Please try again.", "error")
            return render_template("forgot_password.html")
        email = request.form.get("email", "").strip().lower()
        user = query("SELECT * FROM users WHERE role='student' AND lower(email)=?", (email,), one=True)
        if not user:
            flash("Please register an account before resetting your password.", "error")
            return render_template("forgot_password.html")
        raw = serializer.dumps({"user_id": user["id"], "purpose": "password_reset"})
        execute(
            "INSERT INTO password_reset_tokens (user_id, token_hash, created_at) VALUES (?, ?, ?)",
            (user["id"], hashlib.sha256(raw.encode()).hexdigest(), utcnow().isoformat()),
        )
        link = f"{APP_URL.rstrip('/') if APP_URL else request.host_url.rstrip('/')}{url_for('reset_password', token=raw)}"
        body = f"""Hello {user['full_name']},

We received a request to reset your LabVault password. This link is valid for 1 hour and can only be used once.

If you did not request this, you can ignore this email."""
        send_mail(user["email"], "Reset your LabVault password", body, link)
        flash("If that email is registered, a reset link has been sent.", "success")
        return redirect(url_for("login", role="student"))
    return render_template("forgot_password.html")


@app.route("/reset-password/<token>", methods=["GET", "POST"])
def reset_password(token):
    try:
        payload = serializer.loads(token, max_age=60 * 60)
    except (BadSignature, SignatureExpired):
        return render_template("approval.html", error="This password reset link has expired or is invalid.", stage="reset")
    if payload.get("purpose") != "password_reset":
        return render_template("approval.html", error="This password reset link is not valid.", stage="reset")
    row = query(
        "SELECT * FROM password_reset_tokens WHERE token_hash=?",
        (hashlib.sha256(token.encode()).hexdigest(),), one=True,
    )
    if not row or row["status"] != "PENDING":
        return render_template("approval.html", error="This password reset link has already been used.", stage="reset")
    if request.method == "POST":
        if not valid_csrf():
            flash("Your session expired. Please try again.", "error")
            return render_template("reset_password.html", token=token)
        password = request.form.get("password", "")
        confirm = request.form.get("confirm", "")
        if len(password) < 8 or password != confirm:
            flash("Passwords must match and be at least 8 characters.", "error")
            return render_template("reset_password.html", token=token)
        execute("UPDATE users SET password_hash=? WHERE id=?", (generate_password_hash(password), payload["user_id"]))
        execute("UPDATE password_reset_tokens SET status='USED', used_at=? WHERE id=?", (utcnow().isoformat(), row["id"]))
        audit("PASSWORD_RESET", "", query("SELECT * FROM users WHERE id=?", (payload["user_id"],), one=True))
        flash("Password updated. You can sign in now.", "success")
        return redirect(url_for("login", role="student"))
    return render_template("reset_password.html", token=token)

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

    user = current_user()

    if user["is_hod"]:
        return redirect(
            url_for("hod_dashboard")
        )

    if user["role"] == "admin":
        return redirect(
            url_for("admin_dashboard")
        )

    return redirect(
        url_for("student_dashboard")
    )


def check_return_alerts():
    """
    Sends return-date alerts.

    Due today:
        Student -> website + Gmail
        Lab Assistant -> website
        Mentor -> Gmail

    Overdue:
        Student -> website + Gmail
        Lab Assistant -> website
        Mentor -> Gmail

    Each alert is sent only once per issue.
    """

    today = utcnow().date().isoformat()

    rows = query(
        """
        SELECT
            issues.*,
            requests.request_code,
            requests.mentor_name,
            requests.mentor_email,
            users.full_name AS student_name,
            users.email AS student_email
        FROM issues
        JOIN requests
            ON requests.id = issues.request_id
        JOIN users
            ON users.id = issues.student_id
        WHERE issues.returned_at IS NULL
          AND issues.due_date <= ?
          AND issues.extension_status NOT IN (
              'PENDING_ADMIN',
              'PENDING_MENTOR',
              'PENDING_HOD'
          )
        ORDER BY issues.due_date
        """,
        (today,),
    )

    admin_users = query(
        "SELECT id FROM users WHERE role='admin'"
    )

    for row in rows:

        # =====================================================
        # RETURN DUE TODAY
        # =====================================================

        if row["due_date"] == today and not row["due_notified_at"]:

            # -------------------------------------------------
            # STUDENT WEBSITE NOTIFICATION
            # -------------------------------------------------

            notify(
                row["student_id"],
                "Return due today",
                (
                    f"{row['request_code']} is due for return today "
                    f"({row['due_date']}). Please return the components today."
                ),
                "return_due",
            )

            # -------------------------------------------------
            # STUDENT GMAIL
            # -------------------------------------------------

            if row["student_email"]:
                send_mail(
                    row["student_email"],
                    "LabVault — Components Due Today",
                    (
                        f"Your LabVault request {row['request_code']} "
                        f"is due for return today ({row['due_date']}).\n\n"
                        f"Please return the issued components today "
                        f"to avoid an overdue return."
                    ),
                )

            # -------------------------------------------------
            # LAB ASSISTANT WEBSITE
            # -------------------------------------------------

            for admin_row in admin_users:
                notify(
                    admin_row["id"],
                    "Return due today",
                    (
                        f"{row['request_code']} from "
                        f"{row['student_name']} is due for return today "
                        f"({row['due_date']})."
                    ),
                    "return_due",
                )

            # -------------------------------------------------
            # MENTOR GMAIL
            # -------------------------------------------------

            if row["mentor_email"]:
                send_mail(
                    row["mentor_email"],
                    "LabVault — Student Component Return Due Today",
                    (
                        f"LabVault request {row['request_code']} "
                        f"for student {row['student_name']} "
                        f"is due for return today ({row['due_date']}).\n\n"
                        f"Please follow up with the student if required."
                    ),
                )

            # Mark due-today alert as already sent
            execute(
                "UPDATE issues SET due_notified_at=? WHERE id=?",
                (utcnow().isoformat(), row["id"]),
            )

            audit(
                "RETURN_DUE_ALERT_SENT",
                row["request_code"],
            )

        # =====================================================
        # OVERDUE
        # =====================================================

        elif row["due_date"] < today and not row["overdue_notified_at"]:

            # -------------------------------------------------
            # STUDENT WEBSITE NOTIFICATION
            # -------------------------------------------------

            notify(
                row["student_id"],
                "Component overdue",
                (
                    f"{row['request_code']} was due on "
                    f"{row['due_date']} and has not been returned yet."
                ),
                "overdue",
            )

            # -------------------------------------------------
            # STUDENT GMAIL
            # -------------------------------------------------

            if row["student_email"]:
                send_mail(
                    row["student_email"],
                    "LabVault — Component Return Overdue",
                    (
                        f"Your LabVault request {row['request_code']} "
                        f"was due on {row['due_date']} and has not "
                        f"been returned yet.\n\n"
                        f"Please return the components as soon as possible."
                    ),
                )

            # -------------------------------------------------
            # LAB ASSISTANT WEBSITE
            # -------------------------------------------------

            for admin_row in admin_users:
                notify(
                    admin_row["id"],
                    "Overdue return",
                    (
                        f"{row['request_code']} from "
                        f"{row['student_name']} is overdue. "
                        f"Due date: {row['due_date']}."
                    ),
                    "overdue",
                )

            # -------------------------------------------------
            # MENTOR GMAIL
            # -------------------------------------------------

            if row["mentor_email"]:
                send_mail(
                    row["mentor_email"],
                    "LabVault — Overdue Component Return",
                    (
                        f"LabVault request {row['request_code']} "
                        f"for student {row['student_name']} "
                        f"was due on {row['due_date']} and "
                        f"has not been returned yet.\n\n"
                        f"Please follow up with the student."
                    ),
                )

            # Mark overdue alert as already sent
            execute(
                "UPDATE issues SET overdue_notified_at=? WHERE id=?",
                (utcnow().isoformat(), row["id"]),
            )

            audit(
                "OVERDUE_FLAGGED",
                row["request_code"],
            )

def return_alert_worker():
    """
    Background worker that checks return dates every minute.
    """

    while True:
        try:
            with app.app_context():
                check_return_alerts()
        except Exception:
            app.logger.exception(
                "Return alert worker failed."
            )

        time.sleep(60)

@app.route("/student")
@login_required("student")
def student_dashboard():
    check_return_alerts()

    user = current_user()

    # --------------------------------------------------
    # INVENTORY FILTERS
    # --------------------------------------------------

    term = request.args.get("q", "").strip()
    category = request.args.get("category", "")
    availability = request.args.get("availability", "")

    sql = "SELECT * FROM inventory WHERE active=1"
    args = []

    if term:
        sql += """
            AND (
                name LIKE ?
                OR category LIKE ?
                OR stock_number LIKE ?
            )
        """
        args += [
            f"%{term}%",
            f"%{term}%",
            f"%{term}%",
        ]

    if category:
        sql += " AND category=?"
        args.append(category)

    if availability == "available":
        sql += " AND available_qty > 0"

    elif availability == "low":
        sql += """
            AND available_qty <= minimum_stock
            AND available_qty > 0
        """

    sql += " ORDER BY name LIMIT 80"

    items = query(sql, args)

    # --------------------------------------------------
    # CATEGORIES
    # --------------------------------------------------

    categories = query(
        """
        SELECT DISTINCT category
        FROM inventory
        WHERE active=1
        ORDER BY category
        """
    )

    # --------------------------------------------------
    # STUDENT REQUESTS
    # --------------------------------------------------

    show_all_requests = request.args.get("all") == "1"

    if show_all_requests:
        requests = query(
            """
            SELECT
                requests.*,
                EXISTS(
                    SELECT 1
                    FROM issues
                    WHERE issues.request_id = requests.id
                ) AS has_issue
            FROM requests
            WHERE student_id=?
            ORDER BY created_at DESC
            """,
            (user["id"],),
        )
    else:
        requests = query(
            """
            SELECT
                requests.*,
                EXISTS(
                    SELECT 1
                    FROM issues
                    WHERE issues.request_id = requests.id
                ) AS has_issue
            FROM requests
            WHERE student_id=?
            ORDER BY created_at DESC
            LIMIT 6
            """,
            (user["id"],),
        )

    # --------------------------------------------------
    # COMPONENTS TAKEN IN EACH REQUEST
    # --------------------------------------------------

    request_items_map = {
        row["id"]: request_items(row["id"])
        for row in requests
    }

    # --------------------------------------------------
    # CURRENTLY ISSUED COMPONENTS
    # --------------------------------------------------

    issues = query(
        """
        SELECT
            issues.*,
            requests.request_code
        FROM issues
        JOIN requests
            ON requests.id = issues.request_id
        WHERE issues.student_id=?
        AND issues.returned_at IS NULL
        ORDER BY issues.due_date
        """,
        (user["id"],),
    )

    # --------------------------------------------------
    # REQUEST CART
    # --------------------------------------------------

    cart = session.get("cart", {})

    # --------------------------------------------------
    # RENDER STUDENT DASHBOARD
    # --------------------------------------------------

    return render_template(
        "student.html",
        items=items,
        categories=categories,
        requests=requests,
        request_items_map=request_items_map,
        issues=issues,
        cart=cart,
        term=term,
        selected_category=category,
        availability=availability,
        show_all_requests=show_all_requests,
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

    ids = [
        int(key)
        for key in cart.keys()
        if str(key).isdigit()
    ]

    items = (
        query(
            f"""
            SELECT *
            FROM inventory
            WHERE id IN ({','.join('?' for _ in ids)})
            """,
            ids,
        )
        if ids
        else []
    )

    # ---------------------------------------------------------
    # ACTIVE IOT FACULTY
    # Source: faculty table populated from the Excel sync
    # ---------------------------------------------------------

    faculty = query(
        """
        SELECT
            id,
            full_name,
            designation,
            department,
            email,
            active
        FROM faculty
        WHERE active=1
        AND lower(department)=lower('IoT')
        ORDER BY full_name
        """
    )

    # ---------------------------------------------------------
    # HOD DETAILS
    # Taken from the HOD account in the database.
    # No HOD phone is required.
    # ---------------------------------------------------------

    hod_profile = query(
        """
        SELECT
            id,
            full_name,
            email
        FROM users
        WHERE is_hod=1
        LIMIT 1
        """,
        one=True,
    )

    hod_name = (
        hod_profile["full_name"]
        if hod_profile
        else HOD_NAME
    )

    hod_email = (
        hod_profile["email"]
        if hod_profile
        else HOD_EMAIL
    )

    # ---------------------------------------------------------
    # POST
    # ---------------------------------------------------------

    if request.method == "POST":

        if not valid_csrf():

            flash(
                "Your session expired. Please try again.",
                "error",
            )

        elif not items:

            flash(
                "Add at least one available component first.",
                "error",
            )

        elif not request.form.get("responsibility"):

            flash(
                "Please accept the responsibility statement before submitting.",
                "error",
            )

        else:

            # -------------------------------------------------
            # RETURN PERIOD
            # -------------------------------------------------

            try:

                days = min(
                    7,
                    max(
                        1,
                        int(
                            request.form.get(
                                "days",
                                "3",
                            )
                        ),
                    ),
                )

            except (TypeError, ValueError):

                days = 3

            due = (
                utcnow()
                + timedelta(days=days)
            ).date().isoformat()

            year = utcnow().year

            count = (
                query(
                    """
                    SELECT COUNT(*) AS n
                    FROM requests
                    """,
                    one=True,
                )["n"]
                + 1
            )

            code = f"REQ-{year}-{count:06d}"

            user = current_user()

            # -------------------------------------------------
            # MENTOR SELECTION
            # -------------------------------------------------

            mentor_choice = (
                request.form.get(
                    "mentor_choice",
                    "",
                )
                .strip()
            )

            mentor_name = ""
            mentor_email = ""
            mentor_phone = ""

            faculty_id = None
            mentor_type = "faculty"

            # =================================================
            # EXTERNAL / OTHER MENTOR
            # =================================================

            if mentor_choice == "other":

                mentor_type = "external"

                mentor_name = (
                    request.form.get(
                        "mentor_name",
                        "",
                    )
                    .strip()
                )

                mentor_email = (
                    request.form.get(
                        "mentor_email",
                        "",
                    )
                    .strip()
                    .lower()
                )

                mentor_phone = (
                    request.form.get(
                        "mentor_phone",
                        "",
                    )
                    .strip()
                )

                if not mentor_phone:

                        flash(
                            "Please enter the external mentor's phone number.",
                            "error",
                        )

                        return render_template(
                            "request.html",
                            items=items,
                            cart=cart,
                            faculty=faculty,
                            hod_name=hod_name,
                            hod_email=hod_email,
                        )

                if not mentor_name:

                    flash(
                        "Please enter the external mentor's name.",
                        "error",
                    )

                    return render_template(
                        "request.html",
                        items=items,
                        cart=cart,
                        faculty=faculty,
                        hod_name=hod_name,
                        hod_email=hod_email,
                    )

                if not mentor_email:

                    flash(
                        "Please enter the external mentor's email.",
                        "error",
                    )

                    return render_template(
                        "request.html",
                        items=items,
                        cart=cart,
                        faculty=faculty,
                        hod_name=hod_name,
                        hod_email=hod_email,
                    )

                if "@" not in mentor_email:

                    flash(
                        "Please enter a valid external mentor email.",
                        "error",
                    )

                    return render_template(
                        "request.html",
                        items=items,
                        cart=cart,
                        faculty=faculty,
                        hod_name=hod_name,
                        hod_email=hod_email,
                    )

                # External mentor is NOT trusted yet.
                mentor_verified = 0
                mentor_verified_at = None
                mentor_verified_by = None

            # =================================================
            # OFFICIAL IOT FACULTY MENTOR
            # =================================================

            else:

                mentor_type = "faculty"

                try:

                    faculty_id = int(
                        mentor_choice
                    )

                except (
                    TypeError,
                    ValueError,
                ):

                    faculty_id = None

                selected_faculty = None

                if faculty_id is not None:

                    selected_faculty = query(
                        """
                        SELECT
                            id,
                            full_name,
                            designation,
                            department,
                            email,
                            active
                        FROM faculty
                        WHERE id=?
                        AND active=1
                        AND lower(department)=lower('IoT')
                        """,
                        (faculty_id,),
                        one=True,
                    )

                if not selected_faculty:

                    flash(
                        "Please select a valid active IoT faculty mentor.",
                        "error",
                    )

                    return render_template(
                        "request.html",
                        items=items,
                        cart=cart,
                        faculty=faculty,
                        hod_name=hod_name,
                        hod_email=hod_email,
                    )

                # Faculty details come directly from the
                # faculty database record.
                mentor_name = (
                    selected_faculty["full_name"]
                )

                mentor_email = (
                    selected_faculty["email"]
                )

                # No faculty phone exists in the Excel file.
                mentor_phone = ""

                # Official faculty members are already trusted.
                mentor_verified = 1
                mentor_verified_at = (
                    utcnow().isoformat()
                )
                mentor_verified_by = None

            # =================================================
            # CREATE REQUEST
            # =================================================

            req_id = execute(
                """
                INSERT INTO requests
                (
                    request_code,
                    student_id,
                    project_title,
                    purpose,

                    mentor_name,
                    mentor_email,
                    mentor_phone,

                    hod_name,
                    hod_email,
                    hod_phone,

                    faculty_id,
                    mentor_type,
                    mentor_verified,
                    mentor_verified_at,
                    mentor_verified_by,
                    mentor_verification_reason,

                    due_date,
                    return_days,

                    admin_status,
                    mentor_status,
                    hod_status,
                    overall_status,

                    created_at,
                    updated_at
                )

                VALUES
                (
                    ?, ?, ?, ?,
                    ?, ?, ?,
                    ?, ?, ?,
                    ?, ?, ?, ?, ?, ?,
                    ?, ?,
                    'PENDING',
                    'PENDING',
                    'PENDING',
                    'PENDING_ADMIN',
                    ?, ?
                )
                """,
                (
                    code,

                    user["id"],

                    request.form.get(
                        "project_title",
                        "",
                    ).strip(),

                    request.form.get(
                        "purpose",
                        "",
                    ).strip(),

                    # -----------------------------------------
                    # MENTOR
                    # -----------------------------------------

                    mentor_name,
                    mentor_email,
                    mentor_phone,

                    # -----------------------------------------
                    # HOD
                    # HOD phone does not exist.
                    # Keep existing DB column compatible.
                    # -----------------------------------------

                    hod_name,
                    hod_email,
                    "",

                    # -----------------------------------------
                    # FACULTY / VERIFICATION
                    # -----------------------------------------

                    faculty_id,
                    mentor_type,
                    mentor_verified,
                    mentor_verified_at,
                    mentor_verified_by,
                    None,

                    # -----------------------------------------
                    # RETURN
                    # -----------------------------------------

                    due,
                    days,

                    # -----------------------------------------
                    # TIMESTAMPS
                    # -----------------------------------------

                    utcnow().isoformat(),
                    utcnow().isoformat(),
                ),
            )

            # =================================================
            # REQUEST ITEMS
            # =================================================

            for item in items:

                quantity = min(
                    int(
                        cart[
                            str(item["id"])
                        ]
                    ),
                    item["available_qty"],
                )

                if quantity <= 0:
                    continue

                execute(
                    """
                    INSERT INTO request_items
                    (
                        request_id,
                        inventory_id,
                        quantity
                    )
                    VALUES (?, ?, ?)
                    """,
                    (
                        req_id,
                        item["id"],
                        quantity,
                    ),
                )

            # =================================================
            # STUDENT NOTIFICATION
            # =================================================

            notify(
                user["id"],
                "Request submitted",
                (
                    f"{code} is waiting for "
                    f"Lab Assistant approval."
                ),
                "request",
            )

            # =================================================
            # LAB ASSISTANT NOTIFICATION
            # =================================================

            for admin_row in query(
                """
                SELECT id
                FROM users
                WHERE role='admin'
                AND is_hod=0
                """
            ):

                notify(
                    admin_row["id"],
                    "New component request",
                    (
                        f"{code} from "
                        f"{user['full_name']} "
                        f"needs your approval before "
                        f"it goes to the Mentor."
                    ),
                    "request",
                )

            # =================================================
            # EXTERNAL MENTOR NOTICE
            # =================================================

            if mentor_type == "external":

                hod_user = query(
                    """
                    SELECT id
                    FROM users
                    WHERE is_hod=1
                    LIMIT 1
                    """,
                    one=True,
                )

                if hod_user:

                    notify(
                        hod_user["id"],
                        "External mentor submitted",
                        (
                            f"{code} contains an external "
                            f"mentor submitted by "
                            f"{user['full_name']}."
                        ),
                        "mentor_verification",
                    )

            # =================================================
            # AUDIT
            # =================================================

            audit(
                "REQUEST_CREATED",
                code,
            )

            # =================================================
            # CLEAR CART
            # =================================================

            session["cart"] = {}

            flash(
                (
                    f"{code} submitted successfully. "
                    f"Waiting for Lab Assistant approval."
                ),
                "success",
            )

            return redirect(
                url_for(
                    "student_dashboard"
                )
            )

    # =========================================================
    # GET / VALIDATION FAILURE
    # =========================================================

    return render_template(
        "request.html",
        items=items,
        cart=cart,
        faculty=faculty,
        hod_name=hod_name,
        hod_email=hod_email,
    )

@app.post("/student/request/<int:request_id>/contacts")
@login_required("student")
def edit_request_contacts(request_id):

    student = current_user()

    req = query(
        """SELECT *
           FROM requests
           WHERE id=?
           AND student_id=?""",
        (request_id, student["id"]),
        one=True,
    )

    if not req:
        flash("Request not found.", "error")
        return redirect(url_for("student_dashboard"))

    if not valid_csrf():
        flash("Your session expired. Please try again.", "error")
        return redirect(url_for("student_dashboard"))

    if req["overall_status"] == "FULLY_APPROVED":
        flash(
            "Contact details cannot be changed after final approval.",
            "error",
        )
        return redirect(url_for("student_dashboard"))

    if req["overall_status"] == "REJECTED":
        flash(
            "Rejected requests cannot be edited.",
            "error",
        )
        return redirect(url_for("student_dashboard"))

    mentor_name = request.form.get("mentor_name", "").strip()
    mentor_email = request.form.get("mentor_email", "").strip()
    mentor_phone = request.form.get("mentor_phone", "").strip()

    hod_name = request.form.get("hod_name", "").strip()
    hod_email = request.form.get("hod_email", "").strip()
    hod_phone = request.form.get("hod_phone", "").strip()

    if not mentor_name or not mentor_email:
        flash(
            "Mentor name and email are required.",
            "error",
        )
        return redirect(url_for("student_dashboard"))

    if not hod_name or not hod_email:
        flash(
            "HOD name and email are required.",
            "error",
        )
        return redirect(url_for("student_dashboard"))

    if "@" not in mentor_email:
        flash("Please enter a valid Mentor email.", "error")
        return redirect(url_for("student_dashboard"))

    if "@" not in hod_email:
        flash("Please enter a valid HOD email.", "error")
        return redirect(url_for("student_dashboard"))

    mentor_changed = (
        mentor_name != (req["mentor_name"] or "")
        or mentor_email != (req["mentor_email"] or "")
        or mentor_phone != (req["mentor_phone"] or "")
    )

    hod_changed = (
        hod_name != (req["hod_name"] or "")
        or hod_email != (req["hod_email"] or "")
        or hod_phone != (req["hod_phone"] or "")
    )

    now = utcnow().isoformat()

    execute(
        """UPDATE requests
           SET mentor_name=?,
               mentor_email=?,
               mentor_phone=?,
               hod_name=?,
               hod_email=?,
               hod_phone=?,
               updated_at=?
           WHERE id=?""",
        (
            mentor_name,
            mentor_email,
            mentor_phone,
            hod_name,
            hod_email,
            hod_phone,
            now,
            request_id,
        ),
    )

    # --------------------------------------------------
    # PENDING ADMIN
    # No approval token exists yet, so only save changes.
    # --------------------------------------------------

    if req["overall_status"] == "PENDING_ADMIN":

        audit(
            "REQUEST_CONTACTS_UPDATED",
            f"{req['request_code']} Mentor/HOD details corrected by student.",
        )

        flash(
            "Mentor and HOD details updated successfully.",
            "success",
        )

        return redirect(url_for("student_dashboard"))


    # --------------------------------------------------
    # PENDING MENTOR
    # Regenerate Mentor approval token if Mentor details changed.
    # --------------------------------------------------

    if req["overall_status"] == "PENDING_MENTOR" and mentor_changed:

        execute(
            """UPDATE approval_tokens
               SET status='REPLACED',
                   used_at=?
               WHERE request_id=?
               AND stage='mentor'
               AND status='PENDING'""",
            (now, request_id),
        )

        mentor_token = signed_token(
            request_id,
            "mentor",
        )

        link = (
            f"{APP_URL.rstrip('/') if APP_URL else request.host_url.rstrip('/')}"
            f"{url_for('approval', stage='mentor', token=mentor_token)}"
        )

        body = f"""The Mentor contact details for this LabVault request have been corrected by the student.

Student: {student['full_name']} ({student['uid']})
Project: {req['project_title']}
Request ID: {req['request_code']}

Please review the request using the updated approval link."""

        sent = send_mail(
            mentor_email,
            "Updated Mentor Approval Required — LabVault",
            body,
            link,
        )

        notify(
            student["id"],
            "Mentor details updated",
            f"{req['request_code']} Mentor details were updated and the approval link was regenerated.",
            "approval",
        )

        audit(
            "MENTOR_CONTACT_UPDATED",
            f"{req['request_code']} Mentor details corrected and approval link regenerated.",
        )

        flash(
            "Mentor details updated and a new approval link was sent.",
            "success" if sent else "error",
        )

        return redirect(url_for("student_dashboard"))


    # --------------------------------------------------
    # PENDING HOD
    # Regenerate HOD approval token if HOD details changed.
    # --------------------------------------------------

    if req["overall_status"] == "PENDING_HOD" and hod_changed:

        execute(
            """UPDATE approval_tokens
               SET status='REPLACED',
                   used_at=?
               WHERE request_id=?
               AND stage='hod'
               AND status='PENDING'""",
            (now, request_id),
        )

        hod_token = signed_token(
            request_id,
            "hod",
        )

        link = (
            f"{APP_URL.rstrip('/') if APP_URL else request.host_url.rstrip('/')}"
            f"{url_for('approval', stage='hod', token=hod_token)}"
        )

        body = f"""The HOD contact details for this LabVault request have been corrected by the student.

Student: {student['full_name']} ({student['uid']})
Project: {req['project_title']}
Request ID: {req['request_code']}

Please review the request using the updated approval link."""

        sent = send_mail(
            hod_email,
            "Updated HOD Approval Required — LabVault",
            body,
            link,
        )

        notify(
            student["id"],
            "HOD details updated",
            f"{req['request_code']} HOD details were updated and the approval link was regenerated.",
            "approval",
        )

        audit(
            "HOD_CONTACT_UPDATED",
            f"{req['request_code']} HOD details corrected and approval link regenerated.",
        )

        flash(
            "HOD details updated and a new approval link was sent.",
            "success" if sent else "error",
        )

        return redirect(url_for("student_dashboard"))


    # --------------------------------------------------
    # Generic save
    # --------------------------------------------------

    if mentor_changed or hod_changed:

        audit(
            "REQUEST_CONTACTS_UPDATED",
            f"{req['request_code']} Mentor/HOD details corrected by student.",
        )

    flash(
        "Contact details updated successfully.",
        "success",
    )

    return redirect(url_for("student_dashboard"))

@app.post("/admin/requests/<int:request_id>/decide")
@login_required("admin")
def admin_decide_request(request_id):
    if not valid_csrf():
        flash("Your session expired. Please try again.", "error")
        return redirect(url_for("admin_dashboard"))
    req = query(
        """SELECT requests.*, users.full_name AS student_name, users.uid
        FROM requests JOIN users ON users.id=requests.student_id
        WHERE requests.id=? AND requests.overall_status='PENDING_ADMIN'""",
        (request_id,), one=True,
    )
    if not req:
        flash("That request is not waiting for Lab Assistant approval.", "error")
        return redirect(url_for("admin_dashboard"))
    decision = request.form.get("decision")
    now = utcnow().isoformat()
    if decision == "approve":
        execute("UPDATE requests SET admin_status='APPROVED', overall_status='PENDING_MENTOR', updated_at=? WHERE id=?", (now, request_id))
        token = signed_token(request_id, "mentor")
        link = f"{APP_URL.rstrip('/') if APP_URL else request.host_url.rstrip('/')}{url_for('approval', stage='mentor', token=token)}"
        body = f"""A component request requires your review.

Student: {req['student_name']} ({req['uid']})
Project: {req['project_title']}
Return window: {req['return_days']} day(s) from the date of collection
Request ID: {req['request_code']}"""
        sent = send_mail(req["mentor_email"], "Component Request Requires Your Approval", body, link)
        notify(req["student_id"], "Approved by Lab Assistant", f"{req['request_code']} is now waiting for Mentor approval.", "approval")
        audit("ADMIN_APPROVED", req["request_code"])
        flash(f"{req['request_code']} approved. {'Mentor notified.' if sent else 'Mentor email is not configured.'}", "success")
    elif decision == "reject":
        reason = request.form.get("reason", "").strip() or "No reason given."
        execute("UPDATE requests SET admin_status='REJECTED', overall_status='REJECTED', rejection_reason=?, updated_at=? WHERE id=?", (reason, now, request_id))
        notify(req["student_id"], "Request rejected by Lab Assistant", reason, "rejection")
        audit("ADMIN_REJECTED", req["request_code"])
        flash(f"{req['request_code']} rejected.", "success")
    else:
        flash("Choose approve or reject.", "error")
    return redirect(url_for("admin_dashboard"))


@app.route("/approval/<stage>/<token>", methods=["GET", "POST"])
def approval(stage, token):

    if stage not in {"mentor", "hod"}:
        return render_template(
            "approval.html",
            error="That approval page could not be found.",
            stage=stage,
        )

    result, error = token_payload(token, stage)

    if error:
        return render_template(
            "approval.html",
            error=error,
            stage=stage,
        )

    payload, token_row = result

    req = query(
        """
        SELECT
            requests.*,
            users.full_name AS student_name,
            users.uid,
            users.branch,
            users.division,
            users.year
        FROM requests
        JOIN users
            ON users.id = requests.student_id
        WHERE requests.id=?
        """,
        (payload["request_id"],),
        one=True,
    )

    if not req:
        return render_template(
            "approval.html",
            error="That request could not be found.",
            stage=stage,
        )

    items = request_items(req["id"])

    # =========================================================
    # HOD APPROVAL PROTECTION FOR EXTERNAL MENTORS
    # =========================================================
    #
    # An external mentor must be verified by the HOD first.
    # The HOD cannot use the approval page to bypass that step.
    #
    if (
        stage == "hod"
        and req["mentor_type"] == "external"
        and not req["mentor_verified"]
    ):

        return render_template(
            "approval.html",
            req=req,
            items=items,
            stage=stage,
            error=(
                "This request has an external mentor. "
                "The mentor details must be verified before "
                "the HOD can approve or reject the request."
            ),
            external_verification_required=True,
        )

    # =========================================================
    # POST
    # =========================================================

    if request.method == "POST":

        if not valid_csrf():

            flash(
                "Your approval session expired. Please try again.",
                "error",
            )

        else:

            decision = request.form.get("decision")

            if decision not in {
                "approve",
                "reject",
            }:

                flash(
                    "Choose approve or reject.",
                    "error",
                )

            elif (
                decision == "reject"
                and not request.form.get(
                    "reason",
                    "",
                ).strip()
            ):

                flash(
                    "A rejection reason is required.",
                    "error",
                )

            else:

                now = utcnow().isoformat()

                # =================================================
                # MENTOR STAGE
                # =================================================

                if stage == "mentor":

                    if decision == "approve":

                        execute(
                            """
                            UPDATE requests
                            SET
                                mentor_status='APPROVED',
                                overall_status='PENDING_HOD',
                                updated_at=?
                            WHERE id=?
                            """,
                            (
                                now,
                                req["id"],
                            ),
                        )

                        # -------------------------------------------------
                        # EXTERNAL MENTOR
                        # HOD must verify before getting approval link.
                        # -------------------------------------------------

                        if (
                            req["mentor_type"] == "external"
                            and not req["mentor_verified"]
                        ):

                            hod_user = query(
                                """
                                SELECT
                                    id,
                                    full_name,
                                    email
                                FROM users
                                WHERE is_hod=1
                                LIMIT 1
                                """,
                                one=True,
                            )

                            if hod_user:
                                base_url = (
                                    APP_URL.rstrip("/")
                                    if APP_URL
                                    else request.host_url.rstrip("/")
                                )

                                hod_dashboard_url = (
                                    f"{base_url}{url_for('hod_dashboard')}"
                                )

                            send_mail(
                                    hod_user["email"],
                                    "External Mentor Verification Required — LabVault",
                                    (
                                        f"The Mentor has approved "
                                        f"{req['request_code']}.\n\n"
                                        f"Before final HOD approval, "
                                        f"please verify the external mentor details:\n\n"
                                        f"Student: {req['student_name']}\n"
                                        f"Project: {req['project_title']}\n"
                                        f"Mentor: {req['mentor_name']}\n"
                                        f"Email: {req['mentor_email']}\n"
                                        f"Phone: "
                                        f"{req['mentor_phone'] or 'Not provided'}\n\n"
                                        f"Please verify these details "
                                        f"from the HOD dashboard."
                                    ),
                                    hod_dashboard_url,
                                )

                            notify(
                                    hod_user["id"],
                                    "External mentor verification required",
                                    (
                                        f"{req['request_code']} has an "
                                        f"external mentor whose details "
                                        f"must be verified before HOD approval."
                                    ),
                                    "mentor_verification",
                                )

                            notify(
                                req["student_id"],
                                "Mentor approved — HOD verification required",
                                (
                                    f"{req['request_code']} was approved "
                                    f"by the Mentor. The HOD must verify "
                                    f"the external mentor before final approval."
                                ),
                                "approval",
                            )

                            audit(
                                "MENTOR_APPROVED_EXTERNAL_PENDING_VERIFICATION",
                                req["request_code"],
                            )

                            message = (
                                "MENTOR APPROVAL COMPLETED — "
                                "HOD VERIFICATION REQUIRED"
                            )

                        # -------------------------------------------------
                        # NORMAL IOT FACULTY MENTOR
                        # -------------------------------------------------

                        else:

                            hod_token = signed_token(
                                req["id"],
                                "hod",
                            )

                            base_url = (
                                APP_URL.rstrip("/")
                                if APP_URL
                                else request.host_url.rstrip("/")
                            )

                            link = (
                                f"{base_url}"
                                f"{url_for('approval', stage='hod', token=hod_token)}"
                            )

                            send_mail(
                                req["hod_email"],
                                "HOD Approval Required — Mentor Has Approved",
                                (
                                    f"The Mentor has approved "
                                    f"{req['request_code']}.\n\n"
                                    f"Student: {req['student_name']}\n"
                                    f"Project: {req['project_title']}\n\n"
                                    f"Your approval is now required."
                                ),
                                link,
                            )

                            hod_user = query(
                                """
                                SELECT id
                                FROM users
                                WHERE is_hod=1
                                LIMIT 1
                                """,
                                one=True,
                            )

                            if hod_user:

                                notify(
                                    hod_user["id"],
                                    "HOD approval required",
                                    (
                                        f"{req['request_code']} is "
                                        f"waiting for your approval."
                                    ),
                                    "approval",
                                )

                            notify(
                                req["student_id"],
                                "Mentor approved your request",
                                (
                                    f"{req['request_code']} is now "
                                    f"waiting for HOD approval."
                                ),
                                "approval",
                            )

                            audit(
                                "MENTOR_APPROVED",
                                req["request_code"],
                            )

                            message = (
                                "MENTOR APPROVAL COMPLETED"
                            )

                    else:

                        reason = request.form[
                            "reason"
                        ].strip()

                        execute(
                            """
                            UPDATE requests
                            SET
                                mentor_status='REJECTED',
                                overall_status='REJECTED',
                                rejection_reason=?,
                                updated_at=?
                            WHERE id=?
                            """,
                            (
                                reason,
                                now,
                                req["id"],
                            ),
                        )

                        notify(
                            req["student_id"],
                            "Request rejected by mentor",
                            reason,
                            "rejection",
                        )

                        audit(
                            "MENTOR_REJECTED",
                            req["request_code"],
                        )

                        message = (
                            "MENTOR REJECTION RECORDED"
                        )

                # =================================================
                # HOD STAGE
                # =================================================

                else:

                    # -------------------------------------------------
                    # Safety check again for external mentors.
                    # -------------------------------------------------

                    if (
                        req["mentor_type"] == "external"
                        and not req["mentor_verified"]
                    ):

                        return render_template(
                            "approval.html",
                            req=req,
                            items=items,
                            stage=stage,
                            error=(
                                "The external mentor has not "
                                "been verified yet."
                            ),
                            external_verification_required=True,
                        )

                    if decision == "approve":

                        execute(
                            """
                            UPDATE requests
                            SET
                                hod_status='APPROVED',
                                overall_status='FULLY_APPROVED',
                                collection_otp_ready=1,
                                updated_at=?
                            WHERE id=?
                            """,
                            (
                                now,
                                req["id"],
                            ),
                        )

                        notify(
                            req["student_id"],
                            "Ready for collection",
                            (
                                f"{req['request_code']} is approved "
                                f"by Mentor and HOD. Generate an OTP "
                                f"when you are ready to collect it."
                            ),
                            "success",
                        )

                        for admin_row in query(
                            """
                            SELECT id
                            FROM users
                            WHERE role='admin'
                            AND is_hod=0
                            """
                        ):

                            notify(
                                admin_row["id"],
                                "Ready for collection",
                                (
                                    f"{req['request_code']} is fully "
                                    f"approved. The student may arrive "
                                    f"to collect it."
                                ),
                                "success",
                            )

                        # Notify Mentor
                        send_mail(
                            req["mentor_email"],
                            "LabVault — Fully Approved",
                            (
                                f"{req['request_code']} for "
                                f"{req['student_name']} has been "
                                f"fully approved by HOD and is ready "
                                f"for collection."
                            ),
                        )

                        # Notify student by email as well
                        send_mail(
                            req["student_email"]
                            if "student_email" in req.keys()
                            else "",
                            "LabVault — Request Fully Approved",
                            (
                                f"Your request "
                                f"{req['request_code']} has been "
                                f"fully approved by the Mentor and HOD "
                                f"and is ready for collection."
                            ),
                        )

                        audit(
                            "HOD_APPROVED",
                            req["request_code"],
                        )

                        message = (
                            "HOD APPROVAL COMPLETED"
                        )

                    else:

                        reason = request.form[
                            "reason"
                        ].strip()

                        execute(
                            """
                            UPDATE requests
                            SET
                                hod_status='REJECTED',
                                overall_status='REJECTED',
                                rejection_reason=?,
                                updated_at=?
                            WHERE id=?
                            """,
                            (
                                reason,
                                now,
                                req["id"],
                            ),
                        )

                        notify(
                            req["student_id"],
                            "Request rejected by HOD",
                            reason,
                            "rejection",
                        )

                        audit(
                            "HOD_REJECTED",
                            req["request_code"],
                        )

                        message = (
                            "HOD REJECTION RECORDED"
                        )

                # -------------------------------------------------
                # APPROVAL TOKEN USED
                # -------------------------------------------------

                execute(
                    """
                    UPDATE approval_tokens
                    SET
                        status='USED',
                        used_at=?
                    WHERE id=?
                    """,
                    (
                        now,
                        token_row["id"],
                    ),
                )

                return render_template(
                    "approval.html",
                    complete=message,
                    stage=stage,
                    req=req,
                    items=items,
                )

    return render_template(
        "approval.html",
        req=req,
        items=items,
        stage=stage,
    )

def make_otp(request_id, kind):
    raw = f"{secrets.randbelow(1_000_000):06d}"
    now = utcnow()

    req = query(
        "SELECT student_id FROM requests WHERE id=?",
        (request_id,),
        one=True
    )

    student_id = req["student_id"] if req else None

    execute(
        """UPDATE otp_tokens
           SET used_at=?
           WHERE request_id=?
           AND kind=?
           AND used_at IS NULL""",
        (now.isoformat(), request_id, kind)
    )

    execute(
        """INSERT INTO otp_tokens
           (request_id, student_id, kind, otp_hash, expires_at, created_at)
           VALUES (?, ?, ?, ?, ?, ?)""",
        (
            request_id,
            student_id,
            kind,
            hashlib.sha256(raw.encode()).hexdigest(),
            (now + timedelta(minutes=10)).isoformat(),
            now.isoformat(),
        )
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
        for admin_row in query("SELECT id FROM users WHERE role='admin'"):
            notify(admin_row["id"], "Student on the way", f"{req['request_code']} — collection OTP just generated, student is ready to collect.", "otp")
        flash(f"Collection OTP: {code} · valid for 10 minutes", "success")
    return redirect(url_for("student_dashboard"))

@app.post("/admin/student/<int:student_id>/contact")
@login_required("admin")
def admin_student_contact(student_id):
    data = request.get_json(silent=True) or {}

    password = str(data.get("password", ""))

    if not password:
        return jsonify({
            "success": False,
            "message": "Admin password is required."
        }), 400

    admin = current_user()

    if not admin or not check_password_hash(
        admin["password_hash"],
        password
    ):
        return jsonify({
            "success": False,
            "message": "Incorrect admin password."
        }), 403

    student = query(
        """SELECT id, full_name,erp_id, uid, email, phone, branch, division, year
           FROM users
           WHERE id=? AND role='student'""",
        (student_id,),
        one=True,
    )

    if not student:
        return jsonify({
            "success": False,
            "message": "Student not found."
        }), 404

    audit(
        "STUDENT_CONTACT_VIEWED",
        f"Protected contact details viewed for student #{student_id}",
        admin,
    )

    return jsonify({
        "success": True,
        "student": {
            "id": student["id"],
            "full_name": student["full_name"],
            "uid": student["uid"],
            "email": student["email"] or "Not provided",
            "phone": student["phone"] or "Not provided",
            "branch": student["branch"] or "Not provided",
            "division": student["division"] or "Not provided",
            "year": student["year"] or "Not provided",
            "erp_id": student["erp_id"] or "Not provided",
        }
    })

@app.post("/admin/request/<int:request_id>/details")
@login_required("admin")
def admin_request_details(request_id):
    data = request.get_json(silent=True) or {}

    password = str(data.get("password", "")).strip()

    if not password:
        return jsonify({
            "success": False,
            "message": "Admin password is required."
        }), 400

    admin = current_user()

    if not admin or not check_password_hash(
        admin["password_hash"],
        password
    ):
        return jsonify({
            "success": False,
            "message": "Incorrect admin password."
        }), 403

    req = query(
        """
        SELECT
            requests.*,
            users.full_name AS student_name,
            users.erp_id AS student_erp_id,
            users.uid AS student_uid,
            users.email AS student_email,
            users.phone AS student_phone,
            users.branch AS student_branch,
            users.division AS student_division,
            users.year AS student_year
        FROM requests
        JOIN users ON users.id = requests.student_id
        WHERE requests.id=?
        """,
        (request_id,),
        one=True,
    )

    if not req:
        return jsonify({
            "success": False,
            "message": "Request not found."
        }), 404

    items = query(
        """
        SELECT
            i.name,
            i.category,
            ri.quantity,
            i.location
        FROM request_items ri
        JOIN inventory i
            ON i.id = ri.inventory_id
        WHERE ri.request_id=?
        ORDER BY i.name
        """,
        (request_id,),
    )

    audit(
        "REQUEST_DETAILS_VIEWED",
        f"Protected request details viewed for {req['request_code']}",
        admin,
    )

    return jsonify({
        "success": True,
        "request": {
            "request_code": req["request_code"],
            "student_name": req["student_name"],
            "erp_id": req["student_erp_id"] or "Not provided",
            "uid": req["student_uid"] or "Not provided",
            "email": req["student_email"] or "Not provided",
            "phone": req["student_phone"] or "Not provided",
            "branch": req["student_branch"] or "Not provided",
            "division": req["student_division"] or "Not provided",
            "year": req["student_year"] or "Not provided",
            "project_title": req["project_title"],
            "purpose": req["purpose"],
            "due_date": req["due_date"],
            "overall_status": req["overall_status"],
        },
        "items": [
            {
                "name": item["name"],
                "category": item["category"],
                "quantity": item["quantity"],
                "location": item["location"] or "Main Lab",
            }
            for item in items
        ],
    })

@app.route("/hod")
@hod_required
def hod_dashboard():

    hod = current_user()

    pending_hod = query(
        """
        SELECT
            requests.*,
            users.full_name AS student_name,
            users.uid,
            users.erp_id
        FROM requests
        JOIN users
            ON users.id = requests.student_id
        WHERE requests.overall_status='PENDING_HOD'
        ORDER BY requests.created_at DESC
        """
    )

    external_pending = query(
        """
        SELECT
            requests.*,
            users.full_name AS student_name,
            users.uid,
            users.erp_id
        FROM requests
        JOIN users
            ON users.id = requests.student_id
        WHERE requests.overall_status='PENDING_HOD'
        AND requests.mentor_type='external'
        AND requests.mentor_verified=0
        ORDER BY requests.created_at DESC
        """
    )

    faculty = query(
        """
        SELECT *
        FROM faculty
        ORDER BY
            active DESC,
            full_name
        """
    )

    recent_requests = query(
        """
        SELECT
            requests.*,
            users.full_name AS student_name
        FROM requests
        JOIN users
            ON users.id = requests.student_id
        ORDER BY requests.created_at DESC
        LIMIT 20
        """
    )

    return render_template(
        "hod.html",
        hod=hod,
        pending_hod=pending_hod,
        external_pending=external_pending,
        faculty=faculty,
        recent_requests=recent_requests,
    )

@app.post("/hod/faculty/add")
@hod_required
def hod_add_faculty():

    if not valid_csrf():

        flash(
            "Your session expired. Please try again.",
            "error",
        )

        return redirect(
            url_for("hod_dashboard")
        )

    name = request.form.get(
        "full_name",
        "",
    ).strip()

    designation = request.form.get(
        "designation",
        "",
    ).strip()

    email = request.form.get(
        "email",
        "",
    ).strip().lower()

    if not name or not email:

        flash(
            "Faculty name and email are required.",
            "error",
        )

        return redirect(
            url_for("hod_dashboard")
        )

    if "@" not in email:

        flash(
            "Enter a valid faculty email.",
            "error",
        )

        return redirect(
            url_for("hod_dashboard")
        )

    existing = query(
        """
        SELECT id
        FROM faculty
        WHERE lower(email)=?
        """,
        (email,),
        one=True,
    )

    if existing:

        flash(
            "A faculty member with that email already exists.",
            "error",
        )

        return redirect(
            url_for("hod_dashboard")
        )

    now = utcnow().isoformat()

    execute(
        """
        INSERT INTO faculty
        (
            full_name,
            designation,
            department,
            email,
            active,
            created_at,
            updated_at
        )
        VALUES (?, ?, 'IoT', ?, 1, ?, ?)
        """,
        (
            name,
            designation,
            email,
            now,
            now,
        ),
    )

    audit(
        "FACULTY_ADDED",
        f"IoT faculty member added: {name}",
        current_user(),
    )

    flash(
        "Faculty member added successfully.",
        "success",
    )

    return redirect(
        url_for("hod_dashboard")
    )

@app.post("/hod/faculty/<int:faculty_id>/edit")
@hod_required
def hod_edit_faculty(faculty_id):

    if not valid_csrf():

        flash(
            "Your session expired. Please try again.",
            "error",
        )

        return redirect(
            url_for("hod_dashboard")
        )

    name = request.form.get(
        "full_name",
        "",
    ).strip()

    designation = request.form.get(
        "designation",
        "",
    ).strip()

    email = request.form.get(
        "email",
        "",
    ).strip().lower()

    if not name or not email:

        flash(
            "Faculty name and email are required.",
            "error",
        )

        return redirect(
            url_for("hod_dashboard")
        )

    duplicate = query(
        """
        SELECT id
        FROM faculty
        WHERE lower(email)=?
        AND id<>?
        """,
        (
            email,
            faculty_id,
        ),
        one=True,
    )

    if duplicate:

        flash(
            "Another faculty member already uses that email.",
            "error",
        )

        return redirect(
            url_for("hod_dashboard")
        )

    execute(
        """
        UPDATE faculty
        SET
            full_name=?,
            designation=?,
            email=?,
            department='IoT',
            updated_at=?
        WHERE id=?
        """,
        (
            name,
            designation,
            email,
            utcnow().isoformat(),
            faculty_id,
        ),
    )

    audit(
        "FACULTY_UPDATED",
        f"IoT faculty member updated: {name}",
        current_user(),
    )

    flash(
        "Faculty member updated.",
        "success",
    )

    return redirect(
        url_for("hod_dashboard")
    )


@app.post("/hod/faculty/<int:faculty_id>/toggle")
@hod_required
def hod_toggle_faculty(faculty_id):

    if not valid_csrf():
        flash(
            "Your session expired. Please try again.",
            "error",
        )
        return redirect(url_for("hod_dashboard"))

    member = query(
        """
        SELECT *
        FROM faculty
        WHERE id=?
        """,
        (faculty_id,),
        one=True,
    )

    if not member:
        flash(
            "Faculty member not found.",
            "error",
        )
        return redirect(url_for("hod_dashboard"))

    new_status = 0 if member["active"] else 1

    execute(
        """
        UPDATE faculty
        SET
            active=?,
            updated_at=?
        WHERE id=?
        """,
        (
            new_status,
            utcnow().isoformat(),
            faculty_id,
        ),
    )

    audit(
        "FACULTY_STATUS_CHANGED",
        (
            f"{member['full_name']} "
            f"{'activated' if new_status else 'deactivated'}."
        ),
        current_user(),
    )

    flash(
        "Faculty status updated.",
        "success",
    )

    return redirect(url_for("hod_dashboard"))

@app.post("/hod/request/<int:request_id>/verify-mentor")
@hod_required
def hod_verify_external_mentor(request_id):

    if not valid_csrf():
        flash(
            "Your session expired. Please try again.",
            "error",
        )
        return redirect(url_for("hod_dashboard"))

    req = query(
        """
        SELECT *
        FROM requests
        WHERE id=?
        AND overall_status='PENDING_HOD'
        AND mentor_type='external'
        """,
        (request_id,),
        one=True,
    )

    if not req:
        flash(
            "That external mentor verification request was not found.",
            "error",
        )
        return redirect(url_for("hod_dashboard"))

    decision = request.form.get(
        "decision"
    )

    reason = request.form.get(
        "reason",
        "",
    ).strip()

    hod = current_user()

    if decision == "verify":

        now = utcnow().isoformat()

        execute(
            """
            UPDATE requests
            SET
                mentor_verified=1,
                mentor_verified_at=?,
                mentor_verified_by=?,
                mentor_verification_reason=NULL,
                updated_at=?
            WHERE id=?
            """,
            (
                now,
                hod["id"],
                now,
                request_id,
            ),
        )

        token = signed_token(
            request_id,
            "hod",
        )

        link = (
            f"{APP_URL.rstrip('/') if APP_URL else request.host_url.rstrip('/')}"
            f"{url_for('approval', stage='hod', token=token)}"
        )

        send_mail(
            hod["email"],
            "HOD Approval Ready — External Mentor Verified",
            (
                f"External mentor details for "
                f"{req['request_code']} have been verified.\n\n"
                f"Student: {req['student_name'] if 'student_name' in req.keys() else 'Student'}\n"
                f"Mentor: {req['mentor_name']}\n"
                f"Email: {req['mentor_email']}\n\n"
                f"The request is now ready for final HOD approval."
            ),
            link,
        )

        notify(
            req["student_id"],
            "External mentor verified",
            (
                f"The external mentor for "
                f"{req['request_code']} has been verified by HOD. "
                f"The request is now waiting for final HOD approval."
            ),
            "approval",
        )

        audit(
            "EXTERNAL_MENTOR_VERIFIED",
            req["request_code"],
            hod,
        )

        flash(
            "External mentor verified. HOD approval is now available.",
            "success",
        )

    elif decision == "reject":

        if len(reason) < 5:
            flash(
                "Please provide a reason when rejecting an external mentor.",
                "error",
            )
            return redirect(url_for("hod_dashboard"))

        execute(
            """
            UPDATE requests
            SET
                mentor_verified=0,
                mentor_verification_reason=?,
                hod_status='REJECTED',
                overall_status='REJECTED',
                rejection_reason=?,
                updated_at=?
            WHERE id=?
            """,
            (
                reason,
                reason,
                utcnow().isoformat(),
                request_id,
            ),
        )

        notify(
            req["student_id"],
            "External mentor rejected",
            (
                f"The HOD rejected the external mentor "
                f"for {req['request_code']}: {reason}"
            ),
            "rejection",
        )

        audit(
            "EXTERNAL_MENTOR_REJECTED",
            f"{req['request_code']} — {reason}",
            hod,
        )

        flash(
            "External mentor rejected.",
            "success",
        )

    return redirect(
        url_for("hod_dashboard")
    )

@app.route("/hod/change-password", methods=["GET", "POST"])
@hod_required
def hod_change_password():

    hod = current_user()

    if request.method == "POST":

        if not valid_csrf():

            flash(
                "Your session expired. Please try again.",
                "error",
            )

            return redirect(
                url_for("hod_change_password")
            )

        current_password = request.form.get(
            "current_password",
            "",
        )

        new_password = request.form.get(
            "new_password",
            "",
        )

        confirm_password = request.form.get(
            "confirm_password",
            "",
        )

        if not check_password_hash(
            hod["password_hash"],
            current_password,
        ):

            flash(
                "Current HOD password is incorrect.",
                "error",
            )

            return redirect(
                url_for("hod_change_password")
            )

        if len(new_password) < 8:

            flash(
                "New password must contain at least 8 characters.",
                "error",
            )

            return redirect(
                url_for("hod_change_password")
            )

        if new_password != confirm_password:

            flash(
                "The passwords do not match.",
                "error",
            )

            return redirect(
                url_for("hod_change_password")
            )

        execute(
            """
            UPDATE users
            SET password_hash=?
            WHERE id=?
            AND is_hod=1
            """,
            (
                generate_password_hash(
                    new_password
                ),
                hod["id"],
            ),
        )

        audit(
            "HOD_PASSWORD_CHANGED",
            "HOD changed her own password.",
            hod,
        )

        flash(
            "Your HOD password has been changed successfully.",
            "success",
        )

        return redirect(
            url_for("hod_dashboard")
        )

    return render_template(
        "hod_change_password.html",
        hod=hod,
    )

@app.route(
    "/hod/change-admin-password",
    methods=["GET", "POST"]
)
@hod_required
def hod_change_admin_password():

    hod = current_user()

    admin = query(
        """
        SELECT
            id,
            full_name,
            email
        FROM users
        WHERE role='admin'
        AND COALESCE(is_hod, 0)=0
        ORDER BY id
        LIMIT 1
        """,
        one=True,
    )

    if not admin:

        flash(
            "Lab Assistant account could not be found.",
            "error",
        )

        return redirect(
            url_for("hod_dashboard")
        )

    if request.method == "POST":

        if not valid_csrf():

            flash(
                "Your session expired. Please try again.",
                "error",
            )

            return redirect(
                url_for("hod_change_admin_password")
            )

        new_password = request.form.get(
            "new_password",
            "",
        )

        confirm_password = request.form.get(
            "confirm_password",
            "",
        )

        if len(new_password) < 8:

            flash(
                "The new Lab Assistant password must contain at least 8 characters.",
                "error",
            )

            return redirect(
                url_for("hod_change_admin_password")
            )

        if new_password != confirm_password:

            flash(
                "The passwords do not match.",
                "error",
            )

            return redirect(
                url_for("hod_change_admin_password")
            )

        execute(
            """
            UPDATE users
            SET password_hash=?
            WHERE id=?
            AND role='admin'
            AND COALESCE(is_hod, 0)=0
            """,
            (
                generate_password_hash(
                    new_password
                ),
                admin["id"],
            ),
        )

        notify(
            admin["id"],
            "Lab Assistant password changed",
            (
                "Your Lab Assistant password has been changed "
                "by the HOD. Contact the HOD to obtain your new password."
            ),
            "security",
        )

        audit(
            "LAB_ASSISTANT_PASSWORD_CHANGED_BY_HOD",
            (
                f"Lab Assistant password changed by HOD "
                f"for user {admin['id']}."
            ),
            hod,
        )

        flash(
            "Lab Assistant password changed successfully.",
            "success",
        )

        return redirect(
            url_for("hod_dashboard")
        )

    return render_template(
        "hod_change_admin_password.html",
        hod=hod,
        admin=admin,
    )

@app.route("/hod/request/<int:request_id>/review")
@hod_required
def hod_request_review(request_id):

    req = query(
        """
        SELECT
            requests.*,
            users.full_name AS student_name,
            users.uid,
            users.erp_id,
            users.email AS student_email,
            users.phone AS student_phone,
            users.branch,
            users.division,
            users.year
        FROM requests
        JOIN users
            ON users.id = requests.student_id
        WHERE requests.id=?
        AND requests.overall_status='PENDING_HOD'
        """,
        (request_id,),
        one=True,
    )

    if not req:
        flash(
            "That request is not waiting for HOD approval.",
            "error",
        )
        return redirect(
            url_for("hod_dashboard")
        )

    if (
        req["mentor_type"] == "external"
        and not req["mentor_verified"]
    ):
        flash(
            "Verify the external mentor before approving this request.",
            "error",
        )
        return redirect(
            url_for("hod_dashboard")
        )

    token = signed_token(
        request_id,
        "hod",
    )

    return redirect(
        url_for(
            "approval",
            stage="hod",
            token=token,
        )
    )

@app.route("/admin", methods=["GET", "POST"])
@login_required("admin")
def admin_dashboard():
    check_return_alerts()

    if request.method == "POST" and valid_csrf():
        action = request.form.get("action")

        if action == "issue":
            request_code = request.form.get("request_code", "").strip()
            entered_otp = request.form.get("otp", "").strip()
            print("===== OTP DEBUG =====")
            print("Request code received:", repr(request_code))
            print("OTP received:", repr(entered_otp))
            print("=====================")

            req = query(
                """SELECT *
                   FROM requests
                   WHERE request_code=?
                   AND overall_status='FULLY_APPROVED'""",
                (request_code,),
                one=True,
            )

            if not req:
                flash(
                    "Invalid request ID. Please check the request code.",
                    "error"
                )
            else:
                otp = query(
                    """SELECT *
                    FROM otp_tokens
                    WHERE request_id=?
                    AND student_id=?
                    AND kind='collection'
                    AND used_at IS NULL
                    ORDER BY id DESC
                    LIMIT 1""",
                    (req["id"], req["student_id"]),
                    one=True,
                )

                if not otp:
                    flash(
                        "No active collection OTP found for this request. "
                        "Ask the student to generate a new OTP.",
                        "error"
                    )

                else:
                    now = utcnow()

                    try:
                        expires_at = datetime.fromisoformat(
                            otp["expires_at"]
                        )
                    except (ValueError, TypeError):
                        expires_at = now

                    if expires_at < now:
                        execute(
                            "UPDATE otp_tokens SET used_at=? WHERE id=?",
                            (now.isoformat(), otp["id"]),
                        )
                        flash(
                            "This collection OTP has expired. "
                            "Ask the student to generate a new OTP.",
                            "error"
                        )

                    elif otp["attempts"] >= 5:
                        execute(
                            "UPDATE otp_tokens SET used_at=? WHERE id=?",
                            (now.isoformat(), otp["id"]),
                        )
                        flash(
                            "This collection OTP has reached the maximum "
                            "number of attempts. Ask the student to generate "
                            "a new OTP.",
                            "error"
                        )

                    else:
                        entered_hash = hashlib.sha256(
                            entered_otp.encode()
                        ).hexdigest()

                        if not secrets.compare_digest(
                            otp["otp_hash"],
                            entered_hash,
                        ):
                            new_attempts = otp["attempts"] + 1

                            if new_attempts >= 5:
                                execute(
                                    """UPDATE otp_tokens
                                       SET attempts=?, used_at=?
                                       WHERE id=?""",
                                    (
                                        new_attempts,
                                        now.isoformat(),
                                        otp["id"],
                                    ),
                                )
                                flash(
                                    "Incorrect OTP. This OTP has now "
                                    "reached the maximum number of attempts. "
                                    "Ask the student to generate a new OTP.",
                                    "error"
                                )
                            else:
                                execute(
                                    "UPDATE otp_tokens SET attempts=? WHERE id=?",
                                    (new_attempts, otp["id"]),
                                )

                                remaining = 5 - new_attempts

                                flash(
                                    f"Incorrect collection OTP. "
                                    f"{remaining} attempt(s) remaining.",
                                    "error"
                                )

                        else:
                            # OTP IS VALID
                            items = request_items(req["id"])

                            for item in items:
                                execute(
                                    """UPDATE inventory
                                       SET available_qty=MAX(0, available_qty-?),
                                           updated_at=?
                                       WHERE name=?""",
                                    (
                                        item["quantity"],
                                        now.isoformat(),
                                        item["name"],
                                    ),
                                )

                            execute(
                                "UPDATE otp_tokens SET used_at=? WHERE id=?",
                                (now.isoformat(), otp["id"]),
                            )

                            actual_due_date = (
                                now + timedelta(
                                    days=req["return_days"] or 3
                                )
                            ).date().isoformat()

                            issue_id = execute(
                                """INSERT INTO issues
                                   (request_id, student_id, issued_at, due_date)
                                   VALUES (?, ?, ?, ?)""",
                                (
                                    req["id"],
                                    req["student_id"],
                                    now.isoformat(),
                                    actual_due_date,
                                ),
                            )

                            execute(
                                "UPDATE requests SET due_date=? WHERE id=?",
                                (actual_due_date, req["id"]),
                            )

                            notify(
                                req["student_id"],
                                "Components issued",
                                f"{req['request_code']} was collected. "
                                f"Return by {actual_due_date}.",
                                "issue",
                            )

                            audit(
                                "COMPONENTS_ISSUED",
                                req["request_code"],
                            )

                            flash(
                                f"Collection completed for "
                                f"{req['request_code']}. "
                                f"Issue record #{issue_id} created.",
                                "success",
                            )

        elif action == "add_inventory":
            name = request.form.get("name", "").strip()

            try:
                qty = max(
                    0,
                    int(request.form.get("quantity", "0"))
                )
            except ValueError:
                qty = 0

            if name:
                execute(
                    """INSERT INTO inventory
                       (
                           source_file,
                           source_key,
                           name,
                           category,
                           description,
                           location,
                           total_qty,
                           available_qty,
                           minimum_stock,
                           created_at,
                           updated_at
                       )
                       VALUES (
                           'admin',
                           ?,
                           ?,
                           ?,
                           ?,
                           ?,
                           ?,
                           ?,
                           ?,
                           ?,
                           ?
                       )""",
                    (
                        f"admin:{secrets.token_hex(8)}",
                        name,
                        request.form.get("category", "Other"),
                        "Added by Lab Assistant.",
                        request.form.get("location", "Main Lab"),
                        qty,
                        qty,
                        max(
                            1,
                            int(
                                request.form.get(
                                    "minimum",
                                    "1"
                                ) or 1
                            ),
                        ),
                        utcnow().isoformat(),
                        utcnow().isoformat(),
                    ),
                )

                audit("INVENTORY_CREATED", name)
                flash("Inventory item added.", "success")

    metrics = {
        "items": query(
            "SELECT COUNT(*) AS n FROM inventory WHERE active=1",
            one=True,
        )["n"],

        "units": query(
            "SELECT COALESCE(SUM(available_qty),0) AS n "
            "FROM inventory WHERE active=1",
            one=True,
        )["n"],

        "issued": query(
            "SELECT COUNT(*) AS n "
            "FROM issues WHERE returned_at IS NULL",
            one=True,
        )["n"],

        "low": query(
            """SELECT COUNT(*) AS n
               FROM inventory
               WHERE active=1
               AND available_qty <= minimum_stock
               AND available_qty > 0""",
            one=True,
        )["n"],

        "pending": query(
            "SELECT COUNT(*) AS n "
            "FROM requests WHERE overall_status LIKE 'PENDING%%'",
            one=True,
        )["n"],

        "pending_admin": query(
            "SELECT COUNT(*) AS n "
            "FROM requests WHERE overall_status='PENDING_ADMIN'",
            one=True,
        )["n"],

        "overdue": query(
            """SELECT COUNT(*) AS n
               FROM issues
               WHERE returned_at IS NULL
               AND due_date < ?""",
            (utcnow().date().isoformat(),),
            one=True,
        )["n"],
    }

    inventory = query(
        """SELECT *
           FROM inventory
           WHERE active=1
           ORDER BY name
           LIMIT 60"""
    )

    pending_admin_requests = query(
        """SELECT requests.*,
                  users.full_name AS student_name,
                  users.uid
           FROM requests
           JOIN users ON users.id=requests.student_id
           WHERE requests.overall_status='PENDING_ADMIN'
           ORDER BY requests.created_at"""
    )

    pending_admin_items = {
        row["id"]: request_items(row["id"])
        for row in pending_admin_requests
    }

    pending_extension_requests = query(
            """
            SELECT
                issues.id,
                issues.request_id,
                issues.student_id,
                issues.due_date,
                issues.extension_reason,
                issues.extension_requested_days,
                issues.extension_status,
                issues.extension_admin_status,
                issues.extension_mentor_status,
                issues.extension_hod_status,
                issues.extensions_used,
                issues.returned_at,

                requests.request_code,
                requests.project_title,

                users.full_name AS student_name,
                users.uid

            FROM issues

            INNER JOIN requests
                ON requests.id = issues.request_id

            INNER JOIN users
                ON users.id = issues.student_id

            WHERE issues.extension_status = 'PENDING_ADMIN'
            AND issues.returned_at IS NULL

            ORDER BY issues.issued_at DESC
            """
        )

    print("===== EXTENSION DEBUG =====")
    print("Pending extension count:", len(pending_extension_requests))

    for extension in pending_extension_requests:
            print(
                "Extension:",
                extension["id"],
                extension["request_code"],
                extension["student_name"],
                extension["extension_status"],
                extension["extension_reason"],
                extension["extension_requested_days"],
            )
    print("===========================")

    pending_extension_items = {
        row["id"]: request_items(row["request_id"])
        for row in pending_extension_requests
    }

    show_all_requests = request.args.get("all") == "1"

    if show_all_requests:
        recent = query(
            """SELECT
                requests.*,
                users.full_name AS student_name,
                users.uid,
                users.erp_id
            FROM requests
            JOIN users
                ON users.id=requests.student_id
            ORDER BY requests.created_at DESC"""
        )
    else:
        recent = query(
            """SELECT
                requests.*,
                users.full_name AS student_name,
                users.uid,
                users.erp_id
            FROM requests
            JOIN users
                ON users.id=requests.student_id
            ORDER BY requests.created_at DESC
            LIMIT 8"""
        )

    recent_request_items = {
        row["id"]: request_items(row["id"])
        for row in recent
    }

            # =========================================================
    # HISTORICAL USAGE ANALYTICS
    # =========================================================

    # ---------------------------------------------------------
    # LOAD ALL HISTORICAL RECORDS
    # ---------------------------------------------------------

    historical_rows = query(
        """
        SELECT
            student_name,
            student_email,
            year,
            department,
            division,
            project_name,
            component_name,
            quantity,
            issue_date,
            returned_date,
            condition,
            phone
        FROM past_component_usage
        ORDER BY issue_date
        """
    )

    # ---------------------------------------------------------
    # NORMALIZATION HELPERS
    # ---------------------------------------------------------

    def clean_value(value):
        if value is None:
            return ""

        value = str(value).strip()

        if value.lower() in {
            "",
            "-",
            "na",
            "n/a",
            "null",
            "none",
            "nan",
        }:
            return ""

        return value

    def normalize_component_name(value):
        """
        Merge different spellings of the same component.
        """
        value = clean_value(value)

        if not value:
            return ""

        normalized = re.sub(r"\s+", " ", value).strip().lower()

        # Jumper wires
        jumper_variants = {
            "jumperwire",
            "jumper wire",
            "jumper wires",
            "jumperwires",
            "jumper-wire",
            "jumper-wires",
        }

        if normalized in jumper_variants:
            return "Jumper Wires"

        # Common normalization examples
        component_aliases = {
            "esp 32": "ESP32",
            "esp-32": "ESP32",
            "esp8266 board": "ESP8266",
            "arduino uno r3": "Arduino UNO",
            "arduino uno": "Arduino UNO",
            "bread board": "Breadboard",
            "bread board": "Breadboard",
            "usb cable": "USB Cable",
        }

        if normalized in component_aliases:
            return component_aliases[normalized]

        # Normal title formatting for everything else
        return normalized.title()

    def normalize_year(value):
        """
        Only accept 1st, 2nd, 3rd and 4th year.
        Blank / '-' / invalid values are excluded.
        """

        value = clean_value(value).lower()

        year_map = {
        "1": "F.T.",
        "1st": "F.T.",
        "1st year": "F.T.",

        "2": "S.T.",
        "2nd": "S.T.",
        "2nd year": "S.T.",

        "3": "T.T.",
        "3rd": "T.T.",
        "3rd year": "T.T.",

        "4": "B.T.",
        "4th": "B.T.",
        "4th year": "B.T.",
    }

        return year_map.get(value, "")

    def normalize_department(value):
        """
        Merge IoT / IOT / Iot into IoT.
        Exclude blank values.
        """

        value = clean_value(value)

        if not value:
            return ""

        normalized = re.sub(r"\s+", " ", value).strip().lower()

        department_aliases = {
            "iot": "IoT",
            "i.o.t": "IoT",
            "internet of things": "IoT",

            "ecs": "ECS",
            "extc": "EXTC",
            "mech": "MECH",
            "mechanical": "MECH",

            "ai&ds": "AI&DS",
            "ai & ds": "AI&DS",
            "aids": "AI&DS",

            "aiml": "AIML",

            "comp": "COMP",
            "comps": "COMP",

            "mme": "MME",
        }

        return department_aliases.get(
            normalized,
            normalized.upper()
        )

    def normalize_project(value):
        """
        Clean project names and exclude missing values.
        """

        value = clean_value(value)

        if not value:
            return ""

        # Remove repeated spaces
        value = re.sub(r"\s+", " ", value).strip()

        normalized = value.lower()

        project_aliases = {

            "smart irrigation system":
                "Smart Irrigation System",

            "smart solarpowered irrigation system":
                "Smart Solar-Powered Irrigation System",

            "smart solar powered irrigation system":
                "Smart Solar-Powered Irrigation System",

            "iot based attendance system":
                "IoT Based Attendance System",

            "iot based automation controller for stps":
                "IoT Based Automation Controller for STPs",

            "pbl mini-project":
                "PBL Mini-Project",

            "pbl mini project":
                "PBL Mini-Project",

            "smart parking system":
                "Smart Parking System",

            "iot workshop in zephr":
                "IoT Workshop in Zephr",

            "iei outreach":
                "IEI Outreach",

            "iot-enabled supply management system for food industry":
                "IoT-Enabled Supply Management System for Food Industry",
        }

        return project_aliases.get(
            normalized,
            value
        )

    # ---------------------------------------------------------
    # SUMMARY
    # ---------------------------------------------------------

    total_records = len(historical_rows)

    total_quantity = sum(
        int(row["quantity"] or 0)
        for row in historical_rows
    )

    unique_components_set = set()
    unique_students_set = set()

    returned_records = 0

    for row in historical_rows:

        component = normalize_component_name(
            row["component_name"]
        )

        if component:
            unique_components_set.add(component)

        student_name = clean_value(
            row["student_name"]
        )

        if student_name:
            unique_students_set.add(
                student_name.lower()
            )

        returned_date = clean_value(
            row["returned_date"]
        )

        if returned_date:
            returned_records += 1

    historical_summary = {
        "total_records": total_records,
        "total_quantity": total_quantity,
        "unique_components": len(unique_components_set),
        "unique_students": len(unique_students_set),
    }

    if total_records:
        historical_return_percentage = round(
            (returned_records / total_records) * 100,
            1
        )
    else:
        historical_return_percentage = 0

    # ---------------------------------------------------------
    # MOST USED COMPONENTS
    # ---------------------------------------------------------

    component_stats = {}

    for row in historical_rows:

        component = normalize_component_name(
            row["component_name"]
        )

        if not component:
            continue

        if component not in component_stats:
            component_stats[component] = {
                "component_name": component,
                "issue_count": 0,
                "total_quantity": 0,
            }

        component_stats[component]["issue_count"] += 1

        component_stats[component]["total_quantity"] += int(
            row["quantity"] or 0
        )

    historical_components = sorted(
        component_stats.values(),
        key=lambda x: (
            x["total_quantity"],
            x["issue_count"]
        ),
        reverse=True,
    )[:10]

    # ---------------------------------------------------------
    # YEAR-WISE USAGE
    # ---------------------------------------------------------

    year_stats = {
        "F.T.": {
            "year": "F.T.",
            "issue_count": 0,
            "total_quantity": 0,
        },
        "S.T.": {
            "year": "S.T.",
            "issue_count": 0,
            "total_quantity": 0,
        },
        "T.T.": {
            "year": "T.T.",
            "issue_count": 0,
            "total_quantity": 0,
        },
        "B.T.": {
            "year": "B.T.",
            "issue_count": 0,
            "total_quantity": 0,
        },
    }

    year_total_quantity = sum(
            item["total_quantity"]
            for item in historical_years
        )

    for item in historical_years:
            if year_total_quantity > 0:
                item["percentage"] = round(
                    (
                        item["total_quantity"]
                        / year_total_quantity
                    ) * 100,
                    1,
                )
            else:
                item["percentage"] = 0
                
    for row in historical_rows:

        year = normalize_year(
            row["year"]
        )

        # Ignore blank / '-' / invalid year
        if not year:
            continue

        year_stats[year]["issue_count"] += 1

        year_stats[year]["total_quantity"] += int(
            row["quantity"] or 0
        )

    historical_years = [
        year_stats["F.T."],
        year_stats["S.T."],
        year_stats["T.T."],
        year_stats["B.T."],
    ]

    # ---------------------------------------------------------
    # DEPARTMENT-WISE USAGE
    # ---------------------------------------------------------

    department_stats = {}

    for row in historical_rows:

        department = normalize_department(
            row["department"]
        )

        # Ignore blank / '-'
        if not department:
            continue

        if department not in department_stats:
            department_stats[department] = {
                "department": department,
                "issue_count": 0,
                "total_quantity": 0,
            }

        department_stats[department]["issue_count"] += 1

        department_stats[department]["total_quantity"] += int(
            row["quantity"] or 0
        )

    historical_departments = sorted(
        department_stats.values(),
        key=lambda x: (
            x["total_quantity"],
            x["issue_count"]
        ),
        reverse=True,
    )

    # ---------------------------------------------------------
    # PROJECT-WISE USAGE
    # ---------------------------------------------------------

    project_stats = {}

    for row in historical_rows:

        project = normalize_project(
            row["project_name"]
        )

        # Ignore blank / '-'
        if not project:
            continue

        if project not in project_stats:
            project_stats[project] = {
                "project_name": project,
                "issue_count": 0,
                "total_quantity": 0,
            }

        project_stats[project]["issue_count"] += 1

        project_stats[project]["total_quantity"] += int(
            row["quantity"] or 0
        )

    historical_projects = sorted(
        project_stats.values(),
        key=lambda x: (
            x["total_quantity"],
            x["issue_count"]
        ),
        reverse=True,
    )[:10]

    # ---------------------------------------------------------
    # CONDITION SUMMARY
    # ---------------------------------------------------------

    condition_stats = {}

    for row in historical_rows:

        condition = clean_value(
            row["condition"]
        )

        if not condition:
            continue

        condition_key = condition.lower()

        if condition_key not in condition_stats:
            condition_stats[condition_key] = {
                "condition": condition.title(),
                "record_count": 0,
                "total_quantity": 0,
            }

        condition_stats[condition_key]["record_count"] += 1

        condition_stats[condition_key]["total_quantity"] += int(
            row["quantity"] or 0
        )

    historical_conditions = sorted(
        condition_stats.values(),
        key=lambda x: (
            x["record_count"],
            x["total_quantity"]
        ),
        reverse=True,
    )

    # ---------------------------------------------------------
    # CHART DATA
    # ---------------------------------------------------------

    historical_component_chart = [
        {
            "label": item["component_name"],
            "value": item["total_quantity"],
        }
        for item in historical_components
    ]

    historical_year_chart = [
    {
        "label": item["year"],
        "value": item["total_quantity"],
        "percentage": item["percentage"],
    }
    for item in historical_years
]

    historical_department_chart = [
        {
            "label": item["department"],
            "value": item["total_quantity"],
        }
        for item in historical_departments
    ][:10]

    historical_project_chart = [
        {
            "label": item["project_name"],
            "value": item["total_quantity"],
        }
        for item in historical_projects
    ]

    historical_condition_chart = [
        {
            "label": item["condition"],
            "value": item["total_quantity"],
        }
        for item in historical_conditions
    ]

    return render_template(
        "admin.html",
        metrics=metrics,
        inventory=inventory,
        recent=recent,
        recent_request_items=recent_request_items,
        show_all_requests=show_all_requests,
        pending_admin_requests=pending_admin_requests,
        pending_admin_items=pending_admin_items,
        pending_extension_requests=pending_extension_requests,
        pending_extension_items=pending_extension_items,

        # Historical analytics
        historical_summary=historical_summary,
        historical_return_percentage=historical_return_percentage,
        historical_components=historical_components,
        historical_years=historical_years,
        historical_departments=historical_departments,
        historical_projects=historical_projects,
        historical_conditions=historical_conditions,

        # Chart data
        historical_component_chart=historical_component_chart,
        historical_year_chart=historical_year_chart,
        historical_department_chart=historical_department_chart,
        historical_project_chart=historical_project_chart,
        historical_condition_chart=historical_condition_chart,
    )

@app.route("/admin/change-password", methods=["GET", "POST"])
@login_required("admin")
def admin_change_password():
    admin = current_user()

    if request.method == "POST":

        if not valid_csrf():
            flash(
                "Your session expired. Please try again.",
                "error"
            )
            return render_template(
                "admin_change_password.html"
            )

        current_password = request.form.get(
            "current_password",
            ""
        ).strip()

        reason = request.form.get(
            "reason",
            ""
        ).strip()

        # --------------------------------------------------
        # VERIFY CURRENT ADMIN PASSWORD
        # --------------------------------------------------

        if not current_password:
            flash(
                "Enter your current Admin password to continue.",
                "error"
            )
            return render_template(
                "admin_change_password.html"
            )

        if not check_password_hash(
            admin["password_hash"],
            current_password
        ):
            flash(
                "Incorrect current Admin password.",
                "error"
            )
            return render_template(
                "admin_change_password.html"
            )

        if len(reason) < 5:
            flash(
                "Please provide a genuine reason for requesting "
                "a password change.",
                "error"
            )
            return render_template(
                "admin_change_password.html"
            )

        # --------------------------------------------------
        # PREVENT MULTIPLE ACTIVE REQUESTS
        # --------------------------------------------------

        existing = query(
            """
            SELECT id
            FROM admin_password_change_requests
            WHERE admin_id=?
            AND status='PENDING'
            LIMIT 1
            """,
            (admin["id"],),
            one=True,
        )

        if existing:
            flash(
                "A password-change request is already waiting "
                "for HOD authorization.",
                "error"
            )
            return redirect(
                url_for("admin_dashboard")
            )

        # --------------------------------------------------
        # CREATE SECURE HOD AUTHORIZATION TOKEN
        # --------------------------------------------------

        now = utcnow().isoformat()

        # Temporary request ID first
        request_id = execute(
            """
            INSERT INTO admin_password_change_requests
            (
                admin_id,
                reason,
                token_hash,
                status,
                created_at
            )
            VALUES (?, ?, ?, 'PENDING', ?)
            """,
            (
                admin["id"],
                reason,
                "TEMP",
                now,
            ),
        )

        raw_token = serializer.dumps(
            {
                "admin_id": admin["id"],
                "request_id": request_id,
                "purpose": "admin_password_change",
            }
        )

        token_hash = hashlib.sha256(
            raw_token.encode()
        ).hexdigest()

        execute(
            """
            UPDATE admin_password_change_requests
            SET token_hash=?
            WHERE id=?
            """,
            (
                token_hash,
                request_id,
            ),
        )

        # --------------------------------------------------
        # HOD EMAIL LINK
        # --------------------------------------------------

        link = (
            f"{APP_URL.rstrip('/') if APP_URL else request.host_url.rstrip('/')}"
            f"{url_for('authorize_admin_password', token=raw_token)}"
        )

        body = f"""A password-change request has been submitted for the LabVault Admin/Lab Assistant account.

        Admin account:
        {ADMIN_USERNAME}

        Requested on:
        {format_ist(now)}

        Reason provided by the Lab Assistant:
        {reason}

        The Lab Assistant has verified the current Admin password.

        Please open the secure LabVault page below to review this request and, if appropriate, create a new Admin password.

        The new password must be created by the HOD.
        The Lab Assistant cannot choose the new password through this request."""

        sent = send_mail(
            HOD_EMAIL,
            "LabVault — Admin Password Change Request",
            body,
            link,
        )

        audit(
            "ADMIN_PASSWORD_CHANGE_REQUESTED",
            f"HOD password-change request #{request_id}.",
            admin,
        )

        if sent:
            flash(
                "Your password-change request has been sent to the HOD for authorization.",
                "success",
            )
        else:
            flash(
                "The password-change request was created, but the HOD email could not be delivered. Check the mail configuration.",
                "error",
            )

        return redirect(
            url_for("admin_dashboard")
        )

    return render_template(
        "admin_change_password.html"
    )


@app.route(
    "/admin/authorize-password-change/<token>",
    methods=["GET", "POST"]
)
def authorize_admin_password(token):

    # --------------------------------------------------
    # VERIFY SIGNED TOKEN
    # --------------------------------------------------

    try:
        payload = serializer.loads(
            token,
            max_age=60 * 60 * 24
        )
    except (BadSignature, SignatureExpired):
        return render_template(
            "hod_password_authorize.html",
            error=(
                "This password-change authorization link "
                "has expired or is invalid."
            ),
        )

    if payload.get("purpose") != "admin_password_change":
        return render_template(
            "hod_password_authorize.html",
            error=(
                "This password-change authorization link "
                "is not valid."
            ),
        )

    request_id = payload.get("request_id")
    admin_id = payload.get("admin_id")

    if not request_id or not admin_id:
        return render_template(
            "hod_password_authorize.html",
            error="Invalid password-change request.",
        )

    # --------------------------------------------------
    # VERIFY STORED TOKEN
    # --------------------------------------------------

    token_hash = hashlib.sha256(
        token.encode()
    ).hexdigest()

    change_request = query(
        """
        SELECT *
        FROM admin_password_change_requests
        WHERE id=?
        AND admin_id=?
        AND token_hash=?
        """,
        (
            request_id,
            admin_id,
            token_hash,
        ),
        one=True,
    )

    if not change_request:
        return render_template(
            "hod_password_authorize.html",
            error="Password-change request could not be found.",
        )

    if change_request["status"] != "PENDING":
        return render_template(
            "hod_password_authorize.html",
            error=(
                "This password-change request has already "
                "been reviewed."
            ),
        )

    admin = query(
        """
        SELECT id, full_name
        FROM users
        WHERE id=?
        AND role='admin'
        """,
        (admin_id,),
        one=True,
    )

    if not admin:
        return render_template(
            "hod_password_authorize.html",
            error="The LabVault Admin account could not be found.",
        )

    # --------------------------------------------------
    # HOD CREATES NEW PASSWORD
    # --------------------------------------------------

    if request.method == "POST":

        if not valid_csrf():
            flash(
                "Your authorization session expired. Please try again.",
                "error"
            )

            return render_template(
                "hod_password_authorize.html",
                pending=True,
                admin_name=admin["full_name"],
                reason=change_request["reason"],
            )

        new_password = request.form.get(
            "new_password",
            ""
        )

        confirm_password = request.form.get(
            "confirm_password",
            ""
        )

        if len(new_password) < 8:
            flash(
                "The new Admin password must contain at least 8 characters.",
                "error"
            )

            return render_template(
                "hod_password_authorize.html",
                pending=True,
                admin_name=admin["full_name"],
                reason=change_request["reason"],
            )

        if new_password != confirm_password:
            flash(
                "The password confirmation does not match.",
                "error"
            )

            return render_template(
                "hod_password_authorize.html",
                pending=True,
                admin_name=admin["full_name"],
                reason=change_request["reason"],
            )

        # --------------------------------------------------
        # HASH NEW PASSWORD
        # --------------------------------------------------

        new_password_hash = generate_password_hash(
            new_password
        )

        now = utcnow().isoformat()

        execute(
            """
            UPDATE users
            SET password_hash=?
            WHERE id=?
            AND role='admin'
            """,
            (
                new_password_hash,
                admin_id,
            ),
        )

        execute(
            """
            UPDATE admin_password_change_requests
            SET status='AUTHORIZED',
                authorized_at=?
            WHERE id=?
            """,
            (
                now,
                request_id,
            ),
        )

        # --------------------------------------------------
        # NOTIFY ADMIN
        # --------------------------------------------------

        notify(
            admin_id,
            "Admin password changed",
            (
                "The HOD has authorized the password change "
                "for the LabVault Admin account. "
                "The new password must be obtained directly "
                "from the HOD."
            ),
            "security",
        )

        audit(
            "ADMIN_PASSWORD_CHANGED_BY_HOD",
            f"Admin password changed by HOD for request #{request_id}.",
        )

        return render_template(
            "hod_password_authorize.html",
            decision="approved",
        )

    return render_template(
        "hod_password_authorize.html",
        pending=True,
        admin_name=admin["full_name"],
        reason=change_request["reason"],
    )

@app.post("/admin/inventory/<int:item_id>/toggle")
@login_required("admin")
def toggle_inventory(item_id):
    if valid_csrf():
        execute("UPDATE inventory SET active=CASE active WHEN 1 THEN 0 ELSE 1 END, updated_at=? WHERE id=?", (utcnow().isoformat(), item_id))
        flash("Inventory status updated.", "success")
    return redirect(url_for("admin_dashboard"))


@app.post("/admin/inventory/<int:item_id>/edit")
@login_required("admin")
def edit_inventory(item_id):
    if valid_csrf():
        item = query("SELECT * FROM inventory WHERE id=?", (item_id,), one=True)
        if not item:
            flash("Item not found.", "error")
            return redirect(url_for("admin_dashboard"))
        try:
            total_qty = max(0, int(request.form.get("total_qty", item["total_qty"])))
            available_qty = max(0, min(total_qty, int(request.form.get("available_qty", item["available_qty"]))))
            minimum_stock = max(1, int(request.form.get("minimum_stock", item["minimum_stock"])))
        except ValueError:
            flash("Quantities must be whole numbers.", "error")
            return redirect(url_for("admin_dashboard"))
        execute(
            """UPDATE inventory SET category=?, location=?, total_qty=?, available_qty=?,
            minimum_stock=?, condition=?, updated_at=? WHERE id=?""",
            (
                request.form.get("category", item["category"]).strip() or item["category"],
                request.form.get("location", item["location"]).strip() or item["location"],
                total_qty, available_qty, minimum_stock,
                request.form.get("condition", item["condition"]).strip() or item["condition"],
                utcnow().isoformat(), item_id,
            ),
        )
        audit("INVENTORY_UPDATED", item["name"])
        flash(f"{item['name']} updated.", "success")
    return redirect(url_for("admin_dashboard"))


@app.post("/admin/inventory/<int:item_id>/delete")
@login_required("admin")
def delete_inventory(item_id):
    if valid_csrf():
        item = query("SELECT * FROM inventory WHERE id=?", (item_id,), one=True)
        in_use = query("SELECT COUNT(*) AS n FROM request_items WHERE inventory_id=?", (item_id,), one=True)["n"] if item else 0
        if not item:
            flash("Item not found.", "error")
        elif in_use:
            execute("UPDATE inventory SET active=0, updated_at=? WHERE id=?", (utcnow().isoformat(), item_id))
            audit("INVENTORY_SOFT_DELETED", item["name"])
            flash(f"{item['name']} has request history, so it was deactivated instead of deleted.", "success")
        else:
            execute("DELETE FROM inventory WHERE id=?", (item_id,))
            audit("INVENTORY_DELETED", item["name"])
            flash(f"{item['name']} deleted.", "success")
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

@app.post("/student/issue/<int:issue_id>/extend")
@login_required("student")
def extend_due_date(issue_id):
    issue = query("SELECT * FROM issues WHERE id=? AND student_id=? AND returned_at IS NULL", (issue_id, current_user()["id"]), one=True)
    if not issue or not valid_csrf():
        flash("That issue is no longer active.", "error")
        return redirect(url_for("student_dashboard"))
    if issue["extensions_used"] >= 1:
        flash("This item has already been extended once. Contact the Lab Assistant for further extensions.", "error")
        return redirect(url_for("student_dashboard"))
    if issue["extension_status"] not in ("NONE", "REJECTED"):
        flash("An extension request is already in progress for this item.", "error")
        return redirect(url_for("student_dashboard"))
    reason = request.form.get("reason", "").strip()
    if not reason:
        flash("Please give a reason for the extension.", "error")
        return redirect(url_for("student_dashboard"))
    try:
        days = max(1, min(7, int(request.form.get("days", 3))))
    except ValueError:
        days = 3
    execute(
        """UPDATE issues SET extension_reason=?, extension_requested_days=?, extension_status='PENDING_ADMIN',
        extension_admin_status='PENDING', extension_mentor_status='PENDING', extension_hod_status='PENDING' WHERE id=?""",
        (reason, days, issue["id"]),
    )
    req = query("SELECT request_code FROM requests WHERE id=?", (issue["request_id"],), one=True)
    for admin_row in query("SELECT id FROM users WHERE role='admin'"):
        notify(admin_row["id"], "Extension requested", f"{req['request_code']} — student requested a {days}-day extension: \"{reason}\"", "extension")
    audit("EXTENSION_REQUESTED", f"{req['request_code']} requested {days} day(s): {reason}")
    flash("Extension request sent for approval.", "success")
    return redirect(url_for("student_dashboard"))

@app.post("/admin/issues/<int:issue_id>/extension/decide")
@login_required("admin")
def admin_decide_extension(issue_id):
    if not valid_csrf():
        flash("Your session expired. Please try again.", "error")
        return redirect(url_for("admin_dashboard"))
    issue = query(
        """SELECT issues.*, requests.request_code, requests.mentor_email, requests.mentor_name, requests.hod_email
        FROM issues JOIN requests ON requests.id=issues.request_id
        WHERE issues.id=? AND issues.extension_status='PENDING_ADMIN'""",
        (issue_id,), one=True,
    )
    if not issue:
        flash("That extension is not waiting for Lab Assistant approval.", "error")
        return redirect(url_for("admin_dashboard"))
    decision = request.form.get("decision")
    if decision == "approve":
        execute("UPDATE issues SET extension_admin_status='APPROVED', extension_status='PENDING_MENTOR' WHERE id=?", (issue_id,))
        token = signed_token(issue["request_id"], "mentor_extension")
        link = f"{APP_URL.rstrip('/') if APP_URL else request.host_url.rstrip('/')}{url_for('approve_extension', stage='mentor', token=token)}"
        send_mail(issue["mentor_email"], "Extension Request Requires Your Approval",
                   f"{issue['request_code']} — student requested a {issue['extension_requested_days']}-day extension.\nReason: {issue['extension_reason']}", link)
        notify(issue["student_id"], "Extension approved by Lab Assistant", f"{issue['request_code']} extension is now waiting for Mentor approval.", "extension")
        audit("EXTENSION_ADMIN_APPROVED", issue["request_code"])
        flash("Extension approved and sent to Mentor.", "success")
    elif decision == "reject":
        execute("UPDATE issues SET extension_admin_status='REJECTED', extension_status='REJECTED' WHERE id=?", (issue_id,))
        notify(issue["student_id"], "Extension rejected", "Your extension request was rejected by the Lab Assistant.", "rejection")
        audit("EXTENSION_ADMIN_REJECTED", issue["request_code"])
        flash("Extension rejected.", "success")
    return redirect(url_for("admin_dashboard"))

@app.route("/approval/extension/<stage>/<token>", methods=["GET", "POST"])
def approve_extension(stage, token):
    if stage not in {"mentor", "hod"}:
        return render_template("approval.html", error="That approval page could not be found.", stage=stage)
    try:
        payload = serializer.loads(token, max_age=60 * 60 * 24 * 7)
    except (BadSignature, SignatureExpired):
        return render_template("approval.html", error="This approval link has expired or is invalid.", stage=stage)
    stage_key = f"mentor_extension" if stage == "mentor" else "hod_extension"
    if payload.get("stage") != stage_key:
        return render_template("approval.html", error="This approval link is not valid for this stage.", stage=stage)
    token_row = query("SELECT * FROM approval_tokens WHERE token_hash=? AND stage=?", (hashlib.sha256(token.encode()).hexdigest(), stage_key), one=True)
    if not token_row or token_row["status"] != "PENDING":
        return render_template("approval.html", error="This request has already been reviewed.", stage=stage)
    req = query(
            """SELECT requests.*,
                    users.full_name AS student_name,
                    users.uid,
                    users.branch,
                    users.division,
                    users.year,
                    users.email AS student_email,
                    users.phone AS student_phone
            FROM requests
            JOIN users ON users.id = requests.student_id
            WHERE requests.id=?""",
            (payload["request_id"],),
            one=True,
        )
    issue = query("SELECT * FROM issues WHERE request_id=? AND returned_at IS NULL", (req["id"],), one=True)
    expected = "PENDING_MENTOR" if stage == "mentor" else "PENDING_HOD"
    if not issue or issue["extension_status"] != expected:
        return render_template("approval.html", error="This request has already been reviewed.", stage=stage)
    if request.method == "POST":
        if not valid_csrf():
            flash("Your session expired. Please try again.", "error")
        else:
            decision = request.form.get("decision")
            now = utcnow().isoformat()
            if stage == "mentor":
                if decision == "approve":
                    execute("UPDATE issues SET extension_mentor_status='APPROVED', extension_status='PENDING_HOD' WHERE id=?", (issue["id"],))
                    hod_token = signed_token(req["id"], "hod_extension")
                    link = f"{APP_URL.rstrip('/') if APP_URL else request.host_url.rstrip('/')}{url_for('approve_extension', stage='hod', token=hod_token)}"
                    send_mail(req["hod_email"], "HOD Approval Required — Extension Request",
                               f"Mentor approved the extension for {req['request_code']}. Reason: {issue['extension_reason']}", link)
                    notify(issue["student_id"], "Mentor approved extension", f"{req['request_code']} extension is now waiting for HOD approval.", "extension")
                    message = "MENTOR APPROVAL COMPLETED"
                else:
                    execute("UPDATE issues SET extension_mentor_status='REJECTED', extension_status='REJECTED' WHERE id=?", (issue["id"],))
                    notify(issue["student_id"], "Extension rejected by Mentor", "Your extension request was rejected.", "rejection")
                    message = "MENTOR REJECTION RECORDED"
            else:
                if decision == "approve":
                    new_due = (datetime.fromisoformat(issue["due_date"]).date() + timedelta(days=issue["extension_requested_days"])).isoformat()
                    execute(
                            """
                            UPDATE issues
                            SET due_date=?,
                                due_notified_at=NULL,
                                overdue_notified_at=NULL,
                                extensions_used=extensions_used+1,
                                extension_hod_status='APPROVED',
                                extension_status='APPROVED'
                            WHERE id=?
                            """,
                            (new_due, issue["id"]),)                   
                    execute("UPDATE requests SET due_date=? WHERE id=?", (new_due, req["id"]))
                    notify(issue["student_id"], "Return date extended", f"{req['request_code']} return date moved to {new_due}.", "extension")
                    message = "HOD APPROVAL COMPLETED"
                else:
                    execute("UPDATE issues SET extension_hod_status='REJECTED', extension_status='REJECTED' WHERE id=?", (issue["id"],))
                    notify(issue["student_id"], "Extension rejected by HOD", "Your extension request was rejected.", "rejection")
                    message = "HOD REJECTION RECORDED"
            execute("UPDATE approval_tokens SET status='USED', used_at=? WHERE id=?", (now, token_row["id"]))
            return render_template("approval.html", complete=message, stage=stage)
    return render_template("approval.html", stage=stage, extension_reason=issue["extension_reason"], extension_days=issue["extension_requested_days"], req=req)

@app.route("/admin/returns", methods=["GET", "POST"])
@login_required("admin")
def returns():
    if request.method == "POST" and valid_csrf():
        request_code = request.form.get("request_code", "").strip()
        entered_otp = request.form.get("otp", "").strip()

        req = query(
            "SELECT * FROM requests WHERE request_code=?",
            (request_code,),
            one=True,
        )

        otp = (
                query(
                    """SELECT * FROM otp_tokens
                    WHERE request_id=?
                    AND student_id=?
                    AND kind='return'
                    AND used_at IS NULL
                    ORDER BY id DESC
                    LIMIT 1""",
                    (req["id"], req["student_id"]),
                    one=True,
                )
                if req
                else None
            )
    

        if not req:
            flash("Invalid request code.", "error")

        elif not otp:
            flash(
                "No active return OTP found. "
                "Ask the student to generate a new OTP.",
                "error",
            )

        elif datetime.fromisoformat(otp["expires_at"]) < utcnow():
            execute(
                "UPDATE otp_tokens SET used_at=? WHERE id=?",
                (utcnow().isoformat(), otp["id"]),
            )
            flash(
                "This return OTP has expired. "
                "Ask the student to generate a new OTP.",
                "error",
            )

        elif otp["attempts"] >= 5:
            execute(
                "UPDATE otp_tokens SET used_at=? WHERE id=?",
                (utcnow().isoformat(), otp["id"]),
            )
            flash(
                "This return OTP has reached the maximum attempts. "
                "Ask the student to generate a new OTP.",
                "error",
            )

        elif not secrets.compare_digest(
            otp["otp_hash"],
            hashlib.sha256(entered_otp.encode()).hexdigest(),
        ):
            new_attempts = otp["attempts"] + 1

            if new_attempts >= 5:
                execute(
                    """UPDATE otp_tokens
                       SET attempts=?, used_at=?
                       WHERE id=?""",
                    (
                        new_attempts,
                        utcnow().isoformat(),
                        otp["id"],
                    ),
                )
                flash(
                    "Incorrect OTP. This OTP has reached the maximum "
                    "attempts. Ask the student to generate a new OTP.",
                    "error",
                )
            else:
                execute(
                    "UPDATE otp_tokens SET attempts=? WHERE id=?",
                    (new_attempts, otp["id"]),
                )
                flash(
                    f"Incorrect return OTP. "
                    f"{5 - new_attempts} attempt(s) remaining.",
                    "error",
                )

        else:
            # OTP is correct
            issue = query(
                """SELECT * FROM issues
                   WHERE request_id=?
                   AND student_id=?
                   AND returned_at IS NULL""",
                (req["id"], req["student_id"]),
                one=True,
            )

            if not issue:
                flash(
                    "No active issue was found for this request.",
                    "error",
                )
            else:
                condition = request.form.get("condition", "GOOD")

                execute(
                    """UPDATE issues
                       SET returned_at=?,
                           return_condition=?,
                           return_remarks=?
                       WHERE id=?""",
                    (
                        utcnow().isoformat(),
                        condition,
                        request.form.get("remarks", "").strip(),
                        issue["id"],
                    ),
                )

                for item in request_items(req["id"]):
                    execute(
                        """UPDATE inventory
                           SET available_qty=MIN(total_qty, available_qty+?),
                               updated_at=?
                           WHERE name=?""",
                        (
                            item["quantity"],
                            utcnow().isoformat(),
                            item["name"],
                        ),
                    )

                execute(
                    "UPDATE otp_tokens SET used_at=? WHERE id=?",
                    (utcnow().isoformat(), otp["id"]),
                )

                notify(
                    req["student_id"],
                    "Components returned",
                    f"Return completed for {req['request_code']} "
                    f"with condition {condition}.",
                    "return",
                )

                audit(
                    "COMPONENTS_RETURNED",
                    req["request_code"],
                )

                flash(
                    "Return completed and inventory updated.",
                    "success",
                )

    active = query(
        """SELECT issues.*, requests.request_code, users.full_name
           FROM issues
           JOIN requests ON requests.id=issues.request_id
           JOIN users ON users.id=issues.student_id
           WHERE issues.returned_at IS NULL
           ORDER BY issues.due_date"""
    )

    return render_template("returns.html", active=active)


@app.route("/notifications")
@login_required()
def notifications_page():
    user = current_user()
    items = query("SELECT * FROM notifications WHERE user_id=? ORDER BY created_at DESC LIMIT 100", (user["id"],))
    return render_template("notifications.html", items=items)


@app.route("/notifications/read-all", methods=["POST"])
@login_required()
def read_notifications():
    if valid_csrf():
        execute("UPDATE notifications SET is_read=1 WHERE user_id=?", (current_user()["id"],))
    return redirect(request.referrer or url_for("dashboard"))

def normalize_ai_text(text):
    """Normalize project text for matching."""
    if not text:
        return ""

    text = text.lower()

    replacements = {
        "air pollution": " air_quality pollution ",
        "air quality": " air_quality pollution ",
        "gas leakage": " gas leak ",
        "gas leak detection": " gas leak ",
        "smart agriculture": " agriculture farming ",
        "smart farming": " agriculture farming ",
        "weather monitoring": " weather temperature humidity ",
        "health monitoring": " healthcare health ",
        "fire detection": " fire flame temperature ",
        "security system": " security motion alarm ",
        "home automation": " automation smart_home ",
        "water monitoring": " water flow quality ",
        "water level": " water level tank ",
        "soil monitoring": " soil agriculture moisture ",
    }

    for old, new in replacements.items():
        text = text.replace(old, new)

    return text


# Each capability contains words which describe the capability,
# and inventory keywords which are likely to satisfy it.
AI_CAPABILITIES = {
    "air_quality": {
        "triggers": [
            "air",
            "pollution",
            "polluted",
            "air_quality",
            "environmental",
            "smoke",
            "toxic",
            "gas",
            "gas_detection",
            "gas_leak",
            "gas leakage",
        ],
        "keywords": [
            "mq-2",
            "mq2",
            "mq-3",
            "mq3",
            "mq-4",
            "mq4",
            "mq-5",
            "mq5",
            "mq-6",
            "mq6",
            "mq-7",
            "mq7",
            "mq-8",
            "mq8",
            "mq-9",
            "mq9",
            "mq-135",
            "mq135",
            "gas sensor",
            "air quality",
            "air sensor",
            "smoke sensor",
            "pollution sensor",
        ],
        "reason": "Useful for detecting gases, smoke, or air-quality conditions.",
    },

    "temperature": {
        "triggers": [
            "temperature",
            "heat",
            "thermal",
            "hot",
            "cold",
            "environment",
            "weather",
        ],
        "keywords": [
            "temperature sensor",
            "dht11",
            "dht22",
            "ds18",
            "lm35",
            "bmp180",
            "bmp280",
            "bme280",
            "thermistor",
        ],
        "reason": "Useful for measuring temperature or environmental conditions.",
    },

    "humidity": {
        "triggers": [
            "humidity",
            "moisture",
            "humid",
            "environment",
            "weather",
        ],
        "keywords": [
            "dht11",
            "dht22",
            "bme280",
            "humidity sensor",
        ],
        "reason": "Useful for measuring humidity and environmental moisture.",
    },

    "soil_moisture": {
        "triggers": [
            "soil",
            "agriculture",
            "farming",
            "plant",
            "irrigation",
            "crop",
            "moisture",
        ],
        "keywords": [
            "soil moisture",
            "soil sensor",
            "moisture sensor",
        ],
        "reason": "Useful for measuring soil moisture for irrigation and agriculture projects.",
    },

    "water_flow": {
        "triggers": [
            "water",
            "flow",
            "pipeline",
            "pipe",
            "water monitoring",
            "liquid",
        ],
        "keywords": [
            "water flow",
            "flow sensor",
            "yfs201",
            "yfs 201",
            "water sensor",
        ],
        "reason": "Useful for measuring water or liquid flow.",
    },

    "water_level": {
        "triggers": [
            "water level",
            "tank",
            "reservoir",
            "overflow",
            "water monitoring",
        ],
        "keywords": [
            "water level",
            "water sensor",
            "ultrasonic",
            "float sensor",
        ],
        "reason": "Useful for monitoring water level and detecting tank conditions.",
    },

    "motion": {
        "triggers": [
            "motion",
            "movement",
            "intrusion",
            "security",
            "person detection",
            "human detection",
        ],
        "keywords": [
            "pir",
            "pir sensor",
            "motion sensor",
            "ir sensor",
            "infrared sensor",
        ],
        "reason": "Useful for detecting movement or human presence.",
    },

    "distance": {
        "triggers": [
            "distance",
            "obstacle",
            "parking",
            "range",
            "proximity",
            "object detection",
        ],
        "keywords": [
            "ultrasonic",
            "hc-sr04",
            "distance sensor",
            "proximity sensor",
            "ir sensor",
            "infrared sensor",
        ],
        "reason": "Useful for measuring distance or detecting nearby objects.",
    },

    "light": {
        "triggers": [
            "light",
            "brightness",
            "darkness",
            "illumination",
            "automatic light",
        ],
        "keywords": [
            "ldr",
            "light sensor",
            "ldr sensor",
            "photoresistor",
        ],
        "reason": "Useful for detecting ambient light intensity.",
    },

    "flame": {
        "triggers": [
            "fire",
            "flame",
            "burning",
            "fire detection",
        ],
        "keywords": [
            "flame sensor",
            "fire sensor",
            "infrared flame",
        ],
        "reason": "Useful for detecting flame or fire conditions.",
    },

    "security": {
        "triggers": [
            "security",
            "access control",
            "authentication",
            "fingerprint",
            "biometric",
            "door lock",
            "attendance",
        ],
        "keywords": [
            "fingerprint",
            "r307",
            "fingerprint sensor",
            "rfid",
            "rfid module",
            "keypad",
        ],
        "reason": "Useful for identification, access control, or security applications.",
    },

    "display": {
        "triggers": [
            "display",
            "screen",
            "monitor",
            "show readings",
            "visualization",
            "dashboard",
            "lcd",
            "oled",
        ],
        "keywords": [
            "lcd",
            "oled",
            "display",
            "screen",
            "tft",
        ],
        "reason": "Useful for displaying sensor readings, status, or project output.",
    },

    "alert": {
        "triggers": [
            "alert",
            "warning",
            "alarm",
            "notification",
            "buzzer",
            "danger",
        ],
        "keywords": [
            "buzzer",
            "alarm",
            "buzzer module",
            "speaker",
            "led",
            "led module",
        ],
        "reason": "Useful for generating warnings or alerts when a condition is detected.",
    },

    "communication": {
        "triggers": [
            "wifi",
            "wireless",
            "internet",
            "iot",
            "remote monitoring",
            "cloud",
        ],
        "keywords": [
            "esp32",
            "esp8266",
            "wifi",
            "bluetooth",
            "gsm",
            "zigbee",
            "arduino",
        ],
        "reason": "Useful for communication, IoT connectivity, or remote monitoring.",
    },

    "controller": {
        "triggers": [
            "iot",
            "smart system",
            "automation",
            "monitoring",
            "controller",
            "embedded",
            "microcontroller",
        ],
        "keywords": [
            "esp32",
            "esp8266",
            "arduino",
            "uno",
            "nano",
            "microcontroller",
            "development board",
        ],
        "reason": "Provides the controller needed to read sensors and run the application logic.",
    },

    "automation": {
        "triggers": [
            "automation",
            "automatic",
            "smart home",
            "control",
            "switch",
            "relay",
        ],
        "keywords": [
            "relay",
            "arduino",
            "esp32",
            "esp8266",
            "servo",
            "motor",
        ],
        "reason": "Useful for automatically controlling devices or actuators.",
    },

    "motor": {
        "triggers": [
            "motor",
            "robot",
            "robotics",
            "movement",
            "wheel",
            "actuator",
        ],
        "keywords": [
            "motor",
            "servo",
            "stepper",
            "dc motor",
        ],
        "reason": "Useful for mechanical movement and actuation.",
    },

    "position": {
        "triggers": [
            "gps",
            "location",
            "tracking",
            "vehicle tracking",
            "navigation",
        ],
        "keywords": [
            "gps",
            "neo-6m",
            "gps module",
            "location module",
        ],
        "reason": "Useful for location tracking and navigation.",
    },
}


def calculate_component_score(component_name, category, description, project_text):
    """
    Score a real inventory component against the project's
    detected capabilities.

    Returns:
        score, reasons, matched_capabilities
    """
    name = (component_name or "").lower()
    category = (category or "").lower()
    description = (description or "").lower()

    searchable = f"{name} {category} {description}"

    normalized_project = normalize_ai_text(project_text)

    detected_capabilities = []

    for capability, rules in AI_CAPABILITIES.items():
        if any(trigger in normalized_project for trigger in rules["triggers"]):
            detected_capabilities.append(capability)

    # If a project is vague, don't blindly recommend random components.
    if not detected_capabilities:
        return 0, [], []

    score = 0
    reasons = []
    matched = []

    for capability in detected_capabilities:
        rules = AI_CAPABILITIES[capability]

        local_score = 0

        for keyword in rules["keywords"]:
            keyword = keyword.lower()

            if keyword in searchable:
                # Exact component-name matches are strongest.
                if keyword in name:
                    local_score += 12
                elif keyword in category:
                    local_score += 7
                else:
                    local_score += 4

        if local_score > 0:
            matched.append(capability)
            reasons.append(rules["reason"])
            score += local_score

    # Give a small boost to development boards for actual IoT projects.
    if "iot" in normalized_project or "smart" in normalized_project:
        if any(x in searchable for x in [
            "esp32",
            "esp8266",
            "arduino",
            "microcontroller",
            "development board",
        ]):
            score += 6

    return score, reasons, matched


def get_ai_recommendations(project_title, purpose):
    """
    Generate recommendations strictly from the real LabVault inventory.

    Returns:
        recommendations
        unavailable
        alternatives
    """
    project_text = f"{project_title} {purpose}".strip()

    inventory = query(
        """
        SELECT
            id,
            name,
            category,
            description,
            location,
            total_qty,
            available_qty,
            minimum_stock,
            active
        FROM inventory
        WHERE active=1
        ORDER BY name
        """
    )

    scored = []

    for item in inventory:
        score, reasons, matched_capabilities = calculate_component_score(
            item["name"],
            item["category"],
            item["description"],
            project_text,
        )

        if score <= 0:
            continue

        scored.append({
            "id": item["id"],
            "name": item["name"],
            "category": item["category"],
            "description": item["description"],
            "location": item["location"],
            "total_qty": item["total_qty"],
            "available_qty": item["available_qty"],
            "minimum_stock": item["minimum_stock"],
            "score": score,
            "reason": " ".join(dict.fromkeys(reasons)),
            "matched_capabilities": matched_capabilities,
        })

    # Sort strongest project matches first.
    scored.sort(
        key=lambda x: (
            x["score"],
            x["available_qty"] > 0,
            x["available_qty"],
        ),
        reverse=True,
    )

    # Remove duplicate-looking components by normalized name.
    unique = []
    seen_names = set()

    for item in scored:
        normalized_name = " ".join(
            (item["name"] or "").lower().split()
        )

        if normalized_name in seen_names:
            continue

        seen_names.add(normalized_name)
        unique.append(item)

    available = []
    unavailable = []

    for item in unique:
        # Component exists but is currently unavailable.
        if item["available_qty"] <= 0:
            item["status"] = "unavailable"
            unavailable.append(item)
            continue

        # Only genuinely relevant components are shown.
        item["status"] = "available"

        # Suggested quantity based on the role of the component.
        if "controller" in item["matched_capabilities"]:
            quantity = 1
        elif "display" in item["matched_capabilities"]:
            quantity = 1
        elif "communication" in item["matched_capabilities"]:
            quantity = 1
        elif "alert" in item["matched_capabilities"]:
            quantity = 1
        else:
            quantity = 1

        item["quantity"] = min(quantity, item["available_qty"])
        available.append(item)

    # Keep the result useful instead of dumping the whole inventory.
    recommendations = available[:8]
    unavailable = unavailable[:8]

    # Alternatives are real available inventory items that partially
    # satisfy the detected capability but scored lower.
    recommendation_ids = {item["id"] for item in recommendations}

    alternatives = [
        {
            "id": item["id"],
            "name": item["name"],
            "category": item["category"],
            "available_qty": item["available_qty"],
            "reason": item["reason"],
        }
        for item in available[8:16]
        if item["id"] not in recommendation_ids
    ]

    return recommendations, unavailable, alternatives


@app.route("/student/assistant", methods=["POST"])
@login_required("student")
def assistant():
    """
    LabVault AI Project Assistant.

    The assistant ONLY recommends components that exist in the
    real LabVault inventory. It does not modify stock and does
    not issue components.
    """
    title = request.form.get("title", "").strip()
    purpose = request.form.get("purpose", "").strip()

    if not title or not purpose:
        flash(
            "Please provide both a project title and project purpose.",
            "error",
        )
        return redirect(url_for("student_dashboard") + "#assistant")

    try:
        recommendations, unavailable, alternatives = get_ai_recommendations(
            title,
            purpose,
        )

        if not recommendations and not unavailable:
            flash(
                "The assistant could not find relevant components in "
                "the current LabVault inventory. Try describing the "
                "sensors, monitoring task, or hardware you need.",
                "info",
            )

            return render_template(
                "assistant.html",
                title=title,
                purpose=purpose,
                recommendations=[],
                unavailable=[],
                alternatives=[],
            )

        return render_template(
            "assistant.html",
            title=title,
            purpose=purpose,
            recommendations=recommendations,
            unavailable=unavailable,
            alternatives=alternatives,
        )

    except Exception:
        app.logger.exception("AI Project Assistant failed.")

        return render_template(
            "error.html",
            code=500,
            title="Assistant unavailable",
            message=(
                "We could not generate the component recommendations "
                "right now. Please try again."
            ),
        ), 500

@app.errorhandler(500)
def server_error(_error):
    return render_template("error.html", code=500, title="Something went wrong", message="We could not complete that action. Your data has not been removed."), 500

@app.get("/admin/export/master-stock")
@login_required("admin")
def export_master_stock():
    rows = query(
        """SELECT stock_number, name, total_qty, location
           FROM inventory
           WHERE active=1
           AND source_file='master_stock'
           ORDER BY name"""
    )

    output = StringIO()
    writer = csv.writer(output)

    writer.writerow([
        "Sr No",
        "Stock Register No",
        "Name of the Article",
        "Qty",
        "Allocated to",
    ])

    for index, row in enumerate(rows, start=1):
        writer.writerow([
            index,
            row["stock_number"] or "",
            row["name"] or "",
            row["total_qty"] or 0,
            row["location"] or "",
        ])

    csv_data = "\ufeff" + output.getvalue()

    return Response(
        csv_data,
        mimetype="text/csv",
        headers={
            "Content-Disposition":
                "attachment; filename=labvault_master_stock.csv"
        },
    )

@app.get("/admin/export/consumables")
@login_required("admin")
def export_consumables():
    rows = query(
        """SELECT name, total_qty, location
           FROM inventory
           WHERE active=1
           AND source_file='consumables'
           ORDER BY name"""
    )

    output = StringIO()
    writer = csv.writer(output)

    writer.writerow([
        "Sr.No",
        "Name of the Article",
        "QTY",
        "Location",
    ])

    for index, row in enumerate(rows, start=1):
        writer.writerow([
            index,
            row["name"] or "",
            row["total_qty"] or 0,
            row["location"] or "",
        ])

    csv_data = "\ufeff" + output.getvalue()

    return Response(
        csv_data,
        mimetype="text/csv",
        headers={
            "Content-Disposition":
                "attachment; filename=labvault_consumables.csv"
        },
    )

@app.get("/admin/export/inventory")
@login_required("admin")
def export_full_inventory():
    rows = query(
        """SELECT
               source_file,
               stock_number,
               name,
               category,
               location,
               total_qty,
               available_qty,
               minimum_stock,
               condition,
               maintenance_status,
               active
           FROM inventory
           ORDER BY source_file, name"""
    )

    output = StringIO()
    writer = csv.writer(output)

    writer.writerow([
        "Source",
        "Stock Register No",
        "Name of the Article",
        "Category",
        "Location",
        "Total Qty",
        "Available Qty",
        "Minimum Stock",
        "Condition",
        "Maintenance Status",
        "Active",
    ])

    for row in rows:
        writer.writerow([
            row["source_file"] or "",
            row["stock_number"] or "",
            row["name"] or "",
            row["category"] or "",
            row["location"] or "",
            row["total_qty"] or 0,
            row["available_qty"] or 0,
            row["minimum_stock"] or 0,
            row["condition"] or "",
            row["maintenance_status"] or "",
            "Yes" if row["active"] else "No",
        ])

    csv_data = "\ufeff" + output.getvalue()

    return Response(
        csv_data,
        mimetype="text/csv",
        headers={
            "Content-Disposition":
                "attachment; filename=labvault_complete_inventory.csv"
        },
    )

@app.get("/admin/print/inventory")
@login_required("admin")
def print_inventory_report():
    rows = query(
        """SELECT
               source_file,
               stock_number,
               name,
               category,
               location,
               total_qty,
               available_qty,
               minimum_stock,
               condition,
               maintenance_status,
               active
           FROM inventory
           ORDER BY source_file, name"""
    )

    return render_template(
        "inventory_print.html",
        items=rows,
        generated_at=format_ist(utcnow().isoformat()),
    )

with app.app_context():
    init_db()


if __name__ == "__main__":
    port = int(os.getenv("PORT", "5000"))

    alert_thread = threading.Thread(
        target=return_alert_worker,
        daemon=True,
        name="labvault-return-alert-worker",
    )
    alert_thread.start()

    print("LabVault AI Lab Inventory Management")
    print(f"Server running at http://127.0.0.1:{port}")

    app.run(
        host="0.0.0.0",
        port=port,
        debug=False,
    )