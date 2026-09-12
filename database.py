import sqlite3
import os
import json
from datetime import datetime, timezone

DB_PATH = os.path.join(os.environ.get('DATA_DIR', os.path.join(os.path.dirname(__file__), 'data')), 'recon.db')


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute('PRAGMA foreign_keys = ON')
    return conn


def init_db():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = get_conn()
    conn.executescript('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            display_name TEXT,
            is_admin INTEGER DEFAULT 0,
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS runs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            run_label TEXT,
            created_at TEXT NOT NULL,
            summary_json TEXT NOT NULL,
            report_path TEXT NOT NULL,
            detail_path TEXT,
            gl_3493_name TEXT,
            gl_3496_name TEXT,
            gl_345051_name TEXT,
            hdfc_name TEXT,
            FOREIGN KEY (user_id) REFERENCES users(id)
        );
    ''')
    conn.commit()
    # Migration safety-net: add columns that may not exist on databases
    # created before these features, rather than requiring a fresh database.
    existing_run_cols = {row['name'] for row in conn.execute("PRAGMA table_info(runs)").fetchall()}
    if 'detail_path' not in existing_run_cols:
        conn.execute('ALTER TABLE runs ADD COLUMN detail_path TEXT')
        conn.commit()

    existing_user_cols = {row['name'] for row in conn.execute("PRAGMA table_info(users)").fetchall()}
    user_migrations = {
        'must_change_password': 'INTEGER DEFAULT 0',
        'failed_attempts': 'INTEGER DEFAULT 0',
        'is_locked': 'INTEGER DEFAULT 0',
    }
    for col, coltype in user_migrations.items():
        if col not in existing_user_cols:
            conn.execute(f'ALTER TABLE users ADD COLUMN {col} {coltype}')
            conn.commit()
    conn.close()


def create_user(username, password_hash, display_name=None, is_admin=0, must_change_password=1):
    conn = get_conn()
    conn.execute(
        '''INSERT INTO users (username, password_hash, display_name, is_admin, created_at, must_change_password)
           VALUES (?,?,?,?,?,?)''',
        (username, password_hash, display_name or username, is_admin,
         datetime.now(timezone.utc).isoformat(), must_change_password)
    )
    conn.commit()
    conn.close()


def get_user_by_username(username):
    conn = get_conn()
    row = conn.execute('SELECT * FROM users WHERE username = ?', (username,)).fetchone()
    conn.close()
    return row


def get_user_by_id(user_id):
    conn = get_conn()
    row = conn.execute('SELECT * FROM users WHERE id = ?', (user_id,)).fetchone()
    conn.close()
    return row


def any_users_exist():
    conn = get_conn()
    row = conn.execute('SELECT COUNT(*) as c FROM users').fetchone()
    conn.close()
    return row['c'] > 0


def count_admins():
    conn = get_conn()
    row = conn.execute('SELECT COUNT(*) as c FROM users WHERE is_admin = 1').fetchone()
    conn.close()
    return row['c']


def list_users():
    conn = get_conn()
    rows = conn.execute(
        'SELECT id, username, display_name, is_admin, must_change_password, failed_attempts, is_locked, created_at '
        'FROM users ORDER BY id'
    ).fetchall()
    conn.close()
    return rows


def delete_user(user_id):
    conn = get_conn()
    conn.execute('DELETE FROM users WHERE id = ?', (user_id,))
    conn.commit()
    conn.close()


def record_failed_login(user_id, max_attempts=5):
    """Increment the failed-attempt counter; lock the account once it hits
    max_attempts. Returns (attempts_used, is_now_locked)."""
    conn = get_conn()
    conn.execute('UPDATE users SET failed_attempts = failed_attempts + 1 WHERE id = ?', (user_id,))
    row = conn.execute('SELECT failed_attempts FROM users WHERE id = ?', (user_id,)).fetchone()
    attempts = row['failed_attempts']
    now_locked = attempts >= max_attempts
    if now_locked:
        conn.execute('UPDATE users SET is_locked = 1 WHERE id = ?', (user_id,))
    conn.commit()
    conn.close()
    return attempts, now_locked


def reset_failed_attempts(user_id):
    conn = get_conn()
    conn.execute('UPDATE users SET failed_attempts = 0 WHERE id = ?', (user_id,))
    conn.commit()
    conn.close()


def unlock_user(user_id):
    conn = get_conn()
    conn.execute('UPDATE users SET is_locked = 0, failed_attempts = 0 WHERE id = ?', (user_id,))
    conn.commit()
    conn.close()


def set_password(user_id, password_hash, must_change_password=0):
    conn = get_conn()
    conn.execute(
        'UPDATE users SET password_hash = ?, must_change_password = ?, failed_attempts = 0, is_locked = 0 WHERE id = ?',
        (password_hash, must_change_password, user_id)
    )
    conn.commit()
    conn.close()


def save_run(user_id, run_label, summary, report_path, filenames, detail_path=None):
    conn = get_conn()
    cur = conn.execute(
        '''INSERT INTO runs (user_id, run_label, created_at, summary_json, report_path, detail_path,
           gl_3493_name, gl_3496_name, gl_345051_name, hdfc_name)
           VALUES (?,?,?,?,?,?,?,?,?,?)''',
        (user_id, run_label, datetime.now(timezone.utc).isoformat(), json.dumps(summary), report_path, detail_path,
         filenames.get('3493'), filenames.get('3496'), filenames.get('345051'), filenames.get('hdfc'))
    )
    conn.commit()
    run_id = cur.lastrowid
    conn.close()
    return run_id


def update_run_summary(run_id, summary):
    conn = get_conn()
    conn.execute('UPDATE runs SET summary_json = ? WHERE id = ?', (json.dumps(summary), run_id))
    conn.commit()
    conn.close()


def list_runs(user_id=None, limit=100):
    conn = get_conn()
    if user_id is not None:
        rows = conn.execute(
            '''SELECT runs.*, COALESCE(users.display_name, '(deleted user)') as display_name FROM runs
               LEFT JOIN users ON users.id = runs.user_id
               WHERE runs.user_id = ? ORDER BY runs.id DESC LIMIT ?''',
            (user_id, limit)
        ).fetchall()
    else:
        rows = conn.execute(
            '''SELECT runs.*, COALESCE(users.display_name, '(deleted user)') as display_name FROM runs
               LEFT JOIN users ON users.id = runs.user_id
               ORDER BY runs.id DESC LIMIT ?''',
            (limit,)
        ).fetchall()
    conn.close()
    return rows


def get_run(run_id):
    conn = get_conn()
    row = conn.execute(
        '''SELECT runs.*, COALESCE(users.display_name, '(deleted user)') as display_name FROM runs
           LEFT JOIN users ON users.id = runs.user_id
           WHERE runs.id = ?''',
        (run_id,)
    ).fetchone()
    conn.close()
    return row
