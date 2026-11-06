"""Users, projects and each project's migration objects (SQLite, data/app.db); uploaded and generated files under
data/projects/<project id>/<object slug>/. Passwords are stored as scrypt hashes."""
import hashlib
import hmac
import json
import os
import re
import secrets
import shutil
import sqlite3
from contextlib import closing
from datetime import datetime, timezone
from pathlib import Path

DATA = Path(__file__).parent / "data"
DB = DATA / "app.db"
STATUSES = ["Template needed", "Template provided", "Source data uploaded", "Prefilled"]


def _conn():
    DATA.mkdir(parents=True, exist_ok=True)
    c = sqlite3.connect(DB)
    c.row_factory = sqlite3.Row
    c.executescript("""
        CREATE TABLE IF NOT EXISTS users (email TEXT PRIMARY KEY, name TEXT, pw_hash TEXT, salt TEXT,
                                          is_admin INTEGER DEFAULT 0, created TEXT);
        CREATE TABLE IF NOT EXISTS projects (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, target TEXT,
                                             owner TEXT, created TEXT);
        CREATE TABLE IF NOT EXISTS objects (project_id INTEGER, name TEXT, status TEXT, template TEXT,
                                            source TEXT, result TEXT, updated TEXT, PRIMARY KEY (project_id, name));
    """)
    return c


def _now():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


# ---------------------------------------------------------------- users

def _hash(password, salt):
    return hashlib.scrypt(password.encode(), salt=bytes.fromhex(salt), n=2 ** 14, r=8, p=1).hex()


def clean_email(email):
    return (email or "").strip().lower()


def add_user(email, password, name="", is_admin=False):
    salt = secrets.token_hex(16)
    with closing(_conn()) as c, c:
        c.execute("INSERT OR REPLACE INTO users VALUES (?, ?, ?, ?, ?, ?)",
                  (clean_email(email), name, _hash(password, salt), salt, int(is_admin), _now()))


def check_login(email, password):
    with closing(_conn()) as c:
        row = c.execute("SELECT * FROM users WHERE email = ?", (clean_email(email),)).fetchone()
    if row and hmac.compare_digest(row["pw_hash"], _hash(password, row["salt"])):
        return dict(row)
    return None


def users():
    with closing(_conn()) as c:
        return [dict(r) for r in c.execute("SELECT email, name, is_admin, created FROM users ORDER BY created")]


def delete_user(email):
    with closing(_conn()) as c, c:
        c.execute("DELETE FROM users WHERE email = ?", (clean_email(email),))


# ---------------------------------------------------------------- projects

def create_project(name, target, owner, object_names):
    with closing(_conn()) as c, c:
        cur = c.execute("INSERT INTO projects (name, target, owner, created) VALUES (?, ?, ?, ?)",
                        (name.strip(), target, owner, _now()))
        pid = cur.lastrowid
        c.executemany("INSERT INTO objects VALUES (?, ?, ?, NULL, NULL, NULL, ?)",
                      [(pid, n, STATUSES[0], _now()) for n in object_names])
    return pid


def projects():
    with closing(_conn()) as c:
        rows = [dict(r) for r in c.execute("SELECT * FROM projects ORDER BY id DESC")]
        for r in rows:
            r["objects"] = [dict(o) for o in c.execute("SELECT name, status FROM objects WHERE project_id = ?",
                                                        (r["id"],))]
    return rows


def project(pid):
    with closing(_conn()) as c:
        row = c.execute("SELECT * FROM projects WHERE id = ?", (pid,)).fetchone()
    return dict(row) if row else None


def delete_project(pid):
    with closing(_conn()) as c, c:
        c.execute("DELETE FROM objects WHERE project_id = ?", (pid,))
        c.execute("DELETE FROM projects WHERE id = ?", (pid,))
    shutil.rmtree(DATA / "projects" / str(pid), ignore_errors=True)


def add_objects(pid, object_names):
    with closing(_conn()) as c, c:
        c.executemany("INSERT OR IGNORE INTO objects VALUES (?, ?, ?, NULL, NULL, NULL, ?)",
                      [(pid, n, STATUSES[0], _now()) for n in object_names])


def objects(pid):
    with closing(_conn()) as c:
        rows = [dict(r) for r in c.execute("SELECT * FROM objects WHERE project_id = ? ORDER BY rowid", (pid,))]
    for r in rows:
        r["result"] = json.loads(r["result"]) if r["result"] else None
    return rows


def update_object(pid, name, **fields):
    if "result" in fields and fields["result"] is not None:
        fields["result"] = json.dumps(fields["result"], default=str)
    fields["updated"] = _now()
    with closing(_conn()) as c, c:
        c.execute(f"UPDATE objects SET {', '.join(f'{k} = ?' for k in fields)} WHERE project_id = ? AND name = ?",
                  (*fields.values(), pid, name))


def folder(pid, name):
    path = DATA / "projects" / str(pid) / (re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-") or "object")
    path.mkdir(parents=True, exist_ok=True)
    return path


def save_upload(pid, name, kind, filename, content):
    """Store an uploaded file as <kind>__<its own name> in the object's folder (replacing an earlier one)."""
    stem, ext = os.path.splitext(os.path.basename(filename))
    safe = re.sub(r"[^\w.\- ()]+", "_", stem).strip() or kind
    for old in folder(pid, name).glob(f"{kind}__*"):
        old.unlink()
    path = folder(pid, name) / f"{kind}__{safe}{ext.lower()}"
    path.write_bytes(content)
    return str(path)


def display_name(path):
    """The uploaded file's own name (without the <kind>__ prefix)."""
    return Path(path).name.split("__", 1)[-1] if path else ""
