from __future__ import annotations
import sqlite3
from pathlib import Path


def db_path(data_dir: str) -> Path:
    return Path(data_dir) / "data.db"


def connect(data_dir: str) -> sqlite3.Connection:
    path = db_path(data_dir)
    conn = sqlite3.connect(path)
    conn.row_factory = sqlite3.Row
    # performans
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA synchronous=NORMAL;")
    return conn


def migrate(conn: sqlite3.Connection) -> None:
    conn.execute("""
    CREATE TABLE IF NOT EXISTS sequence (
        name TEXT PRIMARY KEY,
        value INTEGER NOT NULL
    );
    """)
    # 1000'den başlat
    conn.execute("""
    INSERT OR IGNORE INTO sequence(name, value) VALUES ('reservation', 1000);
    """)

    conn.execute("""
    CREATE TABLE IF NOT EXISTS reservations (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        advertiser_name TEXT NOT NULL,
        year INTEGER NOT NULL,
        month INTEGER NOT NULL,
        channel_name TEXT NOT NULL,
        status TEXT NOT NULL, -- 'confirmed'
        reservation_no TEXT UNIQUE, -- sadece confirmed için dolu
        payload_json TEXT NOT NULL,
        created_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    """)

    conn.execute("CREATE INDEX IF NOT EXISTS idx_res_adv ON reservations(advertiser_name);")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_res_adv_ym ON reservations(advertiser_name, year, month);")

    conn.commit()


def next_reservation_seq(conn: sqlite3.Connection) -> int:
    # transaction içinde çağır
    row = conn.execute("SELECT value FROM sequence WHERE name='reservation'").fetchone()
    val = int(row["value"])
    conn.execute("UPDATE sequence SET value = value + 1 WHERE name='reservation'")
    return val
