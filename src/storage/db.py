from __future__ import annotations

import sqlite3
from pathlib import Path

SCHEMA_SQL = """
PRAGMA journal_mode=WAL;

CREATE TABLE IF NOT EXISTS meta (
  key TEXT PRIMARY KEY,
  value TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS advertisers (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS reservations (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  reservation_no TEXT UNIQUE,
  advertiser_name TEXT NOT NULL,
  created_at TEXT NOT NULL,
  is_confirmed INTEGER NOT NULL DEFAULT 0,
  payload_json TEXT NOT NULL
);

-- İleride kanal/fiyat tanımı için hazır dursun (MVP'de boş kalabilir)
CREATE TABLE IF NOT EXISTS channels (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL UNIQUE,
  is_active INTEGER NOT NULL DEFAULT 1
);

CREATE TABLE IF NOT EXISTS channel_prices (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  channel_id INTEGER NOT NULL,
  price_dt REAL NOT NULL,
  price_odt REAL NOT NULL,
  FOREIGN KEY(channel_id) REFERENCES channels(id)
);
"""

def ensure_data_folders(data_dir: Path) -> None:
    data_dir.mkdir(parents=True, exist_ok=True)
    (data_dir / "exports").mkdir(parents=True, exist_ok=True)

def connect_db(db_path: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn

def _has_column(conn: sqlite3.Connection, table: str, column: str) -> bool:
    cur = conn.execute(f"PRAGMA table_info({table})")
    return any(r["name"] == column for r in cur.fetchall())

def _ensure_column(conn: sqlite3.Connection, table: str, ddl: str, column: str) -> None:
    # ddl örn: "ALTER TABLE reservations ADD COLUMN is_confirmed INTEGER NOT NULL DEFAULT 0"
    if not _has_column(conn, table, column):
        conn.execute(ddl)

def _table_cols(conn: sqlite3.Connection, table: str) -> dict[str, dict]:
    cur = conn.execute(f"PRAGMA table_info({table})")
    cols = {}
    for r in cur.fetchall():
        cols[r["name"]] = {"notnull": r["notnull"], "dflt": r["dflt_value"]}
    return cols


def _rebuild_reservations_if_legacy(conn: sqlite3.Connection) -> None:
    cols = _table_cols(conn, "reservations")
    if not cols:
        return

    # Legacy şema: year/month/channel_name varsa (veya year NOT NULL gibi eski yapı)
    is_legacy = any(k in cols for k in ("year", "month", "channel_name"))
    if not is_legacy:
        return

    conn.execute("BEGIN")
    try:
        conn.execute("ALTER TABLE reservations RENAME TO reservations_legacy")

        # Yeni doğru tablo
        conn.execute("""
        CREATE TABLE reservations (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          reservation_no TEXT UNIQUE,
          advertiser_name TEXT NOT NULL,
          created_at TEXT NOT NULL,
          is_confirmed INTEGER NOT NULL DEFAULT 0,
          payload_json TEXT NOT NULL
        );
        """)

        # Eski tablodan mümkün olanları taşı
        legacy_cols = _table_cols(conn, "reservations_legacy")

        res_no_expr = "reservation_no" if "reservation_no" in legacy_cols else "NULL"
        adv_expr = "advertiser_name" if "advertiser_name" in legacy_cols else "''"
        payload_expr = "payload_json" if "payload_json" in legacy_cols else ("payload" if "payload" in legacy_cols else "'{}'")

        conn.execute(f"""
            INSERT INTO reservations(reservation_no, advertiser_name, created_at, is_confirmed, payload_json)
            SELECT
              {res_no_expr},
              {adv_expr},
              datetime('now'),
              CASE WHEN {res_no_expr} IS NULL OR {res_no_expr} = '' THEN 0 ELSE 1 END,
              {payload_expr}
            FROM reservations_legacy
        """)

        conn.execute("DROP TABLE reservations_legacy")
        conn.commit()
    except Exception:
        conn.rollback()
        raise


def migrate_and_seed(conn: sqlite3.Connection) -> None:
    conn.executescript(SCHEMA_SQL)

    _rebuild_reservations_if_legacy(conn)  # <-- BUNU EKLE


    # ---- MIGRATION PATCH (eski DB'ler için) ----
    _ensure_column(
        conn,
        "reservations",
        "ALTER TABLE reservations ADD COLUMN is_confirmed INTEGER NOT NULL DEFAULT 0",
        "is_confirmed",
    )
    # payload_json yoksa eklemek de mantıklı (eskide farklı şema olabilir)
    _ensure_column(
        conn,
        "reservations",
        "ALTER TABLE reservations ADD COLUMN payload_json TEXT NOT NULL DEFAULT '{}'",
        "payload_json",
    )

    # default channels seed (MVP)
    conn.executemany(
        "INSERT OR IGNORE INTO channels(name,is_active) VALUES(?,1)",
        [("TRT FM",), ("RADIOSCOPE",), ("POWER FM",)],
    )

    # ---- SEED ----
    cur = conn.execute("SELECT value FROM meta WHERE key = ?", ("reservation_seq",))
    row = cur.fetchone()
    if row is None:
        conn.execute("INSERT INTO meta(key, value) VALUES(?, ?)", ("reservation_seq", "1000"))

    conn.commit()

