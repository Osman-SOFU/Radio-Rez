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
  year INTEGER NOT NULL,
  month INTEGER NOT NULL,        -- 1..12
  channel_id INTEGER NOT NULL,
  price_dt REAL NOT NULL,
  price_odt REAL NOT NULL,
  FOREIGN KEY(channel_id) REFERENCES channels(id)
);

-- NOT: idx_channel_prices_unique index'i migrate_and_seed içinde yaratıyoruz.
-- Çünkü eski DB'lerde channel_prices tablosu year/month kolonları olmadan gelmiş olabilir.
-- SCHEMA_SQL içinde index'i oluşturmak, eski DB'lerde "no such column: year" hatasına sebep olur.

CREATE TABLE IF NOT EXISTS access_example_sets (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  year INTEGER NOT NULL,
  label TEXT NOT NULL,        -- Örn: "Ekim 2025"
  periods TEXT,               -- Örn: "Periods (Regions|...)"
  targets TEXT,               -- Örn: "12+(1)"
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  UNIQUE(year, label)
);

CREATE TABLE IF NOT EXISTS access_example_rows (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  set_id INTEGER NOT NULL,
  channel TEXT NOT NULL,
  universe INTEGER,
  avrch000 INTEGER,
  avrch_pct REAL,
  sort_order INTEGER NOT NULL,
  FOREIGN KEY(set_id) REFERENCES access_example_sets(id) ON DELETE CASCADE
);

-- SPOTLİST+ için yayınlandı durumu (0/1)
-- Her bir yayın satırı, reservations.payload_json içindeki plan_cells'in (day,row_idx)
-- kombinasyonundan türetilir. Bu tablo sadece kullanıcı işaretini saklar.
CREATE TABLE IF NOT EXISTS spotlist_status (
  reservation_id INTEGER NOT NULL,
  day INTEGER NOT NULL,
  row_idx INTEGER NOT NULL,
  published INTEGER NOT NULL DEFAULT 0,
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  PRIMARY KEY(reservation_id, day, row_idx),
  FOREIGN KEY(reservation_id) REFERENCES reservations(id) ON DELETE CASCADE
);

"""

def ensure_data_folders(data_dir: Path) -> None:
    data_dir.mkdir(parents=True, exist_ok=True)
    (data_dir / "exports").mkdir(parents=True, exist_ok=True)

def connect_db(db_path: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys=ON")
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

    # Kanal fiyatları için (eski DB'ler): year/month kolonları yoksa ekle
    _ensure_column(
        conn,
        "channel_prices",
        "ALTER TABLE channel_prices ADD COLUMN year INTEGER NOT NULL DEFAULT 0",
        "year",
    )
    _ensure_column(
        conn,
        "channel_prices",
        "ALTER TABLE channel_prices ADD COLUMN month INTEGER NOT NULL DEFAULT 0",
        "month",
    )

    # Unique index (yıl/ay/kanal)
    conn.execute(
        "CREATE UNIQUE INDEX IF NOT EXISTS idx_channel_prices_unique ON channel_prices(year, month, channel_id)"
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

def access_save_db(self) -> None:
    if not self.repo:
        QMessageBox.critical(self, "Hata", "Repo yok. DB bağlantısı kurulmamış.")
        return

    dates = (self.access_dates.text() or "").strip()
    targets = (self.access_targets.text() or "").strip()

    year = self._parse_year_from_dates(dates)
    label = dates if dates else f"{year}"

    # aynı Dates>> için tekrar kaydedersen UPDATE gibi davranacak (satırları silip yeniden yazar)
    set_id = self.repo.get_or_create_access_set(year=year, label=label, periods="", targets=targets)
    self._access_set_id = set_id

    def _to_int(txt: str):
        t = (txt or "").strip()
        return int(t) if t else None

    def _to_float(txt: str):
        t = (txt or "").strip()
        if not t:
            return None
        return float(t.replace(",", "."))

    rows_out: list[dict] = []
    for r in range(self.access_table.rowCount()):
        ch_item = self.access_table.item(r, 0)
        ch = ch_item.text().strip() if ch_item else ""
        if not ch:
            continue

        u = _to_int(self.access_table.item(r, 1).text() if self.access_table.item(r, 1) else "")
        a0 = _to_int(self.access_table.item(r, 2).text() if self.access_table.item(r, 2) else "")
        ap = _to_float(self.access_table.item(r, 3).text() if self.access_table.item(r, 3) else "")

        # AvRch% boşsa hesapla (opsiyonel ama hayat kurtarır)
        if ap is None and u and a0 is not None and u != 0:
            ap = round((a0 / u) * 100, 2)

        rows_out.append({"channel": ch, "universe": u, "avrch000": a0, "avrch_pct": ap})

    self.repo.save_access_set(set_id=set_id, periods="", targets=targets, rows=rows_out)

    QMessageBox.information(self, "OK", "Erişim örneği DB'ye kaydedildi. Uygulama yeniden açılınca aynen gelecektir.")
