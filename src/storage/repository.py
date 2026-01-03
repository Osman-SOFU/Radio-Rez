from __future__ import annotations

import json
import sqlite3
from dataclasses import dataclass
from datetime import datetime
from typing import Any

@dataclass
class ReservationRecord:
    id: int
    reservation_no: str | None
    advertiser_name: str
    created_at: str
    is_confirmed: int
    payload: dict[str, Any]

class Repository:
    def __init__(self, conn: sqlite3.Connection) -> None:
        self.conn = conn

    def upsert_advertiser(self, name: str) -> None:
        name = name.strip()
        if not name:
            return
        self.conn.execute(
            "INSERT OR IGNORE INTO advertisers(name) VALUES(?)",
            (name,),
        )

    def search_advertisers(self, text: str, limit: int = 30) -> list[str]:
        sql = """
            SELECT DISTINCT advertiser_name
            FROM reservations
            WHERE is_confirmed = 1
            AND advertiser_name IS NOT NULL
            AND advertiser_name != ''
            AND UPPER(advertiser_name) LIKE UPPER(?)
            ORDER BY advertiser_name
            LIMIT ?
        """
        cur = self.conn.execute(sql, (f"%{text}%", limit))
        return [r[0] for r in cur.fetchall()]

    def list_reservations_by_advertiser(self, advertiser_name: str, limit: int = 50) -> list[ReservationRecord]:
        cur = self.conn.execute(
            """
            SELECT * FROM reservations
            WHERE advertiser_name = ?
            ORDER BY datetime(created_at) DESC
            LIMIT ?
            """,
            (advertiser_name, limit),
        )
        out: list[ReservationRecord] = []
        for r in cur.fetchall():
            out.append(
                ReservationRecord(
                    id=r["id"],
                    reservation_no=r["reservation_no"],
                    advertiser_name=r["advertiser_name"],
                    created_at=r["created_at"],
                    is_confirmed=r["is_confirmed"],
                    payload=json.loads(r["payload_json"]),
                )
            )
        return out

    def next_reservation_no(self, advertiser_name: str, when: datetime) -> str:
        """
        Sayaç 1000'den başlar. Sadece confirmed kayıtta artar.
        Format: {ILK_HARF}-{YYYY}W{WEEK}-{SEQ}
        """
        first = advertiser_name.strip()[:1].upper() or "X"
        year = when.isocalendar().year
        week = when.isocalendar().week

        # Transaction içinde seq çek + arttır
        cur = self.conn.execute("SELECT value FROM meta WHERE key=?", ("reservation_seq",))
        row = cur.fetchone()
        if row is None:
            seq = 1000
            self.conn.execute("INSERT INTO meta(key,value) VALUES(?,?)", ("reservation_seq", str(seq)))
        else:
            seq = int(row["value"])

        reservation_no = f"{first}-{year}W{week:02d}-{seq}"

        # arttır
        self.conn.execute(
            "UPDATE meta SET value=? WHERE key=?",
            (str(seq + 1), "reservation_seq"),
        )
        return reservation_no

    def create_reservation(self, advertiser_name: str, payload: dict, confirmed: bool) -> ReservationRecord:
        now = datetime.now().isoformat(timespec="seconds")

        reservation_no = None
        self.conn.execute("BEGIN")
        try:
            if confirmed:
                reservation_no = self.next_reservation_no(advertiser_name, datetime.now())

            self.conn.execute(
                """
                INSERT INTO reservations(reservation_no, advertiser_name, created_at, is_confirmed, payload_json)
                VALUES(?, ?, ?, ?, ?)
                """,
                (reservation_no, advertiser_name, now, 1 if confirmed else 0, json.dumps(payload, ensure_ascii=False)),
            )

            self.upsert_advertiser(advertiser_name)

            rid = self.conn.execute("SELECT last_insert_rowid() AS id").fetchone()["id"]
            self.conn.commit()
        except Exception:
            self.conn.rollback()
            raise

        return ReservationRecord(
            id=rid,
            reservation_no=reservation_no,
            advertiser_name=advertiser_name,
            created_at=now,
            is_confirmed=1 if confirmed else 0,
            payload=payload,
        )

    def list_confirmed_reservations_by_advertiser(self, advertiser_name: str, limit: int = 5000):
        cur = self.conn.execute(
            """
            SELECT * FROM reservations
            WHERE advertiser_name = ? AND is_confirmed = 1
            ORDER BY datetime(created_at) DESC
            LIMIT ?
            """,
            (advertiser_name, limit),
        )
        out = []
        for r in cur.fetchall():
            out.append(
                ReservationRecord(
                    id=r["id"],
                    reservation_no=r["reservation_no"],
                    advertiser_name=r["advertiser_name"],
                    created_at=r["created_at"],
                    is_confirmed=r["is_confirmed"],
                    payload=json.loads(r["payload_json"]),
                )
            )
        return out

    def delete_reservations_by_ids(self, ids: list[int]) -> None:
        if not ids:
            return
        self.conn.execute("BEGIN")
        try:
            # SQLite 999 param sınırı için chunk
            for i in range(0, len(ids), 900):
                chunk = ids[i:i+900]
                q = ",".join(["?"] * len(chunk))
                self.conn.execute(f"DELETE FROM reservations WHERE id IN ({q})", chunk)
            self.conn.commit()
        except Exception:
            self.conn.rollback()
            raise

    def delete_reservations_by_advertiser_and_spot_code(self, advertiser_name: str, spot_code: str) -> int:
        spot_code = (spot_code or "").strip()
        if not spot_code:
            return 0

        recs = self.list_confirmed_reservations_by_advertiser(advertiser_name, limit=50000)
        ids = [
            r.id for r in recs
            if (r.payload.get("spot_code") or "").strip() == spot_code
        ]
        self.delete_reservations_by_ids(ids)
        return len(ids)
    def get_or_create_access_set(self, year: int, label: str, periods: str = "", targets: str = "") -> int:
        label = (label or "").strip()
        if not label:
            label = f"{year}"

        cur = self.conn.execute(
            "SELECT id FROM access_example_sets WHERE year=? AND label=?",
            (year, label),
        )
        row = cur.fetchone()
        if row:
            return int(row["id"])

        self.conn.execute(
            "INSERT INTO access_example_sets(year,label,periods,targets) VALUES(?,?,?,?)",
            (year, label, periods or "", targets or ""),
        )
        return int(self.conn.execute("SELECT last_insert_rowid() AS id").fetchone()["id"])

    def load_access_set(self, set_id: int) -> tuple[dict, list[dict]]:
        meta = self.conn.execute(
            "SELECT * FROM access_example_sets WHERE id=?",
            (set_id,),
        ).fetchone()
        if not meta:
            raise ValueError("Erişim seti bulunamadı.")

        rows = self.conn.execute(
            """
            SELECT * FROM access_example_rows
            WHERE set_id=?
            ORDER BY sort_order ASC, id ASC
            """,
            (set_id,),
        ).fetchall()

        return dict(meta), [dict(r) for r in rows]

    def get_access_rows(self, set_id: int) -> list[dict]:
        """Erişim örneği (Access Example) satırlarını döndürür.

        PLAN ÖZET sayfası dinlenme oranı (AvRch%) bilgisini buradan okur.
        """
        _meta, rows = self.load_access_set(int(set_id))
        return rows


    def save_access_set(self, set_id: int, periods: str, targets: str, rows: list[dict]) -> None:
        # Eğer zaten açık transaction varsa tekrar BEGIN deme (SQLite patlıyor)
        if not self.conn.in_transaction:
            self.conn.execute("BEGIN")
        try:
            self.conn.execute(
                "UPDATE access_example_sets SET periods=?, targets=?, created_at=CURRENT_TIMESTAMP WHERE id=?",
                (periods or "", targets or "", set_id),
            )
            self.conn.execute("DELETE FROM access_example_rows WHERE set_id=?", (set_id,))

            for i, r in enumerate(rows):
                self.conn.execute(
                    """
                    INSERT INTO access_example_rows(set_id, channel, universe, avrch000, avrch_pct, sort_order)
                    VALUES(?,?,?,?,?,?)
                    """,
                    (
                        set_id,
                        (r.get("channel") or "").strip(),
                        r.get("universe"),
                        r.get("avrch000"),
                        r.get("avrch_pct"),
                        i,
                    ),
                )

            self.conn.commit()
        except Exception:
            self.conn.rollback()
            raise


    def get_latest_access_set_id(self) -> int | None:
        row = self.conn.execute(
            """
            SELECT id
            FROM access_example_sets
            ORDER BY datetime(created_at) DESC, id DESC
            LIMIT 1
            """
        ).fetchone()
        return int(row["id"]) if row else None

    def get_latest_access_set_id_for_year(self, year: int) -> int | None:
        """İstenen yıl için en son kaydedilmiş erişim setinin id'si."""
        row = self.conn.execute(
            """
            SELECT id
            FROM access_example_sets
            WHERE year=?
            ORDER BY datetime(created_at) DESC, id DESC
            LIMIT 1
            """,
            (int(year),),
        ).fetchone()
        return int(row["id"]) if row else None

    # ------------------------------
    # Kanal / Fiyat Tanımı (DT-ODT)
    # ------------------------------

    def get_meta(self, key: str) -> str | None:
        row = self.conn.execute("SELECT value FROM meta WHERE key=?", (key,)).fetchone()
        return str(row["value"]) if row else None

    def set_meta(self, key: str, value: str) -> None:
        self.conn.execute(
            "INSERT INTO meta(key, value) VALUES(?, ?) "
            "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (key, value),
        )
        self.conn.commit()

    def list_price_years(self) -> list[int]:
        rows = self.conn.execute(
            "SELECT DISTINCT year FROM channel_prices WHERE year > 0 ORDER BY year DESC"
        ).fetchall()
        return [int(r["year"]) for r in rows]

    def list_channels(self, active_only: bool = True) -> list[dict[str, object]]:
        sql = "SELECT id, name, is_active FROM channels"
        if active_only:
            sql += " WHERE is_active=1"
        sql += " ORDER BY name COLLATE NOCASE"
        rows = self.conn.execute(sql).fetchall()
        return [{"id": int(r["id"]), "name": str(r["name"]), "is_active": int(r["is_active"])} for r in rows]

    def get_or_create_channel(self, name: str) -> int:
        nm = (name or "").strip()
        if not nm:
            raise ValueError("Kanal adı boş olamaz.")

        row = self.conn.execute("SELECT id, is_active FROM channels WHERE name=?", (nm,)).fetchone()
        if row:
            if int(row["is_active"]) == 0:
                self.conn.execute("UPDATE channels SET is_active=1 WHERE id=?", (int(row["id"]),))
                self.conn.commit()
            return int(row["id"])

        self.conn.execute("INSERT INTO channels(name, is_active) VALUES(?, 1)", (nm,))
        cid = int(self.conn.execute("SELECT last_insert_rowid() AS id").fetchone()["id"])
        self.conn.commit()
        return cid

    def update_channel_name(self, channel_id: int, new_name: str) -> None:
        nm = (new_name or "").strip()
        if not nm:
            raise ValueError("Kanal adı boş olamaz.")
        self.conn.execute("UPDATE channels SET name=? WHERE id=?", (nm, int(channel_id)))
        self.conn.commit()

    def deactivate_channel(self, channel_id: int) -> None:
        self.conn.execute("UPDATE channels SET is_active=0 WHERE id=?", (int(channel_id),))
        self.conn.commit()

    def get_channel_prices(self, year: int) -> dict[tuple[int, int], tuple[float, float]]:
        rows = self.conn.execute(
            "SELECT channel_id, month, price_dt, price_odt FROM channel_prices WHERE year=?",
            (int(year),),
        ).fetchall()
        out: dict[tuple[int, int], tuple[float, float]] = {}
        for r in rows:
            out[(int(r["channel_id"]), int(r["month"]))] = (float(r["price_dt"]), float(r["price_odt"]))
        return out

    def upsert_channel_price(self, year: int, month: int, channel_id: int, price_dt: float, price_odt: float) -> None:
        self.conn.execute(
            "INSERT INTO channel_prices(year, month, channel_id, price_dt, price_odt) "
            "VALUES(?,?,?,?,?) "
            "ON CONFLICT(year, month, channel_id) DO UPDATE SET "
            "price_dt=excluded.price_dt, "
            "price_odt=excluded.price_odt",
            (int(year), int(month), int(channel_id), float(price_dt), float(price_odt)),
        )
        self.conn.commit()


    # ------------------------------
    # SPOTLİST+ (yayınlandı durumu)
    # ------------------------------

    def get_spotlist_status_map(self, reservation_ids: list[int]) -> dict[tuple[int, int, int], int]:
        """reservation_id listesi için (reservation_id, day, row_idx) -> published map."""
        if not reservation_ids:
            return {}

        out: dict[tuple[int, int, int], int] = {}
        # SQLite 999 param sınırına göre chunk
        for i in range(0, len(reservation_ids), 900):
            chunk = reservation_ids[i:i+900]
            q = ",".join(["?"] * len(chunk))
            rows = self.conn.execute(
                f"SELECT reservation_id, day, row_idx, published FROM spotlist_status WHERE reservation_id IN ({q})",
                chunk,
            ).fetchall()
            for r in rows:
                out[(int(r["reservation_id"]), int(r["day"]), int(r["row_idx"]))] = int(r["published"])
        return out

    def upsert_spotlist_published(self, reservation_id: int, day: int, row_idx: int, published: int) -> None:
        self.conn.execute(
            "INSERT INTO spotlist_status(reservation_id, day, row_idx, published, updated_at) "
            "VALUES(?,?,?,?, datetime('now')) "
            "ON CONFLICT(reservation_id, day, row_idx) DO UPDATE SET "
            "published=excluded.published, "
            "updated_at=datetime('now')",
            (int(reservation_id), int(day), int(row_idx), 1 if int(published) else 0),
        )
        self.conn.commit()


    def upsert_spotlist_published_many(self, changes: list[tuple[int, int, int, int]]) -> None:
        """Toplu upsert: (reservation_id, day, row_idx, published) listesi."""
        if not changes:
            return

        # Eğer zaten açık transaction varsa tekrar BEGIN deme (SQLite patlıyor)
        if not self.conn.in_transaction:
            self.conn.execute("BEGIN")
        try:
            self.conn.executemany(
                "INSERT INTO spotlist_status(reservation_id, day, row_idx, published, updated_at) "
                "VALUES(?,?,?,?, datetime('now')) "
                "ON CONFLICT(reservation_id, day, row_idx) DO UPDATE SET "
                "published=excluded.published, "
                "updated_at=datetime('now')",
                [(int(rid), int(day), int(row_idx), 1 if int(pub) else 0) for rid, day, row_idx, pub in changes],
            )
            self.conn.commit()
        except Exception:
            self.conn.rollback()
            raise
