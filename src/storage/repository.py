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

    def search_advertisers(self, q: str, limit: int = 20) -> list[str]:
        q = q.strip()
        if not q:
            return []
        cur = self.conn.execute(
            "SELECT name FROM advertisers WHERE name LIKE ? ORDER BY name LIMIT ?",
            (f"%{q}%", limit),
        )
        return [r["name"] for r in cur.fetchall()]

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
