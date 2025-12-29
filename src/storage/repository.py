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
        self.conn.commit()

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
