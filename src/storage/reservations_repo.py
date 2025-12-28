from __future__ import annotations
import json
from dataclasses import dataclass
from datetime import datetime
import sqlite3

from storage.db import next_reservation_seq


@dataclass
class ReservationDraft:
    advertiser_name: str
    year: int
    month: int
    channel_name: str
    payload: dict


def make_reservation_no(advertiser_name: str, year: int, week: int, seq: int) -> str:
    first = advertiser_name.strip()[:1].upper() if advertiser_name.strip() else "X"
    return f"{first}-{year}W{week:02d}-{seq}"


def save_confirmed(conn: sqlite3.Connection, draft: ReservationDraft) -> str:
    # yılın kaçıncı haftası (basit: bugünün haftası)
    week = datetime.now().isocalendar().week

    # Tek transaction
    conn.execute("BEGIN")
    try:
        seq = next_reservation_seq(conn)
        res_no = make_reservation_no(draft.advertiser_name, draft.year, week, seq)

        conn.execute(
            """
            INSERT INTO reservations (advertiser_name, year, month, channel_name, status, reservation_no, payload_json)
            VALUES (?, ?, ?, ?, 'confirmed', ?, ?)
            """,
            (
                draft.advertiser_name.strip(),
                draft.year,
                draft.month,
                draft.channel_name.strip(),
                res_no,
                json.dumps(draft.payload, ensure_ascii=False),
            ),
        )
        conn.commit()
        return res_no
    except Exception:
        conn.rollback()
        raise
