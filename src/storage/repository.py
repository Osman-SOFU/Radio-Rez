from __future__ import annotations

import json
import sqlite3
from dataclasses import dataclass
from datetime import datetime
from typing import Any
import re

@dataclass
class ReservationRecord:
    id: int
    reservation_no: str | None
    advertiser_name: str
    plan_title: str
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


    # ------------------------------
    # PLAN BAŞLIĞI bazlı aramalar
    # ------------------------------
    def search_plan_titles(self, text: str, limit: int = 30) -> list[str]:
        sql = """
            SELECT DISTINCT plan_title
            FROM reservations
            WHERE is_confirmed = 1
            AND plan_title IS NOT NULL
            AND plan_title != ''
            AND UPPER(plan_title) LIKE UPPER(?)
            ORDER BY plan_title
            LIMIT ?
        """
        cur = self.conn.execute(sql, (f"%{text}%", limit))
        return [r[0] for r in cur.fetchall()]

    def list_confirmed_reservations_by_plan_title(self, plan_title: str, limit: int = 5000) -> list[ReservationRecord]:
        pt = (plan_title or "").strip()
        cur = self.conn.execute(
            """
            SELECT * FROM reservations
            WHERE plan_title = ? AND is_confirmed = 1
            ORDER BY datetime(created_at) DESC
            LIMIT ?
            """,
            (pt, limit),
        )
        out: list[ReservationRecord] = []
        for r in cur.fetchall():
            out.append(
                ReservationRecord(
                    id=r["id"],
                    reservation_no=r["reservation_no"],
                    advertiser_name=r["advertiser_name"],
                    plan_title=(r["plan_title"] if "plan_title" in r.keys() else (json.loads(r["payload_json"] or "{}").get("plan_title") or "")),
                    created_at=r["created_at"],
                    is_confirmed=r["is_confirmed"],
                    payload=json.loads(r["payload_json"]),
                )
            )
        return out

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
                    plan_title=(r["plan_title"] if "plan_title" in r.keys() else (json.loads(r["payload_json"] or "{}").get("plan_title") or "")),
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
                INSERT INTO reservations(reservation_no, advertiser_name, plan_title, created_at, is_confirmed, payload_json)
                VALUES(?, ?, ?, ?, ?, ?)
                """,
                (reservation_no, advertiser_name, str(payload.get("plan_title") or "").strip(), now, 1 if confirmed else 0, json.dumps(payload, ensure_ascii=False)),
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
            plan_title=str(payload.get("plan_title") or "").strip(),
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
                    plan_title=(r["plan_title"] if "plan_title" in r.keys() else (json.loads(r["payload_json"] or "{}").get("plan_title") or "")),
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


    def delete_reservations_by_plan_title_and_spot_code(self, plan_title: str, spot_code: str) -> int:
        spot_code = (spot_code or "").strip()
        if not spot_code:
            return 0

        recs = self.list_confirmed_reservations_by_plan_title(plan_title, limit=50000)
        ids = [
            r.id for r in recs
            if (r.payload.get("spot_code") or "").strip() == spot_code
        ]
        self.delete_reservations_by_ids(ids)
        return len(ids)

    def update_reservation_payload(self, reservation_id: int, payload: dict) -> None:
        """Tek bir reservation kaydının payload_json alanını günceller."""
        self.conn.execute(
            "UPDATE reservations SET payload_json=? WHERE id=?",
            (json.dumps(payload, ensure_ascii=False), int(reservation_id)),
        )
        self.conn.commit()

    def remove_code_from_plan_title(self, plan_title: str, spot_code: str) -> int:
        """Belirli plan başlığındaki tüm rezervasyonlardan spot_code tanımını ve hücrelerini siler.

        - code_defs içinden kaldırır
        - plan_cells içinde ilgili kodu geçen hücreleri boşaltır
        - display alanlarını (spot_code, spot_duration_sec, code_definition) yeniden hesaplamaz
          (bu alanlar UI tarafında zaten 'ÇOKLU' kullanılabiliyor)
        """
        spot_code = (spot_code or "").strip().upper()
        if not spot_code:
            return 0

        recs = self.list_confirmed_reservations_by_plan_title(plan_title, limit=50000)
        changed = 0
        for r in recs:
            p = dict(r.payload or {})
            cells = dict(p.get("plan_cells") or {})
            had_any = False

            # hücreleri temizle
            for k, v in list(cells.items()):
                vv = str(v or "").strip().upper()
                if vv == spot_code:
                    cells[k] = ""
                    had_any = True

            # code_defs filtrele
            code_defs = p.get("code_defs")
            if isinstance(code_defs, list):
                new_defs = [d for d in code_defs if str(d.get("code") or "").strip().upper() != spot_code]
                if len(new_defs) != len(code_defs):
                    p["code_defs"] = new_defs
                    had_any = True

            if had_any:
                p["plan_cells"] = cells
                self.update_reservation_payload(r.id, p)
                changed += 1

        return changed

    # ------------------------------
    # Erişim Örneği (Saatlik)
    # ------------------------------
    @staticmethod
    def _norm_hour_label(label: str) -> str:
        """Saat etiketini tek formata indirger.

        - Sondaki parantezli sayımı atar: '08:00-09:00(1)' -> '08:00-09:00'
        - Tire etrafındaki boşlukları temizler: '08:00 - 09:00' -> '08:00-09:00'
        - Saatleri iki haneli yapar: '8:00-9:00' -> '08:00-09:00'
        """
        s = (label or "").strip()
        # sonda '(...)' varsa at
        s = re.sub(r"\([^\)]*\)\s*$", "", s).strip()
        # unicode tireleri standart '-' yap
        s = s.replace("–", "-").replace("—", "-")
        # tire etrafındaki boşlukları temizle
        s = re.sub(r"\s*-\s*", "-", s)

        m = re.match(r"^(\d{1,2}):(\d{2})-(\d{1,2}):(\d{2})$", s)
        if m:
            h1, m1, h2, m2 = (int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4)))
            return f"{h1:02d}:{m1:02d}-{h2:02d}:{m2:02d}"

        return s

    def get_or_create_access_set(
        self,
        year: int,
        label: str,
        periods: str,
        targets: str,
        hours: list[str] | None = None,
    ) -> int:
        year = int(year)
        label = (label or "").strip()
        if not label:
            label = f"{year}"

        cur = self.conn.execute(
            "SELECT id, hours_json FROM access_example_sets WHERE year=? AND label=?",
            (year, label),
        )
        row = cur.fetchone()
        if row:
            set_id = int(row["id"])
            # hours_json boşsa ve yeni saat listesi geldiyse güncelle
            if hours is not None:
                try:
                    existing = json.loads(row["hours_json"] or "[]")
                except Exception:
                    existing = []
                if not existing:
                    self.conn.execute(
                        "UPDATE access_example_sets SET hours_json=? WHERE id=?",
                        (json.dumps(hours, ensure_ascii=False), set_id),
                    )
            return set_id

        self.conn.execute(
            "INSERT INTO access_example_sets(year,label,periods,targets,hours_json) VALUES(?,?,?,?,?)",
            (year, label, periods or "", targets or "", json.dumps(hours or [], ensure_ascii=False)),
        )
        return int(self.conn.execute("SELECT last_insert_rowid() AS id").fetchone()["id"])

    def load_access_set(self, set_id: int) -> tuple[dict, list[dict]]:
        meta = self.conn.execute(
            "SELECT * FROM access_example_sets WHERE id=?",
            (int(set_id),),
        ).fetchone()
        if not meta:
            raise ValueError("Erişim seti bulunamadı.")

        rows = self.conn.execute(
            """
            SELECT * FROM access_example_rows
            WHERE set_id=?
            ORDER BY sort_order ASC, id ASC
            """,
            (int(set_id),),
        ).fetchall()

        meta_dict = dict(meta)
        # hours_json -> list
        hours: list[str] = []
        try:
            hours = json.loads(meta_dict.get("hours_json") or "[]") or []
        except Exception:
            hours = []

        out_rows: list[dict] = []
        for r in rows:
            d = dict(r)
            # values_json -> dict
            vals = {}
            try:
                vals = json.loads(d.get("values_json") or "{}") or {}
            except Exception:
                vals = {}

            d["values"] = vals
            out_rows.append(d)

            # Eğer meta'da hours yoksa, ilk satırdan türet
            if not hours and vals:
                hours = list(vals.keys())

        meta_dict["hours"] = hours
        return meta_dict, out_rows

    def get_access_rows(self, set_id: int) -> list[dict]:
        """Erişim örneği satırlarını döndürür.
        Yeni modelde her satır: {channel: str, values: {hour_label: number}}.
        """
        _meta, rows = self.load_access_set(int(set_id))
        out: list[dict] = []
        for r in rows:
            out.append(
                {
                    "channel": r.get("channel"),
                    "values": r.get("values") or {},
                }
            )
        return out

    def save_access_set(
        self,
        set_id: int,
        periods: str,
        targets: str,
        hours: list[str],
        rows: list[dict],
    ) -> None:
        # Eğer zaten açık transaction varsa tekrar BEGIN deme (SQLite patlıyor)
        if not self.conn.in_transaction:
            self.conn.execute("BEGIN")
        try:
            self.conn.execute(
                "UPDATE access_example_sets SET periods=?, targets=?, hours_json=?, created_at=CURRENT_TIMESTAMP WHERE id=?",
                (periods or "", targets or "", json.dumps(hours or [], ensure_ascii=False), int(set_id)),
            )
            self.conn.execute("DELETE FROM access_example_rows WHERE set_id=?", (int(set_id),))

            for i, r in enumerate(rows):
                ch = (r.get("channel") or "").strip()
                if not ch:
                    continue
                vals = r.get("values") or {}
                self.conn.execute(
                    """
                    INSERT INTO access_example_rows(set_id, channel, values_json, sort_order)
                    VALUES(?,?,?,?)
                    """,
                    (int(set_id), ch, json.dumps(vals, ensure_ascii=False), i),
                )

            self.conn.commit()
        except Exception:
            self.conn.rollback()
            raise

    def get_access_channel_avg_map(self, set_id: int) -> dict[str, float]:
        """normalize(channel)->avg (saatlik değerlerin ortalaması)"""
        _meta, rows = self.load_access_set(int(set_id))
        out: dict[str, float] = {}
        for r in rows:
            ch = str(r.get("channel") or "").strip()
            if not ch:
                continue
            vals = r.get("values") or {}
            nums = []
            for v in vals.values():
                if v is None:
                    continue
                try:
                    nums.append(float(str(v).replace(",", ".")))
                except Exception:
                    continue
            if nums:
                out[ch.upper()] = sum(nums) / len(nums)
        return out

    def get_access_channel_hour_map(self, set_id: int, channel_name: str) -> dict[str, float]:
        """Verilen kanal için normalize(hour)->value döndürür."""
        _meta, rows = self.load_access_set(int(set_id))
        ch_norm = (channel_name or "").strip().upper()
        for r in rows:
            ch = (str(r.get("channel") or "")).strip().upper()
            if ch == ch_norm:
                vals = r.get("values") or {}
                out: dict[str, float] = {}
                for k, v in vals.items():
                    kk = self._norm_hour_label(str(k))
                    try:
                        out[kk] = float(str(v).replace(",", "."))
                    except Exception:
                        # boş/bozuk hücre
                        continue
                return out
        return {}




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
