from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, date, time
from pathlib import Path
from typing import Any

from src.domain.models import ReservationDraft, ConfirmedReservation
from src.domain.time_rules import classify_dt_odt, validate_day

from src.storage.repository import Repository



@dataclass
class ReservationService:
    repo: Repository

    def _row_idx_to_time(self, row_idx: int) -> time:
        """Plan grid satırı -> kuşak başlangıç saati.

        Şablon: 07:00-20:00, 15dk adım.
        """
        mins = 7 * 60 + int(row_idx) * 15
        return time(mins // 60, mins % 60)
    def sanitize_plan_cells(self, plan_cells: dict) -> dict[str, str]:
        fixed: dict[str, str] = {}
        for k, v in (plan_cells or {}).items():
            # tuple key’ler varsa “r,c” formatına çevir (exporter için stabil)
            if isinstance(k, tuple) and len(k) == 2:
                kk = f"{k[0]},{k[1]}"
            else:
                kk = str(k)
            fixed[kk] = "" if v is None else str(v)
        return fixed

    def confirm(self, draft: ReservationDraft, plan_cells: dict) -> ConfirmedReservation:
        adv = draft.advertiser_name.strip()
        if not adv:
            raise ValueError("Reklamveren zorunlu.")

        ok, msg = validate_day(draft.plan_date)
        if not ok:
            raise ValueError(msg)

        dt_odt = classify_dt_odt(draft.spot_time)

        # prepared_by: "İsim - dd.mm.yyyy hh:mm"
        stamp = datetime.now().strftime("%d.%m.%Y %H:%M")
        prepared_by = f"{draft.prepared_by_name.strip()} - {stamp}" if draft.prepared_by_name.strip() else stamp

        cells = self.sanitize_plan_cells(plan_cells)
        adet_total = sum(1 for v in cells.values() if str(v).strip())

        payload: dict[str, Any] = {
            "agency_name": draft.agency_name.strip(),
            "advertiser_name": adv,
            "product_name": draft.product_name.strip(),
            "plan_title": draft.plan_title.strip(),
            "spot_code": draft.spot_code.strip(),
            "spot_duration_sec": int(draft.spot_duration_sec or 0),
            "code_definition": draft.code_definition.strip(),
            "note_text": draft.note_text.strip(),
            "prepared_by": prepared_by,

            "plan_date": draft.plan_date.isoformat(),
            "spot_time": draft.spot_time.strftime("%H:%M"),
            "dt_odt": dt_odt,

            # Kanal + fiyat (plan tarihinin AY/YIL'ına göre)
            "channel_name": (draft.channel_name or "").strip(),
            "channel_price_year": int(draft.plan_date.year),
            "channel_price_month": int(draft.plan_date.month),
            "channel_price_dt": float(draft.channel_price_dt or 0.0),
            "channel_price_odt": float(draft.channel_price_odt or 0.0),

            "plan_cells": cells,
            "adet_total": adet_total,
            
        }
        return ConfirmedReservation(payload=payload)

    def export_test(self, template_path: Path, out_dir: Path, confirmed: ConfirmedReservation) -> Path:
        from src.export.excel_exporter import export_excel

        out_dir.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = out_dir / f"TEST_{ts}.xlsx"

        payload = confirmed.to_payload()
        payload["reservation_no"] = ""
        payload["created_at"] = datetime.now().isoformat(timespec="seconds")

        export_excel(template_path, out_path, payload)
        return out_path

    def save_and_export(self, template_path: Path, out_dir: Path, confirmed: ConfirmedReservation) -> Path:
        from src.export.excel_exporter import export_excel

        payload = confirmed.to_payload()

        # repo create_reservation senin mevcut fonksiyon imzanla uyumlu olmalı:
        # (advertiser_name, reservation_no, payload, confirmed=True/False)
        rec = self.repo.create_reservation(
            advertiser_name=payload["advertiser_name"],
            payload=payload,
            confirmed=True,
        )

        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / f"{rec.reservation_no}.xlsx"

        payload2 = dict(rec.payload)
        payload2["reservation_no"] = rec.reservation_no
        payload2["created_at"] = rec.created_at

        export_excel(template_path, out_path, payload2)
        return out_path

    def get_kod_tanimi_rows(self, advertiser_name: str) -> list[dict]:
        recs = self.repo.list_confirmed_reservations_by_advertiser(advertiser_name)

        grouped: dict[str, dict] = {}
        total = 0

        for r in recs:
            code = (r.payload.get("spot_code") or "").strip()
            if not code:
                continue

            total += 1
            if code not in grouped:
                grouped[code] = {
                    "code": code,
                    "code_desc": (r.payload.get("code_definition") or "").strip(),
                    "length_sn": int(r.payload.get("spot_duration_sec") or 0),
                    "count": 0,
                }

            grouped[code]["count"] += 1

            # aynı koda yeni tanım/süre girildiyse "son kayıt kazansın"
            cd = (r.payload.get("code_definition") or "").strip()
            if cd:
                grouped[code]["code_desc"] = cd

            if r.payload.get("spot_duration_sec") is not None:
                grouped[code]["length_sn"] = int(r.payload.get("spot_duration_sec") or 0)

        rows = list(grouped.values())
        rows.sort(key=lambda x: x["code"])

        for row in rows:
            row["distribution"] = (row["count"] / total) if total > 0 else 0.0

        return rows

    def get_kod_tanimi_avg_len(self, advertiser_name: str) -> float:
        rows = self.get_kod_tanimi_rows(advertiser_name)
        return sum(r["length_sn"] * r["distribution"] for r in rows) if rows else 0.0

    def delete_kod_for_advertiser(self, advertiser_name: str, code: str) -> int:
        return self.repo.delete_reservations_by_advertiser_and_spot_code(advertiser_name, code)

    def export_kod_tanimi_excel(self, out_path, advertiser_name: str) -> None:
        from src.export.excel_exporter import export_kod_tanimi
        rows = self.get_kod_tanimi_rows(advertiser_name)
        export_kod_tanimi(out_path, advertiser_name, rows)

    # ------------------------------
    # SPOTLİST+
    # ------------------------------

    def get_spotlist_rows(self, advertiser_name: str) -> list[dict[str, Any]]:
        """Seçili reklam veren için SPOTLİST+ satırlarını üretir.

        Her satır, confirmed reservation payload'ındaki plan_cells içindeki dolu hücrelerden türetilir.
        """
        adv = (advertiser_name or "").strip()
        if not adv:
            return []

        recs = self.repo.list_confirmed_reservations_by_advertiser(adv, limit=50000)
        if not recs:
            return []

        status_map = self.repo.get_spotlist_status_map([r.id for r in recs])

        rows: list[dict[str, Any]] = []
        for r in recs:
            p = r.payload or {}

            # plan ay/yıl
            try:
                y, m, _d = str(p.get("plan_date") or "").split("-")
                yy = int(y)
                mm = int(m)
            except Exception:
                continue

            channel_name = str(p.get("channel_name") or "").strip()
            spot_code = str(p.get("spot_code") or "").strip()
            duration = int(p.get("spot_duration_sec") or 0)
            price_dt = float(p.get("channel_price_dt") or 0.0)
            price_odt = float(p.get("channel_price_odt") or 0.0)

            cells: dict[str, str] = p.get("plan_cells") or {}
            for k, v in cells.items():
                if not str(v or "").strip():
                    continue
                try:
                    if isinstance(k, str):
                        row_idx_s, day_s = k.split(",")
                        row_idx = int(row_idx_s)
                        day = int(day_s)
                    else:
                        row_idx, day = k
                        row_idx = int(row_idx)
                        day = int(day)
                except Exception:
                    continue

                # gerçek tarih + saat
                try:
                    dt = date(yy, mm, day)
                except Exception:
                    continue

                t0 = self._row_idx_to_time(row_idx)
                dt_odt = classify_dt_odt(t0)
                unit = price_dt if dt_odt == "DT" else price_odt
                budget = float(unit) * float(duration)

                pub = int(status_map.get((r.id, day, row_idx), 0))
                rows.append(
                    {
                        "reservation_id": r.id,
                        "day": day,
                        "row_idx": row_idx,
                        "datetime": datetime(dt.year, dt.month, dt.day, t0.hour, t0.minute),
                        "tarih": dt.strftime("%d.%m.%Y"),
                        "ana_yayin": channel_name,
                        "reklam_firmasi": adv,
                        "adet": 1,
                        "baslangic": t0.strftime("%H:%M"),
                        "sure": duration,
                        "spot_kodu": spot_code,
                        "dt_odt": dt_odt,
                        "birim_saniye": unit,
                        "butce_net": budget,
                        "published": pub,
                    }
                )

        rows.sort(key=lambda x: x["datetime"])
        for i, rr in enumerate(rows, start=1):
            rr["sira"] = i
        return rows

    def set_spotlist_published(self, reservation_id: int, day: int, row_idx: int, published: int) -> None:
        self.repo.upsert_spotlist_published(reservation_id, day, row_idx, published)


    def set_spotlist_published_bulk(self, changes: list[tuple[int, int, int, int]]) -> None:
        """Toplu kaydet: (reservation_id, day, row_idx, published)"""
        self.repo.upsert_spotlist_published_many(changes)

    def export_spotlist_excel_with_rows(self, out_path, advertiser_name: str, rows: list[dict]) -> None:
        """Filtrelenmiş satırlarla SPOTLİST+ Excel çıktısı al."""
        from src.export.excel_exporter import export_spotlist
        export_spotlist(out_path, advertiser_name, rows)

    def export_spotlist_excel(self, out_path, advertiser_name: str) -> None:
        """Tüm satırları çekip SPOTLİST+ Excel çıktısı al."""
        rows = self.get_spotlist_rows(advertiser_name)
        self.export_spotlist_excel_with_rows(out_path, advertiser_name, rows)
