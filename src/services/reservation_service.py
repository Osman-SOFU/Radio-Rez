from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from src.domain.models import ReservationDraft, ConfirmedReservation
from src.domain.time_rules import classify_dt_odt, validate_day

from src.storage.repository import Repository



@dataclass
class ReservationService:
    repo: Repository
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
