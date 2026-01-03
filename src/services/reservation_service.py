from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, date, time
import calendar
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


    # ------------------------------
    # PLAN ÖZET (ay bazlı birleştirilmiş özet)
    # ------------------------------

    _MONTHS_TR = [
        "OCAK",
        "ŞUBAT",
        "MART",
        "NİSAN",
        "MAYIS",
        "HAZİRAN",
        "TEMMUZ",
        "AĞUSTOS",
        "EYLÜL",
        "EKİM",
        "KASIM",
        "ARALIK",
    ]

    @staticmethod
    def _norm_name(s: str) -> str:
        return " ".join((s or "").strip().lower().split())

    @staticmethod
    def _sort_reservation_no(no: str) -> tuple:
        """Rezervasyon no'yu mümkünse (Yıl, Hafta, Seq) şeklinde sıralar."""
        s = (no or "").strip()
        m = __import__("re").match(r"^[A-Z]-([0-9]{4})W([0-9]{2})-([0-9]+)$", s)
        if m:
            return (int(m.group(1)), int(m.group(2)), int(m.group(3)), s)
        # farklı formatlar için lexicographic fallback
        return (9999, 99, 999999999, s)

    def get_plan_ozet_data(self, advertiser_name: str, year: int, month: int) -> dict[str, Any]:
        """Seçili reklamveren + yıl/ay için PLAN ÖZET datasını üretir.

        - Bu özet tek rezervasyon değil; ay içinde aynı reklamverene ait tüm rezervasyonları birleştirir.
        - Eğer ay içinde 1'den fazla rezervasyon varsa Rezervasyon No alanı "ÇOKLU" olur ve numaralar
          sıralı şekilde listelenir.
        """
        adv = (advertiser_name or "").strip()
        if not adv:
            return {
                "header": {},
                "year": int(year),
                "month": int(month),
                "days": 0,
                "rows": [],
                "totals": {},
            }

        yy = int(year)
        mm = int(month)
        days_in_month = int(calendar.monthrange(yy, mm)[1])
        month_name = self._MONTHS_TR[mm - 1]

        # rezervasyonları ay bazlı filtrele
        recs = self.repo.list_confirmed_reservations_by_advertiser(adv, limit=5000)
        month_recs = []
        for r in recs:
            p = r.payload or {}
            dstr = p.get("plan_date")
            try:
                d = datetime.fromisoformat(str(dstr)).date()
            except Exception:
                continue
            if d.year == yy and d.month == mm:
                month_recs.append(r)

        # Header alanları: tekse aynen, çoklaysa "ÇOKLU"
        def _uniq_or_coklu(values: list[str]) -> str:
            vals = [v for v in (v.strip() for v in values) if v]
            if not vals:
                return ""
            u = sorted(set(vals), key=lambda x: x.lower())
            return u[0] if len(u) == 1 else "ÇOKLU"

        agency = _uniq_or_coklu([str((r.payload or {}).get("agency_name") or "") for r in month_recs])
        product = _uniq_or_coklu([str((r.payload or {}).get("product_name") or "") for r in month_recs])
        plan_title = _uniq_or_coklu([str((r.payload or {}).get("plan_title") or "") for r in month_recs])

        res_nos = [str(getattr(r, "reservation_no", "") or "") for r in month_recs if getattr(r, "reservation_no", None)]
        res_nos_sorted = sorted(res_nos, key=self._sort_reservation_no)
        if len(res_nos_sorted) == 0:
            reservation_no_display = ""
        elif len(res_nos_sorted) == 1:
            reservation_no_display = res_nos_sorted[0]
        else:
            # "ÇOKLU" + sıralı rezervasyon noları
            reservation_no_display = "ÇOKLU\n" + "\n".join(res_nos_sorted)

        # Spot süresi: kod tanımı sayfasındaki ort. uzunluk
        spot_len = float(self.get_kod_tanimi_avg_len(adv) or 0.0)

        # Dinlenme oranı (AvRch%) - erişim örneğinden
        access_set_id = self.repo.get_latest_access_set_id_for_year(yy) or self.repo.get_latest_access_set_id()
        access_map: dict[str, str] = {}
        if access_set_id is not None:
            rows = self.repo.get_access_rows(access_set_id)
            for rr in rows:
                ch = self._norm_name(str(rr.get("channel") or ""))
                if not ch:
                    continue
                v = rr.get("avrch_pct")
                access_map[ch] = "NA" if v is None else str(v)

        # Birim sn. fiyatları: fiyat ve kanal tanımı tablosundan (yıl/ay)
        # repo.get_channel_prices(year) -> {(channel_id, month): (dt, odt)}
        price_map = self.repo.get_channel_prices(yy)
        channels = self.repo.list_channels(active_only=False)
        # is_active=0 ama ay içinde rezervasyonda geçen kanalı yine de listele
        used_channels = set(self._norm_name(str((r.payload or {}).get("channel_name") or "")) for r in month_recs)
        used_channels.discard("")

        # name->(id,is_active)
        ch_by_norm: dict[str, dict[str, object]] = {}
        for ch in channels:
            ch_by_norm[self._norm_name(str(ch["name"]))] = ch

        # liste: aktif kanallar + (aktif değil ama kullanılmış) kanallar
        display_channels = []
        for ch in channels:
            if int(ch.get("is_active", 1)) == 1:
                display_channels.append(ch)
        for nm in sorted(used_channels):
            if nm in ch_by_norm and ch_by_norm[nm] not in display_channels:
                display_channels.append(ch_by_norm[nm])

        # sayaç: (norm_channel, dt_odt, day) -> adet
        counts: dict[tuple[str, str, int], int] = {}
        for r in month_recs:
            p = r.payload or {}
            channel_name = self._norm_name(str(p.get("channel_name") or ""))
            if not channel_name:
                continue
            cells: dict[str, str] = p.get("plan_cells") or {}
            for k, v in cells.items():
                if not str(v or "").strip():
                    continue
                try:
                    row_idx_s, day_s = str(k).split(",")
                    row_idx = int(row_idx_s)
                    day = int(day_s)
                except Exception:
                    continue
                if day < 1 or day > days_in_month:
                    continue
                t0 = self._row_idx_to_time(row_idx)
                dt_odt = classify_dt_odt(t0)
                key = (channel_name, dt_odt, day)
                counts[key] = int(counts.get(key, 0)) + 1

        rows_out: list[dict[str, Any]] = []

        def _price_for(ch_id: int, dtodt: str) -> float | None:
            dtp, odtp = price_map.get((int(ch_id), mm), (0.0, 0.0))
            p = float(dtp) if dtodt == "DT" else float(odtp)
            return None if p <= 0 else p

        for ch in sorted(display_channels, key=lambda x: str(x["name"]).lower()):
            ch_name = str(ch["name"])  # gösterimde orijinal case
            ch_norm = self._norm_name(ch_name)
            ch_id = int(ch["id"])

            dinlenme = access_map.get(ch_norm, "NA")
            for dtodt in ("DT", "ODT"):
                day_vals = [int(counts.get((ch_norm, dtodt, d), 0)) for d in range(1, days_in_month + 1)]
                month_adet = int(sum(day_vals))

                # Excel görünümü: 0'ları boş göster
                day_vals_display = ["" if v == 0 else v for v in day_vals]

                month_saniye = float(month_adet) * float(spot_len) if (month_adet and spot_len) else 0.0
                unit_price = _price_for(ch_id, dtodt)
                total_budget = (float(month_saniye) * float(unit_price)) if (unit_price and month_saniye) else 0.0

                rows_out.append(
                    {
                        "channel": ch_name,
                        "publish_group": "",  # kullanıcı elle girecek
                        "dt_odt": dtodt,
                        "dinlenme_orani": dinlenme,
                        "days": day_vals_display,
                        "month_adet": "" if month_adet == 0 else month_adet,
                        "month_saniye": "" if month_saniye == 0 else month_saniye,
                        "unit_price": "" if not unit_price else unit_price,
                        "budget": "" if total_budget == 0 else total_budget,
                    }
                )

        # Totals
        total_day = [0 for _ in range(days_in_month)]
        total_month_adet = 0
        total_month_saniye = 0.0
        total_budget = 0.0
        for rr in rows_out:
            # rr["days"] listesinde boş-string var
            for i, v in enumerate(rr["days"]):
                total_day[i] += int(v) if v not in ("", None) else 0
            if rr["month_adet"] not in ("", None):
                total_month_adet += int(rr["month_adet"])
            if rr["month_saniye"] not in ("", None):
                total_month_saniye += float(rr["month_saniye"])
            if rr["budget"] not in ("", None):
                total_budget += float(rr["budget"])

        totals = {
            "days": ["" if v == 0 else v for v in total_day],
            "month_adet": "" if total_month_adet == 0 else total_month_adet,
            "month_saniye": "" if total_month_saniye == 0 else total_month_saniye,
            "budget": "" if total_budget == 0 else total_budget,
        }

        header = {
            "agency": agency,
            "advertiser": adv,
            "product": product,
            "plan_title": plan_title,
            "reservation_no": reservation_no_display,
            "period": f"{month_name} {yy}",
            "month_name": month_name,
            "spot_len": spot_len,
        }

        return {
            "header": header,
            "year": yy,
            "month": mm,
            "days": days_in_month,
            "rows": rows_out,
            "totals": totals,
        }

    def export_plan_ozet_excel(self, out_path, advertiser_name: str, year: int, month: int) -> None:
        """Plan Özet ekranındaki birleştirilmiş özetin Excel çıktısını üretir."""
        from src.export.excel_exporter import export_plan_ozet
        data = self.get_plan_ozet_data(advertiser_name, int(year), int(month))
        export_plan_ozet(out_path, data)
