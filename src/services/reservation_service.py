from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, date, time
import calendar
import re
from pathlib import Path
from typing import Any

from src.domain.models import ReservationDraft, ConfirmedReservation
from src.domain.time_rules import classify_dt_odt, validate_day

from src.storage.repository import Repository



@dataclass
class ReservationService:
    repo: Repository

    _PREPARED_BY_STAMP_RE = re.compile(r"\s-\s\d{2}\.\d{2}\.\d{4}\s\d{2}:\d{2}.*$")

    def _clean_prepared_by_name(self, raw: str) -> str:
        """UI/DB'den gelebilen 'İSİM - dd.mm.yyyy hh:mm - ...' değerinden sadece ismi al.

        Amaç: Kullanıcı formu tekrar onayladığında ismin yanına tekrar tekrar tarih eklenmesini engellemek.
        """
        s = (raw or "").strip()
        if not s:
            return ""
        return self._PREPARED_BY_STAMP_RE.sub("", s).strip()

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

    def confirm(self, draft: ReservationDraft, plan_cells: dict | None = None) -> ConfirmedReservation:
        if plan_cells is None:
            plan_cells = {}  # UI normalde plan_cells gönderir; yoksa boş kabul et
        adv = (draft.advertiser_name or "").strip()
        pt = (draft.plan_title or "").strip()

        if not pt:
            raise ValueError("Plan başlığı zorunlu.")
        if not adv:
            raise ValueError("Reklamveren zorunlu.")

        ok, msg = validate_day(draft.plan_date)
        if not ok:
            raise ValueError(msg)

        dt_odt = classify_dt_odt(draft.spot_time)

        # prepared_by: "İsim - dd.mm.yyyy hh:mm"
        stamp = datetime.now().strftime("%d.%m.%Y %H:%M")
        pb_name = self._clean_prepared_by_name(draft.prepared_by_name)
        prepared_by = f"{pb_name} - {stamp}" if pb_name else stamp

        cells = self.sanitize_plan_cells(plan_cells)
        adet_total = sum(1 for v in cells.values() if str(v).strip())

        # Çoklu kod desteği
        raw_code_defs = draft.code_defs or []
        code_defs: list[dict[str, Any]] = []
        seen_codes: set[str] = set()
        for rr in raw_code_defs:
            try:
                code = str(rr.get("code") or "").strip().upper()
                if not code:
                    continue
                if code in seen_codes:
                    continue
                seen_codes.add(code)
                desc = str(rr.get("desc") or "").strip()
                dur = int(rr.get("duration_sec") or 0)
                code_defs.append({"code": code, "desc": desc, "duration_sec": dur})
            except Exception:
                continue

        # Geri uyumluluk: tekli alanlardan code_defs üret
        if not code_defs and (draft.spot_code or draft.code_definition or draft.spot_duration_sec):
            code = str(draft.spot_code or "").strip().upper()
            if code:
                code_defs.append({
                    "code": code,
                    "desc": str(draft.code_definition or "").strip(),
                    "duration_sec": int(draft.spot_duration_sec or 0),
                })

        # Hücrelerde kullanılan kodlar
        used_codes = []
        for v in cells.values():
            vv = str(v or "").strip()
            if not vv:
                continue
            used_codes.append(vv.upper())
        used_codes_set = sorted(set(used_codes))

        # Çoklu kodlarda süreyi hesaplamak için adet dağılımı
        code_counts: dict[str, int] = {}
        for c in used_codes:
            code_counts[c] = code_counts.get(c, 0) + 1

        def _lookup_def(c: str) -> dict[str, Any] | None:
            cc = (c or "").strip().upper()
            for d in code_defs:
                if str(d.get("code") or "").strip().upper() == cc:
                    return d
            return None

        # Header alanları için tekli gösterim (Excel export vs.)
        if len(used_codes_set) == 1:
            one = used_codes_set[0]
            dd = _lookup_def(one) or {}
            disp_code = one
            disp_desc = str(dd.get("desc") or "")
            disp_dur = int(dd.get("duration_sec") or 0)
        elif len(used_codes_set) > 1:
            # Çoklu kod: süreyi ağırlıklı ortalama olarak hesapla.
            # (Rezervasyon listesinde ve özet ekranlarında anlamlı bir değer
            # görmek için.)
            disp_code = "ÇOKLU"
            disp_desc = ""
            total = sum(code_counts.get(c, 0) for c in used_codes_set)
            if total > 0:
                wsum = 0.0
                for c in used_codes_set:
                    dd = _lookup_def(c) or {}
                    dur = float(dd.get("duration_sec") or 0)
                    wsum += float(code_counts.get(c, 0)) * dur
                disp_dur = round(wsum / float(total), 2)
            else:
                disp_dur = 0
        else:
            # hiç seçim yoksa, draft'taki değerleri göster
            disp_code = str(draft.spot_code or "").strip()
            disp_desc = str(draft.code_definition or "").strip()
            disp_dur = int(draft.spot_duration_sec or 0)

        payload: dict[str, Any] = {
            "agency_name": draft.agency_name.strip(),
            "advertiser_name": adv,
            "product_name": draft.product_name.strip(),
            "plan_title": pt,
            "spot_code": str(disp_code or "").strip(),
            "spot_duration_sec": int(disp_dur or 0),
            "code_definition": str(disp_desc or "").strip(),
            "code_defs": code_defs,
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

    def get_kod_tanimi_rows(self, plan_title: str) -> list[dict]:
        recs = self.repo.list_confirmed_reservations_by_plan_title(plan_title)

        grouped: dict[str, dict] = {}
        total_spots = 0

        for r in recs:
            p = r.payload or {}
            cells: dict[str, str] = p.get("plan_cells") or {}

            # kod haritası
            code_defs = p.get("code_defs") or []
            code_map = {
                str(d.get("code") or "").strip().upper(): {
                    "desc": str(d.get("desc") or "").strip(),
                    "dur": int(d.get("duration_sec") or 0),
                }
                for d in code_defs
                if str(d.get("code") or "").strip()
            }
            # geri uyumluluk
            if not code_map:
                sc = str(p.get("spot_code") or "").strip().upper()
                if sc:
                    code_map[sc] = {
                        "desc": str(p.get("code_definition") or "").strip(),
                        "dur": int(p.get("spot_duration_sec") or 0),
                    }

            for _k, v in cells.items():
                code = str(v or "").strip().upper()
                if not code:
                    continue
                total_spots += 1

                if code not in grouped:
                    dd = code_map.get(code) or {}
                    grouped[code] = {
                        "code": code,
                        "code_desc": str(dd.get("desc") or ""),
                        "length_sn": int(dd.get("dur") or 0),
                        "count": 0,
                    }

                grouped[code]["count"] += 1

                # güncelleyici (son görülen açıklama/süre kazansın)
                dd = code_map.get(code) or {}
                if dd.get("desc"):
                    grouped[code]["code_desc"] = str(dd.get("desc") or "")
                if dd.get("dur") is not None:
                    grouped[code]["length_sn"] = int(dd.get("dur") or 0)

        rows = list(grouped.values())
        rows.sort(key=lambda x: x["code"])

        for row in rows:
            row["distribution"] = (row["count"] / total_spots) if total_spots > 0 else 0.0

        return rows

    def get_kod_tanimi_avg_len(self, plan_title: str) -> float:
        rows = self.get_kod_tanimi_rows(plan_title)
        return sum(r["length_sn"] * r["distribution"] for r in rows) if rows else 0.0

    def delete_kod_for_plan_title(self, plan_title: str, code: str) -> int:
        # Kod silme: ilgili plan başlığındaki tüm rezervasyon payload'larından kodu kaldırır,
        # kodu kullanan hücreleri boşaltır.
        return self.repo.remove_code_from_plan_title(plan_title, code)

    def export_kod_tanimi_excel(self, out_path, plan_title: str) -> None:
        from src.export.excel_exporter import export_kod_tanimi
        rows = self.get_kod_tanimi_rows(plan_title)
        export_kod_tanimi(out_path, plan_title, rows)

    # ------------------------------
    # SPOTLİST+
    # ------------------------------

    def get_spotlist_rows(self, plan_title: str) -> list[dict[str, Any]]:
        """Seçili plan başlığı için SPOTLİST+ satırlarını üretir.

        Her satır, confirmed reservation payload'ındaki plan_cells içindeki dolu hücrelerden türetilir.
        """
        pt = (plan_title or "").strip()
        if not pt:
            return []

        recs = self.repo.list_confirmed_reservations_by_plan_title(pt, limit=50000)
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

            # kod haritası
            code_defs = p.get("code_defs") or []
            code_map = {
                str(d.get("code") or "").strip().upper(): int(d.get("duration_sec") or 0)
                for d in code_defs
                if str(d.get("code") or "").strip()
            }
            # geri uyumluluk
            if not code_map:
                sc = str(p.get("spot_code") or "").strip().upper()
                if sc:
                    code_map[sc] = int(p.get("spot_duration_sec") or 0)
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

                cell_code = str(v or "").strip().upper()
                duration = int(code_map.get(cell_code, 0))
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
                        "reklam_firmasi": str(p.get("advertiser_name") or "").strip(),
                        "adet": 1,
                        "baslangic": t0.strftime("%H:%M"),
                        "sure": duration,
                        "spot_kodu": cell_code,
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

    def export_spotlist_excel_with_rows(self, out_path, plan_title: str, rows: list[dict]) -> None:
        """Filtrelenmiş satırlarla SPOTLİST+ Excel çıktısı al."""
        from src.export.excel_exporter import export_spotlist
        export_spotlist(out_path, plan_title, rows)

    def export_spotlist_excel(self, out_path, plan_title: str) -> None:
        """Tüm satırları çekip SPOTLİST+ Excel çıktısı al."""
        rows = self.get_spotlist_rows(plan_title)
        self.export_spotlist_excel_with_rows(out_path, plan_title, rows)


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

    def get_plan_ozet_data(self, plan_title: str, year: int, month: int) -> dict[str, Any]:
        """Seçili plan başlığı + yıl/ay için PLAN ÖZET datasını üretir.

        - Bu özet tek rezervasyon değil; ay içinde aynı plan başlığına ait tüm rezervasyonları birleştirir.
        - Eğer ay içinde 1'den fazla rezervasyon varsa Rezervasyon No alanı "ÇOKLU" olur ve numaralar
          sıralı şekilde listelenir.
        """
        pt = (plan_title or "").strip()
        if not pt:
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
        recs = self.repo.list_confirmed_reservations_by_plan_title(pt, limit=5000)
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
        advertiser = _uniq_or_coklu([str((r.payload or {}).get("advertiser_name") or "") for r in month_recs])


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
        spot_len = float(self.get_kod_tanimi_avg_len(pt) or 0.0)

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

        # sayaçlar: (norm_channel, dt_odt, day) -> adet / saniye / bütçe
        counts: dict[tuple[str, str, int], int] = {}
        seconds: dict[tuple[str, str, int], float] = {}
        budgets: dict[tuple[str, str, int], float] = {}
        for r in month_recs:
            p = r.payload or {}
            channel_name = self._norm_name(str(p.get("channel_name") or ""))
            if not channel_name:
                continue

            # kanal id (fiyat için)
            ch_obj = ch_by_norm.get(channel_name)
            ch_id_for_price = int(ch_obj["id"]) if ch_obj and ch_obj.get("id") is not None else None

            # kod haritası (cell içeriğine göre süre)
            code_defs = p.get("code_defs") or []
            code_map = {
                str(d.get("code") or "").strip().upper(): float(d.get("duration_sec") or 0)
                for d in code_defs
                if str(d.get("code") or "").strip()
            }
            if not code_map:
                sc = str(p.get("spot_code") or "").strip().upper()
                if sc:
                    code_map[sc] = float(p.get("spot_duration_sec") or 0)

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

                cell_code = str(v or "").strip().upper()
                dur = float(code_map.get(cell_code, 0.0))
                seconds[key] = float(seconds.get(key, 0.0)) + dur

                # bütçe
                if ch_id_for_price is not None:
                    dtp, odtp = price_map.get((int(ch_id_for_price), mm), (0.0, 0.0))
                    unit_price = float(dtp) if dt_odt == "DT" else float(odtp)
                    budgets[key] = float(budgets.get(key, 0.0)) + (dur * unit_price)

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
                day_secs = [float(seconds.get((ch_norm, dtodt, d), 0.0)) for d in range(1, days_in_month + 1)]
                day_bud = [float(budgets.get((ch_norm, dtodt, d), 0.0)) for d in range(1, days_in_month + 1)]

                month_adet = int(sum(day_vals))
                month_saniye = float(sum(day_secs))
                month_budget = float(sum(day_bud))

                # Excel görünümü: 0'ları boş göster
                day_vals_display = ["" if v == 0 else v for v in day_vals]

                unit_price = _price_for(ch_id, dtodt)
                # unit_price boşsa bütçeyi 0 göster (fiyat tanımı yok)
                total_budget = month_budget if unit_price else 0.0

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
            "advertiser": advertiser,
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

    def export_plan_ozet_excel(self, out_path, plan_title: str, year: int, month: int) -> None:
        """Plan Özet ekranındaki birleştirilmiş özetin Excel çıktısını üretir."""
        from src.export.excel_exporter import export_plan_ozet
        data = self.get_plan_ozet_data(plan_title, int(year), int(month))
        export_plan_ozet(out_path, data)

    def get_plan_ozet_yearly_data(self, plan_title: str, year: int) -> dict[str, Any]:
        """Plan başlığı için seçilen yılın tamamını (Ocak-Aralık) tek bir özet olarak döndürür.

        Not: Aylık ekrandaki gün bazlı kolonlar, yıllık görünümde anlamlı olmadığı için
        days=0 olacak şekilde sadece toplamlar birleştirilir.
        """
        yy = int(year)

        def _merge_value(acc: set, val: str) -> None:
            if val is None:
                return
            s = str(val).strip()
            if not s:
                return
            # "ÇOKLU" gibi placeholder'ları değer setine dahil etme
            if s.upper() == "ÇOKLU":
                return
            acc.add(s)

        agency_set: set[str] = set()
        adv_set: set[str] = set()
        product_set: set[str] = set()
        spot_len_set: set[str] = set()
        resno_set: set[str] = set()

        rows_map: dict[tuple[str, str], dict[str, Any]] = {}

        for mm in range(1, 13):
            data_m = self.get_plan_ozet_data(plan_title, yy, mm)
            hdr = data_m.get("header", {}) or {}
            _merge_value(agency_set, hdr.get("agency", ""))
            _merge_value(adv_set, hdr.get("advertiser", ""))
            _merge_value(product_set, hdr.get("product", ""))
            _merge_value(spot_len_set, str(hdr.get("spot_len", "") or "").strip())

            # Rezervasyon no çoklu görünümü satır satır geliyor; normalize ederek topla
            rn = str(hdr.get("reservation_no", "") or "").strip()
            if rn:
                for line in rn.splitlines():
                    line = line.strip()
                    if not line or line.upper() == "ÇOKLU":
                        continue
                    resno_set.add(line)

            for rr in data_m.get("rows", []) or []:
                key = (str(rr.get("channel", "")), str(rr.get("dt_odt", "")))
                if key not in rows_map:
                    rows_map[key] = {
                        "channel": key[0],
                        "publish_group": "",
                        "dt_odt": key[1],
                        "dinlenme_orani": rr.get("dinlenme_orani", ""),
                        "month_adet": 0,
                        "month_saniye": 0.0,
                        "unit_price_values": set(),
                        "budget": 0.0,
                    }
                agg = rows_map[key]

                # dinlenme oranı ilk dolu değeri al; farklılık varsa "ÇOKLU" göster
                cur_do = str(agg.get("dinlenme_orani", "") or "").strip()
                new_do = str(rr.get("dinlenme_orani", "") or "").strip()
                if not cur_do and new_do:
                    agg["dinlenme_orani"] = new_do
                elif cur_do and new_do and cur_do != new_do:
                    agg["dinlenme_orani"] = "ÇOKLU"

                def _to_int(v):
                    try:
                        return int(v)
                    except Exception:
                        return 0

                def _to_float(v):
                    try:
                        return float(v)
                    except Exception:
                        return 0.0

                if rr.get("month_adet") not in ("", None):
                    agg["month_adet"] += _to_int(rr.get("month_adet"))
                if rr.get("month_saniye") not in ("", None):
                    agg["month_saniye"] += _to_float(rr.get("month_saniye"))
                if rr.get("budget") not in ("", None):
                    agg["budget"] += _to_float(rr.get("budget"))

                up = rr.get("unit_price")
                if up not in ("", None):
                    agg["unit_price_values"].add(str(up))

        # Normalize header values
        def _uniq_or_coklu(values: set[str]) -> str:
            if not values:
                return ""
            if len(values) == 1:
                return next(iter(values))
            return "ÇOKLU"

        reservation_no_display = ""
        if resno_set:
            # Çokluysa başa "ÇOKLU" yaz, altına numaraları koy
            res_lines = sorted(resno_set)
            if len(res_lines) == 1:
                reservation_no_display = res_lines[0]
            else:
                reservation_no_display = "ÇOKLU\n" + "\n".join(res_lines)

        header = {
            "agency": _uniq_or_coklu(agency_set),
            "advertiser": _uniq_or_coklu(adv_set),
            "product": _uniq_or_coklu(product_set),
            "plan_title": plan_title,
            "reservation_no": reservation_no_display,
            "period": f"{yy} (YILLIK)",
            "month_name": "YIL",
            "spot_len": _uniq_or_coklu(spot_len_set),
        }

        # finalize rows
        rows_out: list[dict[str, Any]] = []
        for key, agg in sorted(rows_map.items(), key=lambda kv: (kv[0][0], kv[0][1])):
            upv = agg.pop("unit_price_values", set())
            unit_price = "" if not upv else (next(iter(upv)) if len(upv) == 1 else "ÇOKLU")
            month_adet = agg.get("month_adet", 0)
            month_saniye = agg.get("month_saniye", 0.0)
            budget = agg.get("budget", 0.0)

            rows_out.append(
                {
                    "channel": agg.get("channel", ""),
                    "publish_group": "",
                    "dt_odt": agg.get("dt_odt", ""),
                    "dinlenme_orani": agg.get("dinlenme_orani", ""),
                    "days": [],
                    "month_adet": "" if month_adet == 0 else month_adet,
                    "month_saniye": "" if month_saniye == 0 else month_saniye,
                    "unit_price": unit_price,
                    "budget": "" if budget == 0 else budget,
                }
            )

        # totals
        total_month_adet = 0
        total_month_saniye = 0.0
        total_budget = 0.0
        for rr in rows_out:
            if rr.get("month_adet") not in ("", None):
                total_month_adet += int(rr["month_adet"])
            if rr.get("month_saniye") not in ("", None):
                total_month_saniye += float(rr["month_saniye"])
            if rr.get("budget") not in ("", None):
                total_budget += float(rr["budget"])

        totals = {
            "days": [],
            "month_adet": "" if total_month_adet == 0 else total_month_adet,
            "month_saniye": "" if total_month_saniye == 0 else total_month_saniye,
            "budget": "" if total_budget == 0 else total_budget,
        }

        return {
            "header": header,
            "year": yy,
            "month": 0,
            "days": 0,
            "rows": rows_out,
            "totals": totals,
        }

    def export_plan_ozet_yearly_excel(self, out_path, plan_title: str, year: int) -> None:
        """Plan Özet (YILLIK) çıktısı üretir."""
        from src.export.excel_exporter import export_plan_ozet_yearly
        data = self.get_plan_ozet_yearly_data(plan_title, int(year))
        export_plan_ozet_yearly(out_path, data)