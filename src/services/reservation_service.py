from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, date, time, timedelta
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


    def _parse_iso_date(self, s: Any) -> date | None:
        try:
            if s is None:
                return None
            return datetime.fromisoformat(str(s)).date()
        except Exception:
            return None

    def _is_span_payload(self, p: dict[str, Any]) -> bool:
        return bool(p.get("is_span")) or bool(p.get("span_start")) or bool(p.get("span_end")) or bool(p.get("span_month_matrices"))

    def _get_span_month_matrices(self, p: dict[str, Any]) -> dict[tuple[int, int], dict[str, str]]:
        """DB payload'ındaki span_month_matrices'i güvenli şekilde (yy,mm)->cells formatına çevir."""
        smm = p.get("span_month_matrices") or {}
        out: dict[tuple[int, int], dict[str, str]] = {}
        if not isinstance(smm, dict):
            return out
        for k, cells in smm.items():
            try:
                if isinstance(k, (tuple, list)) and len(k) == 2:
                    yy, mm = int(k[0]), int(k[1])
                else:
                    ks = str(k).strip()
                    if not ks:
                        continue
                    if "-" in ks:
                        a, b = ks.split("-", 1)
                        yy, mm = int(a), int(b)
                    elif "," in ks:
                        a, b = ks.split(",", 1)
                        yy, mm = int(a), int(b)
                    else:
                        continue
                out[(yy, mm)] = self.sanitize_plan_cells(cells or {})
            except Exception:
                continue
        return out

    def _iter_cells(self, payload: dict[str, Any]):
        """Yield (yy,mm,row_idx,day,code,is_span) over all non-empty cells in payload."""
        p = payload or {}
        is_span = self._is_span_payload(p)

        if is_span:
            mats = self._get_span_month_matrices(p)
            for (yy, mm), cells in mats.items():
                for k, v in (cells or {}).items():
                    code = str(v or "").strip().upper()
                    if not code:
                        continue
                    try:
                        row_s, day_s = str(k).split(",", 1)
                        row_idx = int(row_s)
                        day = int(day_s)
                    except Exception:
                        continue
                    yield int(yy), int(mm), int(row_idx), int(day), code, True
        else:
            d = self._parse_iso_date(p.get("plan_date"))
            if not d:
                return
            yy, mm = int(d.year), int(d.month)
            cells = self.sanitize_plan_cells(p.get("plan_cells") or {})
            for k, v in (cells or {}).items():
                code = str(v or "").strip().upper()
                if not code:
                    continue
                try:
                    row_s, day_s = str(k).split(",", 1)
                    row_idx = int(row_s)
                    day = int(day_s)
                except Exception:
                    continue
                yield int(yy), int(mm), int(row_idx), int(day), code, False

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

            # Erişim örneği: saatlik dinlenme oranları (kanal bazlı)
            "access_hour_map": self._get_access_hour_map_for_channel(draft.channel_name, draft.plan_date.year),

            "plan_cells": cells,
            "adet_total": adet_total,

            # Excel'de Ajans Komisyonu % oranı (AR62) dinamik.
            "agency_commission_pct": int(getattr(draft, "agency_commission_pct", 10) or 0),
            
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

            for _yy, _mm, _row_idx, _day, code, _is_span in self._iter_cells(p):
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

    def get_kod_tanimi_len_display(self, plan_title: str) -> str:
        """Kod tanımı süreleri tek değer ise onu, birden fazlaysa "ÇOKLU" döndürür.

        Not: Plan özet çıktılarında ortalama süre yazmak yanıltıcı olabildiği için
        (her kodun süresi farklı olabilir) bu değer sadece bilgilendirme amaçlıdır.
        """
        rows = self.get_kod_tanimi_rows(plan_title)
        lens = sorted({int(r.get("length_sn") or 0) for r in rows if int(r.get("length_sn") or 0) > 0})
        if not lens:
            return ""
        if len(lens) == 1:
            return str(lens[0])
        return "ÇOKLU"

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

        Tek ay kayıtları + span (tarih aralığı) kayıtları desteklenir.
        Span kayıtlarında 'day' alanı çakışmayı engellemek için YYYYMMDD (int) tutulur.
        """
        pt = (plan_title or "").strip()
        if not pt:
            return []

        recs = self.repo.list_confirmed_reservations_by_plan_title(pt, limit=50000)
        if not recs:
            return []

        status_map = self.repo.get_spotlist_status_map([r.id for r in recs])

        # fiyatlar (yıl/ay bazlı) - span kayıtlarında aylara göre değişebileceği için repo'dan okunur
        ch_id_map: dict[str, int] = {}
        try:
            for ch in self.repo.list_channels(active_only=False):
                nm = str(ch.get("name") or "").strip().lower()
                if nm:
                    ch_id_map[nm] = int(ch.get("id") or 0)
        except Exception:
            pass
        price_cache: dict[int, dict[tuple[int, int], tuple[float, float]]] = {}

        rows: list[dict[str, Any]] = []
        for r in recs:
            p = r.payload or {}

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

            adv_name = str(p.get("advertiser_name") or "").strip()

            for yy, mm, row_idx, day, cell_code, is_span in self._iter_cells(p):
                # gerçek tarih + saat
                try:
                    dt = date(int(yy), int(mm), int(day))
                except Exception:
                    continue

                t0 = self._row_idx_to_time(int(row_idx))
                dt_odt = classify_dt_odt(t0)
                # span kayıtlarında fiyat ay bazlı değişebilir: repo fiyatını tercih et
                ch_id = ch_id_map.get(channel_name.strip().lower())
                cache_key = (adv_name.casefold(), int(yy))
                if ch_id and cache_key not in price_cache:
                    try:
                        price_cache[cache_key] = self.repo.get_channel_prices(int(yy), adv_name)
                    except Exception:
                        price_cache[cache_key] = {}
                dt_p, odt_p = (price_dt, price_odt)
                if ch_id:
                    dt_p, odt_p = price_cache.get(cache_key, {}).get((int(ch_id), int(mm)), (price_dt, price_odt))
                unit = float(dt_p) if dt_odt == "DT" else float(odt_p)

                duration = int(code_map.get(str(cell_code).strip().upper(), 0))
                budget = float(unit) * float(duration)

                day_key = int(dt.strftime("%Y%m%d")) if is_span else int(day)
                pub = int(status_map.get((int(r.id), int(day_key), int(row_idx)), 0) or 0)

                rows.append(
                    {
                        "reservation_id": int(r.id),
                        "day": int(day_key),
                        "row_idx": int(row_idx),
                        "datetime": datetime(dt.year, dt.month, dt.day, t0.hour, t0.minute),
                        "tarih": dt.strftime("%d.%m.%Y"),
                        "ana_yayin": channel_name,
                        "reklam_firmasi": str(p.get("advertiser_name") or "").strip(),
                        "adet": 1,
                        "baslangic": t0.strftime("%H:%M"),
                        "sure": duration,
                        "spot_kodu": str(cell_code).strip().upper(),
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


    def _get_access_hour_map_for_channel(self, channel_name: str, year: int) -> dict:
        """Seçili kanal için saatlik erişim değerlerini döndürür (normalize edilmiş saat etiketleri ile)."""
        if not self.repo:
            return {}
        try:
            set_id = self.repo.get_latest_access_set_id_for_year(int(year)) or self.repo.get_latest_access_set_id()
            if not set_id:
                return {}
            return self.repo.get_access_channel_hour_map(int(set_id), (channel_name or "").strip())
        except Exception:
            return {}

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

        # rezervasyonları ay bazlı filtrele (tek ay kayıtları + span kayıtları)
        recs = self.repo.list_confirmed_reservations_by_plan_title(pt, limit=5000)
        month_recs: list[Any] = []
        month_cells_by_id: dict[int, dict[str, str]] = {}

        month_start = date(yy, mm, 1)
        month_end = date(yy, mm, days_in_month)

        for r in recs:
            p = r.payload or {}
            rid = int(getattr(r, "id", 0) or 0)
            if rid <= 0:
                continue

            if self._is_span_payload(p):
                mats = self._get_span_month_matrices(p)
                cells_m = mats.get((yy, mm)) or {}
                # bu ay için hiç hücre yoksa bu ay özetine dahil etmeyelim
                if not any(str(v).strip() for v in cells_m.values()):
                    continue

                s0 = self._parse_iso_date(p.get("span_start")) or month_start
                s1 = self._parse_iso_date(p.get("span_end")) or month_end
                if s1 < month_start or s0 > month_end:
                    continue

                month_recs.append(r)
                month_cells_by_id[rid] = self.sanitize_plan_cells(cells_m)
            else:
                d = self._parse_iso_date(p.get("plan_date"))
                if not d:
                    continue
                if d.year == yy and d.month == mm:
                    month_recs.append(r)
                    month_cells_by_id[rid] = self.sanitize_plan_cells(p.get("plan_cells") or {})

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

        # Spot süresi: ortalama yerine tek/çoklu (bilgilendirme)
        spot_len = self.get_kod_tanimi_len_display(pt)

        # Dinlenme oranı - Erişim örneğindeki saatlik değerlerin kanal bazlı ortalaması
        access_set_id = self.repo.get_latest_access_set_id_for_year(yy) or self.repo.get_latest_access_set_id()
        access_map: dict[str, str] = {}
        if access_set_id is not None:
            try:
                rows = self.repo.get_access_rows(int(access_set_id))
                for rr in rows:
                    ch = self._norm_name(str(rr.get("channel") or ""))
                    if not ch:
                        continue
                    vals = rr.get("values") or {}
                    nums = []
                    for v in vals.values():
                        if v is None or str(v).strip() == "":
                            continue
                        try:
                            nums.append(float(str(v).replace(",", ".")))
                        except Exception:
                            continue
                    if nums:
                        # ortalama (2 hane)
                        access_map[ch] = str(round(sum(nums) / len(nums), 2))
                    else:
                        access_map[ch] = "NA"
            except Exception:
                access_map = {}

        # Birim sn. fiyatları: fiyat ve kanal tanımı tablosundan (reklam veren + yıl/ay)
        # repo.get_channel_prices(year, advertiser) -> {(channel_id, month): (dt, odt)}
        adv_name = ""
        try:
            adv_name = str(((month_recs[0].payload or {}) if month_recs else {}).get("advertiser_name") or "").strip()
        except Exception:
            adv_name = ""
        price_map = self.repo.get_channel_prices(yy, adv_name)
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

    def get_plan_ozet_range_data(self, plan_title: str, start: date, end: date) -> dict[str, Any]:
        """Plan Özet datasını seçili tarih aralığına göre üretir (tek tip).

        Kurallar:
        - Gün kolonları aralık kaç günse o kadar (start..end dahil).
        - Fiyatlar ay bazında değişebilir: bütçe, her günün ait olduğu ayın fiyatıyla hesaplanır.
        - Gün hücreleri 0 ise boş gösterilir.
        - Header.period: 'dd.mm.yyyy-dd.mm.yyyy'
        """
        pt = (plan_title or "").strip()
        if not pt or start is None or end is None:
            return {
                "header": {},
                "start": start,
                "end": end,
                "dates": [],
                "months": [],
                "month_headers": [],
                "rows": [],
                "totals": {},
            }

        rs = start if start <= end else end
        re = end if start <= end else start

        dates: list[date] = []
        d = rs
        while d <= re:
            dates.append(d)
            d = d + timedelta(days=1)

        # aralıkta geçen aylar (sıralı)
        months: list[tuple[int, int]] = []
        for dd in dates:
            ym = (dd.year, dd.month)
            if not months or months[-1] != ym:
                months.append(ym)

        # tüm plan başlığı rezervasyonları
        recs = self.repo.list_confirmed_reservations_by_plan_title(pt, limit=5000)

        # aralıkla kesişen rezervasyonlar + hücreler
        rel_recs: list[Any] = []
        cells_by_id_month: dict[int, dict[tuple[int, int], dict[str, str]]] = {}

        for r in recs:
            p = r.payload or {}
            rid = int(getattr(r, "id", 0) or 0)
            if rid <= 0:
                continue

            if self._is_span_payload(p):
                s0 = self._parse_iso_date(p.get("span_start")) or rs
                s1 = self._parse_iso_date(p.get("span_end")) or re
                if s1 < rs or s0 > re:
                    continue

                mats = self._get_span_month_matrices(p)  # {(yy,mm): cells}
                # sadece aralıkta geçen ayları al
                mmaps: dict[tuple[int, int], dict[str, str]] = {}
                has_any = False
                for (yy, mm), _cells in mats.items():
                    if (yy, mm) in months:
                        fixed = self.sanitize_plan_cells(_cells or {})
                        # tamamen boş ay hücrelerini geç
                        if any(str(v).strip() for v in fixed.values()):
                            mmaps[(yy, mm)] = fixed
                            has_any = True
                if not has_any:
                    continue

                rel_recs.append(r)
                cells_by_id_month[rid] = mmaps
            else:
                d0 = self._parse_iso_date(p.get("plan_date"))
                if not d0:
                    continue
                if d0 < rs or d0 > re:
                    continue

                rel_recs.append(r)
                # tek gün kaydında plan_cells o günün gün numarasıyla geliyor
                fixed = self.sanitize_plan_cells(p.get("plan_cells") or {})
                cells_by_id_month[rid] = {(d0.year, d0.month): fixed}

        # Header alanları: tekse aynen, çoklaysa "ÇOKLU"
        def _pick_single(values: list[str]) -> str:
            vals = [str(v or "").strip() for v in values if str(v or "").strip()]
            if not vals:
                return ""
            uniq = []
            for v in vals:
                if v not in uniq:
                    uniq.append(v)
            if len(uniq) == 1:
                return uniq[0]
            return "ÇOKLU"

        # Header bilgileri: aralık içinde hiç kayıt yoksa bile (ör. yanlış ay seçildi)
        # üst bilgi boş görünmesin diye plan başlığının tüm rezervasyonlarından derle.
        hdr_recs = rel_recs if rel_recs else recs

        agencies = []
        advertisers = []
        products = []
        resnos = []
        for r in hdr_recs:
            p = r.payload or {}
            agencies.append(str(p.get("agency_name") or p.get("agency") or ""))
            advertisers.append(str(p.get("advertiser_name") or p.get("advertiser") or ""))
            products.append(str(p.get("product_name") or p.get("product") or ""))

            # Rezervasyon no: payload'da olmayabilir, DB kolonundan al
            rn = str(getattr(r, "reservation_no", "") or p.get("reservation_no") or "").strip()
            if rn:
                resnos.append(rn)


        agency = _pick_single(agencies)
        advertiser = _pick_single(advertisers)
        product = _pick_single(products)

        # Rezervasyon no: tekse aynen, çoklaysa "ÇOKLU" + liste
        res_nos_sorted = sorted(list({r for r in resnos if r.strip()}))
        if not res_nos_sorted:
            reservation_no_display = ""
        elif len(res_nos_sorted) == 1:
            reservation_no_display = res_nos_sorted[0]
        else:
            reservation_no_display = "ÇOKLU\n" + "\n".join(res_nos_sorted)

        # Spot süresi: ortalama yerine tek/çoklu (bilgilendirme)
        spot_len = self.get_kod_tanimi_len_display(pt)

        # Dinlenme oranı (kanal bazlı ortalama) - 2026 varsa 2026 setini tercih et
        years_in_range = sorted(list({yy for (yy, _mm) in months}))
        access_set_id = None
        if 2026 in years_in_range:
            access_set_id = self.repo.get_latest_access_set_id_for_year(2026)
        if access_set_id is None and years_in_range:
            access_set_id = self.repo.get_latest_access_set_id_for_year(int(years_in_range[-1]))
        if access_set_id is None:
            access_set_id = self.repo.get_latest_access_set_id()

        access_map: dict[str, str] = {}
        if access_set_id is not None:
            try:
                rows = self.repo.get_access_rows(int(access_set_id))
                for rr in rows:
                    ch = self._norm_name(str(rr.get("channel") or ""))
                    if not ch:
                        continue
                    vals = rr.get("values") or {}
                    nums = []
                    for v in vals.values():
                        if v is None or str(v).strip() == "":
                            continue
                        try:
                            nums.append(float(str(v).replace(",", ".")))
                        except Exception:
                            continue
                    if nums:
                        access_map[ch] = str(round(sum(nums) / len(nums), 2))
                    else:
                        access_map[ch] = "NA"
            except Exception:
                access_map = {}

        # Birim sn fiyatları: reklam veren + yıl bazında çekilir, ay bazında kullanılır
        # repo.get_channel_prices(year, advertiser) -> {(channel_id, month): (dt, odt)}
        adv_name = ""
        try:
            adv_name = str(((rel_recs[0].payload or {}) if rel_recs else {}).get("advertiser_name") or "").strip()
        except Exception:
            adv_name = ""

        price_maps: dict[int, Any] = {}
        for yy in years_in_range or [rs.year]:
            try:
                price_maps[int(yy)] = self.repo.get_channel_prices(int(yy), adv_name)
            except Exception:
                price_maps[int(yy)] = {}

        channels = self.repo.list_channels(active_only=False)

        # aralıkta kullanılan kanalları yakala (aktif değilse bile)
        used_channels = set()
        for r in rel_recs:
            p = r.payload or {}
            used_channels.add(self._norm_name(str(p.get("channel_name") or "")))
        used_channels.discard("")

        ch_by_norm: dict[str, dict[str, object]] = {}
        for ch in channels:
            ch_by_norm[self._norm_name(str(ch["name"]))] = ch

        display_channels = []
        for ch in channels:
            if int(ch.get("is_active", 1)) == 1:
                display_channels.append(ch)
        for nm in sorted(used_channels):
            if nm in ch_by_norm and ch_by_norm[nm] not in display_channels:
                display_channels.append(ch_by_norm[nm])

        # sayaçlar: (norm_channel, dt_odt, date) -> adet/saniye/bütçe
        counts: dict[tuple[str, str, date], int] = {}
        seconds: dict[tuple[str, str, date], float] = {}
        budgets: dict[tuple[str, str, date], float] = {}

        # hızlı index
        date_set = set(dates)

        for r in rel_recs:
            p = r.payload or {}
            channel_norm = self._norm_name(str(p.get("channel_name") or ""))
            if not channel_norm:
                continue

            ch_obj = ch_by_norm.get(channel_norm)
            ch_id_for_price = int(ch_obj["id"]) if ch_obj and ch_obj.get("id") is not None else None

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

            rid = int(getattr(r, "id", 0) or 0)
            mmaps = cells_by_id_month.get(rid) or {}

            for (yy, mm), cells in mmaps.items():
                if not cells:
                    continue
                for k, v in (cells or {}).items():
                    if not str(v or "").strip():
                        continue
                    try:
                        row_idx_s, day_s = str(k).split(",")
                        row_idx = int(row_idx_s)
                        day = int(day_s)
                    except Exception:
                        continue

                    try:
                        dd = date(int(yy), int(mm), int(day))
                    except Exception:
                        continue

                    if dd not in date_set:
                        continue

                    t0 = self._row_idx_to_time(row_idx)
                    dt_odt = classify_dt_odt(t0)
                    key = (channel_norm, dt_odt, dd)
                    counts[key] = int(counts.get(key, 0)) + 1

                    cell_code = str(v or "").strip().upper()
                    dur = float(code_map.get(cell_code, 0.0))
                    seconds[key] = float(seconds.get(key, 0.0)) + dur

                    # bütçe: günün ayına göre fiyat uygula
                    if ch_id_for_price is not None:
                        pm = price_maps.get(int(dd.year), {}) or {}
                        dtp, odtp = pm.get((int(ch_id_for_price), int(dd.month)), (0.0, 0.0))
                        unit_price = float(dtp) if dt_odt == "DT" else float(odtp)
                        budgets[key] = float(budgets.get(key, 0.0)) + (dur * unit_price)

        # ay başlıkları
        month_headers: list[str] = []
        for (yy, mm) in months:
            mname = self._MONTHS_TR[int(mm) - 1]
            month_headers.append(f"{mname} Adet")
            month_headers.append(f"{mname} Saniye")

        rows_out: list[dict[str, Any]] = []

        def _unit_price_display(ch_id: int, dtodt: str) -> str:
            prices = []
            for (yy, mm) in months:
                pm = price_maps.get(int(yy), {}) or {}
                dtp, odtp = pm.get((int(ch_id), int(mm)), (0.0, 0.0))
                p = float(dtp) if dtodt == "DT" else float(odtp)
                if p > 0:
                    prices.append(round(p, 6))
            if not prices:
                return ""
            uniq = sorted(list({p for p in prices}))
            if len(uniq) == 1:
                return uniq[0]
            return "ÇOKLU"

        for ch in sorted(display_channels, key=lambda x: str(x["name"]).lower()):
            ch_name = str(ch["name"])
            ch_norm = self._norm_name(ch_name)
            ch_id = int(ch["id"])
            dinlenme = access_map.get(ch_norm, "NA")

            for dtodt in ("DT", "ODT"):
                day_vals = []
                day_secs = []
                day_bud = []
                for dd in dates:
                    day_vals.append(int(counts.get((ch_norm, dtodt, dd), 0)))
                    day_secs.append(float(seconds.get((ch_norm, dtodt, dd), 0.0)))
                    day_bud.append(float(budgets.get((ch_norm, dtodt, dd), 0.0)))

                # 0 -> boş
                day_vals_display = ["" if v == 0 else v for v in day_vals]

                # ay kolonları (adet/saniye)
                month_cols: list[Any] = []
                for (yy, mm) in months:
                    idxs = [i for i, dd in enumerate(dates) if dd.year == yy and dd.month == mm]
                    m_adet = int(sum(day_vals[i] for i in idxs))
                    m_san = float(sum(day_secs[i] for i in idxs))
                    month_cols.append("" if m_adet == 0 else m_adet)
                    month_cols.append("" if m_san == 0 else m_san)

                unit_disp = _unit_price_display(ch_id, dtodt)
                total_budget = float(sum(day_bud))
                rows_out.append(
                    {
                        "channel": ch_name,
                        "publish_group": "",
                        "dt_odt": dtodt,
                        "dinlenme_orani": dinlenme,
                        "days": day_vals_display,
                        "month_cols": month_cols,
                        "unit_price": unit_disp,
                        "budget": "" if total_budget == 0 else total_budget,
                    }
                )

        # Totals
        total_day = [0 for _ in range(len(dates))]
        total_month_cols: list[float] = [0.0 for _ in range(len(month_headers))]
        total_budget = 0.0

        for rr in rows_out:
            for i, v in enumerate(rr.get("days") or []):
                total_day[i] += int(v) if v not in ("", None) else 0

            mc = rr.get("month_cols") or []
            for i, v in enumerate(mc):
                if v in ("", None):
                    continue
                try:
                    total_month_cols[i] += float(v)
                except Exception:
                    continue

            if rr.get("budget") not in ("", None):
                total_budget += float(rr.get("budget"))

        totals = {
            "days": ["" if v == 0 else v for v in total_day],
            "month_cols": [
                "" if (i % 2 == 0 and int(v) == 0) or (i % 2 == 1 and float(v) == 0.0) else (int(v) if i % 2 == 0 else float(v))
                for i, v in enumerate(total_month_cols)
            ],
            "budget": "" if total_budget == 0 else total_budget,
        }

        header = {
            "agency": agency,
            "advertiser": advertiser,
            "product": product,
            "plan_title": pt,
            "reservation_no": reservation_no_display,
            "period": f"{rs:%d.%m.%Y}-{re:%d.%m.%Y}",
            "spot_len": spot_len,
        }

        return {
            "header": header,
            "start": rs,
            "end": re,
            "dates": dates,
            "months": months,
            "month_headers": month_headers,
            "rows": rows_out,
            "totals": totals,
        }

    def export_plan_ozet_range_excel(self, out_path, plan_title: str, start: date, end: date) -> None:
        """Plan Özet ekranındaki aralık bazlı özetin Excel çıktısını üretir."""
        from src.export.excel_exporter import export_plan_ozet_range
        data = self.get_plan_ozet_range_data(plan_title, start, end)
        export_plan_ozet_range(out_path, data)

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
        """Plan Özet (YILLIK) çıktısı üretir.

        Çıktı formatı:
        - TOPLAM sheet: 12 ay sütunu + yıl toplamı
        - 12 ay sheet: her ay için klasik Plan Özet şablonu
        """
        from pathlib import Path
        from src.export.excel_exporter import export_plan_ozet_yearly

        # 12 ayın verisini ayrı ayrı çek
        months_data = []
        for m in range(1, 13):
            month_data = self.get_plan_ozet_data(plan_title=plan_title, year=year, month=m)
            # month index bilgisi (exporter sheet adı için)
            month_data["month"] = m
            months_data.append(month_data)

        # TOPLAM sheet için aylık toplamları topla
        # channel bazlı: [ocak..aralık]
        channel_map = {}
        for m_idx, md in enumerate(months_data, start=1):
            for row in md.get("rows", []) or []:
                key = row.get("channel")
                if not key:
                    continue
                if key not in channel_map:
                    channel_map[key] = {
                        "channel": key,
                        "group": row.get("group", ""),
                        "dt_odt": row.get("dt_odt", ""),
                        "ratio": row.get("ratio", "NA"),
                        "unit_price": row.get("unit_price", 0),
                        "months": [0] * 12,
                    }
                # bu ayın toplam adeti = günlerin toplamı
                days = row.get("days", []) or []
                try:
                    month_total = sum(float(x) for x in days if x not in (None, ""))
                except Exception:
                    month_total = 0
                channel_map[key]["months"][m_idx - 1] += month_total
                # unit_price boşsa güncelle
                if not channel_map[key].get("unit_price") and row.get("unit_price"):
                    channel_map[key]["unit_price"] = row.get("unit_price")
                # ratio NA ise ve row'da değer varsa güncelle
                if (channel_map[key].get("ratio") in (None, "", "NA")) and row.get("ratio") not in (None, "", "NA"):
                    channel_map[key]["ratio"] = row.get("ratio")

        total_rows = []
        for ch, rec in sorted(channel_map.items(), key=lambda x: x[0].lower()):
            total_rows.append({
                "channel": rec["channel"],
                "group": rec.get("group", ""),
                "dt_odt": rec.get("dt_odt", ""),
                "ratio": rec.get("ratio", "NA"),
                "days": rec.get("months", [0] * 12),
                "unit_price": rec.get("unit_price", 0),
            })

        # Header bilgilerini ilk ayın header'ından al (boşsa da sorun değil)
        header0 = (months_data[0].get("header") or {}) if months_data else {}
        total_header = {
            "agency": header0.get("agency", ""),
            "advertiser": header0.get("advertiser", ""),
            "product": header0.get("product", ""),
            "plan_title": header0.get("plan_title", plan_title),
            "spot_len": header0.get("spot_len", 0),
            "year": year,
        }

        template_path = Path(__file__).resolve().parents[2] / "assets" / "plan_ozet_template.xlsx"

        export_plan_ozet_yearly(
            out_path,
            {
                "template_path": str(template_path),
                "year": year,
                "total": {
                    "header": total_header,
                    "rows": total_rows,
                    "month_labels": ["OCAK","ŞUBAT","MART","NİSAN","MAYIS","HAZİRAN","TEMMUZ","AĞUSTOS","EYLÜL","EKİM","KASIM","ARALIK"],
                },
                "months": months_data,
            },
        )