from __future__ import annotations

from pathlib import Path
from typing import Any
from datetime import datetime

import openpyxl
from openpyxl.drawing.image import Image

from src.util.paths import resource_path
from copy import copy

import calendar
from datetime import date
from openpyxl.utils import get_column_letter
from collections import Counter

TR_DOW = ["Pt", "Sa", "Ça", "Pş", "Cu", "Ct", "Pa"]  # Monday=0

def export_excel(template_path: Path, out_path: Path, payload: dict[str, Any]) -> Path:
    if not template_path.exists():
        raise FileNotFoundError(f"Template bulunamadı: {template_path}")

    wb = openpyxl.load_workbook(template_path)
    ws = wb["REZERVASYON ve PLANLAMA"]

    # --- Header labels + values ---
    ws["A1"].value = "Ajans:"
    ws["C1"].value = str(payload.get("agency_name", "")).strip()

    ws["A2"].value = "Reklam Veren:"
    ws["C2"].value = str(payload.get("advertiser_name", "")).strip()

    ws["A3"].value = "Ürün:"
    ws["C3"].value = str(payload.get("product_name", "")).strip()

    ws["A4"].value = "Plan Başlığı:"
    ws["C4"].value = str(payload.get("plan_title", "")).strip()

    ws["A5"].value = "Rezervasyon No:"
    ws["C5"].value = str(payload.get("reservation_no", "")).strip()
    
    # --- Header label font fix (A4/A5/A6 da A1 gibi kalın olsun) ---
    hdr_font = copy(ws["A1"].font)
    for addr in ("A4", "A5", "A6"):
        ws[addr].font = hdr_font


    # Dönemi: AY/YIL (plan_date varsa oradan üret)
    period = str(payload.get("period", "")).strip()
    if not period:
        plan_date = str(payload.get("plan_date", "")).strip()  # "YYYY-MM-DD" bekliyoruz
        try:
            y, m, _ = plan_date.split("-")
            period = f"{m}/{y}"
        except Exception:
            period = ""

    ws["A6"].value = "Dönemi:"
    ws["C6"].value = period

    # --- Ay/Gün başlıklarını plan_date'e göre düzelt + hafta sonu sütunlarını boyala ---
    plan_date = str(payload.get("plan_date", "")).strip()  # "YYYY-MM-DD"
    year = month = None
    try:
        y, m, _ = plan_date.split("-")
        year = int(y)
        month = int(m)
    except Exception:
        year = None
        month = None

    if year and month:
        days_in_month = calendar.monthrange(year, month)[1]
        HEADER_ROW = 7          # C7..AG7
        FIRST_DAY_COL = 3       # C
        MAX_DAYS = 31

        GRID_START_ROW = 8
        GRID_ROWS = 52

        # --- Şablondan "weekday/weekend" fill örneklerini otomatik yakala ---
        # (ODT bir satır seçelim ki DT satır grisiyle karışmasın)
        sample_row = GRID_START_ROW + 20

        # Grid alanında en sık görülen fill = weekday, daha az görülen = weekend                
        def fill_sig_from_cell(cell):
            f = cell.fill
            # StyleProxy'yi hash'e sokmayacağız; sadece primitive imza
            pt = getattr(f, "patternType", None)
            fg = getattr(getattr(f, "fgColor", None), "rgb", None)
            bg = getattr(getattr(f, "bgColor", None), "rgb", None)
            return (pt, fg, bg)

        # Grid örnek fill'leri (sadece signature sayacağız)
        grid_cells = []
        for d in range(1, MAX_DAYS + 1):
            c = FIRST_DAY_COL + (d - 1)
            grid_cells.append(ws.cell(sample_row, c))

        grid_sigs = [fill_sig_from_cell(cell) for cell in grid_cells]
        g_cnt = Counter(grid_sigs).most_common()

        weekday_sig = g_cnt[0][0] if g_cnt else None
        weekend_sig = g_cnt[1][0] if len(g_cnt) > 1 else weekday_sig

        # Header örnek fill'leri
        header_cells = []
        for d in range(1, MAX_DAYS + 1):
            c = FIRST_DAY_COL + (d - 1)
            header_cells.append(ws.cell(HEADER_ROW, c))

        header_sigs = [fill_sig_from_cell(cell) for cell in header_cells]
        h_cnt = Counter(header_sigs).most_common()

        h_weekday_sig = h_cnt[0][0] if h_cnt else None
        h_weekend_sig = h_cnt[1][0] if len(h_cnt) > 1 else h_weekday_sig

        def pick_fill(cells, sig):
            for cell in cells:
                if fill_sig_from_cell(cell) == sig:
                    return copy(cell.fill)  # proxy yerine kopya al
            return None

        weekday_fill = pick_fill(grid_cells, weekday_sig)
        weekend_fill = pick_fill(grid_cells, weekend_sig) or weekday_fill

        header_weekday_fill = pick_fill(header_cells, h_weekday_sig)
        header_weekend_fill = pick_fill(header_cells, h_weekend_sig) or header_weekday_fill

        disabled_fill = weekend_fill or weekday_fill

        for day in range(1, MAX_DAYS + 1):
            col = FIRST_DAY_COL + (day - 1)

            if day <= days_in_month:
                dow = TR_DOW[date(year, month, day).weekday()]
                ws.cell(HEADER_ROW, col).value = f"{dow}\n{day}"

                is_weekend = dow in ("Ct", "Pa")
                hf = header_weekend_fill if is_weekend else header_weekday_fill
                gf = weekend_fill if is_weekend else weekday_fill
            else:
                ws.cell(HEADER_ROW, col).value = None
                hf = header_weekday_fill  # header çok bozmasın
                gf = disabled_fill        # grid'i gri yap

            # Header fill uygula
            if hf is not None:
                ws.cell(HEADER_ROW, col).fill = hf

            # Grid fill uygula (C8..AG59)
            if gf is not None:
                for r in range(GRID_START_ROW, GRID_START_ROW + GRID_ROWS):
                    ws.cell(r, col).fill = gf

        # Bir önceki export'tan gizli kalmış olabilecek kolonlar sadece 29-31.
        # 1..28 gibi kolonlara dokunursak şablonun range genişlikleri bozulabiliyor.
        for day in range(29, 32):
            col_letter = get_column_letter(FIRST_DAY_COL + (day - 1))
            if col_letter in ws.column_dimensions:
                ws.column_dimensions[col_letter].hidden = False

        # Ay dışı gün kolonlarını gizle (29/30/31)
        for day in range(days_in_month + 1, MAX_DAYS + 1):
            col = FIRST_DAY_COL + (day - 1)
            col_letter = get_column_letter(col)
            ws.column_dimensions[col_letter].hidden = True
            # İstersen ekstra: genişliği de sıfıra yakın yap
            # ws.column_dimensions[col_letter].width = 0.1

    # Kanal adı
    ch = str(payload.get("channel_name", "")).strip()
    if ch:
        ws["U2"].value = ch

    # Sabit başlık
    ws["U3"].value = "Rezervasyon Formu"

    # --- PLAN GRID: temizle + doldur ---
    plan_cells = payload.get("plan_cells") or {}

    # Template'te plan grid başlangıcı: C8 (gün 1) varsayımı
    # Satırlar: 07:00-20:00, 15 dk => 52 satır (8..59)
    # Kolonlar: gün 1..31 => C..AG (3..33)
    GRID_START_ROW = 8
    GRID_START_COL = 3   # C
    GRID_ROWS = 52
    GRID_DAYS_MAX = 31

    # 1) Önce tüm grid alanını boşalt (eski A'lar geri gelmesin diye)
    for r in range(GRID_START_ROW, GRID_START_ROW + GRID_ROWS):
        for c in range(GRID_START_COL, GRID_START_COL + GRID_DAYS_MAX):
            ws.cell(r, c).value = None

    # 2) Sonra uygulamada girilenleri bas
    # plan_cells formatı: {(row_idx, day): "A"} veya {"row_idx,day": "A"} gelebilir
    for key, val in plan_cells.items():
        try:
            if isinstance(key, str):
                # "12,5" gibi gelirse
                row_idx_str, day_str = key.split(",")
                row_idx = int(row_idx_str)
                day = int(day_str)
            else:
                row_idx, day = key  # tuple
                row_idx = int(row_idx)
                day = int(day)

            if not (0 <= row_idx < GRID_ROWS and 1 <= day <= GRID_DAYS_MAX):
                continue

            rr = GRID_START_ROW + row_idx
            cc = GRID_START_COL + (day - 1)
            ws.cell(rr, cc).value = str(val)

        except Exception:
            # bozuk key gelirse export'u patlatmayalım
            continue

    # --- Logo ---
    logo_path = template_path.parent / "RADIOSCOPE.PNG"
    if not logo_path.exists():
        logo_path = resource_path("assets/RADIOSCOPE.PNG")

    try:
        ws._images = []  # kalsın, üst üste binmesin
        if not logo_path.exists():
            raise FileNotFoundError(f"Logo bulunamadı: {logo_path}")

        img = Image(str(logo_path))
        img.width = 128
        img.height = 128

        ws.add_image(img, "AO1")

    except Exception as e:
        print("[DEBUG] logo add FAILED:", repr(e))
        pass

    # --- Değiştirilebilir alt alanlar ---
    if payload.get("spot_code") is not None:
        ws["A67"].value = str(payload.get("spot_code", "")).strip()

    if payload.get("spot_duration") is not None:
        try:
            ws["B67"].value = int(payload.get("spot_duration"))
        except Exception:
            pass

    # D67: AH60'daki toplam mantığı -> plan_cells dolu sayısı (parantez içinde)
    adet_total = payload.get("adet_total")
    if adet_total is None:
        adet_total = sum(1 for v in (plan_cells or {}).values() if str(v).strip())
    ws["D67"].value = f"({int(adet_total)})"

    # A76: NOT satırının hemen üstü -> Kod Tanımı
    code_def = str(payload.get("code_definition", "")).strip()
    ws["A76"].value = code_def if code_def else None

    # NOT: sabit, içerik değişebilir
    note = str(payload.get("note_text", "")).strip()
    ws["A77"].value = "NOT:" if not note else f"NOT: {note}"

    # İsim değişebilir
    if payload.get("prepared_by") is not None:
        ws["AK77"].value = str(payload.get("prepared_by", "")).strip()

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return out_path
