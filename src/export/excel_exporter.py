from __future__ import annotations

from pathlib import Path
from typing import Any
from datetime import datetime, time

import openpyxl
from openpyxl.drawing.image import Image

from src.util.paths import resource_path
from copy import copy

import calendar
from datetime import date
from openpyxl.utils import get_column_letter, column_index_from_string
from collections import Counter
from openpyxl import Workbook

from src.domain.time_rules import classify_dt_odt

TR_DOW = ["Pt", "Sa", "Ça", "Pş", "Cu", "Ct", "Pa"]  # Monday=0


def _row_idx_to_time(row_idx: int) -> time:
    """Grid satırı -> kuşak başlangıç saati.
    Şablon: 07:00-20:00, 15dk.
    """
    mins = 7 * 60 + int(row_idx) * 15
    return time(mins // 60, mins % 60)

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
        # Not: DT satırlarının şablonda ayrı bir tonu var. Önceki yaklaşım tek bir satırdan
        # örnek aldığı için DT satırı tonunu ezip geçiyordu. Bu yüzden DT ve ODT için ayrı ayrı
        # örnekleyip, export sırasında satır tipine göre fill basıyoruz.

        # Grid alanında en sık görülen fill = weekday, daha az görülen = weekend
        def fill_sig_from_cell(cell):
            f = cell.fill
            # StyleProxy'yi hash'e sokmayacağız; sadece primitive imza
            pt = getattr(f, "patternType", None)
            fg = getattr(getattr(f, "fgColor", None), "rgb", None)
            bg = getattr(getattr(f, "bgColor", None), "rgb", None)
            return (pt, fg, bg)

        def analyze_grid_row(sample_row: int) -> tuple[Any, Any]:
            """Verilen satır için (weekday_fill, weekend_fill) döndürür."""
            cells = []
            for d in range(1, MAX_DAYS + 1):
                c = FIRST_DAY_COL + (d - 1)
                cells.append(ws.cell(sample_row, c))

            sigs = [fill_sig_from_cell(cell) for cell in cells]
            cnt = Counter(sigs).most_common()
            weekday_sig = cnt[0][0] if cnt else None
            weekend_sig = cnt[1][0] if len(cnt) > 1 else weekday_sig

            def pick_fill(cells2, sig):
                for cell in cells2:
                    if fill_sig_from_cell(cell) == sig:
                        return copy(cell.fill)
                return None

            weekday_fill = pick_fill(cells, weekday_sig)
            weekend_fill = pick_fill(cells, weekend_sig) or weekday_fill
            return weekday_fill, weekend_fill

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

        header_weekday_fill = pick_fill(header_cells, h_weekday_sig)
        header_weekend_fill = pick_fill(header_cells, h_weekend_sig) or header_weekday_fill

        # DT için örnek satır: 07:00 (row_idx=0)
        dt_row = GRID_START_ROW + 0
        # ODT için örnek satır: 11:00 (row_idx=16)  -> 10:15 sonrası ODT
        odt_row = GRID_START_ROW + 16

        dt_weekday_fill, dt_weekend_fill = analyze_grid_row(dt_row)
        odt_weekday_fill, odt_weekend_fill = analyze_grid_row(odt_row)

        disabled_fill = odt_weekday_fill or dt_weekday_fill

        for day in range(1, MAX_DAYS + 1):
            col = FIRST_DAY_COL + (day - 1)

            if day <= days_in_month:
                dow = TR_DOW[date(year, month, day).weekday()]
                ws.cell(HEADER_ROW, col).value = f"{dow}\n{day}"

                is_weekend = dow in ("Ct", "Pa")
                hf = header_weekend_fill if is_weekend else header_weekday_fill
            else:
                ws.cell(HEADER_ROW, col).value = None
                hf = header_weekday_fill  # header çok bozmasın
                gf = disabled_fill        # grid'i gri yap

            # Header fill uygula
            if hf is not None:
                ws.cell(HEADER_ROW, col).fill = hf

            # Grid fill uygula (C8..AG59) - satır tipine göre DT/ODT tonu koru
            for row_idx in range(0, GRID_ROWS):
                rr = GRID_START_ROW + row_idx

                if day > days_in_month:
                    if disabled_fill is not None:
                        ws.cell(rr, col).fill = disabled_fill
                    continue

                slot_type = classify_dt_odt(_row_idx_to_time(row_idx))
                if is_weekend:
                    gf2 = dt_weekend_fill if slot_type == "DT" else odt_weekend_fill
                else:
                    gf2 = dt_weekday_fill if slot_type == "DT" else odt_weekday_fill

                if gf2 is not None:
                    ws.cell(rr, col).fill = gf2

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

    # Kanal adı (Excel'de T2)
    ch = str(payload.get("channel_name", "")).strip()
    if ch:
        ws["T2"].value = ch

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

    # --- Birim fiyatı (AO sütunu) ---
    # DT/ODT fiyatları: seçilen kanalın, plan tarihinin AY/YIL'ına göre.
    try:
        unit_price_col = column_index_from_string("AO")
        dt_price = float(payload.get("channel_price_dt") or 0)
        odt_price = float(payload.get("channel_price_odt") or 0)

        for row_idx in range(GRID_ROWS):
            rr = GRID_START_ROW + row_idx
            slot_type = classify_dt_odt(_row_idx_to_time(row_idx))
            ws.cell(rr, unit_price_col).value = dt_price if slot_type == "DT" else odt_price
    except Exception:
        # birim fiyatı basılamazsa export'u patlatma
        pass

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

    # B67: süre (sn)
    dur_val = payload.get("spot_duration_sec", None)
    if dur_val is None:
        # geriye dönük uyumluluk
        dur_val = payload.get("spot_duration", None)
    if dur_val is not None:
        try:
            ws["B67"].value = int(dur_val)
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

def export_kod_tanimi(out_path, advertiser_name: str, rows: list[dict]) -> None:
    """
    KOD TANIMI çıktısını, assets/kod_tanimi_template.xlsx şablonunun görsel stilini koruyarak üretir.
    Tablo yerleşimi şablondaki gibi C5:F13 aralığındadır.
    """
    out_path = Path(out_path)

    template_path = resource_path("assets/kod_tanimi_template.xlsx")
    if template_path.exists():
        wb = openpyxl.load_workbook(template_path)
        ws = wb["KOPYA TANIMI"] if "KOPYA TANIMI" in wb.sheetnames else wb.active
    else:
        # Fallback: basit boş dosya
        wb = Workbook()
        ws = wb.active
        ws.title = "KOPYA TANIMI"
        # Header
        ws["C5"].value = "Kod"
        ws["D5"].value = "Kod Tanımı"
        ws["E5"].value = "Kod Uzunluğu (SN)"
        ws["F5"].value = "Dağılım"

    start_row = 6
    base_rows = 7  # şablonda 6..12
    total_row = start_row + base_rows  # 13

    # --- satır sayısını ihtiyaca göre büyüt (7'yi aşarsa) ---
    n = len(rows)
    if n > base_rows:
        insert_count = n - base_rows
        ws.insert_rows(total_row, amount=insert_count)

        # 12. satırın (son veri satırı) stilini yeni satırlara kopyala
        src_row = total_row - 1  # eski 12
        for i in range(insert_count):
            dst_row = src_row + 1 + i
            for col in range(3, 7):  # C..F
                src = ws.cell(src_row, col)
                dst = ws.cell(dst_row, col)
                dst._style = copy(src._style)
                dst.number_format = src.number_format
                dst.font = copy(src.font)
                dst.border = copy(src.border)
                dst.fill = copy(src.fill)
                dst.alignment = copy(src.alignment)
                dst.protection = copy(src.protection)
                dst.comment = None

        total_row += insert_count

    # --- önce tabloyu temizle (şablondaki örnek veriler varsa sil) ---
    max_data_rows = max(base_rows, n)
    for r in range(start_row, start_row + max_data_rows):
        for c in range(3, 7):  # C..F
            ws.cell(r, c).value = None

    # --- verileri yaz ---
    for i, r in enumerate(rows):
        rr = start_row + i
        ws.cell(rr, 3).value = r.get("code", "")
        ws.cell(rr, 4).value = r.get("code_desc", "")
        ws.cell(rr, 5).value = int(r.get("length_sn", 0) or 0)
        ws.cell(rr, 6).value = float(r.get("distribution", 0.0) or 0.0)

        # Dağılım yüzde formatında görünsün
        ws.cell(rr, 6).number_format = "0%"

    # --- kalan satırları boş bırak ama format kalsın ---
    # (şablon zaten formatlı, biz sadece value None bıraktık)

    # --- Ort.Uzun. formülleri: dinamik aralık ---
    last_data_row = start_row + max_data_rows - 1
    ws.cell(total_row, 3).value = "Ort.Uzun."
    ws.cell(total_row, 5).value = f"=SUMPRODUCT(E{start_row}:E{last_data_row},F{start_row}:F{last_data_row})"
    ws.cell(total_row, 6).value = f"=SUM(F{start_row}:F{last_data_row})"

    # toplam dağılım hücresi yüzde görünsün
    ws.cell(total_row, 6).number_format = "0%"

    # küçük bir başlık (dosya adı vs.) istersen buraya eklenir; şimdilik dokunmuyoruz.
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def export_spotlist(out_path, advertiser_name: str, rows: list[dict]) -> None:
    """SPOTLİST+ çıktısını şablonun stilini bozmadan üretir.

    Şablon: assets/spotlist_template.xlsx (sheet: "SPOTLİST+")
    Data satırları: 2. satırdan itibaren.
    """
    out_path = Path(out_path)

    template_path = resource_path("assets/spotlist_template.xlsx")
    if template_path.exists():
        wb = openpyxl.load_workbook(template_path)
        ws = wb["SPOTLİST+"] if "SPOTLİST+" in wb.sheetnames else wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "SPOTLİST+"
        headers = [
            "Sıra",
            "TARIH",
            "ANA YAYIN",
            "REKLAMIN FIRMASI",
            "ADET",
            "BASLANGIC",
            "SURE",
            "Spot Kodu",
            "DT-ODT",
            "Birim Saniye ",
            "Bütçe  Net TL",
            "Yayınlandı Durum",
        ]
        for i, h in enumerate(headers, start=1):
            ws.cell(1, i).value = h

    start_row = 2
    n = len(rows)

    # Şablonda hazır satır yoksa (veya yetmezse) satır ekle ve stil kopyala
    if ws.max_row < start_row:
        ws.insert_rows(start_row, amount=1)

    # Stil kopyalama için referans satır: start_row
    style_row = start_row if ws.max_row >= start_row else 1

    needed_last = start_row + max(n, 0) - 1
    if needed_last > ws.max_row:
        insert_count = needed_last - ws.max_row
        ws.insert_rows(ws.max_row + 1, amount=insert_count)
        # son mevcut stil satırını kopyala
        src_row = style_row
        for i in range(0, insert_count):
            dst_row = ws.max_row - insert_count + 1 + i
            for col in range(1, 13):
                src = ws.cell(src_row, col)
                dst = ws.cell(dst_row, col)
                dst._style = copy(src._style)
                dst.number_format = src.number_format
                dst.font = copy(src.font)
                dst.border = copy(src.border)
                dst.fill = copy(src.fill)
                dst.alignment = copy(src.alignment)
                dst.protection = copy(src.protection)
                dst.comment = None

    # Önce eski değerleri temizle (önceki export'tan kalmasın)
    # Template çok büyük olabiliyor, o yüzden sadece "kullanılmış" aralığı temizliyoruz.
    last_used = 1
    for r in range(start_row, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v is not None and str(v).strip() != "":
            last_used = r
    clear_last = max(last_used, needed_last)
    for r in range(start_row, clear_last + 1):
        for c in range(1, 13):
            ws.cell(r, c).value = None

    # Veriyi bas
    for i, rr in enumerate(rows):
        r = start_row + i
        ws.cell(r, 1).value = int(rr.get("sira", i + 1) or (i + 1))
        ws.cell(r, 2).value = rr.get("tarih", "")
        ws.cell(r, 3).value = rr.get("ana_yayin", "")
        ws.cell(r, 4).value = rr.get("reklam_firmasi", advertiser_name)
        ws.cell(r, 5).value = int(rr.get("adet", 1) or 1)
        ws.cell(r, 6).value = rr.get("baslangic", "")
        ws.cell(r, 7).value = int(rr.get("sure", 0) or 0)
        ws.cell(r, 8).value = rr.get("spot_kodu", "")
        ws.cell(r, 9).value = rr.get("dt_odt", "")
        ws.cell(r, 10).value = float(rr.get("birim_saniye", 0.0) or 0.0)
        ws.cell(r, 11).value = float(rr.get("butce_net", 0.0) or 0.0)
        ws.cell(r, 12).value = int(rr.get("published", 0) or 0)

    
    # Özet satırı (filtrelenmiş listeye göre)
    total_adet = sum(int(r.get("adet", 1) or 1) for r in rows) if rows else 0
    total_budget = sum(float(r.get("butce_net", 0.0) or 0.0) for r in rows) if rows else 0.0
    durations = [int(r.get("sure", 0) or 0) for r in rows] if rows else []
    avg_duration = (sum(durations) / len(durations)) if durations else 0.0

    sum_row = start_row + n
    if sum_row > ws.max_row:
        ws.insert_rows(ws.max_row + 1, amount=(sum_row - ws.max_row))

    # Stil kopyala (kenarlıklar vs kalsın)
    for col in range(1, 13):
        src = ws.cell(style_row, col)
        dst = ws.cell(sum_row, col)
        dst._style = copy(src._style)
        dst.number_format = src.number_format
        dst.font = copy(src.font)
        dst.border = copy(src.border)
        dst.fill = copy(src.fill)
        dst.alignment = copy(src.alignment)
        dst.protection = copy(src.protection)
        dst.comment = None

    # Değerleri yaz
    for c in range(1, 13):
        ws.cell(sum_row, c).value = None

    ws.cell(sum_row, 2).value = "TOPLAM"
    ws.cell(sum_row, 5).value = int(total_adet)
    ws.cell(sum_row, 7).value = float(avg_duration)
    ws.cell(sum_row, 11).value = float(total_budget)

    # Kalın yaz
    for c in range(1, 13):
        cell = ws.cell(sum_row, c)
        cell.font = copy(cell.font)
        cell.font = cell.font.copy(bold=True)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
