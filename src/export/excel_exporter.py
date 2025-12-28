from __future__ import annotations

from pathlib import Path
from datetime import datetime, date
import calendar
import shutil
import zipfile

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage


_TR_DOW = {
    0: "Pt",
    1: "Sa",
    2: "Ça",
    3: "Pş",
    4: "Cu",
    5: "Ct",
    6: "Pa",
}


def _pick_ws(wb):
    for name in ("REZERVASYON ve PLANLAMA", "REZERVASYONve PLANLAMA"):
        if name in wb.sheetnames:
            return wb[name]
    return wb.active


def export_reservation_excel(
    *,
    template_path: str | Path,
    output_dir: str | Path,
    advertiser_name: str,
    channel_name: str,
    year: int,
    month: int,
    reservation_no: str | None,

    # --- yeni alanlar ---
    agency_name: str = "",
    product_name: str = "",
    plan_title: str = "",
    # A67 / B67 serbest
    a67_text: str | None = None,
    b67_text: str | None = None,
    # alt imza/not alanları
    a76_text: str | None = None,
    ak77_name: str | None = None,
    note_text: str | None = None,
    # logo
    logo_path: str | Path | None = None,
    logo_anchor: str = "AO2",
    logo_width: int | None = None,
    logo_height: int | None = None,
) -> Path:
    template_path = Path(template_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    if not template_path.exists():
        raise FileNotFoundError(f"Template bulunamadı: {template_path}")

    if not zipfile.is_zipfile(template_path):
        raise ValueError(
            f"Template dosyası gerçek .xlsx değil (zip değil): {template_path}\n"
            "Excel'de açıp 'Farklı Kaydet' ile tekrar .xlsx olarak kaydetmeyi dene."
        )

    safe_adv = (advertiser_name or "BILINMEYEN").strip().replace("/", "-")
    safe_channel = (channel_name or "KANAL").strip().replace("/", "-")
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    prefix = reservation_no if reservation_no else "TEST"
    out_name = f"{prefix}_{safe_adv}_{safe_channel}_{year}-{month:02d}_{stamp}.xlsx"
    out_path = output_dir / out_name

    shutil.copyfile(template_path, out_path)

    wb = load_workbook(out_path)
    ws = _pick_ws(wb)

    # --- İstediğin mapping ---
    # A1 label / C1 value mantığı: biz sadece C kolonuna yazıyoruz.
    ws["C1"].value = (agency_name or "").strip()          # Ajans
    ws["C2"].value = (advertiser_name or "").strip()      # Reklam Veren
    ws["C3"].value = (product_name or "").strip()         # Ürün
    ws["C4"].value = (plan_title or "").strip()           # Plan Başlığı
    ws["C5"].value = reservation_no or ""                 # Rezervasyon No
    ws["C6"].value = f"{month:02d}/{year}"                # Dönemi (AY/YIL)

    ws["U2"].value = (channel_name or "").strip()         # Kanal adı
    ws["U3"].value = "Rezervasyon Formu"                  # Sabit başlık

    # D67: A sayısı (Excel açılınca otomatik hesaplar)
    ws["D67"].value = '=COUNTIF(C8:AG59,"A")'

    # A67 / B67 değişebilir
    if a67_text is not None:
        ws["A67"].value = a67_text
    if b67_text is not None:
        ws["B67"].value = b67_text

    # A76 / AK77 / NOT alanları
    if a76_text is not None:
        ws["A76"].value = a76_text
    if ak77_name is not None:
        ws["AK77"].value = ak77_name

    # NOT: sabit, içerik değişken
    nt = (note_text or "").strip()
    ws["A77"].value = "NOT:" if not nt else f"NOT: {nt}"

    # Logo (AO/AP civarı)
    if logo_path:
        lp = Path(logo_path)
        if lp.exists():
            img = XLImage(str(lp))
            img.anchor = logo_anchor  # default AO2
            if logo_width is not None:
                img.width = logo_width
            if logo_height is not None:
                img.height = logo_height
            ws.add_image(img)

    # --- Gün başlıkları (C7=1.gün) ---
    days_in_month = calendar.monthrange(year, month)[1]
    for day in range(1, 32):
        col_idx = 2 + day  # C=3 => day=1
        col_letter = get_column_letter(col_idx)
        header_cell = f"{col_letter}7"

        if day <= days_in_month:
            dow = _TR_DOW[date(year, month, day).weekday()]
            ws[header_cell].value = f"{dow}\n{day}"
        else:
            ws[header_cell].value = None
            for r in range(8, 60):
                ws[f"{col_letter}{r}"].value = None

    wb.save(out_path)
    wb.close()
    return out_path
