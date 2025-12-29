from __future__ import annotations

from pathlib import Path
from typing import Any
from datetime import datetime

import openpyxl
from openpyxl.drawing.image import Image

from src.util.paths import resource_path


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

    # Kanal adı
    ch = str(payload.get("channel_name", "")).strip()
    if ch:
        ws["U2"].value = ch

    # Sabit başlık
    ws["U3"].value = "Rezervasyon Formu"

    # --- Logo: openpyxl kaydederken uçtuğu için yeniden ekliyoruz ---
    logo_path = resource_path("assets/RADIOSCOPE.PNG")
    try:
        ws._images = []  # tekrarlı basmayı önlemek için
        if logo_path.exists():
            img = Image(str(logo_path))
            img.width = 128
            img.height = 128
            img.anchor = "AO1"   # AO/AP civarı
            ws.add_image(img)
    except Exception:
        # Logo basılamasa da export’u çöpe atmayalım
        pass

    # --- Değiştirilebilir alt alanlar ---
    if payload.get("spot_code") is not None:
        ws["A67"].value = str(payload.get("spot_code", "")).strip()

    if payload.get("spot_duration") is not None:
        try:
            ws["B67"].value = int(payload.get("spot_duration"))
        except Exception:
            pass

    # D67 adet: şimdilik payload'dan, plan grid gelince otomatik saydırırız
    if payload.get("adet_total") is not None:
        try:
            ws["D67"].value = f"({int(payload.get('adet_total'))} Adet )"
        except Exception:
            pass

    if payload.get("client_name") is not None:
        ws["A76"].value = str(payload.get("client_name", "")).strip()

    # NOT: sabit, içerik değişebilir
    note = str(payload.get("note_text", "")).strip()
    ws["A77"].value = "NOT:" if not note else f"NOT: {note}"

    # İsim değişebilir
    if payload.get("prepared_by") is not None:
        ws["AK77"].value = str(payload.get("prepared_by", "")).strip()

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return out_path
