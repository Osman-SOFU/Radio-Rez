from __future__ import annotations
from datetime import date, time


def classify_dt_odt(t: time) -> str:
    """DT: 07:00–10:00 ve 17:00–20:00 (sınırlar dahil)."""
    mins = t.hour * 60 + t.minute
    dt1 = 7 * 60 <= mins < 10 * 60
    dt2 = 17 * 60 <= mins < 20 * 60
    return "DT" if (dt1 or dt2) else "ODT"


def validate_day(plan_date: date) -> tuple[bool, str]:
    # QDate zaten geçersiz gün seçtirmez; yine de güvenlik.
    try:
        _ = plan_date.toordinal()
        return True, ""
    except Exception:
        return False, "Geçersiz tarih."
