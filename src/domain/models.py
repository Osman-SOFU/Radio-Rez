from __future__ import annotations

from dataclasses import dataclass
from datetime import date, time
from typing import Any


@dataclass(frozen=True)
class ReservationDraft:
    advertiser_name: str
    plan_date: date
    spot_time: time

    # Rezervasyonun bağlandığı kanal (Fiyat ve Kanal Tanımı tab'ından)
    channel_name: str = ""
    # Seçilen kanalın, plan tarihinin (yıl/ay) fiyatları
    channel_price_dt: float = 0.0
    channel_price_odt: float = 0.0

    agency_name: str = ""
    product_name: str = ""
    plan_title: str = ""
    spot_code: str = ""
    spot_duration_sec: int = 0
    code_definition: str = ""
    note_text: str = ""
    prepared_by_name: str = ""


@dataclass(frozen=True)
class ConfirmedReservation:
    payload: dict[str, Any]

    def to_payload(self) -> dict[str, Any]:
        return dict(self.payload)
