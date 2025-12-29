from dataclasses import dataclass
from datetime import date, time

@dataclass
class ReservationDraft:
    advertiser_name: str
    plan_date: date
    spot_time: time
