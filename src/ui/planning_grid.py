from __future__ import annotations

import calendar
from dataclasses import dataclass
from datetime import datetime, timedelta, time as dtime

from PySide6.QtCore import Qt
from PySide6.QtGui import QBrush, QColor, QFont
from PySide6.QtWidgets import QWidget, QVBoxLayout, QTableWidget, QTableWidgetItem

from src.domain.time_rules import classify_dt_odt

TR_DOW = ["Pt", "Sa", "Ça", "Pş", "Cu", "Ct", "Pa"]  # Monday=0

def build_timeslots(start="07:00", end="20:00", step_min=15):
    sh, sm = map(int, start.split(":"))
    eh, em = map(int, end.split(":"))
    cur = datetime(2000, 1, 1, sh, sm)
    last = datetime(2000, 1, 1, eh, em)
    out = []
    while cur < last:
        nxt = cur + timedelta(minutes=step_min)
        out.append((cur.time(), f"{cur:%H:%M}-{nxt:%H:%M}"))
        cur = nxt
    return out


class PlanningGrid(QWidget):
    """
    Excel-benzeri plan grid'i:
    - Satırlar: 07:00-20:00 (15dk)
    - Kolonlar: Kuşak, Dolar Kuru + ayın günleri (1..N)
    - DT saat satırları daha koyu, ODT normal
    - Hafta sonu kolonları farklı renkte
    - Seçilen gün kolonu header'ı vurgulanır
    """
    def __init__(self, parent=None):
        super().__init__(parent)

        self.times = build_timeslots()
        self.year = datetime.now().year
        self.month = datetime.now().month
        self.selected_day: int | None = None

        self.table = QTableWidget(self)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.addWidget(self.table)

        self.table.setAlternatingRowColors(False)
        self.table.setSortingEnabled(False)

        self.set_month(self.year, self.month, None)

    def set_month(self, year: int, month: int, selected_day: int | None):
        self.year = year
        self.month = month
        self.selected_day = selected_day

        days_in_month = calendar.monthrange(year, month)[1]
        total_cols = 2 + days_in_month  # Kuşak + Dolar + Günler

        self.table.clear()
        self.table.setRowCount(len(self.times))
        self.table.setColumnCount(total_cols)

        # Header
        headers = ["Kuşak", "Dolar\nKuru"]
        for d in range(1, days_in_month + 1):
            dow = TR_DOW[datetime(year, month, d).weekday()]
            headers.append(f"{dow}\n{d}")
        self.table.setHorizontalHeaderLabels(headers)

        # Kolon genişlikleri (istersen sonra ayarlarız)
        self.table.setColumnWidth(0, 140)
        self.table.setColumnWidth(1, 80)
        for c in range(2, total_cols):
            self.table.setColumnWidth(c, 60)

        # Satırları doldur
        dt_row_brush = QBrush(QColor("#e0e0e0"))
        weekend_brush = QBrush(QColor("#f3f3f3"))
        normal_brush = QBrush(QColor("#ffffff"))

        for r, (start_time, slot_text) in enumerate(self.times):
            # 1) Kuşak (non-editable)
            it0 = QTableWidgetItem(slot_text)
            it0.setFlags(it0.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(r, 0, it0)

            # 2) Dolar (şimdilik sabit “2”, non-editable)
            it1 = QTableWidgetItem("2")
            it1.setTextAlignment(Qt.AlignCenter)
            it1.setFlags(it1.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(r, 1, it1)

            # DT/ODT satır rengi
            row_is_dt = (classify_dt_odt(start_time) == "DT")
            row_brush = dt_row_brush if row_is_dt else normal_brush

            # Gün hücreleri (editable)
            for day in range(1, days_in_month + 1):
                col = 1 + day  # 2.. (2 + days_in_month-1)
                cell = QTableWidgetItem("")
                self.table.setItem(r, col, cell)

                # Hafta sonu kolon rengi
                dow = TR_DOW[datetime(year, month, day).weekday()]
                is_weekend = dow in ("Ct", "Pa")

                if is_weekend:
                    cell.setBackground(weekend_brush)
                else:
                    cell.setBackground(row_brush)

        # Seçili gün header vurgusu
        if self.selected_day and 1 <= self.selected_day <= days_in_month:
            sel_col = 1 + self.selected_day
            font = self.table.horizontalHeader().font()
            font.setBold(True)
            self.table.horizontalHeaderItem(sel_col).setFont(font)
            self.table.horizontalHeaderItem(sel_col).setBackground(QBrush(QColor("#d0d0d0")))

    def get_matrix(self) -> dict[str, str]:
        out: dict[str, str] = {}
        days_in_month = calendar.monthrange(self.year, self.month)[1]
        for r in range(self.table.rowCount()):
            for day in range(1, days_in_month + 1):
                c = 1 + day
                it = self.table.item(r, c)
                if not it:
                    continue
                v = it.text().strip()
                if v:
                    out[f"{r},{day}"] = v   # <-- JSON safe
        return out
