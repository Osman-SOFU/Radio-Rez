from __future__ import annotations

import calendar
from dataclasses import dataclass
from datetime import datetime, timedelta, time as dtime, date

from PySide6.QtCore import Qt
from PySide6.QtGui import QBrush, QColor, QFont
from PySide6.QtWidgets import QWidget, QVBoxLayout, QTableWidgetItem, QAbstractItemView

from src.ui.excel_table import ExcelTableWidget

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

        # Display mode
        # - "month": classic 1..N days of a single month
        # - "span": arbitrary inclusive date range (can cross months)
        self._mode: str = "month"
        self._span_dates: list[date] = []
        # Span mode: start/end are kept as a safety net so we can rebuild span dates
        # even if the internal cache gets cleared by other UI flows.
        self._span_start: date | None = None
        self._span_end: date | None = None
        self._span_month_slices: dict[tuple[int, int], tuple[int, int]] = {}
        self._selected_date: date | None = None

        # Optional date range limitation (inclusive). When set, day columns outside
        # the range are hidden for the currently displayed month.
        self._range_start: date | None = None
        self._range_end: date | None = None

        self.table = ExcelTableWidget(self)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.addWidget(self.table)

        self.table.setAlternatingRowColors(False)
        self.table.setSortingEnabled(False)

        # Excel-like selection (row/column/block)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setSelectionBehavior(QAbstractItemView.SelectItems)

        # Daha kompakt görünüm
        f = self.table.font()
        try:
            f.setPointSize(8)
        except Exception:
            pass
        self.table.setFont(f)

        # dışarıdan okunabilirlik: kayıtlı rezervasyonu sadece görmek için
        self._read_only: bool = False

        # Sizing preferences ("no-scroll" hedefi için dinamik daraltma)
        self._fixed_col_widths = {0: 140, 1: 70}
        self._day_col_min = 28
        self._day_col_max = 60
        self._row_min = 14
        self._row_max = 22

        self.table.horizontalHeader().setMinimumSectionSize(self._day_col_min)

        self.set_month(self.year, self.month, None)

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._apply_dynamic_sizes()

    def is_span_mode(self) -> bool:
        return self._mode == "span"

    def _apply_dynamic_sizes(self) -> None:
        """Try to fit the whole month horizontally; reduce row height for less vertical scroll."""
        try:
            if self._mode == "span":
                day_count = max(0, len(self._span_dates))
            else:
                day_count = calendar.monthrange(self.year, self.month)[1]
            total_cols = 2 + day_count

            # fixed columns
            fixed_total = 0
            for c, w in self._fixed_col_widths.items():
                self.table.setColumnWidth(c, w)
                fixed_total += w

            # remaining width for day columns
            avail = max(0, self.table.viewport().width() - fixed_total - 2)
            if day_count > 0:
                day_w = int(avail / day_count) if avail > 0 else self._day_col_min
                day_w = max(self._day_col_min, min(self._day_col_max, day_w))
            else:
                day_w = self._day_col_min

            for c in range(2, total_cols):
                self.table.setColumnWidth(c, day_w)

            # row height attempt
            rows = self.table.rowCount() or 1
            rh_avail = max(0, self.table.viewport().height() - self.table.horizontalHeader().height() - 2)
            row_h = int(rh_avail / rows) if rh_avail > 0 else self._row_min
            row_h = max(self._row_min, min(self._row_max, row_h))
            self.table.verticalHeader().setDefaultSectionSize(row_h)
        except Exception:
            pass

    def set_read_only(self, read_only: bool) -> None:
        """Gün hücrelerini düzenlenemez yapar (Kuşak/Dolar zaten sabit)."""
        self._read_only = bool(read_only)
        self._apply_read_only_flags()

    def set_date_range(self, start: date | None, end: date | None) -> None:
        """Grid'i bir tarih aralığına (dahil) sınırlar.

        - Aralık set edilirse, o ay görünümünde aralık dışı gün kolonları gizlenir.
        - start/end None verilirse aralık temizlenir ve tüm gün kolonları görünür olur.
        """
        if start is None or end is None:
            self._range_start = None
            self._range_end = None
        else:
            # normalize
            if end < start:
                start, end = end, start
            self._range_start = start
            self._range_end = end

        self._apply_range_visibility()

    def _apply_range_visibility(self) -> None:
        """Seçili ay görünümünde, aralık dışındaki gün kolonlarını gizle/göster."""
        try:
            days_in_month = calendar.monthrange(self.year, self.month)[1]
            # Önce tümünü aç
            for day in range(1, days_in_month + 1):
                col = 1 + day
                if col < self.table.columnCount():
                    self.table.setColumnHidden(col, False)

            if not self._range_start or not self._range_end:
                return

            for day in range(1, days_in_month + 1):
                dt = date(self.year, self.month, day)
                in_range = (self._range_start <= dt <= self._range_end)
                col = 1 + day
                if col < self.table.columnCount():
                    self.table.setColumnHidden(col, not in_range)
        except Exception:
            # UI tarafında range uygular... (sessiz geç)
            return

    def clear_matrix(self) -> None:
        """Sadece gün hücrelerini temizle."""
        if self._mode == "span":
            day_count = len(self._span_dates)
            for r in range(self.table.rowCount()):
                for i in range(day_count):
                    c = 2 + i
                    it = self.table.item(r, c)
                    if it:
                        it.setText("")
            return

        days_in_month = calendar.monthrange(self.year, self.month)[1]
        for r in range(self.table.rowCount()):
            for day in range(1, days_in_month + 1):
                c = 1 + day
                it = self.table.item(r, c)
                if it:
                    it.setText("")

    def set_matrix(self, plan_cells: dict) -> None:
        """DB'den gelen plan_cells'i grid'e basar.

        plan_cells anahtar formatı: "row_idx,day" (string) veya (row_idx, day)
        """
        # Backward compatible entry point.
        # - month mode: expects row,day keys for the current month
        # - span mode: we treat given matrix as belonging to the span start month
        if self._mode == "span":
            self.set_span_month_matrices({(self.year, self.month): (plan_cells or {})})
            return

        self.clear_matrix()
        if not plan_cells:
            self._apply_read_only_flags()
            return

        days_in_month = calendar.monthrange(self.year, self.month)[1]

        for k, v in (plan_cells or {}).items():
            if not str(v or "").strip():
                continue
            try:
                if isinstance(k, str):
                    r_s, d_s = k.split(",")
                    r = int(r_s)
                    day = int(d_s)
                else:
                    r, day = k
                    r = int(r)
                    day = int(day)
            except Exception:
                continue

            if r < 0 or r >= self.table.rowCount():
                continue
            if day < 1 or day > days_in_month:
                continue

            c = 1 + day
            it = self.table.item(r, c)
            if not it:
                it = QTableWidgetItem("")
                self.table.setItem(r, c, it)
            it.setText(str(v))

        self._apply_read_only_flags()
        # Aralık kısıtı varsa kolon görünürlüğünü güncelle
        self._apply_range_visibility()

    def _apply_read_only_flags(self) -> None:
        """Sadece gün hücrelerinde editable flag kontrolü."""
        if self._mode == "span":
            day_count = len(self._span_dates)
            for r in range(self.table.rowCount()):
                for i in range(day_count):
                    c = 2 + i
                    it = self.table.item(r, c)
                    if not it:
                        continue
                    flags = it.flags()
                    if self._read_only:
                        it.setFlags(flags & ~Qt.ItemIsEditable)
                    else:
                        it.setFlags(flags | Qt.ItemIsEditable)
            return

        days_in_month = calendar.monthrange(self.year, self.month)[1]
        for r in range(self.table.rowCount()):
            for day in range(1, days_in_month + 1):
                c = 1 + day
                it = self.table.item(r, c)
                if not it:
                    continue
                flags = it.flags()
                if self._read_only:
                    it.setFlags(flags & ~Qt.ItemIsEditable)
                else:
                    it.setFlags(flags | Qt.ItemIsEditable)

    def set_month(self, year: int, month: int, selected_day: int | None):
        self._mode = "month"
        self._span_dates = []
        self._span_start = None
        self._span_end = None
        self._span_month_slices = {}
        self._selected_date = None
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

        # Başlangıç kolon genişlikleri; resizeEvent ile dinamik ayarlanır.
        for c, w in self._fixed_col_widths.items():
            if c < total_cols:
                self.table.setColumnWidth(int(c), int(w))
        for c in range(2, total_cols):
            self.table.setColumnWidth(c, self._day_col_max)
        for c in range(2, total_cols):
            self.table.setColumnWidth(c, self._day_col_max)

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

        # read-only mod açıksa hücre flag'lerini tekrar uygula
        self._apply_read_only_flags()

        # Aralık kısıtı varsa, yeni ay için kolon görünürlüğünü güncelle
        self._apply_range_visibility()

        # mümkün olduğunca scroll ihtiyacını azalt
        self._apply_dynamic_sizes()

    # ------------------------------
    # Span (date range) mode
    # ------------------------------

    def set_date_span(self, start: date, end: date, selected_date: date | None = None) -> None:
        """Render a single grid for an inclusive date range (can cross months).

        Internally still uses day numbers per month. Higher layers can split
        the grid into per-month matrices via get_span_month_matrices().
        """
        if start and end and start > end:
            start, end = end, start

        # Keep boundaries (used by fallback rebuild if internal span cache is cleared).
        self._span_start = start
        self._span_end = end

        self._mode = "span"
        self._range_start = None
        self._range_end = None
        self.selected_day = None

        # Build date list
        dates: list[date] = []
        cur = start
        while cur <= end:
            dates.append(cur)
            cur = cur.fromordinal(cur.toordinal() + 1)

        self._span_dates = dates
        self._selected_date = selected_date if (selected_date in dates) else (dates[0] if dates else None)

        # Keep a representative (year,month) for compatibility / fallback
        if dates:
            self.year = dates[0].year
            self.month = dates[0].month

        total_cols = 2 + len(dates)
        self.table.clear()
        self.table.setRowCount(len(self.times))
        self.table.setColumnCount(total_cols)

        headers = ["Kuşak", "Dolar\nKuru"]
        for d in dates:
            dow = TR_DOW[datetime(d.year, d.month, d.day).weekday()]
            headers.append(f"{dow}\n{d.day:02d}.{d.month:02d}")
        self.table.setHorizontalHeaderLabels(headers)

        # Default widths; resizeEvent will refine
        for c, w in self._fixed_col_widths.items():
            if c < total_cols:
                self.table.setColumnWidth(int(c), int(w))
        for c in range(2, total_cols):
            self.table.setColumnWidth(c, self._day_col_max)

        dt_row_brush = QBrush(QColor("#e0e0e0"))
        weekend_brush = QBrush(QColor("#f3f3f3"))
        normal_brush = QBrush(QColor("#ffffff"))

        for r, (start_time, slot_text) in enumerate(self.times):
            it0 = QTableWidgetItem(slot_text)
            it0.setFlags(it0.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(r, 0, it0)

            it1 = QTableWidgetItem("2")
            it1.setTextAlignment(Qt.AlignCenter)
            it1.setFlags(it1.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(r, 1, it1)

            row_is_dt = (classify_dt_odt(start_time) == "DT")
            row_brush = dt_row_brush if row_is_dt else normal_brush

            for i, d in enumerate(dates):
                col = 2 + i
                cell = QTableWidgetItem("")
                self.table.setItem(r, col, cell)

                dow = TR_DOW[datetime(d.year, d.month, d.day).weekday()]
                is_weekend = dow in ("Ct", "Pa")
                if is_weekend:
                    cell.setBackground(weekend_brush)
                else:
                    cell.setBackground(row_brush)

        # Highlight selected date column
        self._apply_selected_date_highlight()
        self._apply_read_only_flags()
        self._apply_dynamic_sizes()

    def _apply_selected_date_highlight(self) -> None:
        if self._mode != "span" or not self._span_dates:
            return
        # reset fonts
        base_font = self.table.horizontalHeader().font()
        for i in range(len(self._span_dates)):
            col = 2 + i
            it = self.table.horizontalHeaderItem(col)
            if it:
                f = QFont(base_font)
                f.setBold(False)
                it.setFont(f)
                it.setBackground(QBrush())

        if not self._selected_date:
            return
        try:
            idx = self._span_dates.index(self._selected_date)
        except ValueError:
            return
        sel_col = 2 + idx
        it = self.table.horizontalHeaderItem(sel_col)
        if it:
            f = QFont(base_font)
            f.setBold(True)
            it.setFont(f)
            it.setBackground(QBrush(QColor("#d0d0d0")))

    def set_selected_date(self, d: date | None) -> None:
        """Update header highlight only."""
        if self._mode != "span":
            # month mode uses selected_day
            if d is None:
                return
            if d.year == self.year and d.month == self.month:
                self.selected_day = d.day
                self.set_month(self.year, self.month, self.selected_day)
            return
        if d and d in self._span_dates:
            self._selected_date = d
            self._apply_selected_date_highlight()

    def get_span_month_matrices(self) -> dict[tuple[int, int], dict[str, str]]:
        """Return per-month matrices keyed as (year,month) -> {"row,day": code}."""
        out: dict[tuple[int, int], dict[str, str]] = {}
        if self._mode != "span" or not self._span_dates:
            # fallback to month mode matrix
            return {(self.year, self.month): self.get_matrix()}

        for r in range(self.table.rowCount()):
            for i, d in enumerate(self._span_dates):
                c = 2 + i
                it = self.table.item(r, c)
                if not it:
                    continue
                v = it.text().strip()
                if not v:
                    continue
                key = (d.year, d.month)
                mm = out.setdefault(key, {})
                mm[f"{r},{d.day}"] = v
        return out

    def set_span_month_matrices(self, month_mats: dict[tuple[int, int], dict]) -> None:
        """Populate span grid from per-month matrices."""
        if self._mode != "span" or not self._span_dates:
            # treat as month mode
            self.set_matrix((month_mats or {}).get((self.year, self.month), {}))
            return

        self.clear_matrix()
        if not month_mats:
            self._apply_read_only_flags()
            return

        for r in range(self.table.rowCount()):
            for i, d in enumerate(self._span_dates):
                c = 2 + i
                v = ""
                try:
                    mm = month_mats.get((d.year, d.month)) or {}
                    v = mm.get(f"{r},{d.day}", "")
                except Exception:
                    v = ""
                it = self.table.item(r, c)
                if it is None:
                    it = QTableWidgetItem("")
                    self.table.setItem(r, c, it)
                it.setText(str(v or ""))

        self._apply_read_only_flags()
        self._apply_dynamic_sizes()

    def get_matrix(self) -> dict[str, str]:
        # month mode matrix. In span mode, returns the matrix for self.year/self.month.
        if self._mode == "span":
            mm = self.get_span_month_matrices().get((self.year, self.month), {})
            return dict(mm)

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
