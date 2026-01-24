from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QKeyEvent, QKeySequence
from PySide6.QtWidgets import QApplication, QTableWidget, QTableWidgetItem


class ExcelTableWidget(QTableWidget):
    """QTableWidget with Excel-like clipboard behavior.

    Supported:
    - Copy: Ctrl+C / Ctrl+Insert
    - Paste: Ctrl+V / Shift+Insert
    - Cut: Ctrl+X
    - Clear selection: Delete / Backspace

    Notes
    - Copy uses TSV (tab-separated values) so pasting to/from Excel works.
    - Paste starts from the current cell. If a range is selected, its top-left is used.
    - Non-editable cells are skipped on paste/clear (e.g., locked columns).
    """

    def keyPressEvent(self, event: QKeyEvent) -> None:
        if event.matches(QKeySequence.Copy):
            self.copy_selection()
            event.accept()
            return
        if event.matches(QKeySequence.Cut):
            self.cut_selection()
            event.accept()
            return
        if event.matches(QKeySequence.Paste):
            self.paste_from_clipboard()
            event.accept()
            return

        if event.key() in (Qt.Key_Delete, Qt.Key_Backspace):
            self.clear_selection()
            event.accept()
            return

        super().keyPressEvent(event)

    def _selection_rect(self) -> tuple[int, int, int, int] | None:
        """Return (topRow, leftCol, bottomRow, rightCol) or None."""
        ranges = self.selectedRanges()
        if ranges:
            r = ranges[0]
            return r.topRow(), r.leftColumn(), r.bottomRow(), r.rightColumn()

        cr = self.currentRow()
        cc = self.currentColumn()
        if cr < 0 or cc < 0:
            return None
        return cr, cc, cr, cc

    def copy_selection(self) -> None:
        rect = self._selection_rect()
        if not rect:
            return
        top, left, bottom, right = rect

        lines: list[str] = []
        for r in range(top, bottom + 1):
            row_vals: list[str] = []
            for c in range(left, right + 1):
                it = self.item(r, c)
                row_vals.append(it.text() if it else "")
            lines.append("\t".join(row_vals))

        QApplication.clipboard().setText("\n".join(lines))

    def cut_selection(self) -> None:
        self.copy_selection()
        self.clear_selection()

    def clear_selection(self) -> None:
        rect = self._selection_rect()
        if not rect:
            return
        top, left, bottom, right = rect

        for r in range(top, bottom + 1):
            for c in range(left, right + 1):
                it = self.item(r, c)
                if not it:
                    continue
                if not (it.flags() & Qt.ItemIsEditable):
                    continue
                it.setText("")

    def paste_from_clipboard(self) -> None:
        text = QApplication.clipboard().text()
        if not text:
            return

        # normalize line endings and strip trailing empty lines
        rows = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
        while rows and rows[-1] == "":
            rows.pop()
        if not rows:
            return

        grid = [r.split("\t") for r in rows]

        rect = self._selection_rect()
        if rect:
            start_row, start_col = rect[0], rect[1]
        else:
            start_row, start_col = self.currentRow(), self.currentColumn()

        if start_row < 0 or start_col < 0:
            return

        max_r = self.rowCount()
        max_c = self.columnCount()

        for r_off, row_vals in enumerate(grid):
            r = start_row + r_off
            if r >= max_r:
                break
            for c_off, val in enumerate(row_vals):
                c = start_col + c_off
                if c >= max_c:
                    break

                it = self.item(r, c)
                if it is None:
                    it = QTableWidgetItem("")
                    self.setItem(r, c, it)
                if not (it.flags() & Qt.ItemIsEditable):
                    continue
                it.setText(val)
