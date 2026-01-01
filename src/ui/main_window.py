from __future__ import annotations

from pathlib import Path
from datetime import time, datetime

from PySide6.QtCore import Qt, QTime, QDate, QEvent
from PySide6.QtGui import QColor, QBrush, QFont, QKeySequence
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QTabWidget, QFileDialog, QMessageBox, QListWidget,
    QDateEdit, QTimeEdit, QGroupBox, QSpinBox, QTableWidget, QTableWidgetItem, QAbstractItemView, QHeaderView, QComboBox, QApplication, QInputDialog
)

from src.settings.app_settings import SettingsService, AppSettings
from src.storage.db import ensure_data_folders, connect_db, migrate_and_seed
from src.storage.repository import Repository
from src.ui.planning_grid import PlanningGrid


from src.domain.models import ReservationDraft, ConfirmedReservation
from src.services.reservation_service import ReservationService
from src.domain.time_rules import classify_dt_odt  # label güncellemek için

import re




TAB_NAMES = [
    "REZERVASYON ve PLANLAMA",
    "PLAN ÖZET",
    "SPOTLİST+",
    "KOD TANIMI",
    "Fiyat ve Kanal Tanımı",
    "DT-ODT",
    "Erişim Örneği",
]

class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Radio-Rez (MVP)")
        self.resize(1200, 700)

        self.settings_service = SettingsService()
        self.app_settings: AppSettings = self.settings_service.build()

        self.repo: Repository | None = None
        self.service: ReservationService | None = None
        self.current_confirmed: ConfirmedReservation | None = None
        
        root = QWidget()
        self.setCentralWidget(root)
        main = QVBoxLayout(root)

        # Top bar
        top = QHBoxLayout()
        main.addLayout(top)

        top.addWidget(QLabel("Reklam veren Ara:"))
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Örn: MURATBEY")
        top.addWidget(self.search_edit, 2)

        self.data_dir_label = QLabel(f"Veri Klasörü: {self.app_settings.data_dir}")
        top.addWidget(self.data_dir_label, 3)

        self.btn_pick_folder = QPushButton("Veri Klasörü Seç")
        top.addWidget(self.btn_pick_folder)

        # Mid: results list
        mid = QHBoxLayout()
        main.addLayout(mid)

        left_box = QGroupBox("Arama Sonuçları (Reklam veren)")
        left_box.setMaximumWidth(280)  # 240-320 arası idealdir
        left_layout = QVBoxLayout(left_box)
        self.list_advertisers = QListWidget()
        left_layout.addWidget(self.list_advertisers)
        mid.addWidget(left_box, 1)

        # Tabs
        self.tabs = QTabWidget()
        mid.addWidget(self.tabs, 6)

        self.tab_widgets: dict[str, QWidget] = {}
        for name in TAB_NAMES:
            w = QWidget()
            self.tab_widgets[name] = w
            self.tabs.addTab(w, name)

        # First tab content (minimal form)
        self._build_first_tab()

        # Bottom buttons
        bottom = QHBoxLayout()
        main.addLayout(bottom)

        self.btn_confirm = QPushButton("Onayla")
        bottom.addWidget(self.btn_confirm)

        bottom.addStretch(1)

        self.btn_test_export = QPushButton("Test Olarak İndir")
        self.btn_test_export.setEnabled(False)
        bottom.addWidget(self.btn_test_export)

        self.btn_save_export = QPushButton("Kaydet ve Çıktı Al")
        self.btn_save_export.setEnabled(False)
        bottom.addWidget(self.btn_save_export)

        # Wire
        self.btn_pick_folder.clicked.connect(self.pick_data_folder)
        self.search_edit.textChanged.connect(self.on_search_changed)
        self.list_advertisers.itemClicked.connect(self.on_advertiser_selected)

        self.btn_confirm.clicked.connect(self.on_confirm)
        self.btn_test_export.clicked.connect(self.on_test_export)
        self.btn_save_export.clicked.connect(self.on_save_export)

        # Bootstrap storage
        self.bootstrap_storage()

        self.tabs.currentChanged.connect(self.on_tab_changed)
        self._access_set_id: int | None = None
        self._build_kod_tanimi_tab()
        self._build_price_channel_tab()
        self._build_access_example_tab()
        


    def _build_first_tab(self) -> None:
        tab = self.tab_widgets["REZERVASYON ve PLANLAMA"]
        layout = QVBoxLayout(tab)

        row0 = QHBoxLayout()
        layout.addLayout(row0)

        row0.addWidget(QLabel("Ajans:"))
        self.in_agency = QLineEdit()
        row0.addWidget(self.in_agency, 2)

        row0.addWidget(QLabel("Ürün:"))
        self.in_product = QLineEdit()
        row0.addWidget(self.in_product, 2)

        row0b = QHBoxLayout()
        layout.addLayout(row0b)

        row0b.addWidget(QLabel("Plan Başlığı:"))
        self.in_plan_title = QLineEdit()
        row0b.addWidget(self.in_plan_title, 2)

        row0b.addWidget(QLabel("Kod:"))
        self.in_spot_code = QLineEdit()
        row0b.addWidget(self.in_spot_code, 1)

        row0b.addWidget(QLabel("Süre (sn):"))
        self.in_spot_duration = QSpinBox()
        self.in_spot_duration.setRange(0, 9999)
        row0b.addWidget(self.in_spot_duration, 1)

        row_code_def = QHBoxLayout()
        layout.addLayout(row_code_def)

        row_code_def.addWidget(QLabel("Kod Tanımı:"))
        self.in_code_definition = QLineEdit()
        row_code_def.addWidget(self.in_code_definition, 6)


        row0c = QHBoxLayout()
        layout.addLayout(row0c)

        row0c.addWidget(QLabel("Not:"))
        self.in_note = QLineEdit()
        row0c.addWidget(self.in_note, 4)

        row0c.addWidget(QLabel("Formu Oluşturan:"))
        self.in_prepared_by = QLineEdit()
        row0c.addWidget(self.in_prepared_by, 2)
        row1 = QHBoxLayout()
        layout.addLayout(row1)
        
        row1.addWidget(QLabel("Reklam veren:"))
        self.in_advertiser = QLineEdit()
        row1.addWidget(self.in_advertiser, 2)

        row2 = QHBoxLayout()
        layout.addLayout(row2)
        row2.addWidget(QLabel("Plan Tarihi:"))
        self.in_date = QDateEdit()
        self.in_date.setCalendarPopup(True)
        self.in_date.setDate(QDate.currentDate())
        self.in_date.setFixedWidth(120)
        row2.addWidget(self.in_date)

        row2.addWidget(QLabel("Spot Saat:"))
        self.in_time = QTimeEdit()
        self.in_time.setTime(QTime.currentTime())
        self.in_time.setFixedWidth(80)
        row2.addWidget(self.in_time)

        self.lbl_dt_odt = QLabel("DT/ODT: -")
        self.lbl_dt_odt.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.lbl_dt_odt.setMinimumWidth(80)
        row2.addWidget(self.lbl_dt_odt)
        row2.addStretch(1)

        self.in_time.timeChanged.connect(self.on_time_changed)
        self.on_time_changed(self.in_time.time())
        # --- Excel benzeri plan grid ---
        self.plan_grid = PlanningGrid()
        layout.addWidget(self.plan_grid, 1)

        # Tarih değişince grid ay/gün vurgusunu güncelle
        self.in_date.dateChanged.connect(self.on_plan_date_changed)

        # ilk açılışta da set et
        self.on_plan_date_changed(self.in_date.date())       

    def on_time_changed(self, qt: QTime) -> None:
        t = time(qt.hour(), qt.minute(), qt.second())
        self.lbl_dt_odt.setText(f"DT/ODT: {classify_dt_odt(t)}")

    def bootstrap_storage(self) -> None:
        ensure_data_folders(self.app_settings.data_dir)
        db_path = self.app_settings.data_dir / "data.db"
        conn = connect_db(db_path)
        migrate_and_seed(conn)
        self.repo = Repository(conn)
        self.service = ReservationService(self.repo)

    def pick_data_folder(self) -> None:
        p = QFileDialog.getExistingDirectory(self, "Veri klasörünü seç")
        if not p:
            return
        data_dir = Path(p)
        self.settings_service.set_data_dir(data_dir)
        self.app_settings = self.settings_service.build()
        self.data_dir_label.setText(f"Veri Klasörü: {self.app_settings.data_dir}")

        # Re-bootstrap
        self.bootstrap_storage()
        QMessageBox.information(self, "OK", "Veri klasörü kaydedildi ve DB hazırlandı.")

    def on_search_changed(self, text: str) -> None:
        self.list_advertisers.clear()
        if not self.repo:
            return
        for name in self.repo.search_advertisers(text, limit=30):
            self.list_advertisers.addItem(name)

    def on_advertiser_selected(self, item) -> None:
        if not item:
            return
        self.in_advertiser.setText(item.text())

    def on_plan_date_changed(self, qd: QDate) -> None:
        d = qd.toPython()
        self.plan_grid.set_month(d.year, d.month, d.day)

    def on_confirm(self) -> None:
        if not self.service:
            QMessageBox.warning(self, "Hata", "Servis hazır değil (DB bağlantısı yok).")
            return

        try:
            qt = self.in_time.time()
            spot_t = time(qt.hour(), qt.minute(), qt.second())

            draft = ReservationDraft(
                advertiser_name=self.in_advertiser.text(),
                plan_date=self.in_date.date().toPython(),
                spot_time=spot_t,
                agency_name=self.in_agency.text(),
                product_name=self.in_product.text(),
                plan_title=self.in_plan_title.text(),
                spot_code=self.in_spot_code.text(),
                spot_duration_sec=int(self.in_spot_duration.value()),
                code_definition=self.in_code_definition.text(),
                note_text=self.in_note.text(),
                prepared_by_name=self.in_prepared_by.text(),
            )

            # ✅ Grid verisi buradan alınacak
            plan_cells = self.plan_grid.get_matrix()

            self.current_confirmed = self.service.confirm(draft, plan_cells)

            self.btn_test_export.setEnabled(True)
            self.btn_save_export.setEnabled(True)

            QMessageBox.information(self, "OK", "Onaylandı. Artık test/kayıt çıktısı alabilirsin.")
        except Exception as e:
            QMessageBox.warning(self, "Hata", str(e))


    def _resolve_template_path(self) -> Path:
        # AppSettings içinde template_path varsa onu kullan, yoksa assets/template.xlsx'e düş
        tp = getattr(self.app_settings, "template_path", None)
        if tp:
            return Path(tp)
        return Path("assets") / "template.xlsx"


    def on_test_export(self) -> None:
        if not self.service or not self.current_confirmed:
            QMessageBox.warning(self, "Hata", "Önce Onayla.")
            return

        try:
            template_path = self._resolve_template_path()
            out_dir = self.app_settings.data_dir / "exports"
            out_path = self.service.export_test(template_path, out_dir, self.current_confirmed)

            QMessageBox.information(self, "OK", f"Test çıktısı üretildi:\n{out_path}")
        except Exception as e:
            QMessageBox.critical(self, "Hata", str(e))


    def on_save_export(self) -> None:
        if not self.service or not self.current_confirmed:
            QMessageBox.warning(self, "Hata", "Önce Onayla.")
            return

        try:
            template_path = self._resolve_template_path()
            out_dir = self.app_settings.data_dir / "exports"
            out_path = self.service.save_and_export(template_path, out_dir, self.current_confirmed)

            QMessageBox.information(self, "OK", f"Kaydedildi ve çıktı alındı:\n{out_path}")
        except Exception as e:
            QMessageBox.critical(self, "Hata", str(e))

    def on_tab_changed(self, idx: int) -> None:
        tab_name = self.tabs.tabText(idx)
        if tab_name == "KOD TANIMI":
            self.refresh_kod_tanimi()
        elif tab_name == "Fiyat ve Kanal Tanımı":
            self.refresh_price_channel_tab()

    def _build_kod_tanimi_tab(self) -> None:
        tab = self.tab_widgets["KOD TANIMI"]
        layout = QVBoxLayout(tab)

        btn_row = QHBoxLayout()
        layout.addLayout(btn_row)

        self.btn_kod_refresh = QPushButton("Yenile")
        self.btn_kod_delete = QPushButton("Seçili Kodu Sil")
        self.btn_kod_export = QPushButton("Excel Çıktısı")

        btn_row.addWidget(self.btn_kod_refresh)
        btn_row.addWidget(self.btn_kod_delete)
        btn_row.addWidget(self.btn_kod_export)
        btn_row.addStretch(1)

        self.kod_table = QTableWidget()
        self.kod_table.setColumnCount(4)
        self.kod_table.setHorizontalHeaderLabels(["Kod", "Kod Tanımı", "Kod Uzunluğu (SN)", "Dağılım"])
        # Görsel stil (Excel'e yakın)
        self.kod_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.kod_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.kod_table.setAlternatingRowColors(True)
        self.kod_table.verticalHeader().setVisible(False)
        self.kod_table.setShowGrid(True)

        hdr = self.kod_table.horizontalHeader()
        hdr.setStretchLastSection(False)
        hdr.setSectionResizeMode(0, QHeaderView.ResizeToContents)  # Kod
        hdr.setSectionResizeMode(1, QHeaderView.Stretch)           # Kod Tanımı
        hdr.setSectionResizeMode(2, QHeaderView.ResizeToContents)  # Uzunluk
        hdr.setSectionResizeMode(3, QHeaderView.ResizeToContents)  # Dağılım

        self.kod_table.setStyleSheet("""
            QHeaderView::section {
                background-color: #F28C28;
                color: white;
                font-weight: bold;
                padding: 6px;
                border: 1px solid #B56A1E;
            }
            QTableWidget {
                gridline-color: #A0A0A0;
                selection-background-color: #CFE8FF;
            }
        """)

        layout.addWidget(self.kod_table, 1)

        self.btn_kod_refresh.clicked.connect(self.refresh_kod_tanimi)
        self.btn_kod_delete.clicked.connect(self.delete_selected_kod)
        self.btn_kod_export.clicked.connect(self.export_kod_tanimi_excel)

    def refresh_kod_tanimi(self) -> None:
        if not self.service:
            return
        adv = self.in_advertiser.text().strip()
        if not adv:
            return

        rows = self.service.get_kod_tanimi_rows(adv)
        avg_len = self.service.get_kod_tanimi_avg_len(adv)

        data_count = max(len(rows), 7)
        self.kod_table.setRowCount(data_count + 1)

        # Veri satırları
        for i, r in enumerate(rows):
            it0 = QTableWidgetItem(r["code"])
            it0.setTextAlignment(Qt.AlignCenter)
            self.kod_table.setItem(i, 0, it0)

            it1 = QTableWidgetItem(r["code_desc"])
            f_italic = QFont()
            f_italic.setItalic(True)
            it1.setFont(f_italic)
            self.kod_table.setItem(i, 1, it1)

            it2 = QTableWidgetItem(str(int(r["length_sn"])))
            it2.setTextAlignment(Qt.AlignCenter)
            self.kod_table.setItem(i, 2, it2)

            it3 = QTableWidgetItem(f"{r['distribution']:.0%}")
            it3.setTextAlignment(Qt.AlignCenter)
            self.kod_table.setItem(i, 3, it3)

        # Şablon gibi 7 satıra kadar boş satır göster
        for rr in range(len(rows), data_count):
            for cc in range(4):
                it = QTableWidgetItem("")
                it.setTextAlignment(Qt.AlignCenter if cc != 1 else Qt.AlignLeft | Qt.AlignVCenter)
                self.kod_table.setItem(rr, cc, it)

        # Toplam / Ortalama satırı
        last = data_count
        f_bi = QFont()
        f_bi.setBold(True)
        f_bi.setItalic(True)

        it0 = QTableWidgetItem("Ort.Uzun.")
        it0.setFont(f_bi)
        self.kod_table.setItem(last, 0, it0)

        it2 = QTableWidgetItem(f"{avg_len:.2f}")
        it2.setTextAlignment(Qt.AlignCenter)
        it2.setFont(f_bi)
        self.kod_table.setItem(last, 2, it2)

        it3 = QTableWidgetItem(f"{sum(r['distribution'] for r in rows):.0%}")
        it3.setTextAlignment(Qt.AlignCenter)
        it3.setFont(f_bi)
        it3.setBackground(QBrush(QColor("#8BC34A")))
        self.kod_table.setItem(last, 3, it3)

    def delete_selected_kod(self) -> None:
        if not self.service:
            return
        adv = self.in_advertiser.text().strip()
        row = self.kod_table.currentRow()
        if row < 0:
            return
        code_item = self.kod_table.item(row, 0)
        if not code_item:
            return
        code = code_item.text().strip()
        if not code or code == "Ort.Uzun.":
            return

        deleted = self.service.delete_kod_for_advertiser(adv, code)
        QMessageBox.information(self, "OK", f"{code} koduna ait {deleted} kayıt silindi.")
        self.refresh_kod_tanimi()

    def export_kod_tanimi_excel(self) -> None:
        if not self.service:
            return
        adv = self.in_advertiser.text().strip()
        if not adv:
            return

        path, _ = QFileDialog.getSaveFileName(self, "KOD TANIMI Excel", f"{adv}_KOD_TANIMI.xlsx", "Excel Files (*.xlsx)")
        if not path:
            return

        self.service.export_kod_tanimi_excel(path, adv)
        QMessageBox.information(self, "OK", f"Excel çıktısı oluşturuldu:\n{path}")

    # ------------------------------
    # Fiyat ve Kanal Tanımı
    # ------------------------------
    def _build_price_channel_tab(self) -> None:
        tab = self.tab_widgets["Fiyat ve Kanal Tanımı"]
        layout = QVBoxLayout(tab)

        top = QHBoxLayout()
        layout.addLayout(top)

        top.addWidget(QLabel("Yıl:"))
        self.price_year = QSpinBox()
        self.price_year.setRange(2000, 2100)
        self.price_year.setValue(datetime.now().year)
        top.addWidget(self.price_year)

        self.btn_price_refresh = QPushButton("Yenile")
        self.btn_price_save = QPushButton("Kaydet")
        self.btn_channel_add = QPushButton("Kanal Ekle")
        self.btn_channel_delete = QPushButton("Seçili Kanalı Sil")

        top.addStretch(1)
        top.addWidget(self.btn_channel_add)
        top.addWidget(self.btn_channel_delete)
        top.addWidget(self.btn_price_refresh)
        top.addWidget(self.btn_price_save)

        self.price_table = QTableWidget()
        self.price_table.setAlternatingRowColors(True)
        self.price_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.price_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.price_table.setEditTriggers(
            QAbstractItemView.DoubleClicked
            | QAbstractItemView.EditKeyPressed
            | QAbstractItemView.AnyKeyPressed
        )
        layout.addWidget(self.price_table, 1)

        # Wire
        self.btn_price_refresh.clicked.connect(self.refresh_price_channel_tab)
        self.price_year.valueChanged.connect(lambda _v: self.refresh_price_channel_tab())
        self.btn_price_save.clicked.connect(self.save_price_channel_tab)
        self.btn_channel_add.clicked.connect(self.add_channel_dialog)
        self.btn_channel_delete.clicked.connect(self.delete_selected_channel)

    def _month_names_tr(self) -> list[str]:
        return ["OCAK", "ŞUBAT", "MART", "NİSAN", "MAYIS", "HAZİRAN", "TEMMUZ", "AĞUSTOS", "EYLÜL", "EKİM", "KASIM", "ARALIK"]

    def refresh_price_channel_tab(self) -> None:
        if not self.repo:
            return

        # Meta'dan son seçilen yılı çek
        meta_year = self.repo.get_meta("price_year")
        if meta_year and meta_year.isdigit():
            my = int(meta_year)
            if self.price_year.value() != my:
                self.price_year.blockSignals(True)
                self.price_year.setValue(my)
                self.price_year.blockSignals(False)

        year = int(self.price_year.value())
        try:
            self.repo.set_meta("price_year", str(year))
        except Exception:
            pass

        months = self._month_names_tr()

        headers = ["KANAL"]
        for mn in months:
            headers.append(f"{mn}\nDT")
            headers.append(f"{mn}\nODT")

        self.price_table.clear()
        self.price_table.setColumnCount(len(headers))
        self.price_table.setHorizontalHeaderLabels(headers)

        self.price_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        for c in range(1, len(headers)):
            self.price_table.horizontalHeader().setSectionResizeMode(c, QHeaderView.ResizeToContents)

        channels = self.repo.list_channels(active_only=True)
        prices = self.repo.get_channel_prices(year)

        self.price_table.setRowCount(len(channels))

        for r, ch in enumerate(channels):
            cid = int(ch["id"])
            name = str(ch["name"])

            it_name = QTableWidgetItem(name)
            it_name.setData(Qt.UserRole, cid)
            self.price_table.setItem(r, 0, it_name)

            col = 1
            for m in range(1, 13):
                dt, odt = prices.get((cid, m), (0.0, 0.0))

                it_dt = QTableWidgetItem("" if dt == 0 else f"{dt:g}")
                it_dt.setTextAlignment(Qt.AlignCenter)
                self.price_table.setItem(r, col, it_dt)
                col += 1

                it_odt = QTableWidgetItem("" if odt == 0 else f"{odt:g}")
                it_odt.setTextAlignment(Qt.AlignCenter)
                self.price_table.setItem(r, col, it_odt)
                col += 1

    def _parse_float_cell(self, item: QTableWidgetItem | None) -> float:
        if not item:
            return 0.0
        t = (item.text() or "").strip()
        if not t:
            return 0.0
        t = t.replace(",", ".")
        try:
            return float(t)
        except ValueError:
            return 0.0

    def save_price_channel_tab(self) -> None:
        if not self.repo:
            QMessageBox.warning(self, "Hata", "DB bağlantısı yok.")
            return

        year = int(self.price_year.value())

        try:
            for r in range(self.price_table.rowCount()):
                it_name = self.price_table.item(r, 0)
                name = (it_name.text() if it_name else "").strip()
                if not name:
                    continue

                cid = it_name.data(Qt.UserRole) if it_name else None
                if cid:
                    self.repo.update_channel_name(int(cid), name)
                    channel_id = int(cid)
                else:
                    channel_id = self.repo.get_or_create_channel(name)
                    it_name.setData(Qt.UserRole, channel_id)

                col = 1
                for m in range(1, 13):
                    price_dt = self._parse_float_cell(self.price_table.item(r, col))
                    price_odt = self._parse_float_cell(self.price_table.item(r, col + 1))
                    self.repo.upsert_channel_price(year, m, channel_id, price_dt, price_odt)
                    col += 2

            self.repo.set_meta("price_year", str(year))
            QMessageBox.information(self, "Tamam", "Fiyatlar kaydedildi.")
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Kayıt sırasında hata: {e}")

    def add_channel_dialog(self) -> None:
        if not self.repo:
            QMessageBox.warning(self, "Hata", "DB bağlantısı yok.")
            return

        name, ok = QInputDialog.getText(self, "Kanal Ekle", "Kanal adı:")
        if not ok:
            return
        name = (name or "").strip()
        if not name:
            return

        try:
            self.repo.get_or_create_channel(name)
            self.refresh_price_channel_tab()
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Kanal eklenemedi: {e}")

    def delete_selected_channel(self) -> None:
        if not self.repo:
            QMessageBox.warning(self, "Hata", "DB bağlantısı yok.")
            return

        r = self.price_table.currentRow()
        if r < 0:
            QMessageBox.information(self, "Bilgi", "Silmek için bir kanal seç.")
            return

        it = self.price_table.item(r, 0)
        if not it:
            return

        cid = it.data(Qt.UserRole)
        name = it.text()

        if not cid:
            self.price_table.removeRow(r)
            return

        ans = QMessageBox.question(
            self,
            "Onay",
            f"'{name}' kanalı silinsin mi? (DB'de pasif yapılacak)",
            QMessageBox.Yes | QMessageBox.No,
        )
        if ans != QMessageBox.Yes:
            return

        try:
            self.repo.deactivate_channel(int(cid))
            self.refresh_price_channel_tab()
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Silinemedi: {e}")


    def _build_access_example_tab(self) -> None:
        tab = self.tab_widgets["Erişim Örneği"]
        layout = QVBoxLayout(tab)

        top = QHBoxLayout()
        layout.addLayout(top)

        top.addWidget(QLabel("Dates >>"))
        self.access_dates = QLineEdit()
        self.access_dates.setPlaceholderText("Örn: Ekim 2025")
        top.addWidget(self.access_dates, 2)

        top.addWidget(QLabel("Targets >>"))
        self.access_targets = QLineEdit()
        self.access_targets.setPlaceholderText("Örn: 12+(1)")
        top.addWidget(self.access_targets, 1)

        self.btn_access_save = QPushButton("Kaydet")
        top.addWidget(self.btn_access_save)

        btns = QHBoxLayout()
        layout.addLayout(btns)

        self.btn_access_add = QPushButton("Kanal Ekle")
        self.btn_access_del = QPushButton("Seçili Kanalı Sil")
        self.btn_access_paste = QPushButton("Excel'den Yapıştır (Ctrl+V)")
        btns.addWidget(self.btn_access_add)
        btns.addWidget(self.btn_access_del)
        btns.addWidget(self.btn_access_paste)
        btns.addStretch(1)

        self.access_table = QTableWidget()
        self.access_table.setColumnCount(4)
        self.access_table.setHorizontalHeaderLabels(["Channels", "Universe", "AvRch(000)", "AvRch%"])
        self.access_table.setAlternatingRowColors(True)
        self.access_table.verticalHeader().setVisible(False)

        hdr = self.access_table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.Stretch)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeToContents)

        self.access_table.setStyleSheet("""
            QHeaderView::section { background-color: #F28C28; color: white; font-weight: bold; padding: 6px; }
            QTableWidget { gridline-color: #A0A0A0; selection-background-color: #CFE8FF; }
        """)

        layout.addWidget(self.access_table, 1)

        # Ctrl+V doğrudan tabloda çalışsın
        self.access_table.installEventFilter(self)

        # events
        self.btn_access_add.clicked.connect(self.access_add_row)
        self.btn_access_del.clicked.connect(self.access_delete_selected_row)
        self.btn_access_paste.clicked.connect(self.access_paste_from_clipboard)
        self.btn_access_save.clicked.connect(self.access_save)
        self.access_load_latest_db()



    def access_open_or_create(self) -> None:
        if not self.repo:
            return

        year = int(self.access_year.value())
        label = str(year)
        label = (self.access_label.text() or "").strip() or f"{year}"
        periods = self.access_periods.text() or ""
        targets = self.access_targets.text() or ""

        set_id = self.repo.get_or_create_access_set(year, label, periods=periods, targets=targets)
        self._access_set_id = set_id

        meta, rows = self.repo.load_access_set(set_id)

        self.access_label.setText(meta.get("label") or "")
        self.access_periods.setText(meta.get("periods") or "")
        self.access_targets.setText(meta.get("targets") or "")

        self.access_table.setRowCount(max(len(rows), 10))

        for i, r in enumerate(rows):
            self._set_access_cell(i, 0, r.get("channel"), align_left=True)
            self._set_access_cell(i, 1, r.get("universe"), center=True)
            self._set_access_cell(i, 2, r.get("avrch000"), center=True)
            self._set_access_cell(i, 3, r.get("avrch_pct"), center=True)

        # kalan boş satırlar
        for rr in range(len(rows), self.access_table.rowCount()):
            self._set_access_cell(rr, 0, "", align_left=True)
            self._set_access_cell(rr, 1, "", center=True)
            self._set_access_cell(rr, 2, "", center=True)
            self._set_access_cell(rr, 3, "", center=True)

    def _set_access_cell(self, r: int, c: int, value, center: bool = False, align_left: bool = False) -> None:
        txt = "" if value is None else str(value)
        it = QTableWidgetItem(txt)
        if center:
            it.setTextAlignment(Qt.AlignCenter)
        elif align_left:
            it.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.access_table.setItem(r, c, it)

    def access_add_row(self) -> None:
        r = self.access_table.rowCount()
        self.access_table.setRowCount(r + 1)

    def access_delete_selected_row(self) -> None:
        row = self.access_table.currentRow()
        if row >= 0:
            self.access_table.removeRow(row)


    def access_save(self) -> None:
        if not self.repo:
            QMessageBox.critical(self, "Hata", "Repo yok. Veri klasörü/DB bağlantısı kurulmamış.")
            return

        dates = (self.access_dates.text() or "").strip()
        targets = (self.access_targets.text() or "").strip()
        periods = ""  # UI'da periods alanı yok -> DB'ye boş yaz

        year = self._parse_year_from_dates(dates)
        label = dates if dates else f"{year}"

        # İlk kayıt: set yoksa otomatik oluştur
        if not self._access_set_id:
            self._access_set_id = self.repo.get_or_create_access_set(
                year=year, label=label, periods=periods, targets=targets
            )

        def _to_int(col: int):
            t = self.access_table.item(r, col).text().strip() if self.access_table.item(r, col) else ""
            return int(t) if t else None

        def _to_float(col: int):
            t = self.access_table.item(r, col).text().strip() if self.access_table.item(r, col) else ""
            return float(t.replace(",", ".")) if t else None

        rows: list[dict] = []
        for r in range(self.access_table.rowCount()):
            ch = (self.access_table.item(r, 0).text().strip() if self.access_table.item(r, 0) else "")
            if not ch:
                continue

            universe = _to_int(1)
            avrch000 = _to_int(2)
            avpct = _to_float(3)

            # AvRch% boşsa hesapla
            if avpct is None and universe and avrch000 is not None and universe != 0:
                avpct = round((avrch000 / universe) * 100, 2)

            rows.append({"channel": ch, "universe": universe, "avrch000": avrch000, "avrch_pct": avpct})

        self.repo.save_access_set(self._access_set_id, periods=periods, targets=targets, rows=rows)
        QMessageBox.information(self, "OK", "Erişim örneği DB'ye kaydedildi. Uygulama yeniden açılınca aynen gelecektir.")

    
    def access_save_db(self) -> None:
        self.access_save()


    def access_paste_from_clipboard(self) -> None:
        try:
            text = (QApplication.clipboard().text() or "").strip()
            if not text:
                QMessageBox.warning(self, "Uyarı", "Panoda veri yok. Excel'de alanı kopyalayıp tekrar dene.")
                return

            lines = [ln for ln in text.splitlines() if ln.strip()]
            if not lines:
                QMessageBox.warning(self, "Uyarı", "Panodaki format okunamadı.")
                return

            # Header satırı gelirse atla
            first = lines[0].lower()
            if "channels" in first and ("universe" in first or "avrch" in first):
                lines = lines[1:]

            def norm_num(s: str) -> str:
                s = (s or "").strip()
                # 52,47 -> 52.47
                s = s.replace(",", ".")
                return s

            # Nereden başlayalım? Seçim yoksa 0
            start_row = self.access_table.currentRow()
            if start_row < 0:
                start_row = 0

            need_rows = start_row + len(lines)
            if self.access_table.rowCount() < need_rows:
                self.access_table.setRowCount(need_rows)

            self.access_table.blockSignals(True)

            for i, ln in enumerate(lines):
                parts = ln.split("\t")
                # Excel bazen sonda tab bırakır
                parts = [p.strip() for p in parts]

                # Beklenen 4 kolon; fazlası gelirse ilk 4'ü al
                while len(parts) < 4:
                    parts.append("")
                parts = parts[:4]

                r = start_row + i
                ch, uni, av0, avp = parts

                self.access_table.setItem(r, 0, QTableWidgetItem(ch))
                self.access_table.setItem(r, 1, QTableWidgetItem(norm_num(uni)))
                self.access_table.setItem(r, 2, QTableWidgetItem(norm_num(av0)))
                self.access_table.setItem(r, 3, QTableWidgetItem(norm_num(avp)))

                # hizalama
                for c in (1, 2, 3):
                    it = self.access_table.item(r, c)
                    it.setTextAlignment(Qt.AlignCenter)

            self.access_table.blockSignals(False)

        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Yapıştırma sırasında hata: {e}")

    def norm_num(s: str) -> str:
        s = (s or "").strip()
        # 52,47 -> 52.47
        s = s.replace(".", "") if s.count(".") > 0 and s.count(",") == 1 else s
        s = s.replace(",", ".")
        return s

        for i, ln in enumerate(lines):
            parts = ln.split("\t")
            # Beklenen: Channels | Universe | AvRch(000) | AvRch%
            while len(parts) < 4:
                parts.append("")

            r = start_row + i
            ch = parts[0].strip()
            uni = parts[1].strip()
            av0 = parts[2].strip()
            avp = parts[3].strip()

            self._set_access_cell(r, 0, ch, align_left=True)
            self._set_access_cell(r, 1, norm_num(uni), center=True)
            self._set_access_cell(r, 2, norm_num(av0), center=True)
            self._set_access_cell(r, 3, norm_num(avp), center=True)

    def eventFilter(self, obj, event):
        if obj is getattr(self, "access_table", None) and event.type() == QEvent.KeyPress:
            if event.matches(QKeySequence.Paste):
                self.access_paste_from_clipboard()
                return True
        return super().eventFilter(obj, event)

    def access_save_to_file(self) -> None:
        try:
            data_dir = self.app_settings.data_dir
            data_dir.mkdir(parents=True, exist_ok=True)
            path = data_dir / "access_example.json"

            dates = (self.access_dates.text() or "").strip()
            targets = (self.access_targets.text() or "").strip()

            rows = []
            for r in range(self.access_table.rowCount()):
                ch = self.access_table.item(r, 0).text().strip() if self.access_table.item(r, 0) else ""
                if not ch:
                    continue

                def get(col):
                    return self.access_table.item(r, col).text().strip() if self.access_table.item(r, col) else ""

                rows.append({
                    "channel": ch,
                    "universe": get(1),
                    "avrch000": get(2),
                    "avrch_pct": get(3),
                })

            payload = {
                "dates": dates,
                "targets": targets,
                "rows": rows,
            }

            path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
            QMessageBox.information(self, "OK", "Erişim örneği kaydedildi (access_example.json).")

        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Kaydetme hatası: {e}")

    def access_load_from_file(self) -> None:
        data_dir = self.app_settings.data_dir
        path = data_dir / "access_example.json"

        # default
        if not path.exists():
            # ilk açılış için tarih otomatik bas
            # (istersen boş da bırakırız)
            self.access_dates.setText(QDate.currentDate().toString("MMMM yyyy"))
            self.access_targets.setText("")
            self.access_table.setRowCount(30)
            return

        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
            self.access_dates.setText(payload.get("dates", ""))
            self.access_targets.setText(payload.get("targets", ""))

            rows = payload.get("rows", []) or []
            self.access_table.setRowCount(max(len(rows), 30))

            for i, r in enumerate(rows):
                self.access_table.setItem(i, 0, QTableWidgetItem(str(r.get("channel",""))))
                self.access_table.setItem(i, 1, QTableWidgetItem(str(r.get("universe",""))))
                self.access_table.setItem(i, 2, QTableWidgetItem(str(r.get("avrch000",""))))
                self.access_table.setItem(i, 3, QTableWidgetItem(str(r.get("avrch_pct",""))))
                for c in (1,2,3):
                    self.access_table.item(i,c).setTextAlignment(Qt.AlignCenter)

        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Erişim verisi okunamadı: {e}")
            self.access_table.setRowCount(30)
    def _parse_year_from_dates(self, dates: str) -> int:
        m = re.search(r"(19\d{2}|20\d{2})", dates or "")
        return int(m.group(1)) if m else datetime.now().year

    def access_load_latest_db(self) -> None:
        if not self.repo:
            return

        latest_id = self.repo.get_latest_access_set_id()
        if not latest_id:
            # ilk kullanım: boş şablon
            self.access_dates.setText(QDate.currentDate().toString("MMMM yyyy"))
            self.access_targets.setText("")
            self.access_table.setRowCount(30)
            self._access_set_id = None
            return

        meta, rows = self.repo.load_access_set(latest_id)
        self._access_set_id = latest_id

        # label = Dates >> gibi kullanacağız
        self.access_dates.setText(meta.get("label") or "")
        self.access_targets.setText(meta.get("targets") or "")

        self.access_table.setRowCount(max(len(rows), 30))

        for i, r in enumerate(rows):
            self.access_table.setItem(i, 0, QTableWidgetItem(str(r.get("channel", ""))))
            self.access_table.setItem(i, 1, QTableWidgetItem("" if r.get("universe") is None else str(r.get("universe"))))
            self.access_table.setItem(i, 2, QTableWidgetItem("" if r.get("avrch000") is None else str(r.get("avrch000"))))
            self.access_table.setItem(i, 3, QTableWidgetItem("" if r.get("avrch_pct") is None else str(r.get("avrch_pct"))))
            for c in (1, 2, 3):
                it = self.access_table.item(i, c)
                if it:
                    it.setTextAlignment(Qt.AlignCenter)