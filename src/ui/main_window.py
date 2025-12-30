from __future__ import annotations

from pathlib import Path
from datetime import time

from PySide6.QtCore import Qt, QTime, QDate
from PySide6.QtGui import QColor, QBrush, QFont
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QTabWidget, QFileDialog, QMessageBox, QListWidget,
    QDateEdit, QTimeEdit, QGroupBox, QSpinBox, QTableWidget, QTableWidgetItem, QAbstractItemView, QHeaderView
)

from src.settings.app_settings import SettingsService, AppSettings
from src.storage.db import ensure_data_folders, connect_db, migrate_and_seed
from src.storage.repository import Repository
from src.ui.planning_grid import PlanningGrid


from src.domain.models import ReservationDraft, ConfirmedReservation
from src.services.reservation_service import ReservationService
from src.domain.time_rules import classify_dt_odt  # label güncellemek için




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
        self._build_kod_tanimi_tab()
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
        if self.tabs.tabText(idx) == "KOD TANIMI":
            self.refresh_kod_tanimi()

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
