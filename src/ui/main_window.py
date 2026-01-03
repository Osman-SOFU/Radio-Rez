from __future__ import annotations

from pathlib import Path
from datetime import time, datetime, date

from PySide6.QtCore import Qt, QDate, QEvent
from PySide6.QtGui import QColor, QBrush, QFont, QKeySequence
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QTabWidget, QFileDialog, QMessageBox, QListWidget,
    QDateEdit, QGroupBox, QSpinBox, QTableWidget, QTableWidgetItem, QAbstractItemView, QAbstractItemDelegate,
    QHeaderView, QComboBox, QApplication, QInputDialog, QPlainTextEdit
)

from src.settings.app_settings import SettingsService, AppSettings
from src.storage.db import ensure_data_folders, connect_db, migrate_and_seed
from src.storage.repository import Repository
from src.ui.planning_grid import PlanningGrid


from src.domain.models import ReservationDraft, ConfirmedReservation
from src.services.reservation_service import ReservationService
from src.domain.time_rules import classify_dt_odt  # exporter tarafında kullanılıyor; burada sadece ortak import

import re




TAB_NAMES = [
    "ANA SAYFA",
    "REZERVASYONLAR",
    "PLAN ÖZET",
    "SPOTLİST+",
    "KOD TANIMI",
    "Fiyat ve Kanal Tanımı",
    "Erişim Örneği",
]

MONTHS_TR = [
    "OCAK",
    "ŞUBAT",
    "MART",
    "NİSAN",
    "MAYIS",
    "HAZİRAN",
    "TEMMUZ",
    "AĞUSTOS",
    "EYLÜL",
    "EKİM",
    "KASIM",
    "ARALIK",
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

        # ANA SAYFA (rezervasyon girişi)
        self._build_home_tab()

        # REZERVASYONLAR (kayıtlı rezervasyonları gör/sil)
        self._build_reservations_tab()

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
        self._po_syncing: bool = False
        self._build_spotlist_tab()
        self._build_plan_ozet_tab()
        self._build_kod_tanimi_tab()
        self._build_price_channel_tab()
        self._build_access_example_tab()
        


    def _build_home_tab(self) -> None:
        tab = self.tab_widgets["ANA SAYFA"]
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

        row2.addSpacing(12)
        row2.addWidget(QLabel("Kanal:"))
        self.in_channel = QComboBox()
        self.in_channel.setMinimumWidth(220)
        row2.addWidget(self.in_channel)
        row2.addStretch(1)
        # --- Excel benzeri plan grid ---
        self.plan_grid = PlanningGrid()
        layout.addWidget(self.plan_grid, 1)

        # Tarih değişince grid ay/gün vurgusunu güncelle
        self.in_date.dateChanged.connect(self.on_plan_date_changed)

        # ilk açılışta da set et
        self.on_plan_date_changed(self.in_date.date())       

        # formda yüklü olan kayıt (kayıtlı rezervasyonları görüntülerken işimize yarıyor)
        self._loaded_reservation_id: int | None = None

    def refresh_channel_combo(self) -> None:
        """Rezervasyon sekmesindeki kanal listesini DB'den yeniler."""
        if not getattr(self, "repo", None) or not hasattr(self, "in_channel"):
            return
        self.in_channel.blockSignals(True)
        try:
            current = self.in_channel.currentText().strip()
            self.in_channel.clear()
            self.in_channel.addItem("", None)

            for ch in self.repo.list_channels(active_only=True):
                self.in_channel.addItem(str(ch["name"]), int(ch["id"]))

            # mümkünse eski seçimi geri yükle
            if current:
                idx = self.in_channel.findText(current)
                if idx >= 0:
                    self.in_channel.setCurrentIndex(idx)
        finally:
            self.in_channel.blockSignals(False)

    def bootstrap_storage(self) -> None:
        ensure_data_folders(self.app_settings.data_dir)
        db_path = self.app_settings.data_dir / "data.db"
        conn = connect_db(db_path)
        migrate_and_seed(conn)
        self.repo = Repository(conn)
        self.service = ReservationService(self.repo)

        # UI bağımlı listeleri yenile
        self.refresh_channel_combo()
        try:
            self.refresh_reservation_channel_filter()
        except Exception:
            pass

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
        self.refresh_price_channel_tab()
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
        name = item.text()
        self.in_advertiser.setText(name)

        # Seçili reklam veren için son kaydı forma geri getir
        if not getattr(self, "repo", None):
            return
        recs = self.repo.list_confirmed_reservations_by_advertiser(name, limit=1)
        if not recs:
            return
        p = recs[0].payload or {}

        # Basit alanlar
        self.in_agency.setText(str(p.get("agency_name", "") or ""))
        self.in_product.setText(str(p.get("product_name", "") or ""))
        self.in_plan_title.setText(str(p.get("plan_title", "") or ""))
        self.in_spot_code.setText(str(p.get("spot_code", "") or ""))
        self.in_code_definition.setText(str(p.get("code_definition", "") or ""))
        self.in_note.setText(str(p.get("note_text", "") or ""))
        # payload'ta prepared_by olarak tutuluyor
        self.in_prepared_by.setText(str(p.get("prepared_by", "") or ""))

        try:
            self.in_spot_duration.setValue(int(p.get("spot_duration_sec", 0) or 0))
        except Exception:
            pass

        # Tarih + grid
        try:
            dstr = p.get("plan_date")
            if dstr:
                d = datetime.fromisoformat(str(dstr)).date()
                # dateChanged sinyali grid'i resetlediği için önce tarihi set edip sonra matrix basıyoruz
                self.in_date.blockSignals(True)
                self.in_date.setDate(QDate(d.year, d.month, d.day))
                self.in_date.blockSignals(False)
                self.on_plan_date_changed(self.in_date.date())
        except Exception:
            pass

        # Plan hücrelerini de getir (en büyük eksik buydu)
        try:
            self.plan_grid.set_read_only(False)
            self.plan_grid.set_matrix(p.get("plan_cells") or {})
        except Exception:
            pass

        # Bu formda yüklü olan kayıt id'si
        try:
            self._loaded_reservation_id = int(recs[0].id)
        except Exception:
            self._loaded_reservation_id = None

        # Kanal seçimi
        self.refresh_channel_combo()
        ch = str(p.get("channel_name", "") or "").strip()
        if ch:
            idx = self.in_channel.findText(ch)
            if idx >= 0:
                self.in_channel.setCurrentIndex(idx)

        # Eğer kullanıcı başka sekmedeyse, seçime göre anında tazele
        try:
            current_tab = self.tabs.tabText(self.tabs.currentIndex())
            if current_tab == "SPOTLİST+":
                self.refresh_spotlist()
            elif current_tab == "REZERVASYONLAR":
                self.refresh_reservations_tab()
            elif current_tab == "PLAN ÖZET":
                self._set_plan_ozet_period_from_latest(name)
                self.refresh_plan_ozet()
            elif current_tab == "KOD TANIMI":
                self.refresh_kod_tanimi()
            elif current_tab == "Fiyat ve Kanal Tanımı":
                self.refresh_price_channel_tab()
        except Exception:
            pass

        # rezervasyonlar tabı açık değilse bile, arama ile reklamvereni seçince liste tazelensin
        try:
            self.refresh_reservations_tab()
        except Exception:
            pass

    def on_plan_date_changed(self, qd: QDate) -> None:
        d = qd.toPython()
        self.plan_grid.set_month(d.year, d.month, d.day)

    def on_confirm(self) -> None:
        if not self.service:
            QMessageBox.warning(self, "Hata", "Servis hazır değil (DB bağlantısı yok).")
            return

        try:
            # Spot saati kullanıcı girmiyor; Excel zaten çıktıda zaman damgasını basıyor.
            spot_t = datetime.now().time().replace(microsecond=0)

            # Kanal + plan tarihine göre fiyatlar
            plan_d = self.in_date.date().toPython()
            ch_name = self.in_channel.currentText().strip()
            ch_id = self.in_channel.currentData()
            dt_price = 0.0
            odt_price = 0.0
            if ch_id is not None:
                try:
                    price_map = self.repo.get_channel_prices(plan_d.year)
                    dt_price, odt_price = price_map.get((int(ch_id), int(plan_d.month)), (0.0, 0.0))
                except Exception:
                    dt_price, odt_price = (0.0, 0.0)

            draft = ReservationDraft(
                advertiser_name=self.in_advertiser.text(),
                plan_date=plan_d,
                spot_time=spot_t,
                channel_name=ch_name,
                channel_price_dt=float(dt_price),
                channel_price_odt=float(odt_price),
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
        if tab_name == "REZERVASYONLAR":
            self.refresh_reservations_tab()
        if tab_name == "KOD TANIMI":
            self.refresh_kod_tanimi()
        elif tab_name == "SPOTLİST+":
            self.refresh_spotlist()
        elif tab_name == "PLAN ÖZET":
            self._set_plan_ozet_period_from_latest(self.in_advertiser.text())
            self.refresh_plan_ozet()
        elif tab_name == "Fiyat ve Kanal Tanımı":
            self.refresh_price_channel_tab()

    # ------------------------------
    # REZERVASYONLAR (kayıtlı rezervasyonları gör/sil)
    # ------------------------------

    def _build_reservations_tab(self) -> None:
        tab = self.tab_widgets["REZERVASYONLAR"]
        layout = QVBoxLayout(tab)

        top = QHBoxLayout()
        layout.addLayout(top)

        self.btn_res_refresh = QPushButton("Yenile")
        self.btn_res_open = QPushButton("Seçileni Aç")
        self.btn_res_delete = QPushButton("Seçileni Sil")

        top.addWidget(self.btn_res_refresh)
        top.addWidget(self.btn_res_open)
        top.addWidget(self.btn_res_delete)
        top.addSpacing(16)

        top.addWidget(QLabel("Yıl:"))
        self.res_year = QSpinBox()
        self.res_year.setRange(2000, 2100)
        self.res_year.setValue(QDate.currentDate().year())
        self.res_year.setFixedWidth(90)
        top.addWidget(self.res_year)

        top.addWidget(QLabel("Ay:"))
        self.res_month = QComboBox()
        self.res_month.addItem("TÜMÜ", 0)
        for i, m in enumerate(MONTHS_TR, start=1):
            self.res_month.addItem(m, i)
        self.res_month.setFixedWidth(120)
        top.addWidget(self.res_month)

        top.addSpacing(12)
        top.addWidget(QLabel("Kanal:"))
        self.res_channel = QComboBox()
        self.res_channel.setMinimumWidth(220)
        top.addWidget(self.res_channel)
        top.addStretch(1)

        # Liste
        self.res_table = QTableWidget()
        self.res_table.setColumnCount(7)
        self.res_table.setHorizontalHeaderLabels(
            [
                "Rezervasyon No",
                "Plan Tarihi",
                "Kanal",
                "Spot Kodu",
                "Süre (sn)",
                "Adet",
                "Kayıt Zamanı",
            ]
        )
        self.res_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.res_table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.res_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.res_table.verticalHeader().setVisible(False)
        self.res_table.setAlternatingRowColors(True)
        self.res_table.horizontalHeader().setStretchLastSection(True)
        self.res_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        layout.addWidget(self.res_table, 2)

        # Preview
        self.res_preview_title = QLabel("")
        font = self.res_preview_title.font()
        font.setBold(True)
        self.res_preview_title.setFont(font)
        layout.addWidget(self.res_preview_title)

        self.res_preview_grid = PlanningGrid()
        self.res_preview_grid.set_read_only(True)
        layout.addWidget(self.res_preview_grid, 3)

        self._res_records = []  # ReservationRecord[]

        # Signals
        self.btn_res_refresh.clicked.connect(self.refresh_reservations_tab)
        self.btn_res_open.clicked.connect(self.on_reservation_open)
        self.btn_res_delete.clicked.connect(self.on_reservation_delete)
        self.res_table.itemSelectionChanged.connect(self.on_reservation_selected)

        self.res_year.valueChanged.connect(self.refresh_reservations_tab)
        self.res_month.currentIndexChanged.connect(self.refresh_reservations_tab)
        self.res_channel.currentIndexChanged.connect(self.refresh_reservations_tab)

        self.refresh_reservation_channel_filter()

    def refresh_reservation_channel_filter(self) -> None:
        if not getattr(self, "repo", None) or not hasattr(self, "res_channel"):
            return
        current = self.res_channel.currentText().strip() if self.res_channel.count() else ""
        self.res_channel.blockSignals(True)
        try:
            self.res_channel.clear()
            self.res_channel.addItem("TÜMÜ", "")
            for ch in self.repo.list_channels(active_only=True):
                self.res_channel.addItem(str(ch["name"]), str(ch["name"]))
            if current:
                idx = self.res_channel.findText(current)
                if idx >= 0:
                    self.res_channel.setCurrentIndex(idx)
        finally:
            self.res_channel.blockSignals(False)

    def refresh_reservations_tab(self) -> None:
        if not getattr(self, "repo", None):
            return

        adv = (self.in_advertiser.text() or "").strip()
        self._res_records = []
        self.res_table.setRowCount(0)
        self.res_preview_title.setText("")
        try:
            self.res_preview_grid.clear_matrix()
        except Exception:
            pass

        if not adv:
            return

        year = int(self.res_year.value())
        month = int(self.res_month.currentData() or 0)
        ch_filter = str(self.res_channel.currentData() or "").strip()

        recs = self.repo.list_confirmed_reservations_by_advertiser(adv, limit=50000)
        # filtrele
        filtered = []
        for r in recs:
            p = r.payload or {}
            try:
                y, m, d = str(p.get("plan_date") or "").split("-")
                yy = int(y)
                mm = int(m)
            except Exception:
                continue

            if yy != year:
                continue
            if month and mm != month:
                continue
            if ch_filter:
                if str(p.get("channel_name") or "").strip() != ch_filter:
                    continue
            filtered.append(r)

        # yeni -> eski
        filtered.sort(key=lambda x: x.created_at, reverse=True)

        self._res_records = filtered
        self.res_table.setRowCount(len(filtered))

        for i, r in enumerate(filtered):
            p = r.payload or {}
            res_no = str(r.reservation_no or "")
            plan_date = str(p.get("plan_date") or "")
            channel = str(p.get("channel_name") or "")
            spot_code = str(p.get("spot_code") or "")
            duration = str(p.get("spot_duration_sec") or "")
            adet = str(p.get("adet_total") or "")
            created = str(r.created_at or "")

            for col, val in enumerate([res_no, plan_date, channel, spot_code, duration, adet, created]):
                it = QTableWidgetItem(str(val))
                self.res_table.setItem(i, col, it)

    def _get_selected_reservation_records(self) -> list:
        rows = {idx.row() for idx in self.res_table.selectionModel().selectedRows()}
        out = []
        for r in sorted(rows):
            if 0 <= r < len(self._res_records):
                out.append(self._res_records[r])
        return out

    def on_reservation_selected(self) -> None:
        recs = self._get_selected_reservation_records()
        if len(recs) != 1:
            self.res_preview_title.setText("")
            try:
                self.res_preview_grid.clear_matrix()
            except Exception:
                pass
            return

        r = recs[0]
        p = r.payload or {}
        res_no = str(r.reservation_no or "")
        channel = str(p.get("channel_name") or "")
        plan_date = str(p.get("plan_date") or "")
        self.res_preview_title.setText(f"{res_no}  |  {channel}  |  {plan_date}")

        try:
            y, m, d = plan_date.split("-")
            self.res_preview_grid.set_month(int(y), int(m), int(d))
            self.res_preview_grid.set_matrix(p.get("plan_cells") or {})
            self.res_preview_grid.set_read_only(True)
        except Exception:
            pass

    def on_reservation_open(self) -> None:
        recs = self._get_selected_reservation_records()
        if len(recs) != 1:
            QMessageBox.information(self, "Bilgi", "Lütfen tek bir rezervasyon seç.")
            return
        self._load_reservation_into_form(recs[0])
        # ANA SAYFA'ya geç
        idx = self.tabs.indexOf(self.tab_widgets["ANA SAYFA"])
        if idx >= 0:
            self.tabs.setCurrentIndex(idx)

    def _load_reservation_into_form(self, rec) -> None:
        """Kayıtlı rezervasyonu ANA SAYFA formuna basar."""
        p = rec.payload or {}

        self.in_advertiser.setText(str(p.get("advertiser_name", "") or ""))
        self.in_agency.setText(str(p.get("agency_name", "") or ""))
        self.in_product.setText(str(p.get("product_name", "") or ""))
        self.in_plan_title.setText(str(p.get("plan_title", "") or ""))
        self.in_spot_code.setText(str(p.get("spot_code", "") or ""))
        self.in_code_definition.setText(str(p.get("code_definition", "") or ""))
        self.in_note.setText(str(p.get("note_text", "") or ""))

        # prepared_by payload'ta "prepared_by" adıyla saklanıyor
        self.in_prepared_by.setText(str(p.get("prepared_by", "") or ""))

        try:
            self.in_spot_duration.setValue(int(p.get("spot_duration_sec", 0) or 0))
        except Exception:
            pass

        try:
            dstr = p.get("plan_date")
            if dstr:
                d = datetime.fromisoformat(str(dstr)).date()
                self.in_date.blockSignals(True)
                self.in_date.setDate(QDate(d.year, d.month, d.day))
                self.in_date.blockSignals(False)
                self.on_plan_date_changed(self.in_date.date())
        except Exception:
            pass

        self.refresh_channel_combo()
        ch = str(p.get("channel_name", "") or "").strip()
        if ch:
            idx = self.in_channel.findText(ch)
            if idx >= 0:
                self.in_channel.setCurrentIndex(idx)

        try:
            self.plan_grid.set_read_only(False)
            self.plan_grid.set_matrix(p.get("plan_cells") or {})
        except Exception:
            pass

        try:
            self._loaded_reservation_id = int(rec.id)
        except Exception:
            self._loaded_reservation_id = None

    def on_reservation_delete(self) -> None:
        if not getattr(self, "repo", None):
            return
        recs = self._get_selected_reservation_records()
        if not recs:
            QMessageBox.information(self, "Bilgi", "Silmek için en az 1 rezervasyon seç.")
            return

        # kullanıcıya net liste göster
        lines = [str(r.reservation_no or "") for r in recs]
        lines = [x for x in lines if x]
        msg = "\n".join(lines[:15])
        if len(lines) > 15:
            msg += f"\n... (+{len(lines)-15} adet)"

        ok = QMessageBox.question(
            self,
            "Silme Onayı",
            f"Seçili rezervasyon(lar) kalıcı olarak silinecek. Emin misin?\n\n{msg}",
        )
        if ok != QMessageBox.StandardButton.Yes:
            return

        ids = [int(r.id) for r in recs]
        self.repo.delete_reservations_by_ids(ids)

        # eğer formda yüklü olan kayıt silindiyse form id'sini temizle
        if self._loaded_reservation_id in ids:
            self._loaded_reservation_id = None

        # tüm sayfalar DB'den okuduğu için refresh yeterli
        self.refresh_reservations_tab()
        try:
            self.refresh_spotlist()
        except Exception:
            pass
        try:
            self.refresh_kod_tanimi()
        except Exception:
            pass
        try:
            self.refresh_plan_ozet()
        except Exception:
            pass
        
        self.on_search_changed(self.search_edit.text())
        QMessageBox.information(self, "OK", "Seçili rezervasyon(lar) silindi.")

    def _build_spotlist_tab(self) -> None:
        tab = self.tab_widgets["SPOTLİST+"]
        layout = QVBoxLayout(tab)

        # --- Üst bar (butonlar + filtreler) ---
        top = QHBoxLayout()
        layout.addLayout(top)

        self.btn_spot_refresh = QPushButton("Yenile")
        self.btn_spot_save = QPushButton("Kaydet")
        self.btn_spot_export = QPushButton("Excel Çıktısı")
        self.btn_spot_save.setEnabled(False)

        top.addWidget(self.btn_spot_refresh)
        top.addWidget(self.btn_spot_save)
        top.addWidget(self.btn_spot_export)
        top.addSpacing(16)

        top.addWidget(QLabel("Tarih:"))

        self.spot_from = QDateEdit()
        self.spot_from.setCalendarPopup(True)
        self.spot_from.setDisplayFormat("dd.MM.yyyy")

        self.spot_to = QDateEdit()
        self.spot_to.setCalendarPopup(True)
        self.spot_to.setDisplayFormat("dd.MM.yyyy")

        top.addWidget(self.spot_from)
        top.addWidget(QLabel(" - "))
        top.addWidget(self.spot_to)

        top.addSpacing(12)
        top.addWidget(QLabel("Durum:"))
        self.spot_pub_filter = QComboBox()
        self.spot_pub_filter.addItems(["Tümü", "Yayınlandı (1)", "Yayınlanmadı (0)"])
        top.addWidget(self.spot_pub_filter)

        self.btn_spot_clear_filters = QPushButton("Filtreyi Temizle")
        top.addWidget(self.btn_spot_clear_filters)
        top.addStretch(1)

        # --- Tablo ---
        self.spot_table = QTableWidget()
        self.spot_table.setColumnCount(12)
        self.spot_table.setHorizontalHeaderLabels(
            [
                "Sıra",
                "TARİH",
                "ANA YAYIN",
                "REKLAMIN FIRMASI",
                "ADET",
                "BAŞLANGIÇ",
                "SÜRE",
                "Spot Kodu",
                "DT-ODT",
                "Birim Saniye",
                "Bütçe Net TL",
                "Yayınlandı Durum",
            ]
        )

        self.spot_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.spot_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.spot_table.setAlternatingRowColors(True)
        self.spot_table.verticalHeader().setVisible(False)
        self.spot_table.setShowGrid(True)

        header = self.spot_table.horizontalHeader()
        header.setStretchLastSection(True)

        # Excel'e yakın: bazı kolonlar sabit, bazıları esnek
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)  # sıra
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)  # tarih
        header.setSectionResizeMode(2, QHeaderView.Stretch)           # ana yayın
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)  # firma
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)  # adet
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)  # başlangıç
        header.setSectionResizeMode(6, QHeaderView.ResizeToContents)  # süre
        header.setSectionResizeMode(7, QHeaderView.ResizeToContents)  # spot kodu
        header.setSectionResizeMode(8, QHeaderView.ResizeToContents)  # dt-odt
        header.setSectionResizeMode(9, QHeaderView.ResizeToContents)  # birim saniye
        header.setSectionResizeMode(10, QHeaderView.ResizeToContents) # bütçe
        header.setSectionResizeMode(11, QHeaderView.Stretch)          # yayınlandı

        # Stil
        self.spot_table.setStyleSheet('''
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
        ''')

        # --- Özet bar ---
        self.spot_summary = QLabel("")
        font = self.spot_summary.font()
        font.setBold(True)
        self.spot_summary.setFont(font)

        layout.addWidget(self.spot_table, 1)
        layout.addWidget(self.spot_summary)

        # --- State ---
        self.spot_all_rows = []
        self.spot_dirty = {}              # (reservation_id, day, row_idx) -> 0/1
        self.spot_current_adv = ""
        self.spot_filters_initialized = False

        # --- Signals ---
        self.btn_spot_refresh.clicked.connect(self.refresh_spotlist)
        self.btn_spot_export.clicked.connect(self.on_spotlist_export)
        self.btn_spot_save.clicked.connect(self.on_spotlist_save)

        self.spot_from.dateChanged.connect(self._apply_spotlist_filters)
        self.spot_to.dateChanged.connect(self._apply_spotlist_filters)
        self.spot_pub_filter.currentIndexChanged.connect(self._apply_spotlist_filters)
        self.btn_spot_clear_filters.clicked.connect(self._spotlist_clear_filters)

    def refresh_spotlist(self) -> None:
        if not self.service:
            return

        adv = (self.in_advertiser.text() or "").strip()
        if not adv:
            self.spot_table.setRowCount(0)
            self.spot_summary.setText("")
            return

        # reklam veren değiştiyse dirty temizle (karışmasın)
        if adv != self.spot_current_adv:
            self.spot_dirty.clear()
            self.btn_spot_save.setEnabled(False)
            self.spot_filters_initialized = False
            self.spot_current_adv = adv

        self.spot_all_rows = self.service.get_spotlist_rows(adv)

        # Tarih filtre aralığını ilk yüklemede dataya göre ayarla
        if self.spot_all_rows and not self.spot_filters_initialized:
            try:
                dmin = min(r["datetime"].date() for r in self.spot_all_rows)
                dmax = max(r["datetime"].date() for r in self.spot_all_rows)
            except Exception:
                dmin = date.today()
                dmax = date.today()

            self.spot_from.blockSignals(True)
            self.spot_to.blockSignals(True)
            self.spot_from.setDate(QDate(dmin.year, dmin.month, dmin.day))
            self.spot_to.setDate(QDate(dmax.year, dmax.month, dmax.day))
            self.spot_from.blockSignals(False)
            self.spot_to.blockSignals(False)

            self.spot_pub_filter.setCurrentIndex(0)
            self.spot_filters_initialized = True

        self._apply_spotlist_filters()

    def _spotlist_clear_filters(self) -> None:
        if not self.spot_all_rows:
            return
        try:
            dmin = min(r["datetime"].date() for r in self.spot_all_rows)
            dmax = max(r["datetime"].date() for r in self.spot_all_rows)
        except Exception:
            dmin = date.today()
            dmax = date.today()

        self.spot_from.blockSignals(True)
        self.spot_to.blockSignals(True)
        self.spot_pub_filter.blockSignals(True)

        self.spot_from.setDate(QDate(dmin.year, dmin.month, dmin.day))
        self.spot_to.setDate(QDate(dmax.year, dmax.month, dmax.day))
        self.spot_pub_filter.setCurrentIndex(0)

        self.spot_from.blockSignals(False)
        self.spot_to.blockSignals(False)
        self.spot_pub_filter.blockSignals(False)

        self._apply_spotlist_filters()

    def _filtered_spotlist_rows(self) -> list[dict]:
        if not self.spot_all_rows:
            return []

        d1 = self.spot_from.date().toPython()
        d2 = self.spot_to.date().toPython()
        if d1 > d2:
            d1, d2 = d2, d1

        mode = self.spot_pub_filter.currentIndex()  # 0 all, 1 pub=1, 2 pub=0

        out = []
        for rr in self.spot_all_rows:
            dtv = rr.get("datetime")
            if not dtv:
                continue
            dd = dtv.date()
            if dd < d1 or dd > d2:
                continue

            key = (int(rr.get("reservation_id")), int(rr.get("day")), int(rr.get("row_idx")))
            pub = int(self.spot_dirty.get(key, rr.get("published", 0) or 0))
            rr2 = dict(rr)
            rr2["published"] = pub

            if mode == 1 and pub != 1:
                continue
            if mode == 2 and pub != 0:
                continue

            out.append(rr2)

        for i, rr in enumerate(out, start=1):
            rr["sira"] = i
        return out

    def _apply_spotlist_filters(self, *args) -> None:
        rows = self._filtered_spotlist_rows()
        self._render_spotlist(rows)
        self._update_spotlist_summary(rows)

    def _render_spotlist(self, rows: list[dict]) -> None:
        self.spot_table.setRowCount(len(rows))

        for r_idx, rr in enumerate(rows):
            vals = [
                rr.get("sira", ""),
                rr.get("tarih", ""),
                rr.get("ana_yayin", ""),
                rr.get("reklam_firmasi", ""),
                rr.get("adet", ""),
                rr.get("baslangic", ""),
                rr.get("sure", ""),
                rr.get("spot_kodu", ""),
                rr.get("dt_odt", ""),
                rr.get("birim_saniye", 0.0),
                rr.get("butce_net", 0.0),
            ]

            for c in range(0, 11):
                item = QTableWidgetItem()
                if c in (0, 4, 6):  # int
                    try:
                        item.setText(str(int(vals[c])))
                    except Exception:
                        item.setText(str(vals[c]))
                    item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
                elif c in (9, 10):  # float
                    try:
                        item.setText(f"{float(vals[c]):.2f}")
                    except Exception:
                        item.setText(str(vals[c]))
                    item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
                else:
                    item.setText(str(vals[c]))
                    item.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)

                item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                self.spot_table.setItem(r_idx, c, item)

            key = (int(rr.get("reservation_id")), int(rr.get("day")), int(rr.get("row_idx")))
            pub_val = 1 if int(rr.get("published", 0) or 0) else 0

            combo = QComboBox()
            combo.addItems(["0", "1"])
            combo.blockSignals(True)
            combo.setCurrentText("1" if pub_val == 1 else "0")
            combo.blockSignals(False)
            combo.currentTextChanged.connect(lambda txt, k=key: self.on_spot_published_changed(k, txt))
            self.spot_table.setCellWidget(r_idx, 11, combo)

        self.spot_table.resizeRowsToContents()

    def _update_spotlist_summary(self, rows: list[dict]) -> None:
        if not rows:
            self.spot_summary.setText("Kayıt yok.")
            return

        total_adet = sum(int(r.get("adet", 1) or 1) for r in rows)
        total_budget = sum(float(r.get("butce_net", 0.0) or 0.0) for r in rows)
        durations = [int(r.get("sure", 0) or 0) for r in rows]
        avg_duration = (sum(durations) / len(durations)) if durations else 0.0

        self.spot_summary.setText(
            f"Toplam Satır: {len(rows)}    "
            f"Toplam Adet: {total_adet}    "
            f"Toplam Bütçe: {total_budget:,.2f} TL    "
            f"Ortalama Süre: {avg_duration:.1f} sn"
        )

    def on_spot_published_changed(self, key: tuple[int, int, int], txt: str) -> None:
        try:
            val = 1 if str(txt).strip() == "1" else 0
            self.spot_dirty[key] = val
            self.btn_spot_save.setEnabled(True)

            # Filtre durumuna göre görünürlük değişebilir
            if self.spot_pub_filter.currentIndex() != 0:
                self._apply_spotlist_filters()
            else:
                self._update_spotlist_summary(self._filtered_spotlist_rows())
        except Exception as e:
            QMessageBox.warning(self, "Hata", str(e))

    def on_spotlist_save(self) -> None:
        if not self.service:
            return
        if not self.spot_dirty:
            QMessageBox.information(self, "Bilgi", "Kaydedilecek değişiklik yok.")
            return

        try:
            changes = [(k[0], k[1], k[2], int(v)) for k, v in self.spot_dirty.items()]
            self.service.set_spotlist_published_bulk(changes)
            self.spot_dirty.clear()
            self.btn_spot_save.setEnabled(False)

            # DB'den güncel değerleri tekrar çek
            self.spot_all_rows = self.service.get_spotlist_rows(self.spot_current_adv)
            self._apply_spotlist_filters()

            QMessageBox.information(self, "OK", "Değişiklikler kaydedildi.")
        except Exception as e:
            QMessageBox.critical(self, "Hata", str(e))

    def on_spotlist_export(self) -> None:
        if not self.service:
            QMessageBox.warning(self, "Hata", "Servis hazır değil.")
            return
        adv = (self.in_advertiser.text() or "").strip()
        if not adv:
            QMessageBox.warning(self, "Hata", "Önce bir reklam veren seç.")
            return

        default_name = f"SPOTLIST_{adv}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        default_path = str((self.app_settings.data_dir / "exports" / default_name).resolve())

        path, _ = QFileDialog.getSaveFileName(
            self,
            "SPOTLİST+ Excel Çıktısı",
            default_path,
            "Excel Files (*.xlsx)"
        )
        if not path:
            return

        try:
            rows = self._filtered_spotlist_rows()
            self.service.export_spotlist_excel_with_rows(path, adv, rows)
            QMessageBox.information(self, "OK", f"Excel çıktısı üretildi:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Hata", str(e))


    # ------------------------------
    # PLAN ÖZET (ay bazlı birleştirilmiş özet)
    # ------------------------------

    def _build_plan_ozet_tab(self) -> None:
        tab = self.tab_widgets["PLAN ÖZET"]
        layout = QVBoxLayout(tab)

        # Üst bar: yıl/ay seçimi
        top = QHBoxLayout()
        layout.addLayout(top)

        top.addWidget(QLabel("Yıl:"))
        self.po_year = QSpinBox()
        self.po_year.setRange(2000, 2100)
        self.po_year.setValue(date.today().year)
        top.addWidget(self.po_year)

        top.addWidget(QLabel("Ay:"))
        self.po_month = QComboBox()
        self.po_month.addItems([
            "OCAK", "ŞUBAT", "MART", "NİSAN", "MAYIS", "HAZİRAN",
            "TEMMUZ", "AĞUSTOS", "EYLÜL", "EKİM", "KASIM", "ARALIK"
        ])
        self.po_month.setCurrentIndex(date.today().month - 1)
        top.addWidget(self.po_month)

        self.btn_po_refresh = QPushButton("Yenile")
        self.btn_po_export = QPushButton("Excel Çıktısı")
        top.addWidget(self.btn_po_export)
        top.addWidget(self.btn_po_refresh)
        top.addStretch(1)

        # Header alanları (Excel'deki gibi)
        header_box = QGroupBox()
        header_box.setTitle("")
        header_layout = QVBoxLayout(header_box)
        layout.addWidget(header_box)

        def _row(label: str, widget: QWidget) -> None:
            r = QHBoxLayout()
            lab = QLabel(label)
            lab.setFixedWidth(160)
            lab.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
            lab.setStyleSheet("background:#0b2f6b;color:white;padding:6px;font-weight:bold;")
            r.addWidget(lab)
            r.addWidget(widget, 1)
            header_layout.addLayout(r)

        self.po_agency = QLineEdit(); self.po_agency.setReadOnly(True)
        self.po_advertiser = QLineEdit(); self.po_advertiser.setReadOnly(True)
        self.po_product = QLineEdit(); self.po_product.setReadOnly(True)
        self.po_plan_title = QLineEdit(); self.po_plan_title.setReadOnly(True)

        # Rezervasyon no: çoklu olunca listelenecek, o yüzden multi-line
        self.po_resno = QPlainTextEdit()
        self.po_resno.setReadOnly(True)
        self.po_resno.setMaximumHeight(72)

        self.po_period = QLineEdit(); self.po_period.setReadOnly(True)
        self.po_spot_len = QLineEdit(); self.po_spot_len.setReadOnly(True)

        _row("Ajans", self.po_agency)
        _row("Reklamveren", self.po_advertiser)
        _row("Ürün", self.po_product)
        _row("Plan Başlığı", self.po_plan_title)
        _row("Rezervasyon No", self.po_resno)
        _row("Dönemi", self.po_period)
        _row("Spot Süresi -Sn", self.po_spot_len)

        # Tablo
        self.po_table = QTableWidget()
        self.po_table.setAlternatingRowColors(True)
        self.po_table.verticalHeader().setVisible(False)
        self.po_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.po_table.setShowGrid(True)
        self.po_table.horizontalHeader().setStretchLastSection(True)
        self.po_table.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed)
        self.po_table.horizontalHeader().setStyleSheet(
            "QHeaderView::section{background:#e67e22;color:white;font-weight:bold;border:1px solid #c66a1d;padding:6px;}"
        )
        layout.addWidget(self.po_table, 1)

        # Wire
        self.btn_po_refresh.clicked.connect(self.refresh_plan_ozet)
        self.btn_po_export.clicked.connect(self.export_plan_ozet_excel)
        self.po_year.valueChanged.connect(lambda *_: self.refresh_plan_ozet())
        self.po_month.currentIndexChanged.connect(lambda *_: self.refresh_plan_ozet())
        self.po_table.cellChanged.connect(self._po_on_cell_changed)

    def _set_plan_ozet_period_from_latest(self, advertiser_name: str) -> None:
        """Reklamveren seçilince plan özetin yıl/ayını en son rezervasyona göre ayarla."""
        adv = (advertiser_name or "").strip()
        if not adv or not getattr(self, "repo", None):
            return
        try:
            recs = self.repo.list_confirmed_reservations_by_advertiser(adv, limit=1)
            if not recs:
                return
            p = recs[0].payload or {}
            dstr = p.get("plan_date")
            if not dstr:
                return
            d = datetime.fromisoformat(str(dstr)).date()
            self.po_year.blockSignals(True)
            self.po_month.blockSignals(True)
            self.po_year.setValue(d.year)
            self.po_month.setCurrentIndex(d.month - 1)
        except Exception:
            pass
        finally:
            try:
                self.po_year.blockSignals(False)
                self.po_month.blockSignals(False)
            except Exception:
                pass

    def _po_on_cell_changed(self, row: int, col: int) -> None:
        """Yayın grubu tek yerden girilsin diye aynı kanalın DT/ODT satırlarını senkron tut."""
        if self._po_syncing:
            return
        if col != 1:
            return
        if not hasattr(self, "po_table"):
            return
        # Toplam satırını karıştırma
        if row >= self.po_table.rowCount() - 1:
            return
        ch_item = self.po_table.item(row, 0)
        if not ch_item:
            return
        ch = ch_item.text()
        val_item = self.po_table.item(row, 1)
        val = val_item.text() if val_item else ""

        try:
            self._po_syncing = True
            for r in range(self.po_table.rowCount() - 1):
                if r == row:
                    continue
                it = self.po_table.item(r, 0)
                if it and it.text() == ch:
                    tgt = self.po_table.item(r, 1)
                    if tgt is None:
                        tgt = QTableWidgetItem(val)
                        self.po_table.setItem(r, 1, tgt)
                    else:
                        tgt.setText(val)
        finally:
            self._po_syncing = False

    def refresh_plan_ozet(self) -> None:
        if not getattr(self, "service", None):
            return
        adv = (self.in_advertiser.text() or "").strip()
        if not adv:
            return

        yy = int(self.po_year.value())
        mm = int(self.po_month.currentIndex()) + 1

        try:
            data = self.service.get_plan_ozet_data(adv, yy, mm)
            header = data.get("header") or {}
            rows = data.get("rows") or []
            days = int(data.get("days") or 0)

            # Header alanları
            self.po_agency.setText(str(header.get("agency", "") or ""))
            self.po_advertiser.setText(str(header.get("advertiser", adv) or adv))
            self.po_product.setText(str(header.get("product", "") or ""))
            self.po_plan_title.setText(str(header.get("plan_title", "") or ""))
            self.po_resno.setPlainText(str(header.get("reservation_no", "") or ""))
            self.po_period.setText(str(header.get("period", "") or ""))
            sl = header.get("spot_len", 0) or 0
            try:
                sl = float(sl)
                self.po_spot_len.setText("" if sl == 0 else f"{sl:.2f}")
            except Exception:
                self.po_spot_len.setText(str(sl))

            month_name = str(header.get("month_name", "") or "")

            # Kolonlar
            base_headers = ["KANAL", "YAYIN GRUBU", "DT/ODT", "DİNLENME ORANI"]
            day_headers = [str(i) for i in range(1, days + 1)]
            tail_headers = [
                f"{month_name} Adet" if month_name else "Ay Adet",
                f"{month_name} Saniye" if month_name else "Ay Saniye",
                "Birim sn. (TL)",
                "Toplam Bütçe Net TL",
            ]
            headers = base_headers + day_headers + tail_headers

            self.po_table.blockSignals(True)
            self.po_table.clear()
            self.po_table.setColumnCount(len(headers))
            self.po_table.setHorizontalHeaderLabels(headers)

            totals = data.get("totals") or {}
            total_rows = len(rows) + 1  # + toplam satırı
            self.po_table.setRowCount(total_rows)

            def _fmt(v):
                if v in (None, ""):
                    return ""
                try:
                    f = float(v)
                    if abs(f - int(f)) < 1e-9:
                        return str(int(f))
                    return f"{f:.2f}"
                except Exception:
                    return str(v)

            # Data satırları
            for r_i, rr in enumerate(rows):
                # Kanal
                it0 = QTableWidgetItem(str(rr.get("channel", "") or ""))
                it0.setFlags(it0.flags() & ~Qt.ItemIsEditable)
                self.po_table.setItem(r_i, 0, it0)

                # Yayın grubu (editable)
                it1 = QTableWidgetItem(str(rr.get("publish_group", "") or ""))
                self.po_table.setItem(r_i, 1, it1)

                # DT/ODT
                it2 = QTableWidgetItem(str(rr.get("dt_odt", "") or ""))
                it2.setFlags(it2.flags() & ~Qt.ItemIsEditable)
                self.po_table.setItem(r_i, 2, it2)

                # Dinlenme oranı
                it3 = QTableWidgetItem(str(rr.get("dinlenme_orani", "") or "NA"))
                it3.setFlags(it3.flags() & ~Qt.ItemIsEditable)
                self.po_table.setItem(r_i, 3, it3)

                # Günler
                day_vals = rr.get("days") or []
                for di in range(days):
                    v = day_vals[di] if di < len(day_vals) else ""
                    it = QTableWidgetItem(_fmt(v))
                    it.setTextAlignment(Qt.AlignCenter)
                    it.setFlags(it.flags() & ~Qt.ItemIsEditable)
                    self.po_table.setItem(r_i, 4 + di, it)

                # Ay toplamları
                base = 4 + days
                tail = [rr.get("month_adet"), rr.get("month_saniye"), rr.get("unit_price"), rr.get("budget")]
                for j, v in enumerate(tail):
                    it = QTableWidgetItem(_fmt(v))
                    it.setTextAlignment(Qt.AlignCenter)
                    it.setFlags(it.flags() & ~Qt.ItemIsEditable)
                    self.po_table.setItem(r_i, base + j, it)

            # Toplam satırı
            t_row = len(rows)
            bold = QFont(); bold.setBold(True)

            it_t0 = QTableWidgetItem("Toplam")
            it_t0.setFont(bold)
            it_t0.setFlags(it_t0.flags() & ~Qt.ItemIsEditable)
            self.po_table.setItem(t_row, 0, it_t0)

            # boş publish group / dtodt / dinlenme
            for c in (1, 2, 3):
                it = QTableWidgetItem("")
                it.setFlags(it.flags() & ~Qt.ItemIsEditable)
                self.po_table.setItem(t_row, c, it)

            t_days = totals.get("days") or []
            for di in range(days):
                v = t_days[di] if di < len(t_days) else ""
                it = QTableWidgetItem(_fmt(v))
                it.setFont(bold)
                it.setTextAlignment(Qt.AlignCenter)
                it.setFlags(it.flags() & ~Qt.ItemIsEditable)
                self.po_table.setItem(t_row, 4 + di, it)

            base = 4 + days
            tail = [totals.get("month_adet"), totals.get("month_saniye"), "", totals.get("budget")]
            for j, v in enumerate(tail):
                it = QTableWidgetItem(_fmt(v))
                it.setFont(bold)
                it.setTextAlignment(Qt.AlignCenter)
                it.setFlags(it.flags() & ~Qt.ItemIsEditable)
                self.po_table.setItem(t_row, base + j, it)

            self.po_table.resizeColumnsToContents()
        except Exception as e:
            QMessageBox.critical(self, "Hata", str(e))
        finally:
            try:
                self.po_table.blockSignals(False)
            except Exception:
                pass

    def export_plan_ozet_excel(self) -> None:
        if not getattr(self, "service", None):
            return
        adv = (self.in_advertiser.text() or "").strip()
        if not adv:
            return

        yy = int(self.po_year.value())
        mm = int(self.po_month.currentIndex()) + 1

        default_name = f"PLAN_OZET_{adv}_{yy}_{mm:02d}.xlsx"
        out_dir = self.app_settings.data_dir / "exports"
        out_dir.mkdir(parents=True, exist_ok=True)
        default_path = str((out_dir / default_name).resolve())

        path, _ = QFileDialog.getSaveFileName(
            self,
            "Plan Özet Excel Kaydet",
            default_path,
            "Excel Files (*.xlsx)"
        )
        if not path:
            return

        try:
            self.service.export_plan_ozet_excel(path, adv, yy, mm)
            QMessageBox.information(self, "OK", f"Excel çıktısı oluşturuldu:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Hata", str(e))


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

        # Hücre editöründe kalmış değer varsa (Enter'a basılmadıysa), kaydetmeden önce commit edelim.
        try:
            editor = QApplication.focusWidget()
            if editor and self.price_table and self.price_table.isAncestorOf(editor):
                self.price_table.closeEditor(editor, QAbstractItemDelegate.SubmitModelCache)
                QApplication.processEvents()
        except Exception:
            pass

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
            self.refresh_channel_combo()
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
            self.refresh_channel_combo()
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
            self.refresh_channel_combo()
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