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
from src.ui.excel_table import ExcelTableWidget


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
        # Daha geniş default: ana sayfada gereksiz scroll ihtiyacını azaltır.
        self.resize(1500, 850)

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

        top.addWidget(QLabel("Plan Başlığı Ara:"))
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Örn: TEST BAŞLIĞI")
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

        # --- State (must be initialized BEFORE any tab builder that reads it) ---
        # Çoklu ay (tarih aralığı) için ay bazlı hücre cache'i
        self._month_cells_cache: dict[tuple[int, int], dict] = {}
        self._current_month_key: tuple[int, int] | None = None

        # Çoklu kanal seçimi
        self._selected_channel_names: list[str] = []

        # Ana sayfa tarih aralığı (opsiyonel). Seçiliyse PlanningGrid gün kolonları aralığa göre gizlenir.
        self._home_range_start: date | None = None
        self._home_range_end: date | None = None

        # ANA SAYFA (rezervasyon girişi)
        self._build_home_tab()

        # REZERVASYONLAR (kayıtlı rezervasyonları gör/sil)
        self._build_reservations_tab()

        # Bottom buttons (tek akış: seçilen tarih aralığını DB'ye kaydet)
        bottom = QHBoxLayout()
        main.addLayout(bottom)

        self.btn_new_reservation = QPushButton("Yeni Rezervasyon")
        self.btn_save = QPushButton("Kaydet")
        bottom.addWidget(self.btn_new_reservation)
        bottom.addWidget(self.btn_save)

        # Wire
        self.btn_pick_folder.clicked.connect(self.pick_data_folder)
        self.search_edit.textChanged.connect(self.on_search_changed)
        self.list_advertisers.itemClicked.connect(self.on_advertiser_selected)

        # Yeni model: tek buton. Seçilen tarih aralığındaki tüm seçili kanalları DB'ye kaydeder.
        self.btn_save.clicked.connect(self.on_save)
        self.btn_new_reservation.clicked.connect(self.reset_form_for_new_reservation)

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
        self.in_spot_code.setMaxLength(10)
        self.in_spot_code.setFixedWidth(90)
        row0b.addWidget(self.in_spot_code, 1)

        row0b.addWidget(QLabel("Süre (sn):"))
        self.in_spot_duration = QSpinBox()
        self.in_spot_duration.setRange(0, 9999)
        self.in_spot_duration.setFixedWidth(90)
        row0b.addWidget(self.in_spot_duration, 1)

        # Ajans komisyonu (%): çıktıda AR62 formülünü dinamik yapmak için.
        row0b.addWidget(QLabel("Ajans Kom. (%):"))
        self.in_agency_commission = QSpinBox()
        self.in_agency_commission.setRange(0, 100)
        self.in_agency_commission.setValue(10)
        self.in_agency_commission.setFixedWidth(70)
        row0b.addWidget(self.in_agency_commission, 1)

        row_code_def = QHBoxLayout()
        layout.addLayout(row_code_def)

        row_code_def.addWidget(QLabel("Kod Tanımı:"))
        self.in_code_definition = QLineEdit()
        row_code_def.addWidget(self.in_code_definition, 6)

        # Çoklu kod ekleme (kodları tanımla; grid'e K/A/B gibi harfleri kullanıcı yazar)
        row_code_btns = QHBoxLayout()
        layout.addLayout(row_code_btns)
        self.btn_add_code = QPushButton("Kodu Ekle")
        self.btn_remove_code = QPushButton("Seçili Kodu Sil")
        self.btn_fill_selected = QPushButton("Seçili Hücrelere Yaz")
        self.cb_active_code = QComboBox()
        self.cb_active_code.setMinimumWidth(120)
        row_code_btns.addWidget(self.btn_add_code)
        row_code_btns.addWidget(self.btn_remove_code)
        row_code_btns.addSpacing(12)
        row_code_btns.addWidget(QLabel("Aktif Kod:"))
        row_code_btns.addWidget(self.cb_active_code)
        row_code_btns.addWidget(self.btn_fill_selected)
        row_code_btns.addStretch(1)

        self.tbl_codes = QTableWidget(0, 3)
        self.tbl_codes.setHorizontalHeaderLabels(["Kod", "Kod Tanımı", "Süre (sn)"])
        self.tbl_codes.verticalHeader().setVisible(False)
        self.tbl_codes.setSelectionBehavior(QTableWidget.SelectRows)
        self.tbl_codes.setSelectionMode(QTableWidget.SingleSelection)
        self.tbl_codes.horizontalHeader().setStretchLastSection(True)
        self.tbl_codes.setMaximumHeight(120)
        layout.addWidget(self.tbl_codes)


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
        self.in_channel.setMinimumWidth(200)
        row2.addWidget(self.in_channel)

        self.btn_select_channels = QPushButton("Kanalları Seç")
        row2.addWidget(self.btn_select_channels)
        self.lbl_selected_channels = QLabel("")
        self.lbl_selected_channels.setMinimumWidth(180)
        row2.addWidget(self.lbl_selected_channels)

        row2.addStretch(1)

        row2b = QHBoxLayout()
        layout.addLayout(row2b)
        row2b.addWidget(QLabel("Tarih Aralığı:"))
        self.in_range_start = QDateEdit()
        self.in_range_start.setCalendarPopup(True)
        self.in_range_start.setFixedWidth(120)
        self.in_range_end = QDateEdit()
        self.in_range_end.setCalendarPopup(True)
        self.in_range_end.setFixedWidth(120)
        row2b.addWidget(self.in_range_start)
        row2b.addWidget(QLabel("-"))
        row2b.addWidget(self.in_range_end)
        self.btn_apply_range = QPushButton("Aralığı Uygula")
        row2b.addWidget(self.btn_apply_range)
        row2b.addStretch(1)
        # --- Excel benzeri plan grid ---
        self.plan_grid = PlanningGrid()
        layout.addWidget(self.plan_grid, 1)

        # Tarih değişince grid ay/gün vurgusunu güncelle
        self.in_date.dateChanged.connect(self.on_plan_date_changed)

        # Kod ve kanal aksiyonları
        self.btn_add_code.clicked.connect(self.on_add_code_def)
        self.btn_remove_code.clicked.connect(self.on_remove_code_def)
        self.btn_fill_selected.clicked.connect(self.on_fill_selected_cells_with_active_code)
        self.cb_active_code.currentIndexChanged.connect(self.on_active_code_changed)
        self.btn_select_channels.clicked.connect(self.on_select_channels)
        self.btn_apply_range.clicked.connect(self.on_apply_date_range)

        # tarih aralığı defaultu: mevcut ay
        today = self.in_date.date()
        first = QDate(today.year(), today.month(), 1)
        last = QDate(today.year(), today.month(), today.daysInMonth())
        self.in_range_start.setDate(first)
        self.in_range_end.setDate(last)
        self.on_apply_date_range()

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
        for name in self.repo.search_plan_titles(text, limit=30):
            self.list_advertisers.addItem(name)

    def on_advertiser_selected(self, item) -> None:
        if not item:
            return
        name = item.text()
        self.in_plan_title.setText(name)

        # Seçili reklam veren için son kaydı forma geri getir
        if not getattr(self, "repo", None):
            return
        recs = self.repo.list_confirmed_reservations_by_plan_title(name, limit=1)
        if not recs:
            return
        p = recs[0].payload or {}
        # Plan başlığı seçimi ile birlikte reklamvereni de payload içinden güncelle
        self.in_advertiser.setText(str(p.get("advertiser_name") or ""))

        # Basit alanlar
        self.in_agency.setText(str(p.get("agency_name", "") or ""))
        self.in_product.setText(str(p.get("product_name", "") or ""))
        self.in_plan_title.setText(str(p.get("plan_title", "") or ""))
        self.in_spot_code.setText(str(p.get("spot_code", "") or ""))
        self.in_code_definition.setText(str(p.get("code_definition", "") or ""))
        self.in_note.setText(str(p.get("note_text", "") or ""))
        # payload'ta prepared_by "İSİM - dd.mm.yyyy hh:mm" şeklinde tutuluyor.
        # Formda ise sadece isim görünsün ki tekrar onaylandığında tarih tekrar tekrar eklenmesin.
        pb_raw = str(p.get("prepared_by", "") or "").strip()
        pb_name = pb_raw.split(" - ", 1)[0].strip() if pb_raw else ""
        self.in_prepared_by.setText(pb_name)

        try:
            self.in_spot_duration.setValue(int(p.get("spot_duration_sec", 0) or 0))
        except Exception:
            pass

        try:
            self.in_agency_commission.setValue(int(float(p.get("agency_commission_pct", 0) or 0)))
        except Exception:
            pass

        # Tarih + grid
        try:
            is_span = bool(p.get("is_span"))
            if is_span and p.get("span_start") and p.get("span_end"):
                ds = datetime.fromisoformat(str(p.get("span_start"))).date()
                de = datetime.fromisoformat(str(p.get("span_end"))).date()

                self.in_range_start.blockSignals(True)
                self.in_range_end.blockSignals(True)
                self.in_range_start.setDate(QDate(ds.year, ds.month, ds.day))
                self.in_range_end.setDate(QDate(de.year, de.month, de.day))
                self.in_range_start.blockSignals(False)
                self.in_range_end.blockSignals(False)

                self.in_date.blockSignals(True)
                self.in_date.setDate(QDate(ds.year, ds.month, ds.day))
                self.in_date.blockSignals(False)

                self.on_apply_date_range()
            else:
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

        # Eğer kullanıcı başka sekmedeyse, seçime göre anında tazele
        try:
            current_tab = self.tabs.tabText(self.tabs.currentIndex())
            if current_tab == "SPOTLİST+":
                self.refresh_spotlist()
            elif current_tab == "REZERVASYONLAR":
                self.refresh_reservations_tab()
            elif current_tab == "PLAN ÖZET":
                # Plan özet dönemi rezervasyon ekranındaki tarih aralığından türetilir.
                # Metot argümansız tasarlanmıştı; bazı çağrılarda yanlışlıkla parametre gönderiliyordu.
                self._set_plan_ozet_period_from_latest()
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

        # Eğer grid tarih aralığı (span) modundaysa, ay değiştirmeden sadece seçili günü vurgula.
        try:
            if getattr(self.plan_grid, "is_span_mode", lambda: False)():
                self.plan_grid.set_selected_date(d)
                return
        except Exception:
            pass

        new_key = (d.year, d.month)

        # Aynı ay içinde gün değişiyorsa cache'e dokunma
        if self._current_month_key is None:
            self._current_month_key = new_key
        elif new_key != self._current_month_key:
            # Eski ayın matrix'ini cache'e yaz
            try:
                self._month_cells_cache[self._current_month_key] = self.plan_grid.get_matrix()
            except Exception:
                pass
            self._current_month_key = new_key

        self.plan_grid.set_month(d.year, d.month, d.day)

        # Yeni ay daha önce planlandıysa geri yükle
        cached = self._month_cells_cache.get(new_key)
        if cached:
            try:
                self.plan_grid.set_matrix(cached)
            except Exception:
                pass


    # -------------------------
    # Çoklu Kod Tanımları (ANA SAYFA)
    # -------------------------
    def _get_code_defs_from_ui(self) -> list[dict]:
        """ANA SAYFA'daki kod tanımı tablosundan code_defs üret."""
        out: list[dict] = []
        seen: set[str] = set()

        # Tablo
        try:
            for r in range(self.tbl_codes.rowCount()):
                code = (self.tbl_codes.item(r, 0).text() if self.tbl_codes.item(r, 0) else "").strip().upper()
                desc = (self.tbl_codes.item(r, 1).text() if self.tbl_codes.item(r, 1) else "").strip()
                dur_txt = (self.tbl_codes.item(r, 2).text() if self.tbl_codes.item(r, 2) else "").strip()
                if not code:
                    continue
                if code in seen:
                    continue
                seen.add(code)
                try:
                    dur = int(float(dur_txt)) if dur_txt else 0
                except Exception:
                    dur = 0
                out.append({"code": code, "desc": desc, "duration_sec": dur})
        except Exception:
            pass

        # Geri uyumluluk: tablo boşsa tekli inputları kullan
        if not out:
            code = (self.in_spot_code.text() or "").strip().upper()
            desc = (self.in_code_definition.text() or "").strip()
            dur = int(self.in_spot_duration.value() or 0)
            if code or desc or dur:
                if code:
                    out.append({"code": code, "desc": desc, "duration_sec": dur})

        return out

    def _sync_active_code_combo(self) -> None:
        """Aktif kod combobox'ını tabloyla senkron tut."""
        try:
            current = self.cb_active_code.currentText().strip().upper()
        except Exception:
            current = ""

        codes = []
        try:
            for r in range(self.tbl_codes.rowCount()):
                code = (self.tbl_codes.item(r, 0).text() if self.tbl_codes.item(r, 0) else "").strip().upper()
                if code and code not in codes:
                    codes.append(code)
        except Exception:
            pass

        self.cb_active_code.blockSignals(True)
        try:
            self.cb_active_code.clear()
            self.cb_active_code.addItem("", "")
            for c in codes:
                self.cb_active_code.addItem(c, c)
            if current:
                idx = self.cb_active_code.findText(current)
                if idx >= 0:
                    self.cb_active_code.setCurrentIndex(idx)
        finally:
            self.cb_active_code.blockSignals(False)

    def on_add_code_def(self) -> None:
        code = (self.in_spot_code.text() or "").strip().upper()
        desc = (self.in_code_definition.text() or "").strip()
        dur = int(self.in_spot_duration.value() or 0)

        if not code:
            QMessageBox.information(self, "Bilgi", "Kod zorunlu. (Örn: K, A, B)")
            return

        # Duplicate kontrol
        for r in range(self.tbl_codes.rowCount()):
            existing = (self.tbl_codes.item(r, 0).text() if self.tbl_codes.item(r, 0) else "").strip().upper()
            if existing == code:
                # Güncelle (desc/dur)
                self.tbl_codes.setItem(r, 1, QTableWidgetItem(desc))
                self.tbl_codes.setItem(r, 2, QTableWidgetItem(str(dur)))
                self._sync_active_code_combo()
                return

        r = self.tbl_codes.rowCount()
        self.tbl_codes.insertRow(r)
        self.tbl_codes.setItem(r, 0, QTableWidgetItem(code))
        self.tbl_codes.setItem(r, 1, QTableWidgetItem(desc))
        self.tbl_codes.setItem(r, 2, QTableWidgetItem(str(dur)))
        self._sync_active_code_combo()

        # Kullanıcı hızlı ekleme yapabilsin
        self.in_spot_code.clear()
        self.in_code_definition.clear()
        try:
            self.in_spot_duration.setValue(0)
        except Exception:
            pass

    def on_remove_code_def(self) -> None:
        row = self.tbl_codes.currentRow()
        if row < 0:
            return
        self.tbl_codes.removeRow(row)
        self._sync_active_code_combo()

    def on_active_code_changed(self) -> None:
        # Manuel grid girişinde zorunlu değil; opsiyonel kullanım.
        return

    def on_fill_selected_cells_with_active_code(self) -> None:
        """Opsiyonel: seçili hücrelere aktif kodu yazar (Excel 'doldur' gibi)."""
        code = (self.cb_active_code.currentText() or "").strip().upper()
        if not code:
            return

        # Day kolonları 2..N; satırlar 0..times
        table = self.plan_grid.table
        for it in table.selectedItems():
            r = it.row()
            c = it.column()
            if c < 2:
                continue
            # read-only modda yazma
            try:
                if self.plan_grid._read_only:
                    continue
            except Exception:
                pass
            it.setText(code)


    # -------------------------
    # Çoklu Kanal Seçimi (ANA SAYFA)
    # -------------------------
    def _update_selected_channels_label(self) -> None:
        names = [n for n in (self._selected_channel_names or []) if str(n).strip()]
        if not names:
            self.lbl_selected_channels.setText("")
            return

        # Hepsi seçili mi?
        try:
            all_names = [c["name"] for c in self.repo.list_channels(active_only=True)] if self.repo else []
        except Exception:
            all_names = []

        if all_names and len(names) == len(all_names):
            self.lbl_selected_channels.setText(f"Tümü ({len(names)})")
            return

        if len(names) <= 3:
            self.lbl_selected_channels.setText(", ".join(names))
        else:
            self.lbl_selected_channels.setText(f"{names[0]}, {names[1]}, {names[2]} ... (+{len(names)-3})")

    def on_select_channels(self) -> None:
        if not self.repo:
            QMessageBox.warning(self, "Hata", "DB hazır değil.")
            return

        from PySide6.QtWidgets import QDialog, QListWidget, QDialogButtonBox

        dlg = QDialog(self)
        dlg.setWindowTitle("Kanalları Seç")
        dlg.resize(420, 520)
        v = QVBoxLayout(dlg)

        lw = QListWidget(dlg)
        lw.setSelectionMode(QAbstractItemView.MultiSelection)
        v.addWidget(lw, 1)

        # Kanalları yükle
        channels = self.repo.list_channels(active_only=True)
        names = [str(c["name"]) for c in channels]
        for n in names:
            lw.addItem(n)

        # mevcut seçimleri işaretle
        selected_set = {s.strip().lower() for s in (self._selected_channel_names or [])}
        for i in range(lw.count()):
            it = lw.item(i)
            if it.text().strip().lower() in selected_set:
                it.setSelected(True)

        # Yardımcı butonlar: Tümü / Temizle
        row = QHBoxLayout()
        v.addLayout(row)
        btn_all = QPushButton("Tümü")
        btn_clear = QPushButton("Temizle")
        row.addWidget(btn_all)
        row.addWidget(btn_clear)
        row.addStretch(1)

        def _select_all():
            for i in range(lw.count()):
                lw.item(i).setSelected(True)

        def _clear():
            for i in range(lw.count()):
                lw.item(i).setSelected(False)

        btn_all.clicked.connect(_select_all)
        btn_clear.clicked.connect(_clear)

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        v.addWidget(bb)
        bb.accepted.connect(dlg.accept)
        bb.rejected.connect(dlg.reject)

        if dlg.exec() != QDialog.Accepted:
            return

        picked = [it.text().strip() for it in lw.selectedItems() if it.text().strip()]
        self._selected_channel_names = picked
        self._update_selected_channels_label()

        # UI kolaylığı: combo'da ilk seçiliyi göster
        if picked:
            idx = self.in_channel.findText(picked[0])
            if idx >= 0:
                self.in_channel.setCurrentIndex(idx)


    # -------------------------
    # Tarih Aralığı (multi-month)
    # -------------------------
    def on_apply_date_range(self) -> None:
        rs = self.in_range_start.date().toPython()
        re = self.in_range_end.date().toPython()
        if rs > re:
            # swap
            rs, re = re, rs
            self.in_range_start.setDate(QDate(rs.year, rs.month, rs.day))
            self.in_range_end.setDate(QDate(re.year, re.month, re.day))

        # state'e yaz (PlanningGrid kolon gizleme bu state'i kullanır)
        self._home_range_start = rs
        self._home_range_end = re

        # Aralık seçilince, tek tabloda tüm günler görünsün (ay kırılımı yok).
        # Not: Onayla anında ay bazında kaydetmeye yine devam ediyoruz.
        try:
            sel = self.in_date.date().toPython()
            if sel < rs or sel > re:
                sel = rs
                self.in_date.setDate(QDate(sel.year, sel.month, sel.day))

            # Eski aylık cache'leri temizle; artık edit tek grid üzerinde.
            self._month_cells_cache = {}
            self._current_month_key = (rs.year, rs.month)

            self.plan_grid.set_date_span(rs, re, selected_date=sel)
        except Exception:
            # Fallback: en azından eski davranış bozulmasın
            try:
                self.plan_grid.set_month(rs.year, rs.month, rs.day)
            except Exception:
                pass


    def on_save(self) -> None:
        """Tek buton akışı: seçili tarih aralığını DB'ye kaydet.

        Notlar:
        - Span modunda ekranda tek tabloda birden fazla ay görünebilir.
        - DB'de *tek kayıt* (kanal başına 1 kayıt) tutulur; ay ay bölünmez.
          Aylık raporlar/Plan Özet gibi ekranlar span kaydın ay kırılımlarını payload'dan okur.
        """

        if not self.service:
            QMessageBox.warning(self, "Hata", "Servis hazır değil (DB bağlantısı yok).")
            return

        if not getattr(self, "repo", None):
            QMessageBox.warning(self, "Hata", "Repo hazır değil (DB bağlantısı yok).")
            return

        try:
            # Rezervasyon Excel çıktıları bu klasöre üretilecek.
            template_path = self._resolve_template_path()
            out_dir = self.app_settings.data_dir / "exports"
            out_paths: list[Path] = []
            export_failures: list[str] = []

            # Spot saati kullanıcı girmiyor; SpotList+ zaten kuşak saatlerini kullanıyor.
            spot_t = datetime.now().time().replace(microsecond=0)

            # Ay bazında matrisleri al
            span_mode = bool(getattr(self.plan_grid, "is_span_mode", lambda: False)())
            if span_mode:
                month_matrices = self.plan_grid.get_span_month_matrices()
            else:
                d0 = self.in_date.date().toPython()
                month_matrices = {(d0.year, d0.month): self.plan_grid.get_matrix()}

            # Span modunda kullanıcı aralığı
            rs = self.in_range_start.date().toPython()
            re = self.in_range_end.date().toPython()

            # Kod tanımları
            code_defs = self._get_code_defs_from_ui()

            # Kanal seçimleri
            selected_channels = list(self._selected_channel_names)
            if not selected_channels:
                # combobox tek kanal seçimi fallback
                ch = self.in_channel.currentText().strip()
                if ch:
                    selected_channels = [ch]

            if not selected_channels:
                QMessageBox.warning(self, "Eksik", "En az 1 kanal seçmelisin.")
                return

            # Zorunlu başlık
            plan_title = self.in_plan_title.text().strip()
            if not plan_title:
                QMessageBox.warning(self, "Eksik", "Plan başlığı boş olamaz.")
                return

            # Kanal id map'i (fiyat çekmek için)
            channels = self.repo.list_channels(active_only=False)
            ch_by_norm = {str(c.get("name") or "").strip().lower(): c for c in channels}

            # Year bazlı fiyat haritaları cache'i
            price_maps: dict[int, dict[tuple[int, int], tuple[float, float]]] = {}

            # DB kayıtları
            created = []

            if span_mode:
                # Span: DB'de kanal başına tek kayıt; ay ay bölme YOK.
                span_month_matrices_db: dict[str, dict[str, str]] = {}
                merged_cells: dict[str, str] = {}

                for (yy, mm), matrix in sorted(month_matrices.items()):
                    # PlanningGrid returns month matrices as dicts of {"r,day": "CODE", ...}.
                    # Some legacy flows may wrap them as {"cells": {...}, "row_times": [...]}.
                    if isinstance(matrix, dict) and "cells" in matrix:
                        cells = matrix.get("cells") or {}
                    else:
                        cells = matrix or {}

                    if not any(str(v).strip() for v in cells.values()):
                        continue

                    span_month_matrices_db[f"{int(yy):04d}-{int(mm):02d}"] = dict(cells)

                    # confirm() için çakışmasız (row,YYYYMMDD) key ile merge
                    for k, v in cells.items():
                        vv = str(v or "").strip()
                        if not vv:
                            continue
                        try:
                            row_s, day_s = str(k).split(",", 1)
                            row_idx = int(row_s)
                            day = int(day_s)
                        except Exception:
                            continue
                        merged_cells[f"{row_idx},{int(yy):04d}{int(mm):02d}{int(day):02d}"] = vv

                if not span_month_matrices_db:
                    QMessageBox.information(self, "Bilgi", "Kaydedilecek bir hücre bulunamadı.")
                    return

                # Başlangıç yılındaki fiyat haritası (draft için)
                if rs.year not in price_maps:
                    try:
                        price_maps[rs.year] = self.repo.get_channel_prices(rs.year)
                    except Exception:
                        price_maps[rs.year] = {}

                for ch_name in selected_channels:
                    ch = ch_by_norm.get(ch_name.strip().lower())
                    ch_id = int(ch.get("id")) if ch else None
                    dt_price = 0.0
                    odt_price = 0.0
                    if ch_id is not None:
                        dt_price, odt_price = price_maps[rs.year].get((ch_id, int(rs.month)), (0.0, 0.0))

                    draft = ReservationDraft(
                        advertiser_name=self.in_advertiser.text().strip(),
                        plan_date=rs,  # referans tarih: aralığın başlangıcı
                        spot_time=spot_t,
                        channel_name=ch_name,
                        channel_price_dt=float(dt_price),
                        channel_price_odt=float(odt_price),
                        agency_name=self.in_agency.text().strip(),
                        product_name=self.in_product.text().strip(),
                        plan_title=plan_title,
                        spot_code=self.in_spot_code.text().strip(),
                        spot_duration_sec=int(self.in_spot_duration.value()),
                        code_definition=self.in_code_definition.text().strip(),
                        note_text=self.in_note.text().strip(),
                        prepared_by_name=self.in_prepared_by.text().strip(),
                        code_defs=code_defs,
                        agency_commission_pct=int(self.in_agency_commission.value()),
                    )

                    confirmed = self.service.confirm(draft, merged_cells)

                    payload2 = dict(confirmed.payload or {})
                    payload2["reservation_no"] = None  # create_reservation dolduracak
                    payload2["created_at"] = None
                    payload2["is_span"] = True
                    payload2["span_start"] = rs.isoformat()
                    payload2["span_end"] = re.isoformat()
                    payload2["span_month_matrices"] = span_month_matrices_db

                    dbrec = self.repo.create_reservation(
                        advertiser_name=str(payload2.get("advertiser_name") or "").strip(),
                        payload=payload2,
                        confirmed=True,
                    )
                    created.append(dbrec)

            else:
                # Month: eski davranış (ay + kanal bazlı kayıt)
                for (yy, mm), matrix in sorted(month_matrices.items()):
                    if isinstance(matrix, dict) and "cells" in matrix:
                        cells = matrix.get("cells") or {}
                    else:
                        cells = matrix or {}

                    if not any(str(v).strip() for v in cells.values()):
                        continue

                    if yy not in price_maps:
                        try:
                            price_maps[yy] = self.repo.get_channel_prices(yy)
                        except Exception:
                            price_maps[yy] = {}

                    plan_date = date(yy, mm, 1)

                    for ch_name in selected_channels:
                        ch = ch_by_norm.get(ch_name.strip().lower())
                        ch_id = int(ch.get("id")) if ch else None
                        dt_price = 0.0
                        odt_price = 0.0
                        if ch_id is not None:
                            dt_price, odt_price = price_maps[yy].get((ch_id, int(mm)), (0.0, 0.0))

                        draft = ReservationDraft(
                            advertiser_name=self.in_advertiser.text().strip(),
                            plan_date=plan_date,
                            spot_time=spot_t,
                            channel_name=ch_name,
                            channel_price_dt=float(dt_price),
                            channel_price_odt=float(odt_price),
                            agency_name=self.in_agency.text().strip(),
                            product_name=self.in_product.text().strip(),
                            plan_title=plan_title,
                            spot_code=self.in_spot_code.text().strip(),
                            spot_duration_sec=int(self.in_spot_duration.value()),
                            code_definition=self.in_code_definition.text().strip(),
                            note_text=self.in_note.text().strip(),
                            prepared_by_name=self.in_prepared_by.text().strip(),
                            code_defs=code_defs,
                        agency_commission_pct=int(self.in_agency_commission.value()),
                        )

                        confirmed = self.service.confirm(draft, cells)
                        dbrec = self.repo.create_reservation(
                            advertiser_name=str(confirmed.payload.get("advertiser_name") or "").strip(),
                            payload=confirmed.payload,
                            confirmed=True,
                        )
                        created.append(dbrec)

                        # Month modunda eski tek dosya davranışı:
                        try:
                            from src.export.excel_exporter import export_excel

                            out_path = out_dir / f"{dbrec.reservation_no}.xlsx"
                            payload2 = dict(dbrec.payload or {})
                            payload2["reservation_no"] = dbrec.reservation_no
                            payload2["created_at"] = dbrec.created_at
                            export_excel(template_path, out_path, payload2)
                            out_paths.append(out_path)
                        except Exception as ex2:
                            export_failures.append(f"{dbrec.reservation_no}: {ex2}")

            if not created:
                QMessageBox.information(self, "Bilgi", "Kaydedilecek bir hücre bulunamadı.")
                return

            # Span modunda: seçili tarih aralığına göre *tek dosya* rezervasyon çıktısı üret.
            # Export: Span modunda tek dosya (çok sheet) + kanal bazında ayrı dosya
            if span_mode:
                try:
                    from src.export.excel_exporter import export_excel_span

                    for r in created:
                        payload_span = dict(getattr(r, "payload", {}) or {})
                        payload_span["reservation_no"] = str(getattr(r, "reservation_no", "") or "").strip()
                        payload_span["created_at"] = getattr(r, "created_at", None)

                        fname = f"{payload_span['reservation_no']}.xlsx"
                        out_path = out_dir / fname
                        export_excel_span(
                            template_path=template_path,
                            out_path=out_path,
                            payload=payload_span,
                            month_matrices=month_matrices,
                            span_start=rs,
                            span_end=re,
                        )
                        out_paths.append(out_path)
                except Exception as exs:
                    export_failures.append(f"SPAN_EXPORT: {exs}")

            QMessageBox.information(
                self,
                "Kaydedildi",
                f"DB'ye {len(created)} kayıt eklendi.\n"
                f"(Kanal x Ay kırılımı)\n\n"
                f"Excel çıktısı: {len(out_paths)} adet\n"
                f"Klasör: {out_dir}\n\n"
                f"İpucu: Rezervasyonlar sekmesinden filtreleyebilirsin.",
            )

            # Export hataları varsa kullanıcıyı bilgilendir.
            if export_failures:
                QMessageBox.warning(
                    self,
                    "Excel çıktısı üretilemedi",
                    "Bazı kayıtlar DB'ye yazıldı ama Excel çıktısı üretilemedi:\n\n"
                    + "\n".join(export_failures[:15])
                    + ("\n..." if len(export_failures) > 15 else ""),
                )

        except Exception as ex:
            QMessageBox.critical(self, "Hata", str(ex))


    def on_confirm(self) -> None:
        if not self.service:
            QMessageBox.warning(self, "Hata", "Servis hazır değil (DB bağlantısı yok).")
            return

        try:
            # Spot saati kullanıcı girmiyor; Excel zaten çıktıda zaman damgasını basıyor.
            spot_t = datetime.now().time().replace(microsecond=0)

            # Grid içeriğini ay bazında matrise çevir.
            # Span modunda tek tabloda birden fazla ay görünür; burada ay ay ayrıştırıyoruz.
            try:
                if getattr(self.plan_grid, "is_span_mode", lambda: False)():
                    self._month_cells_cache = self.plan_grid.get_span_month_matrices()
                else:
                    d0 = self.in_date.date().toPython()
                    self._month_cells_cache = {(d0.year, d0.month): self.plan_grid.get_matrix()}
                    self._current_month_key = (d0.year, d0.month)
            except Exception:
                # Worst-case: boş cache ile devam et
                self._month_cells_cache = {}

            # Kod tanımları
            code_defs = self._get_code_defs_from_ui()

            # Kanal seçimleri
            selected_channels = list(self._selected_channel_names)
            if not selected_channels:
                selected_channels = [self.in_channel.currentText().strip()]

            # Kanal adından id mapping
            channels = self.repo.list_channels()
            ch_by_norm = {str(c.get("name", "")).strip().lower(): c for c in channels}

            # Tarih aralığı (ay bazında)
            rs = self.in_range_start.date().toPython()
            re = self.in_range_end.date().toPython()
            if rs > re:
                rs, re = re, rs

            def month_iter(a: date, b: date) -> list[tuple[int, int]]:
                out: list[tuple[int, int]] = []
                yy, mm = a.year, a.month
                end = (b.year, b.month)
                while (yy, mm) <= end:
                    out.append((yy, mm))
                    mm += 1
                    if mm == 13:
                        mm = 1
                        yy += 1
                return out

            months_in_range = set(month_iter(rs, re))

            # Year bazlı fiyat haritaları cache'i
            price_maps: dict[int, dict[tuple[int, int], tuple[float, float]]] = {}

            created: list = []
            for (yy, mm), cells in list(self._month_cells_cache.items()):
                if (yy, mm) not in months_in_range:
                    continue

                # Bu ayda hiç dolu hücre yoksa atla
                if not any(str(v).strip() for v in (cells or {}).values()):
                    continue

                if yy not in price_maps:
                    try:
                        price_maps[yy] = self.repo.get_channel_prices(yy)
                    except Exception:
                        price_maps[yy] = {}

                for ch_name in selected_channels:
                    ch = ch_by_norm.get(ch_name.strip().lower())
                    ch_id = int(ch.get("id")) if ch else None
                    dt_price = 0.0
                    odt_price = 0.0
                    if ch_id is not None:
                        dt_price, odt_price = price_maps[yy].get((ch_id, int(mm)), (0.0, 0.0))

                    draft = ReservationDraft(
                        advertiser_name=self.in_advertiser.text(),
                        plan_date=date(yy, mm, 1),
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
                        code_defs=code_defs,
                        agency_commission_pct=int(self.in_agency_commission.value()),
                    )

                    created.append(self.service.confirm(draft, cells))

            if not created:
                QMessageBox.information(self, "Bilgi", "Seçili tarih aralığında kaydedilecek plan bulunamadı.")
                return

            # Son kaydı aktif kabul edip butonları enable et
            self.current_confirmed = created[-1]
            # Legacy buttons removed (MVP simplification): Onayla/Test/Çıktı.
            # Keep this block harmless in case the method is re-used later.
            if hasattr(self, "btn_test_export"):
                self.btn_test_export.setEnabled(True)
            if hasattr(self, "btn_save_export"):
                self.btn_save_export.setEnabled(True)

            QMessageBox.information(self, "OK", "Onaylandı. Artık test/kayıt çıktısı alabilirsin.")
        except Exception as e:
            QMessageBox.warning(self, "Hata", str(e))


    def _resolve_template_path(self) -> Path:
        """Rezervasyon Excel şablonunu çöz.

        Not: Uygulama farklı CWD ile açılabildiği ve PyInstaller'da dosyalar _MEIPASS altına
        çıktığı için resource_path fallback'i yapıyoruz.
        """
        from src.util.paths import resource_path

        tp = getattr(self.app_settings, "template_path", None)
        if tp:
            p = Path(tp)
            if p.exists():
                return p
            # relatifse proje köküne göre dene
            p2 = resource_path(str(p).replace('\\', '/'))
            return p2

        # default: assets/reservation_template.xlsx
        p = Path("assets") / "reservation_template.xlsx"
        if p.exists():
            return p
        return resource_path("assets/reservation_template.xlsx")


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

    def reset_form_for_new_reservation(self) -> None:
        self._loaded_reservation_id = None
        self.in_advertiser.clear()
        self.in_agency.clear()
        self.in_product.clear()
        self.in_plan_title.clear()
        self.in_spot_code.clear()
        self.in_code_definition.clear()
        self.in_note.clear()
        self.in_prepared_by.clear()

        self.in_spot_duration.setValue(0)
        self.in_agency_commission.setValue(10)

        today = QDate.currentDate()
        self.in_date.setDate(today)
        first = QDate(today.year(), today.month(), 1)
        last = QDate(today.year(), today.month(), today.daysInMonth())
        self.in_range_start.setDate(first)
        self.in_range_end.setDate(last)

        self.refresh_channel_combo()
        if self.in_channel.count() > 0:
            self.in_channel.setCurrentIndex(0)

        if hasattr(self, 'selected_channels'):
            self.selected_channels = []
            try:
                self._update_selected_channels_label()
            except Exception:
                pass

        try:
            self.plan_grid.clear_matrix()
        except Exception:
            pass

        try:
            self.kod_table.setRowCount(0)
        except Exception:
            pass

    def on_tab_changed(self, idx: int) -> None:
        tab_name = self.tabs.tabText(idx)
        if tab_name == "REZERVASYONLAR":
            self.refresh_reservations_tab()
        if tab_name == "KOD TANIMI":
            self.refresh_kod_tanimi()
        elif tab_name == "SPOTLİST+":
            self.refresh_spotlist()
        elif tab_name == "PLAN ÖZET":
            self._set_plan_ozet_period_from_latest()
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

        self.btn_res_export_selected = QPushButton("Seçileni Excel'e Aktar")
        self.btn_res_export_filtered = QPushButton("Filtredekileri Excel'e Aktar")
        self.btn_res_open_exports = QPushButton("Exports Klasörünü Aç")

        top.addWidget(self.btn_res_refresh)
        top.addWidget(self.btn_res_open)
        top.addWidget(self.btn_res_delete)
        top.addWidget(self.btn_res_export_selected)
        top.addWidget(self.btn_res_export_filtered)
        top.addWidget(self.btn_res_open_exports)
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
        self.res_table.setColumnCount(8)
        self.res_table.setHorizontalHeaderLabels(
            [
                "Rezervasyon No",
                "Plan Başlığı",
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
        self.btn_res_export_selected.clicked.connect(self.on_reservation_export_selected)
        self.btn_res_export_filtered.clicked.connect(self.on_reservation_export_filtered)
        self.btn_res_open_exports.clicked.connect(self.open_exports_folder)
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

        pt = (self.in_plan_title.text() or "").strip()
        self._res_records = []
        self.res_table.setRowCount(0)
        self.res_preview_title.setText("")
        try:
            self.res_preview_grid.clear_matrix()
        except Exception:
            pass

        if not pt:
            return

        year = int(self.res_year.value())
        month = int(self.res_month.currentData() or 0)
        ch_filter = str(self.res_channel.currentData() or "").strip()

        recs = self.repo.list_confirmed_reservations_by_plan_title(pt, limit=50000)
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
            plan_title = str(p.get("plan_title") or "")
            if bool(p.get("is_span")) and p.get("span_start") and p.get("span_end"):
                plan_date = f"{p.get('span_start')} - {p.get('span_end')}"
            else:
                plan_date = str(p.get("plan_date") or "")
            channel = str(p.get("channel_name") or "")
            spot_code = str(p.get("spot_code") or "")
            duration = str(p.get("spot_duration_sec") or "")
            adet = str(p.get("adet_total") or "")
            created = str(r.created_at or "")

            for col, val in enumerate([res_no, plan_title, plan_date, channel, spot_code, duration, adet, created]):
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
        plan_title = str(p.get("plan_title") or "")
        plan_date = str(p.get("plan_date") or "")
        is_span = bool(p.get("is_span"))
        if is_span and p.get("span_start") and p.get("span_end"):
            title_date = f"{p.get('span_start')} - {p.get('span_end')}"
        else:
            title_date = plan_date
        self.res_preview_title.setText(f"{res_no}  |  {channel}  |  {title_date}")

        try:
            if is_span and p.get("span_start"):
                y, m, d = str(p.get("span_start")).split("-")
                self.res_preview_grid.set_month(int(y), int(m), int(d))
                raw = p.get("span_month_matrices") or {}
                month_mats = {}
                for k, cells in raw.items():
                    if isinstance(k, str) and "-" in k:
                        yy, mm = k.split("-", 1)
                        month_mats[(int(yy), int(mm))] = cells or {}
                    elif isinstance(k, (tuple, list)) and len(k) == 2:
                        month_mats[(int(k[0]), int(k[1]))] = cells or {}
                if month_mats:
                    self.res_preview_grid.set_span_month_matrices(month_mats)
                else:
                    self.res_preview_grid.set_matrix(p.get("plan_cells") or {})
            else:
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

        pb_raw = str(p.get("prepared_by", "") or "").strip()
        pb_name = pb_raw.split(" - ", 1)[0].strip() if pb_raw else ""
        self.in_prepared_by.setText(pb_name)

        try:
            self.in_spot_duration.setValue(int(p.get("spot_duration_sec", 0) or 0))
        except Exception:
            pass

        # Tarih/range geri yükleme (span desteği)
        try:
            is_span = bool(p.get("is_span"))
            if is_span and p.get("span_start") and p.get("span_end"):
                ds = datetime.fromisoformat(str(p.get("span_start"))).date()
                de = datetime.fromisoformat(str(p.get("span_end"))).date()

                self.in_range_start.blockSignals(True)
                self.in_range_end.blockSignals(True)
                self.in_range_start.setDate(QDate(ds.year, ds.month, ds.day))
                self.in_range_end.setDate(QDate(de.year, de.month, de.day))
                self.in_range_start.blockSignals(False)
                self.in_range_end.blockSignals(False)

                self.in_date.blockSignals(True)
                self.in_date.setDate(QDate(ds.year, ds.month, ds.day))
                self.in_date.blockSignals(False)

                try:
                    self.on_apply_date_range()
                except Exception:
                    pass
            else:
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

        # Kod/Kod Tanımı/Süreleri geri yükle
        try:
            self.tbl_codes.setRowCount(0)
            code_defs = list(p.get("code_defs") or [])
            for cd in code_defs:
                code = str((cd or {}).get("code") or "").strip().upper()
                desc = str((cd or {}).get("desc") or "").strip()
                dur = (cd or {}).get("duration_sec", 0)
                try:
                    dur = int(float(dur or 0))
                except Exception:
                    dur = 0
                if not code:
                    continue
                r = self.tbl_codes.rowCount()
                self.tbl_codes.insertRow(r)
                self.tbl_codes.setItem(r, 0, QTableWidgetItem(code))
                self.tbl_codes.setItem(r, 1, QTableWidgetItem(desc))
                self.tbl_codes.setItem(r, 2, QTableWidgetItem(str(dur)))

            if code_defs:
                first = code_defs[0] or {}
                self.in_spot_code.setText(str(first.get("code") or "").strip().upper())
                self.in_code_definition.setText(str(first.get("desc") or "").strip())
                try:
                    self.in_spot_duration.setValue(int(float(first.get("duration_sec") or 0)))
                except Exception:
                    pass
            self._sync_active_code_combo()
        except Exception:
            pass

        try:
            self.plan_grid.set_read_only(False)
            if bool(p.get("is_span")):
                raw = p.get("span_month_matrices") or {}
                month_mats = {}
                for k, cells in raw.items():
                    if isinstance(k, str) and "-" in k:
                        yy, mm = k.split("-", 1)
                        month_mats[(int(yy), int(mm))] = cells or {}
                    elif isinstance(k, (tuple, list)) and len(k) == 2:
                        month_mats[(int(k[0]), int(k[1]))] = cells or {}
                if month_mats:
                    self.plan_grid.set_span_month_matrices(month_mats)
                else:
                    self.plan_grid.set_matrix(p.get("plan_cells") or {})
            else:
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

    # ------------------------------
    # Rezervasyon Excel Export (REZERVASYONLAR sekmesi)
    # ------------------------------

    def open_exports_folder(self) -> None:
        """Exports klasörünü açar."""
        try:
            out_dir = (self.app_settings.data_dir / "exports")
            out_dir.mkdir(parents=True, exist_ok=True)

            from PySide6.QtGui import QDesktopServices
            from PySide6.QtCore import QUrl

            QDesktopServices.openUrl(QUrl.fromLocalFile(str(out_dir)))
        except Exception as ex:
            QMessageBox.warning(self, "Hata", f"Exports klasörü açılamadı: {ex}")

    def _export_reservation_records(self, recs: list) -> None:
        if not recs:
            QMessageBox.information(self, "Bilgi", "Export edilecek kayıt yok.")
            return

        try:
            template_path = self._resolve_template_path()
            out_dir = self.app_settings.data_dir / "exports"
            out_dir.mkdir(parents=True, exist_ok=True)

            from src.export.excel_exporter import export_excel

            ok_paths: list[Path] = []
            failures: list[str] = []

            for r in recs:
                res_no = str(getattr(r, "reservation_no", "") or "").strip()
                if not res_no:
                    failures.append("(rezervasyon no yok) -> export atlandı")
                    continue

                payload2 = dict(getattr(r, "payload", {}) or {})
                payload2["reservation_no"] = res_no
                payload2["created_at"] = str(getattr(r, "created_at", "") or "")

                out_path = out_dir / f"{res_no}.xlsx"
                try:
                    if bool(payload2.get("is_span")) and payload2.get("span_start") and payload2.get("span_end"):
                        from src.export.excel_exporter import export_excel_span
                        raw = payload2.get("span_month_matrices") or {}
                        month_matrices = {}
                        for k, cells in raw.items():
                            if isinstance(k, str) and "-" in k:
                                yy, mm = k.split("-", 1)
                                month_matrices[(int(yy), int(mm))] = cells or {}
                        from datetime import date as _d
                        ys, ms, ds = str(payload2.get("span_start")).split("-")
                        ye, me, de = str(payload2.get("span_end")).split("-")
                        export_excel_span(
                            template_path=template_path,
                            out_path=out_path,
                            payload=payload2,
                            month_matrices=month_matrices,
                            span_start=_d(int(ys), int(ms), int(ds)),
                            span_end=_d(int(ye), int(me), int(de)),
                        )
                    else:
                        export_excel(template_path, out_path, payload2)
                    ok_paths.append(out_path)
                except Exception as ex2:
                    failures.append(f"{res_no}: {ex2}")

            # Sonuç mesajı
            msg = (
                f"Excel çıktısı üretildi: {len(ok_paths)} adet\n"
                f"Klasör: {out_dir}\n\n"
                "Not: Aynı rezervasyon no ile dosya varsa üzerine yazar."
            )
            QMessageBox.information(self, "Export", msg)

            if failures:
                QMessageBox.warning(
                    self,
                    "Bazı dosyalar üretilemedi",
                    "Aşağıdaki kayıt(lar) için export başarısız oldu:\n\n"
                    + "\n".join(failures[:20])
                    + ("\n..." if len(failures) > 20 else ""),
                )

        except Exception as ex:
            QMessageBox.critical(self, "Hata", str(ex))

    def on_reservation_export_selected(self) -> None:
        recs = self._get_selected_reservation_records()
        if not recs:
            QMessageBox.information(self, "Bilgi", "Önce export etmek için satır seç.")
            return
        self._export_reservation_records(recs)

    def on_reservation_export_filtered(self) -> None:
        recs = list(getattr(self, "_res_records", []) or [])
        if not recs:
            QMessageBox.information(self, "Bilgi", "Liste boş. Önce filtrele / yenile.")
            return
        self._export_reservation_records(recs)

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
        # Not: SPOTLİST+ artık reklamverene göre değil, plan başlığına göre çalışıyor.
        # Bu state değişkeni de "mevcut plan başlığı" olarak kullanılır.
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

        pt = (self.in_plan_title.text() or "").strip()
        if not pt:
            self.spot_table.setRowCount(0)
            self.spot_summary.setText("")
            return

        # plan başlığı değiştiyse dirty temizle (karışmasın)
        if pt != self.spot_current_adv:
            self.spot_dirty.clear()
            self.btn_spot_save.setEnabled(False)
            self.spot_filters_initialized = False
            self.spot_current_adv = pt

        self.spot_all_rows = self.service.get_spotlist_rows(pt)

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
        self.spot_table.setUpdatesEnabled(False)
        self.spot_table.blockSignals(True)
        self.spot_table.setSortingEnabled(False)
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


        self.spot_table.blockSignals(False)
        self.spot_table.setUpdatesEnabled(True)
        try:
            self.spot_table.setSortingEnabled(True)
        except Exception:
            pass

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
        pt = (self.in_plan_title.text() or "").strip()
        if not pt:
            QMessageBox.warning(self, "Hata", "Önce bir plan başlığı seç.")
            return

        safe_pt = "".join(ch if ch.isalnum() or ch in ("-", "_", " ") else "_" for ch in pt).strip()
        safe_pt = safe_pt.replace(" ", "_")[:60] or "PLAN"
        default_name = f"SPOTLIST_{safe_pt}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
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
            self.service.export_spotlist_excel_with_rows(path, pt, rows)
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
        # "TÜMÜ" seçeneği: seçili plan başlığının yıl içindeki tüm aylarını tek ekranda (yıllık) özetler.
        self.po_month.addItems([
            "TÜMÜ",
            "OCAK", "ŞUBAT", "MART", "NİSAN", "MAYIS", "HAZİRAN",
            "TEMMUZ", "AĞUSTOS", "EYLÜL", "EKİM", "KASIM", "ARALIK"
        ])
        # Varsayılan: içinde bulunulan ay
        self.po_month.setCurrentIndex(date.today().month)  # +1 offset (0: TÜMÜ)
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

    def _set_plan_ozet_period_from_latest(self, *_args, **_kwargs) -> None:
        """Plan Özet üstündeki 'Dönemi' alanını rezervasyon tabındaki tarih aralığına göre doldurur.

        Not: Geçmişte bazı yerlerde bu metoda yanlışlıkla parametre gönderilmişti.
        Güvenli olması için *args/**kwargs kabul ediyor.
        """
        # Öncelik: Rezervasyon tabındaki tarih aralığı (widget'lar mevcutsa)
        rs = None
        re_ = None
        try:
            if hasattr(self, "in_range_start") and hasattr(self, "in_range_end"):
                qd1 = self.in_range_start.date()
                qd2 = self.in_range_end.date()
                rs = date(qd1.year(), qd1.month(), qd1.day())
                re_ = date(qd2.year(), qd2.month(), qd2.day())
        except Exception:
            rs = None
            re_ = None

        # İkinci öncelik: daha önce uygulanan aralık
        if rs is None or re_ is None:
            rs = getattr(self, "_home_range_start", None)
            re_ = getattr(self, "_home_range_end", None)

        # Son çare: Plan Özet yıl/ay seçicileri
        if rs is None or re_ is None:
            try:
                y = int(self.po_year.value())
                m_ix = int(self.po_month.currentIndex()) if hasattr(self, "po_month") else 0
                if m_ix <= 0:
                    rs = date(y, 1, 1)
                    re_ = date(y, 12, 31)
                else:
                    rs = date(y, m_ix, 1)
                    if m_ix == 12:
                        re_ = date(y + 1, 1, 1) - timedelta(days=1)
                    else:
                        re_ = date(y, m_ix + 1, 1) - timedelta(days=1)
            except Exception:
                rs = None
                re_ = None

        # UI
        if hasattr(self, "po_period"):
            if rs is None or re_ is None:
                self.po_period.setText("")
            else:
                # ayrıca sakla (Plan Özet / Excel export aynı aralığı kullansın)
                self._home_range_start = rs
                self._home_range_end = re_
                self.po_period.setText(f"{rs:%d.%m.%Y}-{re_:%d.%m.%Y}")

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
        """Plan Özet tablosunu tarih aralığına göre (tek tip) yeniler."""
        if not hasattr(self, "po_table"):
            return

        pt = (self.in_plan_title.text() or "").strip() if hasattr(self, "in_plan_title") else ""

        # Plan Özet'in dönemi her zaman rezervasyon aralığına bağlı.
        # Kullanıcı aralık seçip "Aralığı Uygula"'ya basmasa bile, rezervasyon tabındaki
        # tarih editörlerinden güncel aralığı çekiyoruz.
        try:
            self._set_plan_ozet_period_from_latest()
        except Exception:
            pass

        rs = getattr(self, "_home_range_start", None)
        re_ = getattr(self, "_home_range_end", None)

        if not pt or rs is None or re_ is None:
            # tabloyu temizle
            self.po_table.setRowCount(0)
            self.po_table.setColumnCount(0)

            # üst bilgileri de temizle
            if hasattr(self, "po_agency"): self.po_agency.setText("")
            if hasattr(self, "po_advertiser"): self.po_advertiser.setText("")
            if hasattr(self, "po_product"): self.po_product.setText("")
            if hasattr(self, "po_plan_title"): self.po_plan_title.setText(pt or "")
            if hasattr(self, "po_resno"): self.po_resno.setPlainText("")
            if hasattr(self, "po_period"): self.po_period.setText("")
            if hasattr(self, "po_spot_len"): self.po_spot_len.setText("")
            return

        try:
            data = self.service.get_plan_ozet_range_data(pt, rs, re_)
        except Exception as e:
            QMessageBox.critical(self, "Hata", str(e))
            return

        header = data.get("header") or {}
        rows = data.get("rows") or []
        totals = data.get("totals") or {}
        dates = data.get("dates") or []
        months = data.get("months") or []

        # Üst bilgiler
        if hasattr(self, "po_agency"):
            self.po_agency.setText(str(header.get("agency", "") or ""))
        if hasattr(self, "po_advertiser"):
            self.po_advertiser.setText(str(header.get("advertiser", "") or ""))
        if hasattr(self, "po_product"):
            self.po_product.setText(str(header.get("product", "") or ""))
        if hasattr(self, "po_plan_title"):
            # header boş gelse bile seçili plan başlığını gösterelim
            self.po_plan_title.setText(str(header.get("plan_title", "") or pt or ""))
        if hasattr(self, "po_resno"):
            self.po_resno.setPlainText(str(header.get("reservation_no", "") or ""))
        if hasattr(self, "po_period"):
            self.po_period.setText(str(header.get("period", "") or ""))
        if hasattr(self, "po_spot_len"):
            self.po_spot_len.setText(str(header.get("spot_len", "") or ""))

        # Gün başlıkları: Pş / Çr dahil
        TR_DOW_UI = ["Pt", "Sa", "Çr", "Pş", "Cu", "Ct", "Pa"]
        day_headers = [f"{TR_DOW_UI[d.weekday()]}\n{d:%d.%m}" for d in dates]

        # Ay başlıkları
        MONTHS_TR = ["OCAK","ŞUBAT","MART","NİSAN","MAYIS","HAZİRAN","TEMMUZ","AĞUSTOS","EYLÜL","EKİM","KASIM","ARALIK"]
        month_headers = []
        for (yy, mm) in months:
            mn = MONTHS_TR[int(mm) - 1]
            month_headers.append(f"{mn} Adet")
            month_headers.append(f"{mn} Saniye")

        headers = ["Kanal", "Yayın Grubu", "DT/ODT", "Dinlenme Oranı"] + day_headers + month_headers + ["Birim sn. (TL)", "Toplam Bütçe\nNet TL"]

        self.po_table.blockSignals(True)
        try:
            self.po_table.setColumnCount(len(headers))
            self.po_table.setHorizontalHeaderLabels(headers)
            self.po_table.setRowCount(len(rows) + 1)  # +Toplam
            self.po_table.verticalHeader().setVisible(False)

            def _item(val, editable: bool = False) -> QTableWidgetItem:
                it = QTableWidgetItem("" if val is None else str(val))
                if not editable:
                    it.setFlags(it.flags() & ~Qt.ItemIsEditable)
                return it

            # satırlar
            for r_i, rr in enumerate(rows):
                self.po_table.setItem(r_i, 0, _item(rr.get("channel", "")))
                # yayın grubu editable
                pg_it = QTableWidgetItem(str(rr.get("publish_group", "") or ""))
                self.po_table.setItem(r_i, 1, pg_it)

                self.po_table.setItem(r_i, 2, _item(rr.get("dt_odt", "")))
                self.po_table.setItem(r_i, 3, _item(rr.get("dinlenme_orani", "NA")))

                # günler
                day_vals = rr.get("days") or []
                base = 4
                for j in range(len(day_headers)):
                    v = day_vals[j] if j < len(day_vals) else ""
                    self.po_table.setItem(r_i, base + j, _item("" if v in ("", None) else v))

                # ay kolonları
                mcols = rr.get("month_cols") or []
                mbase = base + len(day_headers)
                for j in range(len(month_headers)):
                    v = mcols[j] if j < len(mcols) else ""
                    self.po_table.setItem(r_i, mbase + j, _item("" if v in ("", None) else v))

                # unit + budget
                ucol = mbase + len(month_headers)
                bcol = ucol + 1
                self.po_table.setItem(r_i, ucol, _item("" if rr.get("unit_price", "") in ("", None) else rr.get("unit_price")))
                self.po_table.setItem(r_i, bcol, _item("" if rr.get("budget", "") in ("", None) else rr.get("budget")))

            # toplam satırı
            tr = len(rows)
            self.po_table.setItem(tr, 0, _item("Toplam"))
            self.po_table.setItem(tr, 1, _item(""))
            self.po_table.setItem(tr, 2, _item(""))
            self.po_table.setItem(tr, 3, _item(""))

            tdays = totals.get("days") or []
            base = 4
            for j in range(len(day_headers)):
                v = tdays[j] if j < len(tdays) else ""
                self.po_table.setItem(tr, base + j, _item("" if v in ("", None) else v))

            tm = totals.get("month_cols") or []
            mbase = base + len(day_headers)
            for j in range(len(month_headers)):
                v = tm[j] if j < len(tm) else ""
                self.po_table.setItem(tr, mbase + j, _item("" if v in ("", None) else v))

            ucol = mbase + len(month_headers)
            bcol = ucol + 1
            self.po_table.setItem(tr, ucol, _item(""))
            self.po_table.setItem(tr, bcol, _item("" if totals.get("budget", "") in ("", None) else totals.get("budget")))

        finally:
            self.po_table.blockSignals(False)

        try:
            self.po_table.resizeColumnsToContents()
        except Exception:
            pass

    def export_plan_ozet_excel(self) -> None:
        if not getattr(self, "service", None):
            return
        pt = (self.in_plan_title.text() or "").strip()
        if not pt:
            return

        rs = getattr(self, "_home_range_start", None)
        re_ = getattr(self, "_home_range_end", None)
        if rs is None or re_ is None:
            return

        out_dir = self.app_settings.data_dir / "exports"
        out_dir.mkdir(parents=True, exist_ok=True)

        default_name = f"PLAN_OZET_{pt}_{rs:%Y%m%d}_{re_:%Y%m%d}.xlsx"
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
            self.service.export_plan_ozet_range_excel(path, pt, rs, re_)
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

    # Backward-compatible aliases (old method names used in earlier UI wiring)
    def refresh_plan_summary_tab(self) -> None:
        self.refresh_plan_ozet()

    def refresh_code_def_tab(self) -> None:
        self.refresh_kod_tanimi()


    def refresh_kod_tanimi(self) -> None:
        if not self.service:
            return
        pt = self.in_plan_title.text().strip()
        if not pt:
            return

        rows = self.service.get_kod_tanimi_rows(pt)
        avg_len = self.service.get_kod_tanimi_avg_len(pt)

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
        pt = self.in_plan_title.text().strip()
        row = self.kod_table.currentRow()
        if row < 0:
            return
        code_item = self.kod_table.item(row, 0)
        if not code_item:
            return
        code = code_item.text().strip()
        if not code or code == "Ort.Uzun.":
            return

        deleted = self.service.delete_kod_for_plan_title(pt, code)
        QMessageBox.information(self, "OK", f"{code} koduna ait {deleted} kayıt silindi.")
        self.refresh_kod_tanimi()

    def export_kod_tanimi_excel(self) -> None:
        if not self.service:
            return
        pt = self.in_plan_title.text().strip()
        if not pt:
            return

        path, _ = QFileDialog.getSaveFileName(self, "KOD TANIMI Excel", f"{pt}_KOD_TANIMI.xlsx", "Excel Files (*.xlsx)")
        if not path:
            return

        self.service.export_kod_tanimi_excel(path, pt)
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

        self.price_table = ExcelTableWidget()
        self.price_table.setAlternatingRowColors(True)
        self.price_table.setSortingEnabled(True)
        self.price_table.horizontalHeader().setSortIndicatorShown(True)
        self.price_table.horizontalHeader().setSectionsClickable(True)
        # Excel gibi satır/kolon/blok seçimi + kopyala/yapıştır
        self.price_table.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.price_table.setSelectionMode(QAbstractItemView.ExtendedSelection)
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

        # Kanal adı genişleyebilir; fiyat kolonlarını dar tutuyoruz ki mümkün olduğunca yatay scroll istemesin.
        self.price_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        for c in range(1, len(headers)):
            self.price_table.horizontalHeader().setSectionResizeMode(c, QHeaderView.Fixed)
            self.price_table.setColumnWidth(c, 55)

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
                try:
                    it_dt.setData(Qt.EditRole, float(dt))
                except Exception:
                    pass
                it_dt.setTextAlignment(Qt.AlignCenter)
                self.price_table.setItem(r, col, it_dt)
                col += 1

                it_odt = QTableWidgetItem("" if odt == 0 else f"{odt:g}")
                try:
                    it_odt.setData(Qt.EditRole, float(odt))
                except Exception:
                    pass
                it_odt.setTextAlignment(Qt.AlignCenter)
                self.price_table.setItem(r, col, it_odt)
                col += 1

        try:
            self.price_table.sortItems(0, Qt.AscendingOrder)
        except Exception:
            pass

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

        top.addWidget(QLabel("Periods >>"))
        self.access_periods = QLineEdit()
        self.access_periods.setPlaceholderText("Örn: Periods (Regions|Total Day|All|Total|Total)")
        top.addWidget(self.access_periods, 3)

        top.addWidget(QLabel("Dates >>"))
        self.access_dates = QLineEdit()
        self.access_dates.setPlaceholderText("Örn: Aralık 2025")
        top.addWidget(self.access_dates, 2)

        top.addWidget(QLabel("Targets >>"))
        self.access_targets = QLineEdit()
        self.access_targets.setPlaceholderText("Örn: 35-54ABC1C2")
        top.addWidget(self.access_targets, 2)

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
        self.access_table.setAlternatingRowColors(True)
        self.access_table.setSortingEnabled(True)
        self.access_table.horizontalHeader().setSortIndicatorShown(True)
        self.access_table.horizontalHeader().setSectionsClickable(True)
        self.access_table.verticalHeader().setVisible(False)

        # Varsayılan saat kolonları (Excel örneğiyle aynı)
        self._access_hours = ['07:00-08:00',
                            '08:00-09:00',
                            '09:00-10:00',
                            '10:00-11:00',
                            '11:00-12:00',
                            '12:00-13:00',
                            '13:00-14:00',
                            '14:00-15:00',
                            '15:00-16:00',
                            '16:00-17:00',
                            '17:00-18:00',
                            '18:00-19:00',
                            '19:00-20:00']

        self._set_access_table_hours(self._access_hours)

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
        self.btn_access_save.clicked.connect(self.access_save_db)

        # Uygulama açılışında (repo hazırsa) en son kaydı otomatik yükle
        try:
            self.access_load_latest_db()
        except Exception:
            pass

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

        periods = (self.access_periods.text() or "").strip()
        dates = (self.access_dates.text() or "").strip()
        targets = (self.access_targets.text() or "").strip()

        year = self._parse_year_from_dates(dates)
        label = dates if dates else f"{year}"

        # İlk kayıt: set yoksa otomatik oluştur
        if not getattr(self, "_access_set_id", None):
            self._access_set_id = self.repo.get_or_create_access_set(
                year=year, label=label, periods=periods, targets=targets, hours=self._access_hours
            )

        def _to_float(item: QTableWidgetItem | None):
            t = (item.text() or "").strip() if item else ""
            if not t:
                return None
            try:
                return float(t.replace(",", "."))
            except Exception:
                return None

        rows: list[dict] = []
        for r in range(self.access_table.rowCount()):
            ch_item = self.access_table.item(r, 0)
            ch = (ch_item.text().strip() if ch_item else "")
            if not ch:
                continue

            values: dict = { }
            for i, hour in enumerate(self._access_hours, start=1):
                v = _to_float(self.access_table.item(r, i))
                # boş hücreleri yazmaya gerek yok (DB şişmesin)
                if v is None:
                    continue
                values[str(hour)] = v

            rows.append({"channel": ch, "values": values})

        self.repo.save_access_set(
            int(self._access_set_id),
            periods=periods,
            targets=targets,
            hours=self._access_hours,
            rows=rows,
        )
        QMessageBox.information(self, "OK", "Erişim örneği DB'ye kaydedildi. Uygulama yeniden açılınca aynen gelecektir.")

    def access_save_db(self) -> None:
        self.access_save()

    def access_paste_from_clipboard(self) -> None:
        """Excel'den kopyalanan saatlik erişim tablosunu yapıştırır.
        Beklenen format: ilk kolon kanal adı, sonraki kolonlar saatlik değerler.
        """
        try:
            text = (QApplication.clipboard().text() or "").strip()
            if not text:
                QMessageBox.warning(self, "Uyarı", "Panoda veri yok. Excel'de alanı kopyalayıp tekrar dene.")
                return

            lines = [ln for ln in text.splitlines() if ln.strip()]
            if not lines:
                QMessageBox.warning(self, "Uyarı", "Panoda geçerli satır yok.")
                return

            # başlangıç satırı: seçili satır varsa oradan, yoksa 0
            start_row = self.access_table.currentRow()
            if start_row < 0:
                start_row = 0

            r = start_row
            for ln in lines:
                cols = ln.split("\t")
                if not cols:
                    continue

                first = (cols[0] or "").strip()
                if not first:
                    continue

                # Header satırıysa atla
                if first.lower() in ("channels", "channel", "kanal", "kanallar"):
                    continue

                if r >= self.access_table.rowCount():
                    self.access_table.insertRow(r)

                # Kanal adı
                self.access_table.setItem(r, 0, QTableWidgetItem(first))

                # Saatlik değerler
                for i in range(1, 1 + len(self._access_hours)):
                    val = cols[i].strip() if i < len(cols) else ""
                    it = QTableWidgetItem(val)
                    it.setTextAlignment(Qt.AlignCenter)
                    self.access_table.setItem(r, i, it)

                r += 1

        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Yapıştırma başarısız: {e}")

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

    def _set_access_table_hours(self, hours: list[str]) -> None:
        """Erişim tablosu kolonlarını (Channels + saatlik kolonlar) yeniden kurar."""
        hours = hours or []
        self._access_hours = hours

        headers = ["Channels"] + list(hours)

        self.access_table.clear()
        self.access_table.setColumnCount(len(headers))
        self.access_table.setHorizontalHeaderLabels(headers)
        self.access_table.setRowCount(max(self.access_table.rowCount(), 30))
        self.access_table.setAlternatingRowColors(True)
        self.access_table.verticalHeader().setVisible(False)

        hdr = self.access_table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.Stretch)
        for i in range(1, len(headers)):
            hdr.setSectionResizeMode(i, QHeaderView.ResizeToContents)

    def access_load_latest_db(self) -> None:
        if not self.repo:
            return

        latest_id = self.repo.get_latest_access_set_id()
        if not latest_id:
            # ilk kullanım: boş şablon
            self.access_periods.setText("Periods (Regions|Total Day|All|Total|Total)")
            self.access_dates.setText(QDate.currentDate().toString("MMMM yyyy"))
            self.access_targets.setText("")
            self._access_set_id = None

            # default kolonlar + boş satırlar
            self._set_access_table_hours(getattr(self, "_access_hours", None) or self._access_hours)
            self.access_table.setRowCount(30)
            return

        meta, rows = self.repo.load_access_set(int(latest_id))
        self._access_set_id = int(latest_id)

        self.access_periods.setText(str(meta.get("periods") or ""))
        self.access_dates.setText(str(meta.get("label") or ""))
        self.access_targets.setText(str(meta.get("targets") or ""))

        hours = meta.get("hours") or getattr(self, "_access_hours", None) or []
        if hours:
            self._set_access_table_hours(hours)

        self.access_table.setRowCount(max(len(rows), 30))

        def _norm_hour(s: str) -> str:
            return re.sub(r"\([^\)]*\)\s*$", "", (s or "").strip())

        for i, r in enumerate(rows):
            self.access_table.setItem(i, 0, QTableWidgetItem(str(r.get("channel", ""))))

            vals = r.get("values") or {}
            # normalize map for fallback
            norm_map = {}
            for k, v in vals.items():
                norm_map[_norm_hour(str(k))] = v

            for col_idx, hour in enumerate(self._access_hours, start=1):
                v = None
                if str(hour) in vals:
                    v = vals.get(str(hour))
                else:
                    v = norm_map.get(_norm_hour(str(hour)))

                it = QTableWidgetItem("" if v is None else str(v))
                if v is not None:
                    try:
                        it.setData(Qt.EditRole, float(v))
                    except Exception:
                        pass
                it.setTextAlignment(Qt.AlignCenter)
                self.access_table.setItem(i, col_idx, it)

        try:
            self.access_table.sortItems(0, Qt.AscendingOrder)
        except Exception:
            pass