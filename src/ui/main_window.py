from __future__ import annotations

from pathlib import Path
from datetime import datetime, date, time

from PySide6.QtCore import Qt, QTime, QDate
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QTabWidget, QFileDialog, QMessageBox, QListWidget,
    QDateEdit, QTimeEdit, QGroupBox
)

from src.settings.app_settings import SettingsService, AppSettings
from src.storage.db import ensure_data_folders, connect_db, migrate_and_seed
from src.storage.repository import Repository
from src.export.excel_exporter import export_excel

TAB_NAMES = [
    "REZERVASYON ve PLANLAMA",
    "PLAN ÖZET",
    "SPOTLİST+",
    "KOPYA TANIMI",
    "Fiyat ve Kanal Tanımı",
    "DT-ODT",
    "Erişim Örneği",
]

def classify_dt_odt(t: time) -> str:
    # DT: 07:00–10:00 ve 17:00–20:00 (sınırlar dahil varsayıyorum)
    mins = t.hour * 60 + t.minute
    dt1 = 7 * 60 <= mins <= 10 * 60
    dt2 = 17 * 60 <= mins <= 20 * 60
    return "DT" if (dt1 or dt2) else "ODT"

def validate_day(plan_date: date) -> tuple[bool, str]:
    # QDate zaten geçersiz gün seçtirmez; yine de kuralı burada tutuyoruz.
    try:
        _ = plan_date.toordinal()
        return True, ""
    except Exception:
        return False, "Geçersiz tarih."

class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Radio-Rez (MVP)")
        self.resize(1200, 700)

        self.settings_service = SettingsService()
        self.app_settings: AppSettings = self.settings_service.build()

        self.repo: Repository | None = None
        self.current_payload: dict | None = None

        root = QWidget()
        self.setCentralWidget(root)
        main = QVBoxLayout(root)

        # Top bar
        top = QHBoxLayout()
        main.addLayout(top)

        top.addWidget(QLabel("Reklamveren Ara:"))
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

        left_box = QGroupBox("Arama Sonuçları (Reklamveren)")
        left_layout = QVBoxLayout(left_box)
        self.list_advertisers = QListWidget()
        left_layout.addWidget(self.list_advertisers)
        mid.addWidget(left_box, 1)

        # Tabs
        self.tabs = QTabWidget()
        mid.addWidget(self.tabs, 4)

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

    def _build_first_tab(self) -> None:
        tab = self.tab_widgets["REZERVASYON ve PLANLAMA"]
        layout = QVBoxLayout(tab)

        row1 = QHBoxLayout()
        layout.addLayout(row1)
        row1.addWidget(QLabel("Reklamveren:"))
        self.in_advertiser = QLineEdit()
        row1.addWidget(self.in_advertiser, 2)

        row2 = QHBoxLayout()
        layout.addLayout(row2)
        row2.addWidget(QLabel("Plan Tarihi:"))
        self.in_date = QDateEdit()
        self.in_date.setCalendarPopup(True)
        self.in_date.setDate(QDate.currentDate())
        row2.addWidget(self.in_date)

        row2.addWidget(QLabel("Spot Saat:"))
        self.in_time = QTimeEdit()
        self.in_time.setTime(QTime.currentTime())
        row2.addWidget(self.in_time)

        self.lbl_dt_odt = QLabel("DT/ODT: -")
        self.lbl_dt_odt.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        row2.addWidget(self.lbl_dt_odt, 1)

        self.in_time.timeChanged.connect(self.on_time_changed)
        self.on_time_changed(self.in_time.time())

        layout.addStretch(1)

    def on_time_changed(self, qt: QTime) -> None:
        t = time(qt.hour(), qt.minute(), qt.second())
        self.lbl_dt_odt.setText(f"DT/ODT: {classify_dt_odt(t)}")

    def bootstrap_storage(self) -> None:
        ensure_data_folders(self.app_settings.data_dir)
        db_path = self.app_settings.data_dir / "data.db"
        conn = connect_db(db_path)
        migrate_and_seed(conn)
        self.repo = Repository(conn)

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

    def on_advertiser_selected(self) -> None:
        # MVP: seçince sadece reklamveren input'una basıyoruz
        item = self.list_advertisers.currentItem()
        if not item:
            return
        self.in_advertiser.setText(item.text())

    def on_confirm(self) -> None:
        adv = self.in_advertiser.text().strip()
        if not adv:
            QMessageBox.warning(self, "Hata", "Reklamveren zorunlu.")
            return

        d = self.in_date.date().toPython()
        ok, msg = validate_day(d)
        if not ok:
            QMessageBox.warning(self, "Hata", msg)
            return

        qt = self.in_time.time()
        t = time(qt.hour(), qt.minute(), qt.second())
        dt_odt = classify_dt_odt(t)

        self.current_payload = {
            "advertiser_name": adv,
            "plan_date": d.isoformat(),
            "spot_time": f"{t.hour:02d}:{t.minute:02d}",
            "dt_odt": dt_odt,
        }

        self.btn_test_export.setEnabled(True)
        self.btn_save_export.setEnabled(True)
        QMessageBox.information(self, "OK", "Onaylandı. Artık test veya kayıt alabilirsin.")

    def on_test_export(self) -> None:
        if not self.current_payload:
            return
        out_dir = self.app_settings.data_dir / "exports"
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = out_dir / f"TEST_{ts}.xlsx"

        payload = dict(self.current_payload)
        payload["reservation_no"] = ""  # testte yok
        payload["created_at"] = datetime.now().isoformat(timespec="seconds")

        try:
            export_excel(self.app_settings.template_path, out_path, payload)
            QMessageBox.information(self, "OK", f"Test çıktısı üretildi:\n{out_path}")
        except Exception as e:
            QMessageBox.critical(self, "Hata", str(e))

    def on_save_export(self) -> None:
        if not self.current_payload or not self.repo:
            return

        try:
            rec = self.repo.create_reservation(
                advertiser_name=self.current_payload["advertiser_name"],
                payload=self.current_payload,
                confirmed=True,
            )

            out_dir = self.app_settings.data_dir / "exports"
            out_path = out_dir / f"{rec.reservation_no}.xlsx"

            payload = dict(rec.payload)
            payload["reservation_no"] = rec.reservation_no
            payload["created_at"] = rec.created_at

            export_excel(self.app_settings.template_path, out_path, payload)
            QMessageBox.information(self, "OK", f"Kaydedildi ve çıktı alındı:\n{out_path}")
        except Exception as e:
            QMessageBox.critical(self, "Hata", str(e))
