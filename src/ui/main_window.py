from pathlib import Path

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTabWidget,
    QFileDialog, QMessageBox
)
from PySide6.QtCore import Qt

from settings.app_settings import AppSettings
from storage.db import connect, migrate
from storage.reservations_repo import ReservationDraft, save_confirmed
from export.excel_exporter import export_reservation_excel


BASE_DIR = Path(__file__).resolve().parents[2]  # .../Radio-Rez


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Radio-Rez")
        self.resize(1100, 700)

        self.settings = AppSettings.load()

        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)

        # Top bar
        top = QHBoxLayout()
        layout.addLayout(top)

        top.addWidget(QLabel("Reklamveren Ara:"))
        self.search = QLineEdit()
        self.search.setPlaceholderText("Örn: MURATBEY")
        top.addWidget(self.search)

        self.folder_label = QLabel(self._folder_text())
        self.folder_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        top.addWidget(self.folder_label)

        btn_folder = QPushButton("Veri Klasörü Seç")
        btn_folder.clicked.connect(self.pick_folder)
        top.addWidget(btn_folder)

        # Tabs
        tabs = QTabWidget()
        layout.addWidget(tabs)

        for name in [
            "REZERVASYON ve PLANLAMA",
            "PLAN ÖZET",
            "SPOTLİST+",
            "KOPYA TANIMI",
            "Fiyat ve Kanal Tanımı",
            "DT-ODT",
            "Erişim Örneği",
        ]:
            tabs.addTab(self.simple_tab(name), name)

        # Bottom buttons
        bottom = QHBoxLayout()
        layout.addLayout(bottom)

        self.btn_approve = QPushButton("Onayla")
        self.btn_approve.clicked.connect(self.on_approve)
        bottom.addWidget(self.btn_approve)

        bottom.addStretch()

        self.btn_test = QPushButton("Test Olarak İndir")
        self.btn_test.setEnabled(False)
        self.btn_test.clicked.connect(self.on_test)
        bottom.addWidget(self.btn_test)

        self.btn_save = QPushButton("Kaydet ve Çıktı Al")
        self.btn_save.setEnabled(False)
        self.btn_save.clicked.connect(self.on_save)
        bottom.addWidget(self.btn_save)

    def simple_tab(self, text):
        w = QWidget()
        l = QVBoxLayout(w)
        l.addWidget(QLabel(text))
        l.addStretch()
        return w

    def _folder_text(self):
        return f"Veri Klasörü: {self.settings.data_dir or '(seçilmedi)'}"

    def pick_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Veri klasörü seç")
        if folder:
            self.settings.data_dir = folder
            self.settings.save()
            self.folder_label.setText(self._folder_text())

    def on_approve(self):
        if not self.settings.data_dir:
            QMessageBox.warning(self, "Eksik", "Önce veri klasörü seçmelisin.")
            return
        self.btn_test.setEnabled(True)
        self.btn_save.setEnabled(True)
        QMessageBox.information(self, "Onay", "Onaylandı.")

    def on_test(self):
        if not self.settings.data_dir:
            QMessageBox.warning(self, "Eksik", "Önce veri klasörü seçmelisin.")
            return

        advertiser = self.search.text().strip() or "BILINMEYEN"
        year, month = 2025, 12
        channel = "Kanal-1"

        try:
            template = BASE_DIR / "assets" / "template.xlsx"
            out_path = export_reservation_excel(
                template_path=template,
                output_dir=Path(self.settings.data_dir),
                advertiser_name=advertiser,
                channel_name=channel,
                year=year,
                month=month,
                reservation_no=None,
            )
            QMessageBox.information(self, "Test Çıktı", f"Test excel oluşturuldu:\n{out_path}")
        except Exception as e:
            QMessageBox.critical(self, "Hata", str(e))

    def on_save(self):
        if not self.settings.data_dir:
            QMessageBox.warning(self, "Eksik", "Önce veri klasörü seçmelisin.")
            return

        draft = ReservationDraft(
            advertiser_name=self.search.text().strip() or "BILINMEYEN",
            year=2025,
            month=12,
            channel_name="Kanal-1",
            payload={"note": "MVP payload - sonra formdan dolacak"},
        )

        try:
            conn = connect(self.settings.data_dir)
            migrate(conn)
            res_no = save_confirmed(conn, draft)
            conn.close()

            template = BASE_DIR / "assets" / "template.xlsx"
            logo_path = str(Path("assets") / "RADIOSCOPE.PNG")
            out_path = export_reservation_excel(
                template_path=template_path,
                output_dir=self.settings.data_dir,
                advertiser_name=draft.advertiser_name,
                channel_name=draft.channel_name,
                year=draft.year,
                month=draft.month,
                reservation_no=res_no,

                agency_name="",          # şimdilik boş
                product_name="",         # şimdilik boş
                plan_title="",           # şimdilik boş
                a67_text=None,           # dokunma: template neyse kalsın
                b67_text=None,
                a76_text="FLYDEAL",      # ya da UI’dan alırsın
                ak77_name="FERHAT ÇİNDEMİR",
                note_text="",
                logo_path=logo_path,
                logo_anchor="AO2",
                # logo_width=140, logo_height=50,
            )

            QMessageBox.information(self, "Kaydedildi", f"Rez No: {res_no}\nExcel:\n{out_path}")
        except Exception as e:
            QMessageBox.critical(self, "Hata", str(e))
