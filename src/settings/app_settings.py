from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from PySide6.QtCore import QSettings

from src.util.paths import resource_path

ORG = "RadioRez"
APP = "RadioRez"

@dataclass(frozen=True)
class AppSettings:
    data_dir: Path
    template_path: Path

class SettingsService:
    def __init__(self) -> None:
        self._qs = QSettings(ORG, APP)

    def get_data_dir(self) -> Path | None:
        v = self._qs.value("data_dir", "")
        v = str(v).strip()
        return Path(v) if v else None

    def set_data_dir(self, p: Path) -> None:
        self._qs.setValue("data_dir", str(p))

    def get_template_path(self) -> Path:
        # Şablonu app ile birlikte assets/reservation_template.xlsx olarak taşıyoruz.
        # İleride istersen kullanıcı seçtirebilirsin.
        v = str(self._qs.value("template_path", "")).strip()
        if v:
            return Path(v)
        return resource_path("assets/reservation_template.xlsx")

    def set_template_path(self, p: Path) -> None:
        self._qs.setValue("template_path", str(p))

    def build(self) -> AppSettings:
        data_dir = self.get_data_dir()
        if not data_dir:
            # Henüz seçilmemiş olabilir; UI bunun için zaten zorlayacak.
            data_dir = Path.home() / "RadioRezData"
        return AppSettings(
            data_dir=data_dir,
            template_path=self.get_template_path(),
        )
