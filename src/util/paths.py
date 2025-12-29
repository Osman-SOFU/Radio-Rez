import sys
from pathlib import Path

def project_root() -> Path:
    return Path(__file__).resolve().parents[2]

def resource_path(relative: str) -> Path:
    """
    PyInstaller onefile'da dosyalar _MEIPASS altına çıkar.
    Dev ortamında proje kökünden okur.
    """
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / relative  # type: ignore[attr-defined]
    return project_root() / relative
