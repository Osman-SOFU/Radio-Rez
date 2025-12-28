import json
from pathlib import Path


class AppSettings:
    def __init__(self, data_dir=None):
        self.data_dir = data_dir

    @staticmethod
    def path():
        base = Path.home() / ".radio-rez"
        base.mkdir(exist_ok=True)
        return base / "settings.json"

    @classmethod
    def load(cls):
        p = cls.path()
        if not p.exists():
            return cls()
        return cls(**json.loads(p.read_text(encoding="utf-8")))

    def save(self):
        self.path().write_text(
            json.dumps({"data_dir": self.data_dir}, indent=2),
            encoding="utf-8"
        )
