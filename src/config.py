"""アプリケーション設定。

データフォルダ (DB + 写真) の場所はユーザーが選択可能。
OneDrive 等のクラウド同期フォルダを指定することで複数PC間でDBを共有できる。

設定優先順位:
1. settings.json の data_dir
2. プラットフォーム既定 (Windows: %APPDATA%/MikiApp, macOS: ~/Library/Application Support/MikiApp)
"""
from __future__ import annotations
import json
import os
import sys
from pathlib import Path


APP_NAME = "Roster"

DEFAULT_PHOTO_WIDTH_COLS = 4
DEFAULT_PHOTO_HEIGHT_ROWS = 6
DEFAULT_PHOTO_OFFSET_ROWS_UP = 8

PLACEHOLDER_PATTERN = r"\{\{photo:([^}]+)\}\}"


def _resolve_assets_dir() -> Path:
    """同梱アセット（プレースホルダー画像等）の所在を解決。

    開発時はリポジトリ直下 assets/、PyInstaller配布時は _MEIPASS/assets/。
    """
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.argv[0]).parent)) / "assets"
    return Path(__file__).resolve().parent.parent / "assets"


ASSETS_DIR: Path = _resolve_assets_dir()
PLACEHOLDER_PHOTO_PATH: Path = ASSETS_DIR / "placeholder.png"


def _local_config_dir() -> Path:
    """このPC固有の設定（data_dir パス情報）を保存する場所。

    クラウド同期は **しない**。data_dir のパスはPCごとに異なるため。
    """
    if sys.platform == "win32":
        base = Path(os.environ.get("APPDATA", Path.home())) / APP_NAME
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support" / APP_NAME
    else:
        base = Path.home() / ".config" / APP_NAME
    base.mkdir(parents=True, exist_ok=True)
    return base


LOCAL_CONFIG_DIR: Path = _local_config_dir()
SETTINGS_FILE: Path = LOCAL_CONFIG_DIR / "settings.json"


def _default_data_dir() -> Path:
    """data_dir が未設定の場合のフォールバック先。"""
    if getattr(sys, "frozen", False):
        return LOCAL_CONFIG_DIR / "data"
    return Path(__file__).resolve().parent.parent / "data"


class Settings:
    """settings.json の読み書きを担当。"""

    def __init__(self, path: Path = SETTINGS_FILE):
        self.path = path
        self._data: dict = {}
        self._load()

    def _load(self) -> None:
        if self.path.is_file():
            try:
                self._data = json.loads(self.path.read_text(encoding="utf-8"))
            except (json.JSONDecodeError, OSError):
                self._data = {}
        else:
            self._data = {}

    def save(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(
            json.dumps(self._data, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    @property
    def data_dir(self) -> Path:
        raw = self._data.get("data_dir")
        return Path(raw) if raw else _default_data_dir()

    @data_dir.setter
    def data_dir(self, value: str | Path) -> None:
        self._data["data_dir"] = str(Path(value).expanduser().resolve())

    @property
    def is_configured(self) -> bool:
        return "data_dir" in self._data

    @property
    def export_defaults(self) -> dict:
        return self._data.get("export_defaults", {})

    def set_export_defaults(self, **kwargs) -> None:
        self._data.setdefault("export_defaults", {}).update(kwargs)


class DataPaths:
    """data_dir 配下のパス解決をまとめる。"""

    def __init__(self, data_dir: Path):
        self.data_dir = Path(data_dir)
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.photos_dir = self.data_dir / "photos"
        self.photos_dir.mkdir(parents=True, exist_ok=True)
        self.db_path = self.data_dir / "employees.db"
        self.lock_path = self.data_dir / ".lock"
