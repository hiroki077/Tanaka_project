"""Miki アプリ エントリーポイント。

起動シーケンス:
1. settings.json を読み、data_dir が未設定なら DataFolderDialog で選ばせる
2. data_dir の .lock を確認し、他者が使用中なら警告
3. SQLite を初期化
4. MainWindow 起動
"""
from __future__ import annotations
import sys
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QMessageBox, QFileDialog,
)

from .config import Settings, DataPaths
from .db import Database
from .lock import DataLock
from .services import EmployeeService, PhotoService
from .gui.main_window import MainWindow


def _choose_data_folder(settings: Settings) -> Path | None:
    msg = QMessageBox()
    msg.setWindowTitle("データフォルダの選択")
    msg.setIcon(QMessageBox.Information)
    msg.setText(
        "従業員DBと顔写真を保存するフォルダを選択してください。\n\n"
        "OneDrive 等のクラウド同期フォルダを選ぶと\n"
        "複数PCでデータを共有できます。"
    )
    msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
    if msg.exec() != QMessageBox.Ok:
        return None
    folder = QFileDialog.getExistingDirectory(
        None, "データフォルダを選択", str(Path.home())
    )
    if not folder:
        return None
    settings.data_dir = folder
    settings.save()
    return Path(folder)


def _confirm_lock_takeover(holder) -> bool:
    msg = QMessageBox()
    msg.setWindowTitle("データフォルダ使用中")
    msg.setIcon(QMessageBox.Warning)
    msg.setText(
        f"このデータフォルダは別のユーザー/PCが使用中です:\n"
        f"  ホスト: {holder.host}\n"
        f"  PID  : {holder.pid}\n\n"
        f"無視して使用すると DB 破損の恐れがあります。\n"
        f"それでも使用を継続しますか？"
    )
    msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
    msg.setDefaultButton(QMessageBox.No)
    return msg.exec() == QMessageBox.Yes


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("Miki 体制表ジェネレーター")

    settings = Settings()
    if not settings.is_configured:
        chosen = _choose_data_folder(settings)
        if chosen is None:
            return 0

    paths = DataPaths(settings.data_dir)

    lock = DataLock(paths.lock_path)
    ok, holder = lock.acquire()
    if not ok:
        if not _confirm_lock_takeover(holder):
            return 0
        lock.acquire(force=True)

    try:
        db = Database(paths.db_path)
        db.create_all()
        photos = PhotoService(paths.photos_dir)
        employee_service = EmployeeService(db, photos)

        win = MainWindow(
            settings=settings,
            paths=paths,
            db=db,
            employee_service=employee_service,
            photo_service=photos,
        )
        win.show()
        return app.exec()
    finally:
        lock.release()


if __name__ == "__main__":
    sys.exit(main())
