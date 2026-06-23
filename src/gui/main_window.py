from __future__ import annotations
import sys
from pathlib import Path

from PySide6.QtCore import Qt, QThread, QTimer, Signal
from PySide6.QtWidgets import (
    QMainWindow, QTabWidget, QStatusBar, QMenuBar,
    QMessageBox, QFileDialog, QProgressDialog, QApplication,
)
from PySide6.QtGui import QAction

from ..config import Settings, DataPaths
from ..db import Database
from ..services import EmployeeService, PhotoService
from ..services.update_service import (
    UpdateInfo,
    check_for_update,
    download_exe,
    install_and_restart,
    is_frozen,
    current_build_sha,
)
from .employee_tab import EmployeeTab
from .export_tab import ExportTab


class _UpdateCheckThread(QThread):
    """`check_for_update()` をバックグラウンドで実行。

    結果は `done` シグナルで返す。値は UpdateInfo | None | Exception。
    """
    done = Signal(object)

    def run(self) -> None:
        try:
            self.done.emit(check_for_update())
        except Exception as e:  # noqa: BLE001 - ネットワーク何でも来るので広く拾う
            self.done.emit(e)


class _DownloadThread(QThread):
    progress = Signal(int, int)  # written, total
    done = Signal(object)        # None on success, Exception on failure

    def __init__(self, url: str, dest: Path, parent=None) -> None:
        super().__init__(parent)
        self._url = url
        self._dest = dest
        self._cancel = False

    def cancel(self) -> None:
        self._cancel = True

    def run(self) -> None:
        try:
            download_exe(
                self._url, self._dest,
                progress_cb=lambda w, t: self.progress.emit(w, t),
                cancel_cb=lambda: self._cancel,
            )
            self.done.emit(None)
        except Exception as e:  # noqa: BLE001
            self.done.emit(e)


class MainWindow(QMainWindow):
    def __init__(
        self,
        settings: Settings,
        paths: DataPaths,
        db: Database,
        employee_service: EmployeeService,
        photo_service: PhotoService,
    ):
        super().__init__()
        self.settings = settings
        self.paths = paths
        self.db = db
        self.employee_service = employee_service
        self.photo_service = photo_service

        self._update_thread: _UpdateCheckThread | None = None
        self._download_thread: _DownloadThread | None = None

        self.setWindowTitle("Miki 体制表ジェネレーター")
        self.resize(1100, 720)

        self._build_menu()
        self._build_tabs()
        self._build_status()

        # 起動時に自動で1回更新確認（ウィンドウ表示後に走らせる）
        QTimer.singleShot(1500, lambda: self._check_update(manual=False))

    def _build_menu(self) -> None:
        menubar: QMenuBar = self.menuBar()
        file_menu = menubar.addMenu("ファイル(&F)")
        action_change_dir = QAction("データフォルダを変更(&D)…", self)
        action_change_dir.triggered.connect(self._on_change_data_dir)
        file_menu.addAction(action_change_dir)
        file_menu.addSeparator()
        action_quit = QAction("終了(&Q)", self)
        action_quit.triggered.connect(self.close)
        file_menu.addAction(action_quit)

        # メニューバー直下に「更新を確認…」
        action_check_update = QAction("更新を確認(&U)…", self)
        action_check_update.triggered.connect(lambda: self._check_update(manual=True))
        menubar.addAction(action_check_update)

        help_menu = menubar.addMenu("ヘルプ(&H)")
        action_about = QAction("バージョン情報(&A)", self)
        action_about.triggered.connect(self._on_about)
        help_menu.addAction(action_about)

    def _build_tabs(self) -> None:
        tabs = QTabWidget(self)
        self.employee_tab = EmployeeTab(
            employee_service=self.employee_service,
            photo_service=self.photo_service,
        )
        self.export_tab = ExportTab(
            db=self.db,
            employee_service=self.employee_service,
            photo_service=self.photo_service,
            settings=self.settings,
        )
        tabs.addTab(self.employee_tab, "👥 従業員")
        tabs.addTab(self.export_tab, "📤 Excel出力")
        self.setCentralWidget(tabs)

    def _build_status(self) -> None:
        bar = QStatusBar(self)
        bar.showMessage(f"データフォルダ: {self.paths.data_dir}")
        self.setStatusBar(bar)

    def _on_change_data_dir(self) -> None:
        folder = QFileDialog.getExistingDirectory(
            self, "データフォルダを選択", str(self.paths.data_dir)
        )
        if not folder:
            return
        QMessageBox.information(
            self, "再起動が必要",
            "データフォルダを変更しました。\nアプリを再起動してください。"
        )
        self.settings.data_dir = folder
        self.settings.save()
        self.close()

    def _on_about(self) -> None:
        sha = current_build_sha()
        sha_part = f"\nビルド: {sha[:7]}" if sha else "\nビルド: dev"
        QMessageBox.about(
            self, "バージョン情報",
            "Miki 体制表ジェネレーター v0.1\n"
            "顔写真付き Excel 体制表の自動生成ツール"
            + sha_part
        )

    # -------------------- 更新チェック / インストール --------------------

    def _check_update(self, manual: bool) -> None:
        # 開発実行中は何もしない。手動押下時のみメッセージを出す。
        if not is_frozen() or not current_build_sha():
            if manual:
                QMessageBox.information(
                    self, "更新確認",
                    "開発ビルドでは更新確認は無効です。"
                )
            return
        # 二重起動防止
        if self._update_thread is not None and self._update_thread.isRunning():
            return
        if manual:
            self.statusBar().showMessage("更新を確認中…", 3000)
        self._update_thread = _UpdateCheckThread(self)
        self._update_thread.done.connect(
            lambda result: self._on_update_check_done(result, manual)
        )
        self._update_thread.start()

    def _on_update_check_done(self, result, manual: bool) -> None:
        if isinstance(result, Exception):
            if manual:
                QMessageBox.warning(
                    self, "更新確認に失敗",
                    f"ネットワークまたはサーバ側のエラーで更新を確認できませんでした。\n\n{result}"
                )
            return
        info: UpdateInfo | None = result
        if info is None:
            if manual:
                QMessageBox.information(
                    self, "更新確認",
                    f"お使いのバージョンは最新です。\nビルド: {current_build_sha()[:7]}"
                )
            return

        size_mb = info.asset_size / 1024 / 1024 if info.asset_size else 0
        ret = QMessageBox.question(
            self, "更新があります",
            "新しいバージョンが公開されています。\n\n"
            f"現在のビルド : {info.current_sha[:7]}\n"
            f"最新のビルド : {info.latest_sha[:7]}\n"
            f"公開日時     : {info.published_at}\n"
            f"ダウンロード : 約 {size_mb:.1f} MB\n\n"
            "今すぐ更新しますか？\n（ダウンロード後、自動的に再起動します）",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes,
        )
        if ret != QMessageBox.Yes:
            return
        self._start_download(info)

    def _start_download(self, info: UpdateInfo) -> None:
        current_exe = Path(sys.executable).resolve()
        new_exe = current_exe.with_suffix(current_exe.suffix + ".new")

        progress = QProgressDialog(
            "更新ファイルをダウンロード中…", "キャンセル", 0, 100, self
        )
        progress.setWindowTitle("更新中")
        progress.setWindowModality(Qt.ApplicationModal)
        progress.setMinimumDuration(0)
        progress.setAutoClose(False)
        progress.setAutoReset(False)
        progress.setValue(0)
        progress.show()

        thread = _DownloadThread(info.asset_url, new_exe, parent=self)
        self._download_thread = thread

        def on_progress(written: int, total: int) -> None:
            if total > 0:
                progress.setMaximum(100)
                progress.setValue(int(written * 100 / total))
            else:
                # サイズ不明時は不定プログレス
                progress.setMaximum(0)

        def on_canceled() -> None:
            thread.cancel()

        def on_done(result) -> None:
            progress.close()
            if isinstance(result, Exception):
                if isinstance(result, InterruptedError):
                    return  # キャンセル時は何も出さない
                QMessageBox.critical(
                    self, "ダウンロード失敗",
                    f"更新ファイルのダウンロードに失敗しました。\n\n{result}"
                )
                return
            # 成功 → ヘルパ bat 起動して終了
            try:
                install_and_restart(new_exe)
            except Exception as e:  # noqa: BLE001
                QMessageBox.critical(
                    self, "更新失敗",
                    f"更新スクリプトの起動に失敗しました。\n\n{e}"
                )
                return
            QApplication.quit()

        thread.progress.connect(on_progress)
        thread.done.connect(on_done)
        progress.canceled.connect(on_canceled)
        thread.start()
