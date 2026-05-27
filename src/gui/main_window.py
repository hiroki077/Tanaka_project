from __future__ import annotations

from PySide6.QtWidgets import (
    QMainWindow, QTabWidget, QStatusBar, QMenuBar,
    QMessageBox, QFileDialog,
)
from PySide6.QtGui import QAction

from ..config import Settings, DataPaths
from ..db import Database
from ..services import EmployeeService, PhotoService
from .employee_tab import EmployeeTab
from .export_tab import ExportTab


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

        self.setWindowTitle("Miki 体制表ジェネレーター")
        self.resize(1100, 720)

        self._build_menu()
        self._build_tabs()
        self._build_status()

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
        QMessageBox.about(
            self, "バージョン情報",
            "Miki 体制表ジェネレーター v0.1\n"
            "顔写真付き Excel 体制表の自動生成ツール"
        )
