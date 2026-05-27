from __future__ import annotations
from pathlib import Path

from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QPixmap, QIcon
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLineEdit, QComboBox,
    QTableWidget, QTableWidgetItem, QPushButton, QMessageBox,
    QFileDialog, QHeaderView, QAbstractItemView, QLabel,
)

from ..db import EmploymentStatus
from ..services import EmployeeService, PhotoService
from .dialogs.employee_edit_dialog import EmployeeEditDialog


THUMB_SIZE = QSize(48, 60)


class EmployeeTab(QWidget):
    COLUMNS = [
        ("写真", 70),
        ("氏名", 150),
        ("カナ", 160),
        ("照合キー", 110),
        ("入社年", 80),
        ("役職", 100),
        ("雇用", 80),
        ("状態", 80),
    ]

    def __init__(self, employee_service: EmployeeService, photo_service: PhotoService):
        super().__init__()
        self.employee_service = employee_service
        self.photo_service = photo_service

        layout = QVBoxLayout(self)
        layout.addLayout(self._build_filter_bar())
        layout.addWidget(self._build_table(), 1)
        layout.addLayout(self._build_action_bar())

        self.refresh()

    def _build_filter_bar(self) -> QHBoxLayout:
        bar = QHBoxLayout()
        bar.addWidget(QLabel("検索:"))
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("氏名/カナ/照合キー")
        self.search_edit.textChanged.connect(self.refresh)
        bar.addWidget(self.search_edit, 2)

        bar.addWidget(QLabel("状態:"))
        self.status_combo = QComboBox()
        self.status_combo.addItem("（全て）", None)
        for s in EmploymentStatus:
            self.status_combo.addItem(s.value, s.value)
        self.status_combo.currentIndexChanged.connect(self.refresh)
        bar.addWidget(self.status_combo)

        bar.addWidget(QLabel("写真:"))
        self.photo_filter_combo = QComboBox()
        self.photo_filter_combo.addItem("（全て）", "all")
        self.photo_filter_combo.addItem("写真あり", "with")
        self.photo_filter_combo.addItem("写真なし", "without")
        self.photo_filter_combo.addItem("横長（要回転）", "landscape")
        self.photo_filter_combo.currentIndexChanged.connect(self.refresh)
        bar.addWidget(self.photo_filter_combo)

        self.count_label = QLabel("")
        bar.addWidget(self.count_label)
        return bar

    def _build_table(self) -> QTableWidget:
        self.table = QTableWidget(0, len(self.COLUMNS))
        self.table.setHorizontalHeaderLabels([c[0] for c in self.COLUMNS])
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setIconSize(THUMB_SIZE)
        for i, (_, w) in enumerate(self.COLUMNS):
            self.table.setColumnWidth(i, w)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.doubleClicked.connect(lambda _: self._on_edit())
        return self.table

    def _build_action_bar(self) -> QHBoxLayout:
        bar = QHBoxLayout()
        for label, slot in [
            ("新規追加", self._on_add),
            ("編集", self._on_edit),
            ("写真変更", self._on_change_photo),
            ("↺ 左90°", lambda: self._on_rotate(-90)),
            ("↻ 右90°", lambda: self._on_rotate(90)),
            ("⤺ 180°", lambda: self._on_rotate(180)),
            ("削除", self._on_delete),
        ]:
            btn = QPushButton(label)
            btn.clicked.connect(slot)
            bar.addWidget(btn)
        bar.addStretch(1)
        return bar

    def _on_rotate(self, degrees: int) -> None:
        emp_id = self._selected_employee_id()
        if emp_id is None:
            QMessageBox.information(self, "選択なし", "写真を回転する従業員を選択してください。")
            return
        emp = self.employee_service.get(emp_id)
        if emp is None or not emp.photo_path:
            QMessageBox.information(self, "写真なし", "この従業員には写真が登録されていません。")
            return
        try:
            self.employee_service.rotate_photo(emp_id, degrees)
        except Exception as e:
            QMessageBox.critical(self, "回転失敗", str(e))
            return
        self.refresh()

    def refresh(self) -> None:
        keyword = self.search_edit.text().strip() or None
        status = self.status_combo.currentData()
        photo_filter = self.photo_filter_combo.currentData()
        rows = self.employee_service.list(keyword=keyword, status_filter=status)

        if photo_filter == "with":
            rows = [e for e in rows if self.photo_service.resolve(e.photo_path)]
        elif photo_filter == "without":
            rows = [e for e in rows if not self.photo_service.resolve(e.photo_path)]
        elif photo_filter == "landscape":
            rows = [e for e in rows if self._is_landscape(e)]

        self.count_label.setText(f"  {len(rows)}名")
        self.table.setRowCount(len(rows))
        for r, emp in enumerate(rows):
            self._fill_row(r, emp)
            self.table.setRowHeight(r, THUMB_SIZE.height() + 8)

    def _is_landscape(self, emp, ratio_threshold: float = 1.1) -> bool:
        """写真の縦横比が横長（width/height > ratio_threshold）かどうか。"""
        p = self.photo_service.resolve(emp.photo_path)
        if not p:
            return False
        try:
            from PIL import Image
            with Image.open(p) as img:
                w, h = img.size
        except Exception:
            return False
        return w / max(h, 1) > ratio_threshold

    def _fill_row(self, r: int, emp) -> None:
        # 写真サムネ
        item_photo = QTableWidgetItem()
        item_photo.setData(Qt.UserRole, emp.id)
        p = self.photo_service.resolve(emp.photo_path)
        if p:
            pm = QPixmap(str(p)).scaled(
                THUMB_SIZE, Qt.KeepAspectRatio, Qt.SmoothTransformation
            )
            item_photo.setIcon(QIcon(pm))
        self.table.setItem(r, 0, item_photo)

        year_display = emp.join_year_text or (str(emp.join_year) if emp.join_year else "")
        values = [
            emp.name or "",
            emp.name_kana or "",
            emp.match_key or "",
            year_display,
            emp.role or "",
            emp.employment_type or "",
            emp.status or "",
        ]
        for c, v in enumerate(values, start=1):
            it = QTableWidgetItem(v)
            it.setData(Qt.UserRole, emp.id)
            self.table.setItem(r, c, it)

    def _selected_employee_id(self) -> int | None:
        row = self.table.currentRow()
        if row < 0:
            return None
        item = self.table.item(row, 0)
        return item.data(Qt.UserRole) if item else None

    def _on_add(self) -> None:
        dlg = EmployeeEditDialog(self, employee=None, photo_service=self.photo_service)
        if dlg.exec() == EmployeeEditDialog.Accepted:
            data = dlg.collect()
            photo_src = data.pop("_photo_src", None)
            emp = self.employee_service.create(**data)
            if photo_src:
                self.employee_service.set_photo(emp.id, photo_src)
            self.refresh()

    def _on_edit(self) -> None:
        emp_id = self._selected_employee_id()
        if emp_id is None:
            QMessageBox.information(self, "選択なし", "編集する従業員を選択してください。")
            return
        emp = self.employee_service.get(emp_id)
        if emp is None:
            return
        dlg = EmployeeEditDialog(self, employee=emp, photo_service=self.photo_service)
        if dlg.exec() == EmployeeEditDialog.Accepted:
            data = dlg.collect()
            photo_src = data.pop("_photo_src", None)
            self.employee_service.update(emp_id, **data)
            if photo_src:
                self.employee_service.set_photo(emp_id, photo_src)
            self.refresh()

    def _on_change_photo(self) -> None:
        emp_id = self._selected_employee_id()
        if emp_id is None:
            QMessageBox.information(self, "選択なし", "写真を変更する従業員を選択してください。")
            return
        path, _ = QFileDialog.getOpenFileName(
            self, "顔写真を選択", str(Path.home()),
            "画像ファイル (*.jpg *.jpeg *.png)",
        )
        if not path:
            return
        self.employee_service.set_photo(emp_id, path)
        self.refresh()

    def _on_delete(self) -> None:
        emp_id = self._selected_employee_id()
        if emp_id is None:
            return
        emp = self.employee_service.get(emp_id)
        if emp is None:
            return
        if QMessageBox.question(
            self, "削除確認",
            f"「{emp.name}」を削除します。よろしいですか？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
        ) != QMessageBox.Yes:
            return
        self.employee_service.delete(emp_id)
        self.refresh()
