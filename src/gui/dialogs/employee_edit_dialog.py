from __future__ import annotations
import io
import shutil
import tempfile
import uuid
from pathlib import Path

from PIL import Image, ImageOps
from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import (
    QDialog, QFormLayout, QVBoxLayout, QHBoxLayout, QLineEdit, QSpinBox,
    QComboBox, QPushButton, QLabel, QDialogButtonBox, QFileDialog,
    QMessageBox,
)

from ...db import EmploymentStatus
from ...services import PhotoService


PREVIEW_SIZE = (150, 180)


class EmployeeEditDialog(QDialog):
    def __init__(self, parent, employee, photo_service: PhotoService):
        super().__init__(parent)
        self.setWindowTitle("従業員を編集" if employee else "従業員を追加")
        self.employee = employee
        self.photo_service = photo_service
        self._new_photo_src: str | None = None
        # 写真の作業ファイル（保存済みなら photos_dir 内、新規選択なら temp）
        self._working_path: Path | None = None
        self._working_is_temp: bool = False
        self._working_dirty: bool = False  # 新規選択 or 回転で変更されたか

        self.name = QLineEdit()
        self.name_kana = QLineEdit()
        self.match_key = QLineEdit()
        self.match_key.setPlaceholderText("テンプレに書く文字列（例: 藤堂）")
        self.join_year = QSpinBox()
        self.join_year.setRange(1950, 2100)
        self.join_year.setValue(2024)
        self.role = QLineEdit()
        self.role_marks = QLineEdit()
        self.role_marks.setPlaceholderText("☆■◆○▽ などカンマ区切り")
        self.employment_type = QComboBox()
        self.employment_type.addItems(["", "正社員", "契約", "派遣", "再雇用"])
        self.status = QComboBox()
        for s in EmploymentStatus:
            self.status.addItem(s.value, s.value)
        self.note = QLineEdit()

        self.preview = QLabel()
        self.preview.setFixedSize(*PREVIEW_SIZE)
        self.preview.setStyleSheet("border: 1px solid #aaa; background:#fafafa;")
        self.preview.setAlignment(Qt.AlignCenter)
        self.preview.setText("(未設定)")
        self.choose_photo_btn = QPushButton("写真を選択…")
        self.choose_photo_btn.clicked.connect(self._on_choose_photo)

        self.rotate_left_btn = QPushButton("↺ 左90°")
        self.rotate_right_btn = QPushButton("↻ 右90°")
        self.rotate_180_btn = QPushButton("⤺ 180°")
        self.rotate_left_btn.clicked.connect(lambda: self._on_rotate(-90))
        self.rotate_right_btn.clicked.connect(lambda: self._on_rotate(90))
        self.rotate_180_btn.clicked.connect(lambda: self._on_rotate(180))
        for btn in (self.rotate_left_btn, self.rotate_right_btn, self.rotate_180_btn):
            btn.setToolTip("回転は即座にファイルへ反映されます")

        form = QFormLayout()
        form.addRow("氏名（漢字）", self.name)
        form.addRow("氏名（カナ）", self.name_kana)
        form.addRow("照合キー", self.match_key)
        form.addRow("入社年", self.join_year)
        form.addRow("役職", self.role)
        form.addRow("役職記号", self.role_marks)
        form.addRow("雇用区分", self.employment_type)
        form.addRow("ステータス", self.status)
        form.addRow("備考", self.note)
        # 所属（本部・支店・課）はテンプレートExcel側で管理するためフォームから除外

        rotate_row = QHBoxLayout()
        rotate_row.addWidget(self.rotate_left_btn)
        rotate_row.addWidget(self.rotate_right_btn)
        rotate_row.addWidget(self.rotate_180_btn)

        photo_box = QVBoxLayout()
        photo_box.addWidget(self.preview)
        photo_box.addWidget(self.choose_photo_btn)
        photo_box.addLayout(rotate_row)
        photo_box.addStretch(1)

        body = QHBoxLayout()
        body.addLayout(form, 2)
        body.addLayout(photo_box, 1)

        bbox = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        bbox.accepted.connect(self._on_save)
        bbox.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addLayout(body)
        layout.addWidget(bbox)

        if employee:
            self._fill_from(employee)

    def _fill_from(self, emp) -> None:
        self.name.setText(emp.name or "")
        self.name_kana.setText(emp.name_kana or "")
        self.match_key.setText(emp.match_key or "")
        if emp.join_year:
            self.join_year.setValue(emp.join_year)
        self.role.setText(emp.role or "")
        self.role_marks.setText(emp.role_marks or "")
        self.employment_type.setCurrentText(emp.employment_type or "")
        idx = self.status.findData(emp.status)
        if idx >= 0:
            self.status.setCurrentIndex(idx)
        self.note.setText(emp.note or "")
        p = self.photo_service.resolve(emp.photo_path)
        if p:
            self._working_path = p
            self._working_is_temp = False
            self._show_preview(p)
        self._update_rotate_buttons()

    def _show_preview(self, path: Path) -> None:
        # QPixmap はキャッシュを利用するため、回転後の同一パスを再読込する場合は
        # 一度バイト経由で loadFromData する必要がある。
        pm = QPixmap()
        try:
            pm.loadFromData(Path(path).read_bytes())
        except Exception:
            pm = QPixmap(str(path))
        if pm.isNull():
            self.preview.setText("(読込失敗)")
            return
        scaled = pm.scaled(
            self.preview.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation
        )
        self.preview.setPixmap(scaled)

    def _update_rotate_buttons(self) -> None:
        enabled = self._working_path is not None
        for btn in (self.rotate_left_btn, self.rotate_right_btn, self.rotate_180_btn):
            btn.setEnabled(enabled)

    def _on_choose_photo(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "顔写真を選択", str(Path.home()),
            "画像ファイル (*.jpg *.jpeg *.png)",
        )
        if not path:
            return
        # 編集中の回転操作のため、選択ファイルを一時ディレクトリにコピー
        ext = Path(path).suffix or ".jpg"
        tmp = Path(tempfile.gettempdir()) / f"miki_edit_{uuid.uuid4().hex}{ext}"
        shutil.copy2(path, tmp)
        self._cleanup_temp()
        self._working_path = tmp
        self._working_is_temp = True
        self._working_dirty = True
        self._new_photo_src = str(tmp)
        self._show_preview(tmp)
        self._update_rotate_buttons()

    def _on_rotate(self, degrees: int) -> None:
        if self._working_path is None:
            QMessageBox.information(
                self, "写真未設定",
                "回転する写真がありません。先に「写真を選択」してください。"
            )
            return
        try:
            if self._working_is_temp:
                self._rotate_file_inplace(self._working_path, degrees)
            else:
                # 保存済み写真は PhotoService 経由で回転（拡張子変化に対応）
                if self.employee and self.employee.photo_path:
                    new_path = self.photo_service.rotate_photo(
                        self.employee.photo_path, degrees
                    )
                    if new_path:
                        self.employee.photo_path = new_path
                        resolved = self.photo_service.resolve(new_path)
                        if resolved:
                            self._working_path = resolved
        except Exception as e:
            QMessageBox.critical(self, "回転失敗", str(e))
            return
        self._show_preview(self._working_path)

    @staticmethod
    def _rotate_file_inplace(path: Path, degrees: int) -> None:
        img = Image.open(path)
        img = ImageOps.exif_transpose(img)
        rotated = img.rotate(-degrees, expand=True, resample=Image.BICUBIC)
        fmt = (img.format or "JPEG").upper()
        if fmt == "JPG":
            fmt = "JPEG"
        if fmt == "JPEG" and rotated.mode in ("RGBA", "LA", "P"):
            rotated = rotated.convert("RGB")
        if fmt == "JPEG":
            rotated.save(path, format=fmt, quality=92)
        else:
            rotated.save(path, format=fmt)

    def _cleanup_temp(self) -> None:
        if self._working_is_temp and self._working_path:
            try:
                self._working_path.unlink(missing_ok=True)
            except Exception:
                pass

    def _on_save(self) -> None:
        if not self.name.text().strip():
            self.name.setFocus()
            return
        if not self.match_key.text().strip():
            self.match_key.setFocus()
            return
        self.accept()

    def collect(self) -> dict:
        return {
            "name": self.name.text().strip(),
            "name_kana": self.name_kana.text().strip() or None,
            "match_key": self.match_key.text().strip(),
            "join_year": self.join_year.value(),
            "role": self.role.text().strip() or None,
            "role_marks": self.role_marks.text().strip() or None,
            "employment_type": self.employment_type.currentText() or None,
            "status": self.status.currentData(),
            "note": self.note.text().strip() or None,
            "_photo_src": self._new_photo_src,
        }
