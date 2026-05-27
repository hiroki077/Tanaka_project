from __future__ import annotations

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QRadioButton, QButtonGroup,
    QDialogButtonBox, QGroupBox,
)

from ...db import Employee
from ...services import AmbiguousMatch


class AmbiguityDialog(QDialog):
    """同名解決のための候補選択ダイアログ。

    candidates の中から1人を選ぶ。「該当なし（スキップ）」も選択可能。
    """

    def __init__(self, parent, ambiguous: AmbiguousMatch):
        super().__init__(parent)
        self.setWindowTitle("同名解決")
        self.ambiguous = ambiguous
        self.choice: Employee | None = None

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(
            f"テンプレ [{ambiguous.sheet_name}!{ambiguous.cell_address}] の "
            f"「{ambiguous.match_key}」に複数候補があります。\n"
            f"該当者を選択してください（次回以降は自動適用されます）。"
        ))

        group_box = QGroupBox("候補")
        gbox_layout = QVBoxLayout(group_box)
        self.button_group = QButtonGroup(self)
        for i, emp in enumerate(ambiguous.candidates):
            label = self._format_candidate(emp)
            rb = QRadioButton(label)
            if i == 0:
                rb.setChecked(True)
            self.button_group.addButton(rb, i)
            gbox_layout.addWidget(rb)
        rb_skip = QRadioButton("どれでもない（スキップ）")
        self.button_group.addButton(rb_skip, -1)
        gbox_layout.addWidget(rb_skip)
        layout.addWidget(group_box)

        bbox = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bbox.accepted.connect(self._on_accept)
        bbox.rejected.connect(self.reject)
        layout.addWidget(bbox)

    @staticmethod
    def _format_candidate(emp: Employee) -> str:
        parts = [emp.name]
        if emp.name_kana:
            parts.append(emp.name_kana)
        if emp.join_year:
            parts.append(f"{emp.join_year}入社")
        if emp.role:
            parts.append(emp.role)
        return " / ".join(parts)

    def _on_accept(self) -> None:
        idx = self.button_group.checkedId()
        if idx == -1:
            self.choice = None
        else:
            self.choice = self.ambiguous.candidates[idx]
        self.accept()
