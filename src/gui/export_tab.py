"""Excel出力タブ。

マスター体制表 (例: 202605.xlsx) を指定して、支店別シートに展開した
顔写真付き体制図を自動生成する。

入力: マスターxlsx
出力: 支店別シート構成の xlsx
"""
from __future__ import annotations
from pathlib import Path
from datetime import datetime

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFormLayout, QLineEdit, QPushButton,
    QPlainTextEdit, QFileDialog, QGroupBox, QLabel, QMessageBox,
)

from ..config import Settings
from ..db import Database
from ..services import EmployeeService, PhotoService
from ..cli.build_from_master import parse_master, build_workbook


class ExportTab(QWidget):
    def __init__(
        self,
        db: Database,
        employee_service: EmployeeService,
        photo_service: PhotoService,
        settings: Settings,
    ):
        super().__init__()
        self.db = db
        self.employee_service = employee_service
        self.photo_service = photo_service
        self.settings = settings

        layout = QVBoxLayout(self)
        layout.addWidget(self._build_file_group())
        layout.addLayout(self._build_action_bar())
        layout.addWidget(self._build_log_group(), 1)

        self._load_defaults()

    def _build_file_group(self) -> QGroupBox:
        box = QGroupBox("ファイル")
        layout = QFormLayout(box)

        self.master_edit = QLineEdit()
        self.master_edit.setPlaceholderText("マスター体制表 xlsx を選択… (例: 202605.xlsx)")
        btn_m = QPushButton("📁 選択…")
        btn_m.clicked.connect(self._on_pick_master)
        h1 = QHBoxLayout()
        h1.addWidget(self.master_edit, 1)
        h1.addWidget(btn_m)
        layout.addRow("入力 (マスター)", h1)

        self.output_edit = QLineEdit()
        self.output_edit.setPlaceholderText("出力先 xlsx を選択…")
        btn_o = QPushButton("📁 選択…")
        btn_o.clicked.connect(self._on_pick_output)
        h2 = QHBoxLayout()
        h2.addWidget(self.output_edit, 1)
        h2.addWidget(btn_o)
        layout.addRow("出力先", h2)
        return box

    def _build_action_bar(self) -> QHBoxLayout:
        bar = QHBoxLayout()
        bar.addStretch(1)
        self.run_btn = QPushButton("▶ 体制表を生成")
        self.run_btn.setMinimumWidth(180)
        self.run_btn.setStyleSheet(
            "QPushButton { font-weight: bold; padding: 8px 16px; "
            "background: #2E7D5B; color: white; border-radius: 4px; }"
            "QPushButton:hover { background: #1B5E40; }"
        )
        self.run_btn.clicked.connect(self._on_run)
        bar.addWidget(self.run_btn)
        return bar

    def _build_log_group(self) -> QGroupBox:
        box = QGroupBox("実行ログ")
        layout = QVBoxLayout(box)
        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setStyleSheet("font-family: 'Menlo', 'Consolas', monospace; font-size: 11px;")
        layout.addWidget(self.log)
        return box

    def _load_defaults(self) -> None:
        d = self.settings.export_defaults
        if (t := d.get("master_path") or d.get("template_path")):
            if Path(t).is_file():
                self.master_edit.setText(t)
        if (o := d.get("output_path")):
            self.output_edit.setText(o)

    def _save_defaults(self) -> None:
        self.settings.set_export_defaults(
            master_path=self.master_edit.text().strip(),
            output_path=self.output_edit.text().strip(),
        )
        self.settings.save()

    def _on_pick_master(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "マスター体制表 xlsx を選択",
            self.master_edit.text() or str(Path.home()),
            "Excel ファイル (*.xlsx *.xlsm)",
        )
        if not path:
            return
        self.master_edit.setText(path)
        if not self.output_edit.text().strip():
            stem = Path(path).stem
            self.output_edit.setText(
                str(Path(path).parent / f"{stem}_顔写真版_{datetime.now():%Y%m%d_%H%M}.xlsx")
            )

    def _on_pick_output(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self, "出力先 xlsx を選択",
            self.output_edit.text() or str(Path.home()),
            "Excel ファイル (*.xlsx)",
        )
        if path:
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            self.output_edit.setText(path)

    def _on_run(self) -> None:
        master = self.master_edit.text().strip()
        output = self.output_edit.text().strip()
        if not master or not Path(master).is_file():
            QMessageBox.warning(self, "エラー", "マスター xlsx を指定してください。")
            return
        if not output:
            QMessageBox.warning(self, "エラー", "出力先を指定してください。")
            return

        self._save_defaults()
        self.log.clear()
        self._append_log(f"📂 入力: {master}")
        self._append_log(f"💾 出力: {output}")
        self._append_log("")

        try:
            self._append_log("① マスター体制表を解析中…")
            branches = parse_master(Path(master))
            for b in branches:
                total = sum(len(cr.persons) for cs in b.sections.values() for cr in cs)
                extra = len(b.jimu) + len(b.keiyaku) + len(b.haken)
                top = len(b.top_positions)
                self._append_log(
                    f"   - {b.branch_name}: 上位{top} / 総合職{total} / その他{extra}"
                )

            self._append_log("")
            self._append_log("② 写真を配置中…")
            wb, stats = build_workbook(branches, self.employee_service, self.photo_service)

            self._append_log("")
            self._append_log("③ ファイルを保存中…")
            Path(output).parent.mkdir(parents=True, exist_ok=True)
            wb.save(output)

            self._append_log("")
            self._append_log(f"=== 結果 ===")
            self._append_log(f"   写真挿入: {stats['placed']}枚")
            self._append_log(f"   休職スキップ: {len(stats['on_leave'])}名")
            unmatched = sorted(set(stats['unmatched']))
            if unmatched:
                self._append_log(f"   DB未登録: {len(unmatched)}名")
                for n in unmatched[:30]:
                    self._append_log(f"     - {n}")
                if len(unmatched) > 30:
                    self._append_log(f"     ... 他 {len(unmatched) - 30}名")

            self._append_log("")
            self._append_log(f"✅ 出力完了: {output}")
            QMessageBox.information(
                self, "出力完了",
                f"{stats['placed']}枚の写真を配置しました。\n\n{output}"
            )
        except Exception as e:
            import traceback
            self._append_log(f"\n[エラー] {e}")
            self._append_log(traceback.format_exc())
            QMessageBox.critical(self, "出力失敗", str(e))

    def _append_log(self, line: str) -> None:
        self.log.appendPlainText(line)
