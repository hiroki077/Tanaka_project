from __future__ import annotations
from pathlib import Path

from PIL import Image
from PySide6.QtCore import Qt, QPoint, QRect, QSize, Signal
from PySide6.QtGui import QPainter, QPen, QColor, QPixmap, QBrush, QCursor
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QDialogButtonBox, QMessageBox, QWidget, QSpinBox,
)

from ...config import DEFAULT_CROP_ASPECT_W, DEFAULT_CROP_ASPECT_H


# 四隅ハンドルのヒット判定範囲（ビュー座標 px）
HANDLE_HIT = 12
HANDLE_VISUAL = 8


class _CropCanvas(QWidget):
    """画像表示＋固定アスペクト枠の移動／四隅拡縮。

    画像は短辺をビュー内に収まるよう等比スケールで描画。
    枠は「画像座標」で保持し、ウィンドウサイズ変化に追従させる。
    """

    selectionChanged = Signal()

    def __init__(self, parent=None, aspect_w: float = DEFAULT_CROP_ASPECT_W,
                 aspect_h: float = DEFAULT_CROP_ASPECT_H):
        super().__init__(parent)
        self.setMinimumSize(400, 400)
        self.setMouseTracking(True)
        self._pixmap: QPixmap | None = None
        self._img_size: QSize = QSize(0, 0)
        self._sel_img: QRect = QRect()  # 画像座標系
        self._aspect_w = float(aspect_w)
        self._aspect_h = float(aspect_h)
        self._mode: str | None = None  # None | "move" | "nw" | "ne" | "sw" | "se"
        self._press_view: QPoint = QPoint()
        self._press_rect: QRect = QRect()

    # ----- public API -----

    def set_image(self, path: Path) -> bool:
        pm = QPixmap()
        try:
            pm.loadFromData(Path(path).read_bytes())
        except Exception:
            return False
        if pm.isNull():
            return False
        self._pixmap = pm
        self._img_size = pm.size()
        self._reset_selection()
        self.update()
        self.selectionChanged.emit()
        return True

    def set_aspect(self, w: float, h: float) -> None:
        if w <= 0 or h <= 0:
            return
        self._aspect_w = float(w)
        self._aspect_h = float(h)
        if not self._sel_img.isEmpty():
            # 現在の枠中心を保ったままアスペクトを揃え直す
            cx = self._sel_img.center().x()
            cy = self._sel_img.center().y()
            cur_w = self._sel_img.width()
            cur_h = self._sel_img.height()
            new_h = int(cur_w * self._aspect_h / self._aspect_w)
            if new_h <= self._img_size.height():
                new_w = cur_w
            else:
                new_h = cur_h
                new_w = int(cur_h * self._aspect_w / self._aspect_h)
            new_rect = QRect(cx - new_w // 2, cy - new_h // 2, new_w, new_h)
            self._sel_img = self._clamp_rect(new_rect)
        self.update()
        self.selectionChanged.emit()

    def selection_image_rect(self) -> QRect:
        return QRect(self._sel_img)

    # ----- selection helpers -----

    def _reset_selection(self) -> None:
        iw, ih = self._img_size.width(), self._img_size.height()
        if iw <= 0 or ih <= 0:
            self._sel_img = QRect()
            return
        # アスペクトに合う最大サイズの 90%
        max_w = min(iw, int(ih * self._aspect_w / self._aspect_h))
        max_h = int(max_w * self._aspect_h / self._aspect_w)
        if max_h > ih:
            max_h = ih
            max_w = int(max_h * self._aspect_w / self._aspect_h)
        bw = int(max_w * 0.9)
        bh = int(bw * self._aspect_h / self._aspect_w)
        self._sel_img = QRect((iw - bw) // 2, (ih - bh) // 2, bw, bh)

    def _clamp_rect(self, r: QRect) -> QRect:
        iw, ih = self._img_size.width(), self._img_size.height()
        rw, rh = r.width(), r.height()
        rw = max(20, min(rw, iw))
        rh = max(int(20 * self._aspect_h / self._aspect_w), min(rh, ih))
        # アスペクト維持
        if abs(rw / rh - self._aspect_w / self._aspect_h) > 0.001:
            rh = int(rw * self._aspect_h / self._aspect_w)
            if rh > ih:
                rh = ih
                rw = int(rh * self._aspect_w / self._aspect_h)
        x = max(0, min(r.x(), iw - rw))
        y = max(0, min(r.y(), ih - rh))
        return QRect(x, y, rw, rh)

    # ----- coordinate transforms -----

    def _view_rect(self) -> QRect:
        if not self._pixmap:
            return QRect()
        ww, wh = self.width(), self.height()
        iw, ih = self._img_size.width(), self._img_size.height()
        if iw == 0 or ih == 0:
            return QRect()
        scale = min(ww / iw, wh / ih)
        dw, dh = int(iw * scale), int(ih * scale)
        x = (ww - dw) // 2
        y = (wh - dh) // 2
        return QRect(x, y, dw, dh)

    def _img_to_view_point(self, p: QPoint) -> QPoint:
        view = self._view_rect()
        if view.isEmpty():
            return QPoint()
        sx = view.width() / self._img_size.width()
        sy = view.height() / self._img_size.height()
        return QPoint(view.x() + int(p.x() * sx), view.y() + int(p.y() * sy))

    def _img_to_view_rect(self, r: QRect) -> QRect:
        return QRect(
            self._img_to_view_point(r.topLeft()),
            self._img_to_view_point(r.bottomRight()),
        ).normalized()

    def _view_delta_to_img(self, dx_v: int, dy_v: int) -> tuple[int, int]:
        view = self._view_rect()
        if view.isEmpty() or view.width() == 0 or view.height() == 0:
            return (0, 0)
        sx = self._img_size.width() / view.width()
        sy = self._img_size.height() / view.height()
        return (int(dx_v * sx), int(dy_v * sy))

    # ----- paint -----

    def paintEvent(self, _event) -> None:
        p = QPainter(self)
        p.fillRect(self.rect(), QColor("#202020"))
        if not self._pixmap:
            return
        view = self._view_rect()
        p.drawPixmap(view, self._pixmap, self._pixmap.rect())

        if self._sel_img.isEmpty():
            return
        sel_view = self._img_to_view_rect(self._sel_img)

        # 選択外を半透明黒で覆う
        overlay = QColor(0, 0, 0, 130)
        p.setPen(Qt.NoPen)
        p.setBrush(QBrush(overlay))
        p.drawRect(view.x(), view.y(), view.width(), sel_view.y() - view.y())
        p.drawRect(view.x(), sel_view.bottom() + 1,
                   view.width(), view.bottom() - sel_view.bottom())
        p.drawRect(view.x(), sel_view.y(),
                   sel_view.x() - view.x(), sel_view.height())
        p.drawRect(sel_view.right() + 1, sel_view.y(),
                   view.right() - sel_view.right(), sel_view.height())

        # 選択枠
        pen = QPen(QColor("#00bcd4"))
        pen.setWidth(2)
        p.setPen(pen)
        p.setBrush(Qt.NoBrush)
        p.drawRect(sel_view)

        # 四隅ハンドル
        p.setBrush(QBrush(QColor("#00bcd4")))
        p.setPen(QPen(QColor("#ffffff"), 1))
        for c in self._corners_view(sel_view):
            p.drawRect(c.x() - HANDLE_VISUAL // 2, c.y() - HANDLE_VISUAL // 2,
                       HANDLE_VISUAL, HANDLE_VISUAL)

    @staticmethod
    def _corners_view(rv: QRect) -> list[QPoint]:
        return [rv.topLeft(), rv.topRight(), rv.bottomLeft(), rv.bottomRight()]

    # ----- mouse interaction -----

    def _hit_test(self, vp: QPoint) -> str | None:
        if self._sel_img.isEmpty():
            return None
        sv = self._img_to_view_rect(self._sel_img)
        corners = {
            "nw": sv.topLeft(),
            "ne": sv.topRight(),
            "sw": sv.bottomLeft(),
            "se": sv.bottomRight(),
        }
        for name, pt in corners.items():
            if abs(vp.x() - pt.x()) <= HANDLE_HIT and abs(vp.y() - pt.y()) <= HANDLE_HIT:
                return name
        if sv.contains(vp):
            return "move"
        return None

    def mouseMoveEvent(self, event) -> None:
        vp = event.position().toPoint()
        if self._mode is None:
            # ホバーカーソルだけ更新
            self._update_cursor(self._hit_test(vp))
            return

        dx_v = vp.x() - self._press_view.x()
        dy_v = vp.y() - self._press_view.y()
        dx_i, dy_i = self._view_delta_to_img(dx_v, dy_v)

        if self._mode == "move":
            new_rect = QRect(self._press_rect)
            new_rect.translate(dx_i, dy_i)
            self._sel_img = self._clamp_rect(new_rect)
        else:
            self._sel_img = self._resize_from_corner(self._mode, dx_i, dy_i)
        self.update()
        self.selectionChanged.emit()

    def mousePressEvent(self, event) -> None:
        if not self._pixmap or event.button() != Qt.LeftButton:
            return
        vp = event.position().toPoint()
        hit = self._hit_test(vp)
        if hit is None:
            # 枠外クリック → そのポイント中心に枠を持っていく
            ip = self._view_to_img_point(vp)
            new_rect = QRect(self._sel_img)
            new_rect.moveCenter(ip)
            self._sel_img = self._clamp_rect(new_rect)
            self.update()
            self.selectionChanged.emit()
            hit = "move"
        self._mode = hit
        self._press_view = vp
        self._press_rect = QRect(self._sel_img)
        self._update_cursor(hit)

    def mouseReleaseEvent(self, _event) -> None:
        self._mode = None
        self._update_cursor(None)

    def _view_to_img_point(self, vp: QPoint) -> QPoint:
        view = self._view_rect()
        if view.isEmpty():
            return QPoint()
        sx = self._img_size.width() / view.width()
        sy = self._img_size.height() / view.height()
        x = max(0, min(self._img_size.width(),
                       int((vp.x() - view.x()) * sx)))
        y = max(0, min(self._img_size.height(),
                       int((vp.y() - view.y()) * sy)))
        return QPoint(x, y)

    def _resize_from_corner(self, corner: str, dx_i: int, dy_i: int) -> QRect:
        r = QRect(self._press_rect)
        # 反対側の固定点
        if corner == "nw":
            anchor = r.bottomRight()
        elif corner == "ne":
            anchor = r.bottomLeft()
        elif corner == "sw":
            anchor = r.topRight()
        else:  # "se"
            anchor = r.topLeft()
        # 動かす側の新座標（マウス追従）
        if corner == "nw":
            mover = QPoint(r.left() + dx_i, r.top() + dy_i)
        elif corner == "ne":
            mover = QPoint(r.right() + dx_i, r.top() + dy_i)
        elif corner == "sw":
            mover = QPoint(r.left() + dx_i, r.bottom() + dy_i)
        else:
            mover = QPoint(r.right() + dx_i, r.bottom() + dy_i)
        raw_w = abs(mover.x() - anchor.x())
        raw_h = abs(mover.y() - anchor.y())
        # アスペクト固定：大きい方を採用
        unit = max(raw_w / self._aspect_w, raw_h / self._aspect_h)
        new_w = max(20, int(unit * self._aspect_w))
        new_h = max(20, int(unit * self._aspect_h))
        # 画像外に出ないようサイズ上限
        iw, ih = self._img_size.width(), self._img_size.height()
        if "n" in corner:
            new_h = min(new_h, anchor.y())
        else:
            new_h = min(new_h, ih - anchor.y())
        if "w" in corner:
            new_w = min(new_w, anchor.x())
        else:
            new_w = min(new_w, iw - anchor.x())
        # アスペクト再揃え
        unit2 = min(new_w / self._aspect_w, new_h / self._aspect_h)
        new_w = max(20, int(unit2 * self._aspect_w))
        new_h = max(20, int(unit2 * self._aspect_h))
        # 配置
        if "w" in corner:
            x = anchor.x() - new_w
        else:
            x = anchor.x()
        if "n" in corner:
            y = anchor.y() - new_h
        else:
            y = anchor.y()
        return self._clamp_rect(QRect(x, y, new_w, new_h))

    def _update_cursor(self, mode: str | None) -> None:
        if mode in ("nw", "se"):
            self.setCursor(QCursor(Qt.SizeFDiagCursor))
        elif mode in ("ne", "sw"):
            self.setCursor(QCursor(Qt.SizeBDiagCursor))
        elif mode == "move":
            self.setCursor(QCursor(Qt.SizeAllCursor))
        else:
            self.setCursor(QCursor(Qt.ArrowCursor))


class PhotoCropDialog(QDialog):
    """写真の手動トリミングダイアログ。

    Excel 配置時の比率（既定 4:6 = 列数:行数）を固定した枠を表示し、
    ユーザーは枠の移動と四隅での拡縮で範囲を決める。
    """

    def __init__(
        self,
        parent,
        source_path: Path,
        aspect_w: float = DEFAULT_CROP_ASPECT_W,
        aspect_h: float = DEFAULT_CROP_ASPECT_H,
    ):
        super().__init__(parent)
        self.setWindowTitle("写真をトリミング")
        self.resize(720, 800)
        self._source = Path(source_path)
        self.output_path: Path | None = None

        self.canvas = _CropCanvas(self, aspect_w=aspect_w, aspect_h=aspect_h)
        hint = QLabel(
            "枠の内側ドラッグで移動・四隅ドラッグで拡縮（縦横比は固定）。"
            "枠外をクリックすると枠がそこへ移動します。"
        )
        hint.setWordWrap(True)
        self.size_label = QLabel("")
        self.canvas.selectionChanged.connect(self._update_size_label)

        # アスペクト比 入力
        self.w_spin = QSpinBox()
        self.w_spin.setRange(1, 999)
        self.w_spin.setValue(int(aspect_w))
        self.h_spin = QSpinBox()
        self.h_spin.setRange(1, 999)
        self.h_spin.setValue(int(aspect_h))
        self.w_spin.valueChanged.connect(self._on_aspect_changed)
        self.h_spin.valueChanged.connect(self._on_aspect_changed)

        aspect_row = QHBoxLayout()
        aspect_row.addWidget(QLabel("縦横比 (列:行)"))
        aspect_row.addWidget(self.w_spin)
        aspect_row.addWidget(QLabel(":"))
        aspect_row.addWidget(self.h_spin)
        aspect_row.addStretch(1)
        aspect_row.addWidget(self.size_label)

        bbox = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bbox.accepted.connect(self._on_ok)
        bbox.rejected.connect(self.reject)
        self.ok_btn = bbox.button(QDialogButtonBox.Ok)
        self.ok_btn.setText("この範囲で切り抜く")

        layout = QVBoxLayout(self)
        layout.addWidget(hint)
        layout.addLayout(aspect_row)
        layout.addWidget(self.canvas, 1)
        layout.addWidget(bbox)

        if not self.canvas.set_image(self._source):
            QMessageBox.critical(self, "読み込み失敗", f"画像を開けません: {self._source}")
            self.ok_btn.setEnabled(False)
        else:
            self._update_size_label()

    def _on_aspect_changed(self, _value: int) -> None:
        self.canvas.set_aspect(self.w_spin.value(), self.h_spin.value())

    def _update_size_label(self) -> None:
        r = self.canvas.selection_image_rect()
        self.size_label.setText(f"選択範囲: {r.width()} × {r.height()} px")

    def _on_ok(self) -> None:
        r = self.canvas.selection_image_rect()
        if r.width() < 10 or r.height() < 10:
            QMessageBox.information(self, "範囲が狭すぎます", "もう少し大きく選択してください。")
            return
        try:
            img = Image.open(self._source)
            from PIL import ImageOps
            img = ImageOps.exif_transpose(img)
            box = (r.x(), r.y(), r.x() + r.width(), r.y() + r.height())
            cropped = img.crop(box)
            fmt = (img.format or "JPEG").upper()
            if fmt == "JPG":
                fmt = "JPEG"
            if fmt == "JPEG" and cropped.mode in ("RGBA", "LA", "P"):
                cropped = cropped.convert("RGB")
            ext = "jpg" if fmt == "JPEG" else fmt.lower()
            import tempfile, uuid
            out = Path(tempfile.gettempdir()) / f"miki_crop_{uuid.uuid4().hex}.{ext}"
            if fmt == "JPEG":
                cropped.save(out, format=fmt, quality=92, optimize=True)
            else:
                cropped.save(out, format=fmt, optimize=True)
            self.output_path = out
        except Exception as e:
            QMessageBox.critical(self, "切り抜き失敗", str(e))
            return
        self.accept()
