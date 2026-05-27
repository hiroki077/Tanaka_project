from __future__ import annotations
from pathlib import Path
import io
import uuid

from PIL import Image, ImageOps


SAFE_FORMATS = {"JPEG", "PNG"}


def normalize_image_bytes(data: bytes, format_hint: str | None = None) -> tuple[bytes, str]:
    """EXIF Orientation を物理的な回転に反映し、整形済みバイト列＋拡張子を返す。

    スマホで撮影した JPEG は EXIF Orientation で「向き」だけメタデータ化されており、
    ビューワが対応していないと横倒し表示になる。Pillow の exif_transpose で
    実際のピクセル回転に変換し、Excel 等どこでも正しい向きで表示されるようにする。
    """
    img = Image.open(io.BytesIO(data))
    img = ImageOps.exif_transpose(img)
    fmt = (format_hint or img.format or "JPEG").upper()
    if fmt == "JPG":
        fmt = "JPEG"
    if fmt not in SAFE_FORMATS:
        fmt = "JPEG"
    if fmt == "JPEG" and img.mode in ("RGBA", "LA", "P"):
        img = img.convert("RGB")
    out = io.BytesIO()
    if fmt == "JPEG":
        img.save(out, format=fmt, quality=92, optimize=True)
    else:
        img.save(out, format=fmt, optimize=True)
    ext = "jpg" if fmt == "JPEG" else fmt.lower()
    return out.getvalue(), ext


class PhotoService:
    """顔写真ファイルのコピー保存・取得を担当。

    保存時に EXIF Orientation を物理回転に反映するため、どのアプリで表示しても
    正しい向きになる。
    """

    def __init__(self, photos_dir: Path):
        self.photos_dir = Path(photos_dir)
        self.photos_dir.mkdir(parents=True, exist_ok=True)

    def import_photo(self, source_path: str | Path) -> str:
        """元画像を photos/ 配下にコピーし、保存後の相対パス文字列を返す。

        EXIF Orientation を正規化して保存する（横倒し写真の防止）。
        """
        src = Path(source_path)
        if not src.is_file():
            raise FileNotFoundError(f"画像が見つかりません: {src}")

        raw = src.read_bytes()
        format_hint = src.suffix.lstrip(".").lower() if src.suffix else None
        normalized, ext = normalize_image_bytes(raw, format_hint=format_hint)

        dest_name = f"{uuid.uuid4().hex}.{ext}"
        (self.photos_dir / dest_name).write_bytes(normalized)
        return dest_name

    def rotate_photo(self, photo_path: str | None, degrees: int) -> str | None:
        """既存の写真ファイルを指定角度で回転して上書き保存。

        Args:
            photo_path: DB に保存されている相対パス（拡張子付き）
            degrees: 90 / 180 / 270 / -90 など（時計回り）
        Returns:
            新しい photo_path（拡張子変化があるため）。元と同名で上書きできた場合は同じ値を返す。
        """
        if not photo_path:
            return None
        target = self.resolve(photo_path)
        if target is None or not target.is_file():
            raise FileNotFoundError(f"写真が見つかりません: {photo_path}")

        img = Image.open(target)
        img = ImageOps.exif_transpose(img)
        # Excel と同じ「時計回り」で回す ＝ PIL では -degrees
        rotated = img.rotate(-degrees, expand=True, resample=Image.BICUBIC)

        fmt = (img.format or "JPEG").upper()
        if fmt == "JPG":
            fmt = "JPEG"
        if fmt not in SAFE_FORMATS:
            fmt = "JPEG"
        if fmt == "JPEG" and rotated.mode in ("RGBA", "LA", "P"):
            rotated = rotated.convert("RGB")

        ext = "jpg" if fmt == "JPEG" else fmt.lower()
        new_name = f"{target.stem}.{ext}" if target.suffix.lstrip(".").lower() == ext else f"{target.stem}.{ext}"
        new_path = self.photos_dir / new_name

        # 一旦バイトに書き出してから target に書く（同名上書き）
        buf = io.BytesIO()
        if fmt == "JPEG":
            rotated.save(buf, format=fmt, quality=92, optimize=True)
        else:
            rotated.save(buf, format=fmt, optimize=True)
        buf.seek(0)
        new_path.write_bytes(buf.getvalue())
        # 旧ファイルが拡張子違いで残った場合は消す
        if new_path != target:
            try:
                target.unlink()
            except FileNotFoundError:
                pass
        return new_name

    def resolve(self, photo_path: str | None) -> Path | None:
        """DBに保存されている `photo_path` を絶対パスに解決する。"""
        if not photo_path:
            return None
        p = Path(photo_path)
        if p.is_absolute() and p.exists():
            return p
        candidate = self.photos_dir / photo_path
        return candidate if candidate.exists() else None

    def delete(self, photo_path: str | None) -> None:
        target = self.resolve(photo_path)
        if target and target.is_file():
            target.unlink(missing_ok=True)
