"""データフォルダの全写真を EXIF Orientation に従って物理的に回転正規化する。

スマホ撮影写真などで EXIF 回転情報のみ保存され、Excel やビューワで横倒し
表示される問題を一括解決するためのワンショットスクリプト。

実行例:
    python3 -m src.cli.normalize_photos
    python3 -m src.cli.normalize_photos --data-dir /path/to/data
"""
from __future__ import annotations
import argparse
import sys
from pathlib import Path

from ..config import Settings, DataPaths
from ..services.photo_service import normalize_image_bytes


SUPPORTED_EXTS = {".jpg", ".jpeg", ".png"}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--data-dir", type=Path, default=None,
                        help="データフォルダ（既定: settings.json の data_dir）")
    parser.add_argument("--dry-run", action="store_true",
                        help="変換せずに対象ファイルだけ表示")
    args = parser.parse_args()

    settings = Settings()
    data_dir = args.data_dir or settings.data_dir
    paths = DataPaths(data_dir)
    photos_dir = paths.photos_dir

    if not photos_dir.is_dir():
        print(f"[error] photos フォルダが見つかりません: {photos_dir}", file=sys.stderr)
        return 1

    targets = [p for p in photos_dir.iterdir()
               if p.is_file() and p.suffix.lower() in SUPPORTED_EXTS]
    print(f"📂 対象フォルダ: {photos_dir}")
    print(f"   対象ファイル数: {len(targets)}")

    if args.dry_run:
        for p in targets[:10]:
            print(f"  - {p.name}")
        if len(targets) > 10:
            print(f"  ... 他 {len(targets) - 10}件")
        print("\n[dry-run] --dry-run を外すと実変換します。")
        return 0

    rotated = 0
    skipped = 0
    failed = 0
    for p in targets:
        try:
            original = p.read_bytes()
            normalized, _ = normalize_image_bytes(
                original, format_hint=p.suffix.lstrip(".").lower()
            )
            if normalized == original:
                skipped += 1
                continue
            p.write_bytes(normalized)
            rotated += 1
        except Exception as e:
            failed += 1
            print(f"  [warn] {p.name}: {e}", file=sys.stderr)

    print(f"\n✅ 正規化完了")
    print(f"   変換: {rotated}件")
    print(f"   変換不要: {skipped}件")
    if failed:
        print(f"   失敗: {failed}件")
    return 0 if not failed else 2


if __name__ == "__main__":
    sys.exit(main())
