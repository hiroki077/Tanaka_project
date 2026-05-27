"""DB に登録されている横向き写真を検出して一括補正する。

顔写真は通常「縦長 (height > width)」だが、撮影時の向きやファイル取り込み時の
向き情報の喪失により「横長 (width > height)」になっている場合がある。

このスクリプトは:
1. 全従業員の写真を走査し、横長のものを抽出
2. --apply 付きで実行すると、すべて時計回り90度回転で縦長化
3. 既定はドライラン（リスト表示のみ）

注意:
- 一律「時計回り90度」で補正するため、稀に逆方向だった場合は GUI で修正が必要
- 既に正しい縦長写真には触らない

実行例:
    python3 -m src.cli.fix_rotated_photos              # 横長候補の一覧表示
    python3 -m src.cli.fix_rotated_photos --apply      # 一律 時計回り90度で補正
    python3 -m src.cli.fix_rotated_photos --apply --direction left  # 反時計回り
"""
from __future__ import annotations
import argparse
import sys
from pathlib import Path

from PIL import Image

from ..config import Settings, DataPaths
from ..db import Database
from ..services import EmployeeService, PhotoService


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--apply", action="store_true",
                        help="検出した横長写真を実際に回転する（指定なしはリスト表示のみ）")
    parser.add_argument("--direction", choices=["right", "left"], default="right",
                        help="回転方向（right=時計回り、left=反時計回り）。既定: right")
    parser.add_argument("--threshold", type=float, default=1.0,
                        help="幅÷高さ がこの値より大きい場合に『横長』と判定。既定: 1.0")
    parser.add_argument("--data-dir", type=Path, default=None)
    args = parser.parse_args()

    settings = Settings()
    data_dir = args.data_dir or settings.data_dir
    paths = DataPaths(data_dir)
    db = Database(paths.db_path)
    db.create_all()
    photos = PhotoService(paths.photos_dir)
    es = EmployeeService(db, photos)

    employees = es.list()
    candidates: list[tuple[int, str, int, int, float]] = []
    for emp in employees:
        p = photos.resolve(emp.photo_path)
        if not p:
            continue
        try:
            with Image.open(p) as img:
                w, h = img.size
        except Exception as e:
            print(f"  [warn] {emp.name}: 写真読込失敗: {e}", file=sys.stderr)
            continue
        ratio = w / max(h, 1)
        if ratio > args.threshold:
            candidates.append((emp.id, emp.name, w, h, ratio))

    candidates.sort(key=lambda x: -x[4])

    print(f"📂 データフォルダ: {paths.data_dir}")
    print(f"全従業員: {len(employees)}名")
    print(f"横長と判定 (幅/高さ > {args.threshold}): {len(candidates)}名\n")
    if not candidates:
        print("補正対象なし。")
        return 0

    print(f"{'ID':>5}  {'氏名':25}  {'サイズ':12}  幅/高さ")
    for emp_id, name, w, h, ratio in candidates:
        print(f"{emp_id:>5}  {name:25}  {w:4}x{h:4}    {ratio:.2f}")

    if not args.apply:
        print(f"\n[dry-run] --apply を付けると一括補正します。")
        print(f"          回転方向は --direction right (時計回り) または left (反時計回り)")
        return 0

    degrees = 90 if args.direction == "right" else -90
    print(f"\n→ {args.direction} 90度 回転を {len(candidates)}名に適用します...")
    success, failed = 0, 0
    for emp_id, name, *_ in candidates:
        try:
            es.rotate_photo(emp_id, degrees)
            success += 1
        except Exception as e:
            failed += 1
            print(f"  [error] {name}: {e}", file=sys.stderr)
    print(f"\n✅ 完了: 成功 {success}名 / 失敗 {failed}名")
    print(f"逆方向に回した方が正しかった場合は GUI で個別に補正してください。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
