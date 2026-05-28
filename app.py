"""PyInstaller 用のエントリーポイントスクリプト。

src/main.py は相対インポート (`from .config import ...`) を使うため、
直接 PyInstaller のエントリにすると親パッケージが解決できず
ImportError: attempted relative import with no known parent package
が発生する。

このラッパー経由でパッケージとして呼び出すことで問題を回避する。
"""
from __future__ import annotations
import sys
from pathlib import Path

# PyInstaller の onefile 展開先 (_MEIPASS) を sys.path に追加して
# パッケージ src を確実に解決できるようにする
if getattr(sys, "frozen", False):
    base = Path(getattr(sys, "_MEIPASS", Path(sys.argv[0]).parent))
    sys.path.insert(0, str(base))
else:
    sys.path.insert(0, str(Path(__file__).resolve().parent))

from src.main import main


if __name__ == "__main__":
    sys.exit(main())
