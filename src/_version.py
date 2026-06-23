"""ビルド時に CI が上書きする想定のバージョン情報。

ローカル開発時はこのファイルのデフォルト値が使われ、更新チェッカは
無効化される（`__build_sha__` が空のため）。
GitHub Actions の build-windows-exe.yml は PyInstaller の実行前に
このファイルを実 SHA とタイムスタンプで書き換える。
"""

__version__ = "dev"
__build_sha__ = ""
__build_time__ = ""
