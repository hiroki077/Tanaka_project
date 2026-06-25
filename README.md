# 体制表ジェネレーター

顔写真付き Excel 体制表を自動生成する Windows ローカルアプリケーション。

## 概要

毎月の人事異動に伴う「Excel体制表への顔写真貼り付け作業」を自動化します。
従業員DB（写真・所属・役職・ステータス）を管理し、ユーザーが用意した Excel テンプレートに対して、適切な位置に写真を自動配置して出力します。

- **写真配置方式**: プレースホルダー（`{{photo:氏名}}`）と氏名直接検出のハイブリッド
- **DB共有**: SQLite を OneDrive 等のクラウド同期フォルダに配置することで複数PCで共有可能
- **休職フィルタ**: ステータス「休職中」「退職」の従業員は自動的に出力対象外

## ディレクトリ構成

```
Tanaka_application/
├── docs/DESIGN.md       # 設計書
├── src/                 # ソースコード
│   ├── main.py          # エントリーポイント
│   ├── config.py        # 設定（data_dir, settings.json）
│   ├── lock.py          # 共有フォルダ用ロック
│   ├── db/              # DB（SQLAlchemy）
│   ├── services/        # 業務ロジック
│   └── gui/             # PySide6 GUI
├── tests/test_excel_export.py
├── data/                # 開発時のデータフォルダ（gitignore）
├── file/                # サンプル Excel（既存）
└── requirements.txt
```

## 開発環境セットアップ

### 必要なもの
- **Python 3.10 以上**（PySide6 6.10+ が必須要件。Python 3.9 ではインストール時に SyntaxError で失敗します）
- pip

`python --version` で 3.10 以上であることを確認してください。
macOS の場合、システム Python（`/usr/bin/python3`）は 3.9 のことが多いため、pyenv や Homebrew で 3.10+ を導入してください。

### 手順

```bash
cd Tanaka_application
python3 -m venv .venv             # python3 を明示（python だと 3.9 を引く可能性あり）
# Windows: .venv\Scripts\activate
# macOS  : source .venv/bin/activate
python --version                  # 3.10 以上であることを確認
pip install -r requirements.txt
```

### 起動

```bash
python -m src.main
```

初回起動時にデータフォルダの選択ダイアログが出ます。OneDrive で共有する場合は、OneDrive 配下のフォルダを選んでください。

### テスト

```bash
python -m tests.test_excel_export
```

## 配布用 .exe のビルド（Windows）

Windows 環境で以下を実行:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name Roster --icon=app.ico src/main.py
```

`dist/Roster.exe` が生成されます（単一ファイル、Python 不要）。

### PyInstaller 注意点
- macOS/Linux でビルドしたバイナリは Windows で動きません（クロスビルド不可）
- アンチウイルスが誤検知する場合あり → 署名 or 例外設定が必要
- 初回起動が遅い（一時展開のため） → `--onedir` も検討

## OneDrive 共有運用

1. 任意のメンバーが OneDrive にフォルダ（例: `OneDrive/共有/RosterData/`）を作成
2. 他メンバーと共有設定
3. 各PC でアプリを初回起動 → データフォルダにそのフォルダを指定
4. 以降、DB と写真が自動同期される

### 注意
- **同時編集は避けてください**。アプリ起動時にロックファイル (`.lock`) を作成し、他PCで起動中なら警告が出ます
- 同時起動による DB 破損を防ぐため、運用ルールとして「編集前に Teams 等で一声かける」を推奨
- 編集後はアプリを終了し、OneDrive 同期完了を待ってから他者に通知

## テンプレートの書き方

### 方式1: プレースホルダー（推奨・確実）
セルに以下のいずれかを記入:

| 書き方 | 動作 |
|---|---|
| `{{photo:藤堂}}` | 「藤堂」さんの写真を当該セルに貼付 |
| `{{photo:佐藤@東京南}}` | 「東京南」支店所属の「佐藤」さんに限定 |
| `{{photo:藤堂#id=42}}` | 従業員ID=42 の写真に限定 |

### 方式2: 氏名直接検出（既存テンプレ互換）
セルに DB 登録済みの「照合キー」（例: `藤堂`）を書くだけ。
写真は氏名セルから「N行上」（GUI で設定）に貼られます。

## 開発の注意

- **絶対にやってはいけないこと**: `data/` フォルダや DB ファイルを git にコミットする（個人情報含むため）
- 写真のトリミングは事前に手動で済ませた画像のみ登録すること（AIトリミングは未実装）
- Excel テンプレートのレイアウト変更はクライアント側責任
