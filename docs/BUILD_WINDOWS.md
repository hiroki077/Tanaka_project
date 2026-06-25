# Windows向け配布ビルド手順

PyInstaller を使って単一の `.exe` を作成します。Python がインストールされていない PC でも実行可能になります。

## 前提
- Windows 10 / 11
- **Python 3.10 以上**をインストール済み（3.9 以下では PySide6 のインストールに失敗します）
- このリポジトリを `C:\dev\Tanaka_application` などに展開済み

`python --version` で必ず 3.10 以上であることを確認してください。

## 手順

### 1. 仮想環境作成と依存関係インストール

PowerShell を開いて:

```powershell
cd C:\dev\Tanaka_application
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
pip install pyinstaller
```

### 2. アイコン（任意）

`app.ico` を用意する場合はプロジェクト直下に置きます。なくてもビルド可能です。

### 3. ビルド実行

```powershell
pyinstaller `
  --onefile `
  --windowed `
  --name Roster `
  --icon=app.ico `
  --collect-all PySide6 `
  src/main.py
```

オプションの説明:
- `--onefile`: 単一の .exe にまとめる（配布が楽だが起動が少し遅い）
- `--windowed`: コンソール窓を表示しない（GUI専用アプリのため）
- `--collect-all PySide6`: PySide6 の Qt プラグインを漏れなく同梱

成功すると `dist\Roster.exe` が生成されます。

### 4. 動作確認

```powershell
.\dist\Roster.exe
```

初回起動でデータフォルダ選択ダイアログが出ることを確認。

## 配布

`dist\Roster.exe` 1ファイルを社内共有（OneDrive、メール、USB等）すれば配布完了です。

### 推奨配布構成

```
Roster配布用/
├── Roster.exe        # 実行ファイル
├── README.txt         # 利用者向け簡易マニュアル
└── 初期データ/         # （任意）初期写真フォルダ
```

## トラブルシューティング

### `ModuleNotFoundError: No module named 'PySide6.xxx'`
→ `--collect-all PySide6` オプションを追加してください。

### ビルドが終わらない / 巨大ファイルになる
→ `--onefile` を `--onedir` に変えると速くなり、起動も早くなります（配布ファイルがフォルダ単位になります）。

### アンチウイルスが誤検知
→ PyInstaller でビルドした exe は誤検知されやすい既知問題です。
- 社内配布ならホワイトリスト登録を依頼
- 本格運用なら **コード署名証明書** で署名すると改善します

### 初回起動が遅い
→ `--onefile` は実行時に一時フォルダへ展開するため遅くなります。`--onedir` を検討してください。

## バージョン更新時

1. ソースを更新
2. `dist/` と `build/` を削除
3. 上記手順3を再実行
4. 新しい `Roster.exe` を配布

## クロスビルドについて

**macOS / Linux からは Windows 用 .exe をビルドできません。**
必ず Windows 環境で PyInstaller を実行してください。

社内に Windows ビルドマシンがない場合の代替:
- GitHub Actions の `windows-latest` ランナーで自動ビルド
- Windows VM（VirtualBox/Parallels/UTM）を用意
