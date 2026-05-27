# 配布手順 (Mac から Windows 用 EXE をリリース)

Mac では Windows の `.exe` を直接ビルドできないため、**GitHub Actions の Windows ランナーでビルドし、その成果物 (artifact) を Mac でダウンロードしてDBと一緒に zip 化** する流れにしています。

## 全体の流れ

```
[Mac]                       [GitHub]                   [Mac]
1. コードを push        →   2. GHA が Windows で     ←  3. release.sh で
                              MikiApp.exe をビルド       artifact を取得
                                                          + データフォルダを同梱
                                                          + zip 化
                                                              ↓
                                                        4. zip をクライアントに送付
```

## 事前準備（一度だけ）

```bash
# GitHub CLI を導入してログイン
brew install gh
gh auth login    # ブラウザで認証
```

## ステップ 1: ビルドをトリガー

`src/` 以下のコードを変更して push すれば自動的にビルドが走ります。

```bash
git add src/
git commit -m "コード修正"
git push
```

または、手動でビルドだけ走らせたい場合:

```bash
gh workflow run build-windows-exe.yml
```

または GitHub の Web UI で:
- Actions タブ → `Build Windows EXE` → "Run workflow"

## ステップ 2: ビルド完了待ち

```bash
# ビルド進捗を確認
gh run list --workflow build-windows-exe.yml --limit 3

# 最新ビルドが完了するまで待つ（5〜10分）
gh run watch
```

ブラウザで確認する場合:
https://github.com/hiroki077/Tanaka_project/actions

## ステップ 3: Mac で配布用 zip を作成

```bash
bash scripts/release.sh
```

これで以下が自動的に行われます:
1. 最新の成功ビルドから `MikiApp.exe` をダウンロード
2. あなたのデータフォルダ（`settings.json` の `data_dir`）から DB と写真を取得
3. README を生成して同梱
4. `.lock` は除外（受け取った人で起動可能にするため）
5. `dist/MikiApp_配布版_YYYYMMDD_HHMM.zip` を生成

### オプション

```bash
# DBを含めず、exe だけ配布する場合
bash scripts/release.sh --no-data
```

## ステップ 4: クライアントに送付

生成された zip ファイルを:
- メール添付（サイズが許せば）
- OneDrive / Google Drive 等の共有リンク
- USB メモリ

などで配布します。

## 受け取った人の手順

1. zip を任意の場所に展開（例: デスクトップ）
2. `MikiApp.exe` をダブルクリック
3. 初回起動でデータフォルダ選択 → 同梱の `初期データ` フォルダを指定
4. 297名のデータが読み込まれて利用開始

zip 内に同梱した `はじめにお読みください.txt` にも同じ手順が書かれています。

## トラブルシューティング

### `gh: command not found`
→ `brew install gh && gh auth login` で導入＆認証

### `成功したビルドが見つかりません`
→ Actions が完了していない、または失敗している可能性。
   `gh run list --workflow build-windows-exe.yml` で状態確認。

### Windows でアンチウイルス警告が出る
→ PyInstaller ビルドの exe は誤検知されやすい既知問題。
   - 社内配布: ホワイトリスト登録依頼
   - 本格運用: コード署名証明書で署名（年額〜数万円）

### ビルドサイズが大きい
→ PySide6 を丸ごと同梱するため約 150MB 程度になります。
   `--collect-all PySide6` を `--collect-submodules PySide6` 等に絞ると軽量化可能ですが、機能不足のリスクあり。
