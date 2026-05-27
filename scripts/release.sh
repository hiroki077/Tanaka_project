#!/usr/bin/env bash
# Mac から配布用 zip を作るスクリプト
#
# 動作:
#   1. GitHub Actions で最新ビルドされた MikiApp.exe をダウンロード
#   2. 現在のデータフォルダ (settings.json の data_dir) を同梱
#   3. ./dist/MikiApp_配布版_YYYYMMDD_HHMM.zip を生成
#
# 前提:
#   - gh CLI がインストール済み (`brew install gh && gh auth login`)
#   - 直近の GitHub Actions ビルドが成功している
#
# 実行:
#   bash scripts/release.sh
#   bash scripts/release.sh --no-data    # exe だけ配布したい場合

set -euo pipefail

REPO_OWNER="hiroki077"
REPO_NAME="Tanaka_project"
ARTIFACT_NAME="MikiApp-windows-exe"

INCLUDE_DATA=true
for arg in "$@"; do
  case "$arg" in
    --no-data) INCLUDE_DATA=false ;;
    -h|--help)
      echo "Usage: $0 [--no-data]"
      echo "  --no-data : DBと写真を同梱せず、exe のみ配布"
      exit 0
      ;;
  esac
done

cd "$(dirname "$0")/.."

if ! command -v gh >/dev/null 2>&1; then
  echo "❌ gh CLI が見つかりません。'brew install gh && gh auth login' を実行してください。" >&2
  exit 1
fi

TIMESTAMP="$(date +%Y%m%d_%H%M)"
DIST_DIR="dist/MikiApp_配布版_${TIMESTAMP}"
ZIP_PATH="dist/MikiApp_配布版_${TIMESTAMP}.zip"

mkdir -p "$DIST_DIR"

echo "📥 GitHub Actions から最新の Windows EXE をダウンロードします..."
LATEST_RUN_ID="$(gh run list --repo "$REPO_OWNER/$REPO_NAME" \
  --workflow build-windows-exe.yml --status success \
  --limit 1 --json databaseId --jq '.[0].databaseId')"

if [[ -z "$LATEST_RUN_ID" ]]; then
  echo "❌ 成功したビルドが見つかりません。" >&2
  echo "   ブラウザで Actions タブを確認してください:" >&2
  echo "   https://github.com/$REPO_OWNER/$REPO_NAME/actions" >&2
  exit 1
fi

echo "   ビルドID: $LATEST_RUN_ID"
gh run download "$LATEST_RUN_ID" \
  --repo "$REPO_OWNER/$REPO_NAME" \
  --name "$ARTIFACT_NAME" \
  --dir "$DIST_DIR"

if [[ ! -f "$DIST_DIR/MikiApp.exe" ]]; then
  echo "❌ MikiApp.exe のダウンロードに失敗しました。" >&2
  exit 1
fi
EXE_SIZE="$(du -h "$DIST_DIR/MikiApp.exe" | awk '{print $1}')"
echo "✅ MikiApp.exe ダウンロード完了 ($EXE_SIZE)"

# データフォルダの同梱
if $INCLUDE_DATA; then
  SETTINGS_FILE="$HOME/Library/Application Support/MikiApp/settings.json"
  if [[ -f "$SETTINGS_FILE" ]]; then
    DATA_DIR="$(python3 -c "import json,sys; print(json.load(open('$SETTINGS_FILE'))['data_dir'])")"
  else
    DATA_DIR=""
  fi

  if [[ -n "$DATA_DIR" && -d "$DATA_DIR" ]]; then
    echo "📦 データフォルダを同梱します: $DATA_DIR"
    mkdir -p "$DIST_DIR/初期データ"
    # .lock は除外（他PCで起動不能になる）
    rsync -av --exclude '.lock' --exclude '*.db-journal' \
      "$DATA_DIR/" "$DIST_DIR/初期データ/" >/dev/null
    DATA_SIZE="$(du -sh "$DIST_DIR/初期データ" | awk '{print $1}')"
    echo "✅ データフォルダ同梱完了 ($DATA_SIZE)"
  else
    echo "⚠ データフォルダが見つかりません ($SETTINGS_FILE)。EXEのみで配布します。"
  fi
fi

# README を生成
cat > "$DIST_DIR/はじめにお読みください.txt" <<'EOF'
==================================================
  Miki 体制表ジェネレーター - 配布版
==================================================

【起動方法】
1. このフォルダを任意の場所に展開してください
   (例: デスクトップ や C:\MikiApp)
2. MikiApp.exe をダブルクリックで起動

【初回起動時】
データフォルダの選択ダイアログが表示されます。
- 同梱の「初期データ」フォルダがある場合: そのフォルダを指定
- ない場合は新規フォルダを作って指定してください

【データを他のメンバーと共有したい場合】
OneDrive の共有フォルダを「データフォルダ」として指定すると、
複数PCで写真・従業員データを共有できます。
※ 同時起動は DB 破損の恐れがあるので避けてください。

【画面構成】
- 「従業員」タブ: 顔写真と氏名の登録・編集・回転
- 「Excel出力」タブ: マスター体制表(202605.xlsx形式)から
                     支店別の顔写真付き体制図を自動生成

【動作環境】
Windows 10 / 11
※ アンチウイルスが誤検知する場合は許可リストへ追加してください。

【お問い合わせ】
管理者まで
EOF

# zip 化
echo "📦 zip を作成しています..."
cd dist
zip -r "$(basename "$ZIP_PATH")" "$(basename "$DIST_DIR")" >/dev/null
cd - >/dev/null

# 後片付け（中間フォルダを残すか確認）
ZIP_SIZE="$(du -h "$ZIP_PATH" | awk '{print $1}')"
echo ""
echo "🎉 配布用 zip 完成: $ZIP_PATH ($ZIP_SIZE)"
echo ""
echo "次のステップ:"
echo "  1. zip ファイルをクライアントに渡す（メール添付/USB/OneDrive 等）"
echo "  2. クライアントは zip を展開して MikiApp.exe を起動"
