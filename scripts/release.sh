#!/usr/bin/env bash
# Mac から配布用 zip を作るスクリプト
#
# 動作:
#   1. GitHub Release で公開された最新の Roster.exe をダウンロード
#   2. 現在のデータフォルダ (settings.json の data_dir) を同梱
#   3. ./dist/Roster_dist_YYYYMMDD_HHMM.zip を生成
#
# Windows Explorer 文字化け対策: zip 内のフォルダ/ファイル名はすべて ASCII にする
# (中身のテキストは日本語でOK)
#
# 実行:
#   bash scripts/release.sh
#   bash scripts/release.sh --no-data    # exe だけ配布したい場合

set -euo pipefail

REPO_OWNER="hiroki077"
REPO_NAME="Tanaka_project"
RELEASE_TAG="latest"
ASSET_NAME="Roster.zip"

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
DIST_DIR="dist/Roster_dist_${TIMESTAMP}"
ZIP_PATH="dist/Roster_dist_${TIMESTAMP}.zip"

mkdir -p "$DIST_DIR"

echo "📥 GitHub Release から最新の Roster.zip (onedir) をダウンロードします..."
gh release download "$RELEASE_TAG" \
  --repo "$REPO_OWNER/$REPO_NAME" \
  --pattern "$ASSET_NAME" \
  --dir "$DIST_DIR" \
  --clobber

if [[ ! -f "$DIST_DIR/$ASSET_NAME" ]]; then
  echo "❌ $ASSET_NAME のダウンロードに失敗しました。" >&2
  echo "   ブラウザで Releases を確認してください:" >&2
  echo "   https://github.com/$REPO_OWNER/$REPO_NAME/releases" >&2
  exit 1
fi
ZIP_SIZE_RAW="$(du -h "$DIST_DIR/$ASSET_NAME" | awk '{print $1}')"
echo "✅ $ASSET_NAME ダウンロード完了 ($ZIP_SIZE_RAW)"

# zip を解凍 → DIST_DIR/Roster/ ができる
echo "📂 解凍中..."
unzip -q "$DIST_DIR/$ASSET_NAME" -d "$DIST_DIR"
rm "$DIST_DIR/$ASSET_NAME"
if [[ ! -f "$DIST_DIR/Roster/Roster.exe" ]]; then
  echo "❌ 解凍後に Roster/Roster.exe が見つかりません。" >&2
  exit 1
fi
APP_SIZE="$(du -sh "$DIST_DIR/Roster" | awk '{print $1}')"
echo "✅ 解凍完了: Roster/ ($APP_SIZE)"

# データフォルダの同梱（フォルダ名は ASCII で 'data' に統一）
if $INCLUDE_DATA; then
  SETTINGS_FILE="$HOME/Library/Application Support/Roster/settings.json"
  # 旧 MikiApp 設定もフォールバックで参照
  if [[ ! -f "$SETTINGS_FILE" ]]; then
    SETTINGS_FILE="$HOME/Library/Application Support/MikiApp/settings.json"
  fi
  if [[ -f "$SETTINGS_FILE" ]]; then
    DATA_DIR="$(python3 -c "import json,sys; print(json.load(open('$SETTINGS_FILE'))['data_dir'])")"
  else
    DATA_DIR=""
  fi

  if [[ -n "$DATA_DIR" && -d "$DATA_DIR" ]]; then
    echo "📦 データフォルダを同梱します: $DATA_DIR"
    mkdir -p "$DIST_DIR/data"
    rsync -av --exclude '.lock' --exclude '*.db-journal' \
      "$DATA_DIR/" "$DIST_DIR/data/" >/dev/null
    DATA_SIZE="$(du -sh "$DIST_DIR/data" | awk '{print $1}')"
    echo "✅ データフォルダ同梱完了 ($DATA_SIZE)"
  else
    echo "⚠ データフォルダが見つかりません ($SETTINGS_FILE)。EXEのみで配布します。"
  fi
fi

# zip 化（フォルダ/ファイル名はすべて ASCII にしたので文字化けしない）
echo "📦 zip を作成しています..."
cd dist
zip -r "$(basename "$ZIP_PATH")" "$(basename "$DIST_DIR")" >/dev/null
cd - >/dev/null

ZIP_SIZE="$(du -h "$ZIP_PATH" | awk '{print $1}')"
echo ""
echo "🎉 配布用 zip 完成: $ZIP_PATH ($ZIP_SIZE)"
echo ""
echo "次のステップ:"
echo "  1. zip ファイルをクライアントに渡す（メール添付/USB/OneDrive 等）"
echo "  2. クライアントは zip を展開して Roster/Roster.exe を起動（onedir 構成）"
echo "     ※ Roster.exe 単体を別フォルダにコピーしないでください（同フォルダ内の"
echo "       _internal/ などが必要なため、起動できなくなります）"
