"""GitHub Releases を見て新しい Roster (onedir zip) があれば差し替える。

リポジトリは public 前提のため認証は不要。
ローカル開発時（`__build_sha__` が空）はチェックを必ずスキップする。

配布形式は --onedir 構成のフォルダを zip 化したもの（Roster.zip）。
zip 内のルートに `Roster/` ディレクトリがあり、その中に Roster.exe + 依存
ファイル一式が含まれる。
"""
from __future__ import annotations
import json
import os
import shutil
import ssl
import subprocess
import sys
import urllib.request
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

import certifi

from .._version import __build_sha__


def _ssl_context() -> ssl.SSLContext:
    """certifi の CA bundle を使う SSL コンテキストを返す。

    PyInstaller 同梱の Python は OS の証明書ストアを参照できないため、
    certifi をバンドルしてここから読ませる。これをしないと Windows 環境で
    `CERTIFICATE_VERIFY_FAILED` が出る。
    """
    return ssl.create_default_context(cafile=certifi.where())


REPO_OWNER = "hiroki077"
REPO_NAME = "Tanaka_project"
ASSET_NAME = "Roster.zip"
APP_DIR_NAME = "Roster"  # zip 内のルートディレクトリ名 / インストール先のフォルダ名
RELEASES_API = (
    f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/releases/latest"
)
_USER_AGENT = "Roster-Updater"


@dataclass
class UpdateInfo:
    tag: str
    published_at: str
    asset_url: str
    asset_size: int
    latest_sha: str
    current_sha: str


def is_frozen() -> bool:
    """PyInstaller でパッケージ化された exe として動いているか。"""
    return bool(getattr(sys, "frozen", False))


def current_build_sha() -> str:
    return __build_sha__


def _parse_sha_from_body(body: str) -> str:
    """Release ノート本文から `コミット: <sha>` を拾う。

    build-windows-exe.yml が生成する Release 本文に
    `コミット: ${{ github.sha }}` が含まれている前提。
    """
    for raw in body.splitlines():
        line = raw.strip().lstrip("-").strip().strip("`").strip()
        for prefix in ("コミット:", "コミット：", "Commit:", "commit:"):
            if prefix in line:
                _, _, rest = line.partition(prefix)
                return rest.strip().strip("`").strip()
    return ""


def _fetch_latest_release(timeout: float = 10.0) -> dict:
    req = urllib.request.Request(
        RELEASES_API,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": _USER_AGENT,
        },
    )
    with urllib.request.urlopen(req, timeout=timeout, context=_ssl_context()) as resp:
        return json.loads(resp.read().decode("utf-8"))


def check_for_update() -> UpdateInfo | None:
    """最新リリースを取得し、現在のビルドより新しければ UpdateInfo を返す。

    開発実行（`__build_sha__` 空）の場合は常に None。
    Roster.zip アセットが無い／SHA を読み取れない場合も None。
    """
    if not __build_sha__:
        return None

    data = _fetch_latest_release()
    body = data.get("body") or ""
    latest_sha = _parse_sha_from_body(body)
    if not latest_sha:
        return None
    if latest_sha[:7].lower() == __build_sha__[:7].lower():
        return None

    for asset in data.get("assets") or []:
        if asset.get("name") == ASSET_NAME:
            return UpdateInfo(
                tag=data.get("tag_name") or "",
                published_at=data.get("published_at") or "",
                asset_url=asset["browser_download_url"],
                asset_size=int(asset.get("size") or 0),
                latest_sha=latest_sha,
                current_sha=__build_sha__,
            )
    return None


def download_asset(
    url: str,
    dest: Path,
    progress_cb: Callable[[int, int], None] | None = None,
    cancel_cb: Callable[[], bool] | None = None,
    timeout: float = 60.0,
) -> None:
    """指定 URL を `dest` に書き出す。

    `progress_cb(written, total)` を逐次呼ぶ。
    `cancel_cb()` が True を返した時点で中断し、書きかけファイルを削除する。
    """
    req = urllib.request.Request(url, headers={"User-Agent": _USER_AGENT})
    dest.parent.mkdir(parents=True, exist_ok=True)
    with urllib.request.urlopen(req, timeout=timeout, context=_ssl_context()) as resp:
        total = int(resp.headers.get("Content-Length") or 0)
        written = 0
        chunk = 64 * 1024
        try:
            with open(dest, "wb") as f:
                while True:
                    if cancel_cb and cancel_cb():
                        raise InterruptedError("download canceled")
                    buf = resp.read(chunk)
                    if not buf:
                        break
                    f.write(buf)
                    written += len(buf)
                    if progress_cb:
                        progress_cb(written, total)
        except BaseException:
            try:
                dest.unlink(missing_ok=True)
            except OSError:
                pass
            raise


# 後方互換用エイリアス（古い main_window 等が呼んでいた場合のフォールバック）
download_exe = download_asset


def install_and_restart(downloaded_zip: Path) -> None:
    """ダウンロードした Roster.zip を解凍し、フォルダ差し替え方式で再起動する。

    onedir 構成では `sys.executable` が `<install>/Roster/Roster.exe` になる。
    install_root = sys.executable.parent.parent。

    手順:
        1. zip を install_root/_roster_new に展開
           → `_roster_new/Roster/` が出来る
        2. updater.bat を install_root に書き出して起動
        3. bat は親プロセス終了を待ち、`Roster/` を消して `_roster_new/Roster` を
           `Roster/` にリネーム、新 exe を起動して自分自身を削除する

    Windows のみ。frozen でない場合は例外。
    """
    if not is_frozen():
        raise RuntimeError("install_and_restart() は frozen ビルドでのみ使えます")

    current_exe = Path(sys.executable).resolve()
    current_dir = current_exe.parent                 # <install>/Roster
    install_root = current_dir.parent                # <install>
    pid = os.getpid()

    # zip を解凍（旧解凍先が残っていたら一度消す）
    extract_dir = install_root / "_roster_new"
    if extract_dir.exists():
        shutil.rmtree(extract_dir, ignore_errors=True)
    extract_dir.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(str(downloaded_zip), "r") as z:
        z.extractall(str(extract_dir))

    new_app_dir = extract_dir / APP_DIR_NAME
    if not new_app_dir.is_dir():
        raise RuntimeError(
            f"展開先に {APP_DIR_NAME}/ ディレクトリが見つかりません: {new_app_dir}"
        )

    new_exe = new_app_dir / current_exe.name
    if not new_exe.is_file():
        raise RuntimeError(f"展開後に exe が見つかりません: {new_exe}")

    bat_path = install_root / "_roster_updater.bat"
    log_path = install_root / "_roster_updater.log"

    # フォルダ差し替え手順を bat に書き出す。失敗時は最大 30 回（=30 秒）リトライ。
    bat_content = (
        "@echo off\r\n"
        "setlocal\r\n"
        f'set "PID={pid}"\r\n'
        f'set "ROOT={install_root}"\r\n'
        f'set "CUR_DIR={current_dir}"\r\n'
        f'set "NEW_DIR={new_app_dir}"\r\n'
        f'set "EXTRACT={extract_dir}"\r\n'
        f'set "ZIP={downloaded_zip}"\r\n'
        f'set "EXE={new_exe}"\r\n'
        f'set "LOG={log_path}"\r\n'
        'echo [%DATE% %TIME%] updater start pid=%PID% > "%LOG%"\r\n'
        ":wait_proc\r\n"
        'tasklist /FI "PID eq %PID%" 2>NUL | findstr /I "%PID%" >NUL\r\n'
        "if %errorlevel% == 0 (\r\n"
        '  timeout /t 1 /nobreak >NUL\r\n'
        "  goto wait_proc\r\n"
        ")\r\n"
        'timeout /t 1 /nobreak >NUL\r\n'
        "set RETRY=0\r\n"
        ":do_remove\r\n"
        'rmdir /s /q "%CUR_DIR%" >>"%LOG%" 2>&1\r\n'
        'if exist "%CUR_DIR%" (\r\n'
        "  set /a RETRY+=1\r\n"
        "  if %RETRY% GEQ 30 (\r\n"
        '    echo remove old folder failed >>"%LOG%"\r\n'
        "    exit /b 1\r\n"
        "  )\r\n"
        '  timeout /t 1 /nobreak >NUL\r\n'
        "  goto do_remove\r\n"
        ")\r\n"
        ":do_move\r\n"
        'move "%NEW_DIR%" "%CUR_DIR%" >>"%LOG%" 2>&1\r\n'
        "if errorlevel 1 (\r\n"
        '  echo move failed >>"%LOG%"\r\n'
        "  exit /b 1\r\n"
        ")\r\n"
        'rmdir /s /q "%EXTRACT%" >>"%LOG%" 2>&1\r\n'
        'del "%ZIP%" >>"%LOG%" 2>&1\r\n'
        'start "" "%CUR_DIR%\\Roster.exe"\r\n'
        '(goto) 2>nul & del "%~f0"\r\n'
    )
    bat_path.write_text(bat_content, encoding="cp932")

    # DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP で親終了に追従させない
    DETACHED_PROCESS = 0x00000008
    subprocess.Popen(
        ["cmd", "/c", str(bat_path)],
        creationflags=DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP,
        close_fds=True,
    )
