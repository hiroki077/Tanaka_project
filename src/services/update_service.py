"""GitHub Releases を見て新しい Roster.exe があれば差し替える。

リポジトリは public 前提のため認証は不要。
ローカル開発時（`__build_sha__` が空）はチェックを必ずスキップする。
"""
from __future__ import annotations
import json
import os
import subprocess
import sys
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from .._version import __build_sha__


REPO_OWNER = "hiroki077"
REPO_NAME = "Tanaka_project"
ASSET_NAME = "Roster.exe"
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
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))


def check_for_update() -> UpdateInfo | None:
    """最新リリースを取得し、現在のビルドより新しければ UpdateInfo を返す。

    開発実行（`__build_sha__` 空）の場合は常に None。
    Roster.exe アセットが無い／SHA を読み取れない場合も None。
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


def download_exe(
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
    with urllib.request.urlopen(req, timeout=timeout) as resp:
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


def install_and_restart(new_exe: Path) -> None:
    """旧 exe を `new_exe` で差し替えて再起動するヘルパ bat を起動し、
    現プロセスは即座に呼び出し元に戻る（呼び出し元が QApplication.quit()）。

    Windows のみ。frozen でない場合は例外。
    """
    if not is_frozen():
        raise RuntimeError("install_and_restart() は frozen ビルドでのみ使えます")

    current_exe = Path(sys.executable).resolve()
    new_exe = Path(new_exe).resolve()
    pid = os.getpid()
    bat_path = current_exe.parent / "_roster_updater.bat"
    log_path = current_exe.parent / "_roster_updater.log"

    # 旧 exe がロック解放されるまで最大 30 秒リトライ。
    # 成功したら新 exe を起動して bat を自己削除。
    bat_content = (
        "@echo off\r\n"
        "setlocal\r\n"
        f'set "PID={pid}"\r\n'
        f'set "CUR={current_exe}"\r\n'
        f'set "NEW={new_exe}"\r\n'
        f'set "LOG={log_path}"\r\n'
        'echo [%DATE% %TIME%] updater start pid=%PID% > "%LOG%"\r\n'
        ":wait_proc\r\n"
        'tasklist /FI "PID eq %PID%" 2>NUL | findstr /I "%PID%" >NUL\r\n'
        "if %errorlevel% == 0 (\r\n"
        '  timeout /t 1 /nobreak >NUL\r\n'
        "  goto wait_proc\r\n"
        ")\r\n"
        "set RETRY=0\r\n"
        ":do_move\r\n"
        'move /Y "%NEW%" "%CUR%" >>"%LOG%" 2>&1\r\n'
        "if errorlevel 1 (\r\n"
        "  set /a RETRY+=1\r\n"
        "  if %RETRY% GEQ 30 (\r\n"
        '    echo move failed after retries >>"%LOG%"\r\n'
        "    exit /b 1\r\n"
        "  )\r\n"
        '  timeout /t 1 /nobreak >NUL\r\n'
        "  goto do_move\r\n"
        ")\r\n"
        'start "" "%CUR%"\r\n'
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
