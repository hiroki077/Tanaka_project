"""共有データフォルダ用のアドバイザリロック。

OneDrive 等で複数PCがDBを共有する場合、SQLite はバイナリ書き込み中の
同期で破損する可能性があるため、起動時にロックを取得して排他制御する。

ロック方式: 単純なファイルベース（PID とホスト名を書き込む）。
強制ロックではないため、別アプリ経由の書き込みは防げないが、
本アプリの複数起動を検知して警告するには十分。
"""
from __future__ import annotations
import json
import os
import socket
import time
from dataclasses import dataclass
from pathlib import Path


@dataclass
class LockInfo:
    pid: int
    host: str
    acquired_at: float

    def to_json(self) -> str:
        return json.dumps({
            "pid": self.pid,
            "host": self.host,
            "acquired_at": self.acquired_at,
        }, ensure_ascii=False)

    @classmethod
    def from_json(cls, text: str) -> "LockInfo":
        d = json.loads(text)
        return cls(pid=int(d["pid"]), host=str(d["host"]),
                   acquired_at=float(d["acquired_at"]))


class DataLock:
    def __init__(self, lock_path: Path, stale_seconds: float = 12 * 3600):
        self.lock_path = lock_path
        self.stale_seconds = stale_seconds
        self._info: LockInfo | None = None

    def read_existing(self) -> LockInfo | None:
        if not self.lock_path.is_file():
            return None
        try:
            return LockInfo.from_json(self.lock_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError, KeyError, ValueError):
            return None

    def is_stale(self, info: LockInfo) -> bool:
        return (time.time() - info.acquired_at) > self.stale_seconds

    def is_self(self, info: LockInfo) -> bool:
        return info.pid == os.getpid() and info.host == socket.gethostname()

    def acquire(self, force: bool = False) -> tuple[bool, LockInfo | None]:
        """ロックを取得する。

        Returns:
            (取得成功か, 既存ロック情報 or None)
        """
        existing = self.read_existing()
        if existing and not self.is_self(existing) and not self.is_stale(existing) and not force:
            return False, existing
        self._info = LockInfo(
            pid=os.getpid(),
            host=socket.gethostname(),
            acquired_at=time.time(),
        )
        self.lock_path.parent.mkdir(parents=True, exist_ok=True)
        self.lock_path.write_text(self._info.to_json(), encoding="utf-8")
        return True, None

    def release(self) -> None:
        existing = self.read_existing()
        if existing and self.is_self(existing):
            self.lock_path.unlink(missing_ok=True)
        self._info = None

    def __enter__(self) -> "DataLock":
        ok, holder = self.acquire()
        if not ok:
            raise RuntimeError(
                f"データフォルダは {holder.host} (PID {holder.pid}) が使用中です"
            )
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        self.release()
