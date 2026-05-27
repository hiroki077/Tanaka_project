from __future__ import annotations
import re
from pathlib import Path
from contextlib import contextmanager
from typing import Iterator

from sqlalchemy import create_engine, event, select, or_
from sqlalchemy.orm import sessionmaker, Session

from .models import Base, Employee, MappingOverride, OutputHistory, EmploymentStatus


_CONCURRENT_PREFIX_RE = re.compile(r"^[\s　]*兼\s*[)）)）]\s*")


def _normalize_for_match(s: str) -> str:
    """テンプレと DB のテキストを照合用に正規化。

    - 兼務プレフィックス「兼）」「兼)」等を除去（テンプレに残っていても DB と一致）
    - 全角・半角スペースを除去
    """
    s = _CONCURRENT_PREFIX_RE.sub("", s)
    return s.replace("　", "").replace(" ", "").strip()


class Database:
    def __init__(self, db_path: Path):
        self.db_path = db_path
        self.engine = create_engine(
            f"sqlite:///{db_path}",
            echo=False,
            connect_args={"check_same_thread": False},
        )
        self._configure_for_shared_storage()
        self._session_factory = sessionmaker(
            bind=self.engine, expire_on_commit=False
        )

    def _configure_for_shared_storage(self) -> None:
        """OneDrive 等のクラウド同期フォルダで使うため、WALを無効化する。

        WALモードは `*.db-wal` `*.db-shm` 補助ファイルを作成し、これらが
        OneDriveに同期されると整合性が崩れることがある。journal_mode=DELETE
        にすることで補助ファイルが残らず、DB本体のみが同期対象になる。
        """
        @event.listens_for(self.engine, "connect")
        def _set_pragma(dbapi_conn, _conn_record):
            cur = dbapi_conn.cursor()
            try:
                cur.execute("PRAGMA journal_mode=DELETE;")
                cur.execute("PRAGMA synchronous=FULL;")
                cur.execute("PRAGMA foreign_keys=ON;")
            finally:
                cur.close()

    def create_all(self) -> None:
        Base.metadata.create_all(self.engine)
        self._migrate_drop_org_columns()

    def _migrate_drop_org_columns(self) -> None:
        """旧バージョンで作成されたDBから所属カラムを削除する。

        所属（本部・支店・課）はテンプレートExcel側で管理するため、
        DBスキーマからは削除した（A案）。既存DBがあれば DROP COLUMN を実行。
        SQLite 3.35+ 必須。

        ALTER TABLE DROP COLUMN は対象カラムを参照するインデックスを
        自動削除しないため、先に明示的に DROP INDEX する必要がある。
        """
        with self.engine.connect() as conn:
            rows = conn.exec_driver_sql("PRAGMA table_info(employees)").fetchall()
            existing_cols = {r[1] for r in rows}
            drop_cols = [c for c in ("department", "branch", "section") if c in existing_cols]
            if not drop_cols:
                return

            indexes = conn.exec_driver_sql(
                "SELECT name FROM sqlite_master WHERE type='index' AND tbl_name='employees'"
            ).fetchall()
            for (idx_name,) in indexes:
                if any(col in idx_name for col in drop_cols):
                    conn.exec_driver_sql(f"DROP INDEX IF EXISTS {idx_name}")

            for col in drop_cols:
                conn.exec_driver_sql(f"ALTER TABLE employees DROP COLUMN {col}")
            conn.commit()

    @contextmanager
    def session(self) -> Iterator[Session]:
        sess = self._session_factory()
        try:
            yield sess
            sess.commit()
        except Exception:
            sess.rollback()
            raise
        finally:
            sess.close()


class EmployeeRepository:
    def __init__(self, db: Database):
        self.db = db

    def list_all(
        self,
        keyword: str | None = None,
        status_filter: str | None = None,
    ) -> list[Employee]:
        with self.db.session() as s:
            stmt = select(Employee)
            if keyword:
                like = f"%{keyword}%"
                stmt = stmt.where(or_(
                    Employee.name.like(like),
                    Employee.name_kana.like(like),
                    Employee.match_key.like(like),
                ))
            if status_filter:
                stmt = stmt.where(Employee.status == status_filter)
            stmt = stmt.order_by(
                Employee.match_key, Employee.display_order, Employee.id
            )
            return list(s.scalars(stmt))

    def get(self, employee_id: int) -> Employee | None:
        with self.db.session() as s:
            return s.get(Employee, employee_id)

    def find_by_match_key(
        self,
        match_key: str,
        only_active: bool = True,
    ) -> list[Employee]:
        with self.db.session() as s:
            stmt = select(Employee).where(Employee.match_key == match_key)
            if only_active:
                stmt = stmt.where(Employee.status == EmploymentStatus.ACTIVE.value)
            return list(s.scalars(stmt))

    def find_by_text(
        self,
        text: str,
        only_active: bool = True,
    ) -> list[Employee]:
        """セル文字列で従業員を引く。

        マッチング戦略（上から順に試行、ヒットしたら終了）:
        1. 漢字フルネーム または match_key と完全一致
        2. 「苗字＋名前の一部」での前方一致
           例: テンプレ "佐藤俊" → DB name "佐藤　俊基" にマッチ
           ただし target が match_key と同一の場合は前方一致しない（誤検出抑止）

        カナ氏名はマッチ対象から除外（漢字と二重マッチを防止）。
        正規化: 兼務マーカー除去、全角・半角スペース除去。
        """
        target = _normalize_for_match(text)
        if not target:
            return []
        with self.db.session() as s:
            stmt = select(Employee)
            if only_active:
                stmt = stmt.where(Employee.status == EmploymentStatus.ACTIVE.value)
            all_emps = list(s.scalars(stmt))

        # 1) 完全一致
        exact = [
            e for e in all_emps
            if _normalize_for_match(e.name or "") == target
            or _normalize_for_match(e.match_key or "") == target
        ]
        if exact:
            return exact

        # 2) 前方一致（苗字＋名前の一部）
        #    target が match_key と同じ場合は対象外（純粋な苗字検索）
        match_keys = {_normalize_for_match(e.match_key or "") for e in all_emps}
        if target in match_keys:
            return []
        return [
            e for e in all_emps
            if _normalize_for_match(e.name or "").startswith(target)
        ]

    def upsert(self, employee: Employee) -> Employee:
        with self.db.session() as s:
            if employee.id is None:
                s.add(employee)
            else:
                s.merge(employee)
            s.flush()
            return employee

    def delete(self, employee_id: int) -> bool:
        with self.db.session() as s:
            obj = s.get(Employee, employee_id)
            if obj is None:
                return False
            s.delete(obj)
            return True


class MappingOverrideRepository:
    def __init__(self, db: Database):
        self.db = db

    def find(
        self, template_signature: str, sheet_name: str, cell_address: str
    ) -> MappingOverride | None:
        with self.db.session() as s:
            stmt = select(MappingOverride).where(
                MappingOverride.template_signature == template_signature,
                MappingOverride.sheet_name == sheet_name,
                MappingOverride.cell_address == cell_address,
            )
            return s.scalars(stmt).first()

    def save(
        self,
        template_signature: str,
        sheet_name: str,
        cell_address: str,
        match_key: str,
        employee_id: int,
    ) -> None:
        with self.db.session() as s:
            existing = s.scalars(
                select(MappingOverride).where(
                    MappingOverride.template_signature == template_signature,
                    MappingOverride.sheet_name == sheet_name,
                    MappingOverride.cell_address == cell_address,
                )
            ).first()
            if existing:
                existing.match_key = match_key
                existing.employee_id = employee_id
            else:
                s.add(MappingOverride(
                    template_signature=template_signature,
                    sheet_name=sheet_name,
                    cell_address=cell_address,
                    match_key=match_key,
                    employee_id=employee_id,
                ))


class OutputHistoryRepository:
    def __init__(self, db: Database):
        self.db = db

    def record(
        self,
        template_path: str,
        output_path: str,
        sheets_processed: int,
        photos_inserted: int,
        warnings_count: int,
    ) -> None:
        with self.db.session() as s:
            s.add(OutputHistory(
                template_path=template_path,
                output_path=output_path,
                sheets_processed=sheets_processed,
                photos_inserted=photos_inserted,
                warnings_count=warnings_count,
            ))
