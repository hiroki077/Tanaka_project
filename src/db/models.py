from __future__ import annotations
from datetime import datetime
from enum import StrEnum
from sqlalchemy import String, Integer, DateTime, ForeignKey, UniqueConstraint, Text
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, relationship


class Base(DeclarativeBase):
    pass


class EmploymentStatus(StrEnum):
    ACTIVE = "在職"
    ON_LEAVE = "休職中"
    RESIGNED = "退職"


class Employee(Base):
    """人物カタログ。

    所属（本部・支店・課）は **テンプレートExcel側で管理する** ため、DBには持たない。
    異動が発生した際の二重管理を避けるための設計判断。
    """

    __tablename__ = "employees"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    name: Mapped[str] = mapped_column(String(64), nullable=False)
    name_kana: Mapped[str | None] = mapped_column(String(64))
    # 照合キー: マスター体制表のセル値と照合するための名前。
    # カンマ区切りで複数指定可能（例: 「大野,大野翔,大野翔一」）。
    # 最初のキーが主、それ以降は引用名・別名として扱われる。
    match_key: Mapped[str] = mapped_column(String(256), nullable=False, index=True)
    # 旧 reference_name は match_key に統合済み（互換のため残す）
    reference_name: Mapped[str | None] = mapped_column(String(128))
    join_year: Mapped[int | None] = mapped_column(Integer)
    # 入社年の元表記（"M2017", "2016C" 等の記号付きも保持）
    join_year_text: Mapped[str | None] = mapped_column(String(16))
    role: Mapped[str | None] = mapped_column(String(32))
    role_marks: Mapped[str | None] = mapped_column(String(32))
    employment_type: Mapped[str | None] = mapped_column(String(16))
    status: Mapped[str] = mapped_column(
        String(8), nullable=False, default=EmploymentStatus.ACTIVE.value
    )
    photo_path: Mapped[str | None] = mapped_column(Text)
    display_order: Mapped[int] = mapped_column(Integer, default=0)
    note: Mapped[str | None] = mapped_column(Text)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.now)
    updated_at: Mapped[datetime] = mapped_column(
        DateTime, default=datetime.now, onupdate=datetime.now
    )

    overrides: Mapped[list["MappingOverride"]] = relationship(
        back_populates="employee", cascade="all, delete-orphan"
    )

    def __repr__(self) -> str:
        return f"<Employee id={self.id} name={self.name} status={self.status}>"


class MappingOverride(Base):
    __tablename__ = "mapping_overrides"
    __table_args__ = (
        UniqueConstraint(
            "template_signature", "sheet_name", "cell_address",
            name="uq_mapping_template_sheet_cell"
        ),
    )

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    template_signature: Mapped[str] = mapped_column(String(128), nullable=False)
    sheet_name: Mapped[str] = mapped_column(String(64), nullable=False)
    cell_address: Mapped[str] = mapped_column(String(16), nullable=False)
    match_key: Mapped[str] = mapped_column(String(64), nullable=False)
    employee_id: Mapped[int] = mapped_column(
        ForeignKey("employees.id", ondelete="CASCADE"), nullable=False
    )
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.now)

    employee: Mapped[Employee] = relationship(back_populates="overrides")


class OutputHistory(Base):
    __tablename__ = "output_history"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    template_path: Mapped[str] = mapped_column(Text, nullable=False)
    output_path: Mapped[str] = mapped_column(Text, nullable=False)
    sheets_processed: Mapped[int] = mapped_column(Integer, default=0)
    photos_inserted: Mapped[int] = mapped_column(Integer, default=0)
    warnings_count: Mapped[int] = mapped_column(Integer, default=0)
    executed_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.now)
