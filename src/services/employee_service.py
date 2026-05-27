from __future__ import annotations
from typing import Iterable

from ..db import Database, Employee, EmployeeRepository, EmploymentStatus
from .photo_service import PhotoService


class EmployeeService:
    def __init__(self, db: Database, photo_service: PhotoService):
        self.repo = EmployeeRepository(db)
        self.photos = photo_service

    def list(
        self,
        keyword: str | None = None,
        status_filter: str | None = None,
    ) -> list[Employee]:
        return self.repo.list_all(keyword=keyword, status_filter=status_filter)

    def get(self, employee_id: int) -> Employee | None:
        return self.repo.get(employee_id)

    def create(self, **fields) -> Employee:
        emp = Employee(**fields)
        if emp.status is None:
            emp.status = EmploymentStatus.ACTIVE.value
        return self.repo.upsert(emp)

    def update(self, employee_id: int, **fields) -> Employee | None:
        emp = self.repo.get(employee_id)
        if emp is None:
            return None
        for k, v in fields.items():
            if hasattr(emp, k):
                setattr(emp, k, v)
        return self.repo.upsert(emp)

    def delete(self, employee_id: int) -> bool:
        emp = self.repo.get(employee_id)
        if emp is None:
            return False
        self.photos.delete(emp.photo_path)
        return self.repo.delete(employee_id)

    def set_photo(self, employee_id: int, source_path: str) -> Employee | None:
        emp = self.repo.get(employee_id)
        if emp is None:
            return None
        if emp.photo_path:
            self.photos.delete(emp.photo_path)
        emp.photo_path = self.photos.import_photo(source_path)
        return self.repo.upsert(emp)

    def rotate_photo(self, employee_id: int, degrees: int) -> Employee | None:
        """従業員の登録写真を時計回りに degrees だけ回転する。"""
        emp = self.repo.get(employee_id)
        if emp is None or not emp.photo_path:
            return None
        new_path = self.photos.rotate_photo(emp.photo_path, degrees)
        if new_path and new_path != emp.photo_path:
            emp.photo_path = new_path
            return self.repo.upsert(emp)
        return emp
