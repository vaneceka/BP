# assignment/excel_assignment_model.py
from dataclasses import dataclass
from typing import Any


@dataclass
class ExcelCellSpec:
    address: str
    input: Any | None = None
    expression: str | None = None
    style: dict | None = None
    conditional_format: list[dict] | None = None


@dataclass
class ExcelAssignment:
    cells: dict[str, ExcelCellSpec]
    borders: list[dict]
    chart: dict | None