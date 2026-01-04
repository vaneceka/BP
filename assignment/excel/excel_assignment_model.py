from dataclasses import dataclass
from typing import Any


@dataclass
class ExcelCellSpec:
    address: str
    input: Any | None = None
    expression: str | None = None
    style: dict | None = None
    conditionalFormat: list[dict] | None = None


@dataclass
class ExcelAssignment:
    cells: dict[str, ExcelCellSpec]
    borders: list[dict]
    chart: dict | None