from abc import ABC, abstractmethod
from pathlib import Path
from typing import Iterable


class SpreadsheetDocument(ABC):
    @staticmethod
    def from_path(path: str | Path) -> "SpreadsheetDocument":
        path = Path(path)
        suffix = path.suffix.lower()

        if suffix == ".xlsx":
            from document.excel_document import ExcelDocument
            return ExcelDocument(str(path))

        if suffix == ".ods":
            from document.calc_document import CalcDocument
            return CalcDocument(str(path))

        raise ValueError(f"Nepodporovaný tabulkový formát: {path}")

    @abstractmethod
    def save_debug_xml(self, out_dir: str):
        ...

    @abstractmethod
    def has_sheet(self, name: str) -> bool:
        ...

    @abstractmethod
    def sheet_names(self) -> list[str]:
        ...

    @abstractmethod
    def get_cell(self, ref: str) -> dict | None:
        ...

    @abstractmethod
    def get_cell_style(self, sheet: str, addr: str) -> dict | None:
        ...

    @abstractmethod
    def get_cell_value(self, sheet: str, addr: str):
        ...

    @abstractmethod
    def has_formula(self, sheet: str, addr: str) -> bool:
        ...

    @abstractmethod
    def iter_cells(self, sheet: str) -> Iterable[str]:
        ...

    @abstractmethod
    def iter_formulas(self):
        ...

    @abstractmethod
    def normalize_formula(self, f: str | None) -> str:
        ...

    @abstractmethod
    def get_defined_names(self) -> set[str]:
        ...

    @abstractmethod
    def has_chart(self, sheet: str) -> bool:
        ...

    @abstractmethod
    def chart_type(self, sheet: str) -> str | None:
        ...

    @abstractmethod
    def has_3d_chart(self, sheet: str) -> bool:
        ...

    @abstractmethod
    def chart_title(self, sheet: str) -> str | None:
        ...

    @abstractmethod
    def chart_x_label(self, sheet: str) -> str | None:
        ...

    @abstractmethod
    def chart_y_label(self, sheet: str) -> str | None:
        ...

    @abstractmethod
    def chart_has_data_labels(self, sheet: str) -> bool:
        ...