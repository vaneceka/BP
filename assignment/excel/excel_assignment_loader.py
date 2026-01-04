# assignment/excel_assignment_loader.py
import json
from .excel_assignment_model import ExcelAssignment, ExcelCellSpec


def load_excel_assignment(path: str) -> ExcelAssignment:
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    cells: dict[str, ExcelCellSpec] = {}

    for addr, spec in data.get("cells", {}).items():
        cells[addr] = ExcelCellSpec(
            address=addr,
            input=spec.get("input"),
            expression=spec.get("expression"),
            style=spec.get("style"),
            conditionalFormat=spec.get("conditionalFormat"),
        )

    return ExcelAssignment(
        cells=cells,
        borders=data.get("borders", []),
        chart=data.get("chart"),
    )