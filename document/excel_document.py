import zipfile
from pathlib import Path
import xml.dom.minidom as minidom
import xml.etree.ElementTree as ET
from openpyxl import load_workbook


class ExcelDocument:
    def __init__(self, path: str):
        self.path = path
        self.wb = load_workbook(path, data_only=False)

    # ---------------------------
    # ZÁKLADNÍ API
    # ---------------------------

    def sheet_names(self) -> list[str]:
        return self.wb.sheetnames

    def has_formula(self) -> bool:
        return any(
            cell.data_type == "f"
            for ws in self.wb.worksheets
            for row in ws.iter_rows()
            for cell in row
        )

    def cells_with_formulas(self):
        return [
            {
                "sheet": ws.title,
                "address": cell.coordinate,
                "formula": cell.value,
            }
            for ws in self.wb.worksheets
            for row in ws.iter_rows()
            for cell in row
            if cell.data_type == "f"
        ]

    def has_chart(self) -> bool:
        return any(ws._charts for ws in self.wb.worksheets)

    # ---------------------------
    # DEBUG XML (DŮLEŽITÉ)
    # ---------------------------

    def save_xml(self, out_dir: str | Path = "debug_excel_xml"):
        """
        Rozbalí XLSX a uloží všechna XML v čitelné podobě.
        """
        out_dir = Path(out_dir)
        out_dir.mkdir(parents=True, exist_ok=True)

        with zipfile.ZipFile(self.path) as z:
            for name in z.namelist():
                if not name.endswith(".xml"):
                    continue

                try:
                    raw = z.read(name)
                    root = ET.fromstring(raw)
                except Exception:
                    # není validní XML (rare, ale může se stát)
                    continue

                pretty = minidom.parseString(
                    ET.tostring(root, encoding="utf-8")
                ).toprettyxml(indent="  ", encoding="utf-8")

                target = out_dir / name
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(pretty)