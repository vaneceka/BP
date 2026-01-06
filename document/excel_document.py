import zipfile
from pathlib import Path
import xml.dom.minidom as minidom
import xml.etree.ElementTree as ET
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.chart.bar_chart import BarChart

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


class ExcelDocument:
    def __init__(self, path: str):
        self.path = path
        self.wb = load_workbook(path, data_only=False)
        self.wb_values = load_workbook(path, data_only=True) 

        self.NS = NS
        self._zip = zipfile.ZipFile(path)
        self.workbook_xml = self._load_xml("xl/workbook.xml")
        self.sheets = self._load_sheets_xml()    


    def _load_xml(self, name: str) -> ET.Element:
        with self._zip.open(name) as f:
            return ET.fromstring(f.read())

    def _load_sheets_xml(self) -> dict:
        sheets = {}

        rels = self._load_xml("xl/_rels/workbook.xml.rels")
        rel_map = {
            r.attrib["Id"]: r.attrib["Target"]
            for r in rels.findall("rel:Relationship", self.NS)
        }

        for sheet in self.workbook_xml.findall(".//main:sheet", self.NS):
            name = sheet.attrib["name"]
            r_id = sheet.attrib.get(
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
            )
            target = rel_map.get(r_id)
            if not target:
                continue

            xml = self._load_xml(f"xl/{target}")
            sheets[name] = {"xml": xml, "path": target}

        return sheets

    def get_cell_value_cached(self, sheet: str, addr: str):
        return self.wb_values[sheet][addr].value

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
        cells = []

        for ws in self.wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if cell.data_type == "f":
                        cells.append({
                            "sheet": ws.title,
                            "address": cell.coordinate,
                            "formula": cell.value,
                            "raw_cell": cell,
                        })

        return cells

    def get_cell(self, address: str, *, include_value=False):
        if "!" not in address:
            raise ValueError("Cell address must be in format 'sheet!A1'")

        sheet, addr = address.split("!", 1)

        if sheet not in self.wb.sheetnames:
            return None

        ws = self.wb[sheet]
        cell = ws[addr]

        if cell.value is None and cell.data_type != "f":
            return None

        data = {
            "sheet": sheet,
            "address": addr,
            "raw_cell": cell,
            "formula": cell.value if cell.data_type == "f" else None,
            "value_cached": self.wb_values[sheet][addr].value,
        }

        if include_value:
            data["value"] = cell.value

        return data

    def has_chart(self) -> bool:
        return any(ws._charts for ws in self.wb.worksheets)
    
    def is_outer_cell(row, col, min_row, max_row, min_col, max_col):
        return (
            row == min_row or row == max_row or
            col == min_col or col == max_col
        )

    def defined_names(self) -> set[str]:
        return {name.upper() for name in self.wb.defined_names.keys()}

    def save_xml(self, out_dir: str | Path = "debug_excel_xml"):
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
                    continue

                pretty = minidom.parseString(
                    ET.tostring(root, encoding="utf-8")
                ).toprettyxml(indent="  ", encoding="utf-8")

                target = out_dir / name
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(pretty)