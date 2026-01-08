# document/calc_document.py
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET
import re
import xml.dom.minidom as minidom

class CalcDocument:
    NS = {
        "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
        "calcext": "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
    }

    def __init__(self, path: str):
        self.path = path
        self._zip = zipfile.ZipFile(path)
        self.content = self._load_xml("content.xml")
        self.styles = self._load_xml("styles.xml")

    def _load_xml(self, name: str):
        with self._zip.open(name) as f:
            return ET.fromstring(f.read())
        
    def save_debug_xml(self, out_dir: str | Path = "debug_calc_xml"):
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
                    # nevalidní XML → přeskočíme
                    continue

                pretty = minidom.parseString(
                    ET.tostring(root, encoding="utf-8")
                ).toprettyxml(
                    indent="  ",
                    encoding="utf-8"
                )

                target = out_dir / name
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(pretty)
        
    def sheet_names(self) -> list[str]:
        names = []
        for sheet in self.content.findall(".//table:table", self.NS):
            name = sheet.attrib.get(f"{{{self.NS['table']}}}name")
            if name:
                names.append(name)
        return names
    

    def _addr_to_row_col(self, addr: str) -> tuple[int, int]:
        m = re.fullmatch(r"([A-Z]+)(\d+)", addr.upper())
        if not m:
            raise ValueError(f"Neplatná adresa buňky: {addr}")

        col_letters, row = m.groups()
        row = int(row)

        col = 0
        for c in col_letters:
            col = col * 26 + (ord(c) - ord("A") + 1)

        return row, col
    
    def _find_cell(self, sheet_name: str, addr: str) -> dict | None:
        row_target, col_target = self._addr_to_row_col(addr)

        for sheet in self.content.findall(".//table:table", self.NS):
            name = sheet.attrib.get(f"{{{self.NS['table']}}}name")
            if name != sheet_name:
                continue

            row_idx = 0
            for row in sheet.findall("table:table-row", self.NS):
                row_repeat = int(
                    row.attrib.get(
                        "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}number-rows-repeated",
                        "1",
                    )
                )

                for _ in range(row_repeat):
                    row_idx += 1
                    if row_idx != row_target:
                        continue

                    col_idx = 0
                    for cell in row.findall("table:table-cell", self.NS):
                        repeat = int(
                            cell.attrib.get(
                                "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}number-columns-repeated",
                                "1",
                            )
                        )

                        for _ in range(repeat):
                            col_idx += 1
                            if col_idx == col_target:
                                formula = cell.attrib.get(
                                    "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}formula"
                                )

                                value = cell.attrib.get(
                                    "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}value"
                                )

                                return {
                                    "sheet": sheet_name,
                                    "address": addr,
                                    "formula": formula,
                                    "value_cached": value,
                                    "raw_cell": cell,
                                }

                    return None

        return None

    #----------------
        
    def get_array_formula_cells(self) -> list[str]:
        cells = []

        for sheet in self.content.findall(".//table:table", self.NS):
            sheet_name = sheet.attrib.get(
                f"{{{self.NS['table']}}}name", "Sheet"
            )

            row_idx = 0
            for row in sheet.findall("table:table-row", self.NS):
                row_idx += 1
                col_idx = 0

                for cell in row.findall("table:table-cell", self.NS):
                    col_idx += 1

                    rows_span = cell.attrib.get(
                        f"{{{self.NS['table']}}}number-matrix-rows-spanned"
                    )
                    cols_span = cell.attrib.get(
                        f"{{{self.NS['table']}}}number-matrix-columns-spanned"
                    )

                    if rows_span or cols_span:
                        col_letter = chr(ord("A") + col_idx - 1)
                        cells.append(f"{sheet_name}!{col_letter}{row_idx}")

        return cells
    
    # def get_cell_info(self, sheet: str, addr: str):
    #     cell = self._find_cell(sheet, addr)
    #     if cell is None:
    #         return None

    #     raw = cell["raw_cell"]

    #     formula = raw.attrib.get(
    #         "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}formula"
    #     )
    #     value_type = raw.attrib.get(
    #         "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}value-type"
    #     )

    #     return {
    #         "has_formula": formula is not None,
    #         "has_cached_value": value_type is not None,
    #     }

    
    def get_cell_info(self, sheet: str, addr: str):
        cell = self._find_cell(sheet, addr)
        if cell is None:
            return None

        raw = cell["raw_cell"]

        formula = raw.attrib.get(
            "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}formula"
        )

        value = raw.attrib.get(
            "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}value"
        )

        # Calc používá of:= → normalizace
        if formula and formula.startswith("of:="):
            formula = "=" + formula[4:]

        return {
            "exists": True,
            "formula": formula,
            "value_cached": value,
            "is_error": value is not None and isinstance(value, str) and value.startswith("#"),
        }
    
    def iter_formulas(self):
        for sheet in self.content.findall(".//table:table", self.NS):
            sheet_name = sheet.attrib.get(f"{{{self.NS['table']}}}name")

            for cell in sheet.findall(".//table:table-cell", self.NS):
                formula = cell.attrib.get(
                    "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}formula"
                )
                if formula:
                    # ODF má "of:=AVERAGE(...)" → normalizuj
                    formula = formula.replace("of:=", "=")

                    yield {
                        "sheet": sheet_name,
                        "formula": formula,
                    }

    def normalize_formula(self, f: str | None) -> str:
        if not f:
            return ""

        f = f.strip()

        # ODF prefix
        if f.startswith("of:="):
            f = "=" + f[4:]

        # [.E2] → E2
        f = re.sub(r"\[\.(\w+\$?\d+)\]", r"\1", f)

        # [.C2:.C23] → C2:C23
        f = re.sub(r"\[\.(\w+):\.(\w+)\]", r"\1:\2", f)

        # ; → ,
        f = f.replace(";", ",")

        # sjednotit uvozovky a texty
        f = re.sub(r'"([^"]+)"', lambda m: f'"{m.group(1).upper()}"', f)

        # pryč s mezerami
        f = re.sub(r"\s+", "", f)

        return f.upper()

    def get_defined_names(self) -> set[str]:
        names = set()

        for ne in self.content.findall(".//table:named-range", self.NS):
            name = ne.attrib.get(f"{{{self.NS['table']}}}name")
            if name:
                names.add(name.upper())

        return names

    def cells_with_formulas(self):
        out = []

        for sheet in self.content.findall(".//table:table", self.NS):
            sheet_name = sheet.attrib.get(f"{{{self.NS['table']}}}name", "Sheet")

            row_idx = 0
            for row in sheet.findall("table:table-row", self.NS):
                row_idx += 1
                col_idx = 0

                for cell in row.findall("table:table-cell", self.NS):
                    col_idx += 1

                    formula = cell.attrib.get(f"{{{self.NS['table']}}}formula")
                    if not formula:
                        continue

                    # A1 / B3...
                    col_letter = self._col_to_letters(col_idx)
                    addr = f"{col_letter}{row_idx}"

                    out.append({
                        "sheet": sheet_name,
                        "address": addr,
                        "formula": formula,
                    })

        return out

    def _col_to_letters(self, col: int) -> str:
        s = ""
        while col:
            col, r = divmod(col - 1, 26)
            s = chr(65 + r) + s
        return s
    
    def iter_used_rows(self, sheet: str) -> list[int]:
        rows = set()

        for table in self.content.findall(".//table:table", self.NS):
            name = table.attrib.get(f"{{{self.NS['table']}}}name")
            if name != sheet:
                continue

            row_idx = 0
            for row in table.findall("table:table-row", self.NS):
                repeat = int(
                    row.attrib.get(
                        "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}number-rows-repeated",
                        "1"
                    )
                )

                for _ in range(repeat):
                    row_idx += 1
                    for cell in row.findall("table:table-cell", self.NS):
                        if (
                            cell.attrib.get("{urn:oasis:names:tc:opendocument:xmlns:table:1.0}formula")
                            or cell.attrib.get("{urn:oasis:names:tc:opendocument:xmlns:office:1.0}value")
                        ):
                            rows.add(row_idx)
                            break

        return sorted(rows)
    
    def merged_ranges(self, sheet: str):
        ranges = []

        for table in self.content.findall(".//table:table", self.NS):
            name = table.attrib.get(
                "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}name"
            )
            if name != sheet:
                continue

            row_idx = 0
            for row in table.findall("table:table-row", self.NS):
                row_idx += 1
                col_idx = 0

                for cell in row.findall("table:table-cell", self.NS):
                    col_idx += 1

                    rows = int(cell.attrib.get(
                        "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}number-rows-spanned",
                        "1"
                    ))
                    cols = int(cell.attrib.get(
                        "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}number-columns-spanned",
                        "1"
                    ))

                    if rows > 1 or cols > 1:
                        ranges.append((
                            col_idx,
                            row_idx,
                            col_idx + cols - 1,
                            row_idx + rows - 1
                        ))

        return ranges
    
    def has_conditional_formatting(self, sheet: str) -> bool:
        return bool(
            self.content.findall(
                ".//style:map",
                {
                    "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
                },
            )
        )
    
   
    def ods_cf_values(self, sheet: str) -> set[tuple[str, float]]:
        values = set()

        VAL_RE = re.compile(r"^\s*([<>]=?)\s*(-?\d+(?:\.\d+)?)\s*$")

        for cf in self.content.iter():
            if not cf.tag.endswith("conditional-format"):
                continue

            target = None
            for k, v in cf.attrib.items():
                if k.endswith("target-range-address"):
                    target = v
                    break

            if not target or not target.startswith(sheet + "."):
                continue

            for cond in cf:
                if not cond.tag.endswith("condition"):
                    continue

                for k, v in cond.attrib.items():
                    if k.endswith("value"):
                        m = VAL_RE.match(v)
                        if m:
                            op, num = m.groups()
                            values.add((op, float(num)))

        return values
    
    def check_conditional_formatting(self, sheet: str, expected: list[dict]) -> list[str]:
        found = self.ods_cf_values(sheet)  

        missing = []

        for exp in expected:
            want_op = ">" if exp["operator"] == "greaterThan" else "<"
            want_val = float(exp["value"])

            if not any(op == want_op and abs(val - want_val) < 0.01 for op, val in found):
                missing.append(f"{exp['operator']} {want_val} (ODS)")

        return missing

    