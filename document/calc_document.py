from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET
import re
import xml.dom.minidom as minidom

from document.spreadsheet_document import SpreadsheetDocument

class CalcDocument(SpreadsheetDocument):
    NS = {
        "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
        "calcext": "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
        "number": "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",

        "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
        "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
        "fo": "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",

        "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",

        "chart": "urn:oasis:names:tc:opendocument:xmlns:chart:1.0",
        "dr3d": "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0",
    }

    def __init__(self, path: str):
        self.path = path
        self._zip = zipfile.ZipFile(path)
        self.content = self._load_xml("content.xml")
        self.styles = self._load_xml("styles.xml")

        self.number_styles = {}

        def collect_number_styles(root):
            if root is None:
                return

            for ns in root.findall(f".//{{{self.NS["number"]}}}number-style"):
                name = ns.attrib.get(f"{{{self.NS["style"]}}}name")
                if not name:
                    continue

                for child in ns.findall(f"{{{self.NS["number"]}}}number"):
                    dp = child.attrib.get(f"{{{self.NS["number"]}}}decimal-places")
                    if dp is not None:
                        self.number_styles[name] = int(dp)

        collect_number_styles(self.content)
        collect_number_styles(self.styles)

        self.objects = {}

        for name in self._zip.namelist():
            if name.startswith("Object") and name.endswith("content.xml"):
                with self._zip.open(name) as f:
                    self.objects[name] = ET.fromstring(f.read())

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
                        f"{{{self.NS["table"]}}}number-rows-repeated",
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
                                f"{{{self.NS["table"]}}}number-columns-repeated",
                                "1",
                            )
                        )

                        for _ in range(repeat):
                            col_idx += 1
                            if col_idx == col_target:
                                formula = cell.attrib.get(
                                   f"{{{self.NS["table"]}}}formula"
                                )

                                value = cell.attrib.get(
                                   f"{{{self.NS["office"]}}}value"
                                )

                                col_defaults = self._column_default_styles(sheet)
                                col_default = (
                                    col_defaults[col_target - 1]
                                    if col_target - 1 < len(col_defaults)
                                    else None
                                )

                                return {
                                    "sheet": sheet_name,
                                    "address": addr,
                                    "formula": formula,
                                    "value_cached": value,
                                    "raw_cell": cell,
                                    "col_default_style": col_default,
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
      
    def get_cell_info(self, sheet: str, addr: str):
        cell = self._find_cell(sheet, addr)
        if cell is None:
            return None

        raw = cell["raw_cell"]

        formula = raw.attrib.get(
            f"{{{self.NS['table']}}}formula"
        )

        value = raw.attrib.get(
            f"{{{self.NS['office']}}}value"
        )

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
                    f"{{{self.NS['table']}}}formula"
                )
                if formula:
                    formula = formula.replace("of:=", "=")

                    yield {
                        "sheet": sheet_name,
                        "formula": formula,
                    }

    def normalize_formula(self, f: str | None) -> str:
        if not f:
            return ""

        f = f.strip()

        if f.startswith("of:="):
            f = "=" + f[4:]

        # [.E2] -> E2
        f = re.sub(r"\[\.(\w+\$?\d+)\]", r"\1", f)

        # [.C2:.C23] -> C2:C23
        f = re.sub(r"\[\.(\w+):\.(\w+)\]", r"\1:\2", f)

        f = f.replace(";", ",")

        f = re.sub(r'"([^"]+)"', lambda m: f'"{m.group(1).upper()}"', f)

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
        
    def merged_ranges(self, sheet: str):
        ranges = []

        for table in self.content.findall(".//table:table", self.NS):
            name = table.attrib.get(
                f"{{{self.NS['table']}}}name"
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
                        f"{{{self.NS['table']}}}number-rows-spanned",
                        "1"
                    ))
                    cols = int(cell.attrib.get(
                       f"{{{self.NS['table']}}}number-columns-spanned",
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
            self.content.findall(".//style:map", self.NS)
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

    def get_cell_style(self, sheet: str, addr: str) -> dict | None:
        cell = self._find_cell(sheet, addr)
        if not cell:
            return None

        raw = cell["raw_cell"]

        style_name = raw.attrib.get(f"{{{self.NS['table']}}}style-name")

        if not style_name:
            style_name = cell.get("col_default_style")

        style = self._find_style(style_name) if style_name else {}

        return {
            "number_format": style.get("number_format"),
            "decimal_places": style.get("decimal_places"),
            "align_h": style.get("align_h"),
            "bold": style.get("bold", False),
            "wrap": style.get("wrap"),        
        }
        
    def _find_style(self, style_name: str) -> dict:
        style_el = self._find_style_element(style_name)
        if style_el is None:
            return {
                "bold": False,
                "align_h": None,
                "number_format": None,
                "decimal_places": None,
                "wrap": None,
            }

        bold = False
        align_h = None
        wrap = None

        number_format = style_el.attrib.get(f"{{{self.NS["style"]}}}data-style-name")
        decimal_places = self.number_styles.get(number_format)

        for el in style_el.iter():

            if el.tag.endswith("text-properties"):
                if el.attrib.get(f"{{{self.NS["fo"]}}}font-weight") == "bold":
                    bold = True

            if el.tag.endswith("paragraph-properties"):
                align_h = el.attrib.get(f"{{{self.NS["fo"]}}}text-align")

            if el.tag.endswith("table-cell-properties"):
                wo = el.attrib.get(f"{{{self.NS["fo"]}}}wrap-option")
                if wo is not None:
                    wrap = (wo != "no-wrap")

        parent = style_el.attrib.get(f"{{{self.NS["style"]}}}parent-style-name")
        if parent:
            p = self._find_style(parent)
            bold = bold or p["bold"]
            align_h = align_h or p["align_h"]
            number_format = number_format or p["number_format"]
            decimal_places = decimal_places if decimal_places is not None else p["decimal_places"]
            wrap = wrap if wrap is not None else p["wrap"]

        return {
            "bold": bold,
            "align_h": align_h,
            "number_format": number_format,
            "decimal_places": decimal_places,
            "wrap": wrap,
        }
    
    def _find_style_element(self, name: str):
        STYLE_NAME_ATTR = f"{{{self.NS["style"]}}}name"
        for s in self.content.iter():
            if s.attrib.get(STYLE_NAME_ATTR) == name:
                return s

        if self.styles is not None:
            for s in self.styles.iter():
                if s.attrib.get(STYLE_NAME_ATTR) == name:
                    return s

        return None
    
    def _column_default_styles(self, sheet) -> list[str | None]:
        defaults = []

        for col in sheet.findall("table:table-column", self.NS):
            rep = int(col.attrib.get(
                f"{{{self.NS["table"]}}}number-columns-repeated",
                "1"
            ))
            style = col.attrib.get(
                f"{{{self.NS["table"]}}}default-cell-style-name"
            )
            defaults.extend([style] * rep)

        return defaults
        
    def iter_cells(self, sheet: str):
        for table in self.content.findall(".//table:table", self.NS):
            if table.attrib.get(f"{{{self.NS['table']}}}name") != sheet:
                continue

            row_idx = 0
            for row in table.findall("table:table-row", self.NS):
                repeat_row = int(row.attrib.get(
                    f"{{{self.NS["table"]}}}number-rows-repeated",
                    "1"
                ))

                for _ in range(repeat_row):
                    row_idx += 1
                    col_idx = 0

                    for cell in row.findall("table:table-cell", self.NS):
                        repeat_col = int(cell.attrib.get(
                            f"{{{self.NS["table"]}}}number-columns-repeated",
                            "1"
                        ))

                        for _ in range(repeat_col):
                            col_idx += 1

                            text = "".join(cell.itertext()).strip()
                            value = cell.attrib.get(
                                f"{{{self.NS["office"]}}}value"
                            )

                            if not text and value is None:
                                continue

                            yield f"{self._col_to_letters(col_idx)}{row_idx}"

    def get_cell_value(self, sheet: str, addr: str):
        cell = self._find_cell(sheet, addr)
        if cell is None:
            return None

        raw = cell["raw_cell"]

        text = "".join(raw.itertext()).strip()
        if text:
            return text

        return raw.attrib.get(
            f"{{{self.NS["office"]}}}value"
        )
    
    def has_formula(self, sheet: str, addr: str) -> bool:
        cell = self._find_cell(sheet, addr)
        if not cell:
            return False
        return cell["formula"] is not None
    
    def iter_charts(self):
        for ch in self.content.findall(".//chart:chart", self.NS):
            yield ch, self.content

        for obj in self.objects.values():
            for ch in obj.findall(".//chart:chart", self.NS):
                yield ch, obj

    def has_chart(self, sheet: str) -> bool:
        return any(True for _ in self.iter_charts())

    def chart_title(self, sheet: str) -> str | None:
        for chart, _root in self.iter_charts():
            p = chart.find(".//chart:title//text:p", self.NS)
            if p is not None:
                return "".join(p.itertext()).strip()
        return None

    def chart_x_label(self, sheet: str) -> str | None:
        for chart, _root in self.iter_charts():
            p = chart.find(
                ".//chart:axis[@chart:dimension='x']//chart:title//text:p",
                self.NS,
            )
            if p is not None:
                return "".join(p.itertext()).strip()
        return None

    def chart_y_label(self, sheet: str) -> str | None:
        for chart, _root in self.iter_charts():
            p = chart.find(
                ".//chart:axis[@chart:dimension='y']//chart:title//text:p",
                self.NS,
            )
            if p is not None:
                return "".join(p.itertext()).strip()
        return None

    def chart_has_data_labels(self, sheet: str) -> bool:
        for chart, root in self.iter_charts():

            for series in chart.findall(".//chart:series", self.NS):
                style_name = series.attrib.get(
                    f"{{{self.NS['chart']}}}style-name"
                )
                if not style_name:
                    continue

                style = root.find(
                    f".//style:style[@style:name='{style_name}']",
                    self.NS,
                )
                if style is None:
                    continue

                props = style.find("style:chart-properties", self.NS)
                if props is None:
                    continue

                if (
                    props.attrib.get(f"{{{self.NS['chart']}}}data-label-number") == "value"
                    or props.attrib.get(f"{{{self.NS['chart']}}}display-label") == "true"
                ):
                    return True

        return False
    
    def chart_type(self, sheet: str) -> str | None:
        for chart, _root in self.iter_charts():
            chart_class = chart.attrib.get(
                f"{{{self.NS['chart']}}}class"
            )
            if not chart_class:
                continue

            if ":" in chart_class:
                return chart_class.split(":")[1]

            return chart_class

        return None
    
    def has_3d_chart(self, sheet: str) -> bool:
        for chart, _source in self.iter_charts():

            if chart.attrib.get(
                f"{{{self.NS['chart']}}}three-dimensional"
            ) == "true":
                return True

            for el in chart.iter():
                if el.tag.startswith(f"{{{self.NS['dr3d']}}}"):
                    return True

        return False
    

    def get_cell(self, ref: str) -> dict | None:
        if "!" not in ref:
            return None

        sheet, addr = ref.split("!", 1)

        info = self.get_cell_info(sheet, addr)
        if not info:
            return None

        return {
            "exists": True,
            "formula": info.get("formula"),
            "value_cached": info.get("value_cached"),
            "is_error": info.get("is_error", False),
        }
       
    def has_sheet(self, name: str) -> bool:
        return name in self.sheet_names()