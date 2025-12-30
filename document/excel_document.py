import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import xml.dom.minidom as minidom


NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "drawing": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "chart": "http://schemas.openxmlformats.org/drawingml/2006/chart",
}


class ExcelDocument:
    def __init__(self, path: str):
        self.path = path
        self.NS = NS
        self._zip = zipfile.ZipFile(path)

        # základní XML
        self.workbook_xml = self._load("xl/workbook.xml")
        self.styles_xml = self._load("xl/styles.xml")
        self.shared_strings = self._load_shared_strings()

        # sheets
        self.sheets = self._load_sheets()

    def _load(self, name: str) -> ET.Element:
        with self._zip.open(name) as f:
            return ET.fromstring(f.read())
        
    def save_xml(self, out_dir: str | Path = "debug_excel_xml"):
        out_dir = Path(out_dir)
        out_dir.mkdir(parents=True, exist_ok=True)

        self._save_pretty(self.workbook_xml, out_dir / "workbook.xml")
        self._save_pretty(self.styles_xml, out_dir / "styles.xml")

        for name, sheet in self.sheets.items():
            self._save_pretty(sheet["xml"], out_dir / f"{name}.xml")

    def _save_pretty(self, root: ET.Element, path: Path):
        rough = ET.tostring(root, encoding="utf-8")
        pretty = minidom.parseString(rough).toprettyxml(
            indent="  ", encoding="utf-8"
        )
        path.write_bytes(pretty)

    def _load_sheets(self):
        sheets = {}

        # rels k workbooku
        rels = self._load("xl/_rels/workbook.xml.rels")

        rel_map = {
            r.attrib["Id"]: r.attrib["Target"]
            for r in rels.findall("rel:Relationship", self.NS)
        }

        for sheet in self.workbook_xml.findall(".//main:sheet", self.NS):
            name = sheet.attrib["name"]
            r_id = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            target = rel_map.get(r_id)

            if not target:
                continue

            xml = self._load(f"xl/{target}")
            sheets[name] = {
                "xml": xml,
                "path": target
            }

        return sheets
    
    def _load_shared_strings(self):
        try:
            xml = self._load("xl/sharedStrings.xml")
        except KeyError:
            return []

        strings = []
        for si in xml.findall(".//main:si", self.NS):
            texts = [t.text for t in si.findall(".//main:t", self.NS) if t.text]
            strings.append("".join(texts))
        return strings
    
    # ----------------------------
    def iter_cells(self):
        for sheet_name, data in self.sheets.items():
            sheet = data["xml"]

            for c in sheet.findall(".//main:c", self.NS):
                addr = c.attrib.get("r")
                formula = c.find("main:f", self.NS)
                value = c.find("main:v", self.NS)

                yield {
                    "sheet": sheet_name,
                    "address": addr,
                    "formula": formula.text if formula is not None else None,
                    "value": self._resolve_value(c, value),
                }

    def _resolve_value(self, c, v):
        if v is None:
            return None

        if c.attrib.get("t") == "s":  # shared string
            return self.shared_strings[int(v.text)]

        return v.text
    
    def sheet_names(self) -> list[str]:
        return list(self.sheets.keys())

    def cells_with_formulas(self):
        return [c for c in self.iter_cells() if c["formula"]]

    def has_formula(self) -> bool:
        return any(c["formula"] for c in self.iter_cells())

    def has_chart(self) -> bool:
        return any(
            name.startswith("xl/charts/")
            for name in self._zip.namelist()
        )