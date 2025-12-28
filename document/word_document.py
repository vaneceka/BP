from typing import Iterable
import zipfile
import xml.etree.ElementTree as ET
import re

from assignment.assignment_model import StyleSpec
from pathlib import Path
import xml.dom.minidom as minidom

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math"
    }

BUILTIN_STYLE_NAMES = {
    "normal",
    "heading 1",
    "heading 2",
    "heading 3",
    "heading 4",
    "caption",
    "bibliography",
    "toc heading",
    "table of contents",
    "content heading",
}


class WordDocument:
    def __init__(self, path: str):
        self.path = path
        self.NS = NS  
        self._zip = zipfile.ZipFile(path)
        self._xml = self._load("word/document.xml")
        self._styles_xml = self._load("word/styles.xml")
        self._sections = self._split_into_sections()

    def save_xml(self, out_dir: str | Path = "debug_xml"):
        """
        Uloží načtené XML soubory (document.xml, styles.xml) do složky.
        Slouží pouze pro debug / analýzu.
        """
        out_dir = Path(out_dir)
        out_dir.mkdir(parents=True, exist_ok=True)

        self._save_pretty_xml(self._xml, out_dir / "document.xml")
        self._save_pretty_xml(self._styles_xml, out_dir / "styles.xml")

    def _save_pretty_xml(self, root: ET.Element, path: Path):
        rough_string = ET.tostring(root, encoding="utf-8")
        reparsed = minidom.parseString(rough_string)
        pretty = reparsed.toprettyxml(indent="  ", encoding="utf-8")

        with open(path, "wb") as f:
            f.write(pretty)

    def _load(self, name):
        with self._zip.open(name) as f:
            return ET.fromstring(f.read())
        
    def iter_paragraphs(self) -> Iterable[ET.Element]:
        return self._xml.findall(".//w:p", self.NS)

    def iter_runs(self) -> Iterable[ET.Element]:
        return self._xml.findall(".//w:r", self.NS)

    def iter_texts(self) -> Iterable[str]:
        for t in self._xml.findall(".//w:t", self.NS):
            if t.text:
                yield t.text
    def is_builtin_style_name(self, name: str) -> bool:
        return name.strip().lower() in BUILTIN_STYLE_NAMES
    
    def split_assignment_styles(self, assignment):
        """
        Rozdělí styly ze zadání na:
        - custom (vlastní)
        - builtin (vestavěné Word styly)
        """
        custom = {}
        builtin = {}

        for name, spec in assignment.styles.items():
            if self.is_builtin_style_name(name):
                builtin[name] = spec
            else:
                custom[name] = spec

        return custom, builtin

    # ==========================================================
    # ODDÍLY
    # ==========================================================

    def _split_into_sections(self):
        """
        Rozdělí dokument na oddíly podle <w:sectPr>.
        Podporuje:
        - <w:sectPr> jako dítě <w:body>
        - <w:sectPr> uvnitř <w:pPr>
        """
        sections = []
        current = []

        body = self._xml.find("w:body", NS)

        for el in body:
            current.append(el)

            # 1) sectPr jako přímé dítě body
            if el.tag.endswith("}sectPr"):
                sections.append(current)
                current = []
                continue

            # 2) sectPr uvnitř odstavce
            if el.tag.endswith("}p"):
                ppr = el.find("w:pPr", NS)
                if ppr is not None and ppr.find("w:sectPr", NS) is not None:
                    sections.append(current)
                    current = []

        if current:
            sections.append(current)

        return sections

    def section_count(self) -> int:
        return len(self._sections)

    def section(self, index: int):
        if index < 0 or index >= len(self._sections):
            return []
        return self._sections[index]

    # ==========================================================
    # OBECNÉ POMOCNÉ METODY
    # ==========================================================

    def _find_instr_texts(self, section):
        texts = []
        for el in section:
            for instr in el.findall(".//w:instrText", NS):
                if instr.text:
                    texts.append(instr.text.strip())
        return texts

    # ==========================================================
    # TOC / SEZNAMY
    # ==========================================================

    def has_toc_in_section(self, section_index: int) -> bool:
        for instr in self._find_instr_texts(self.section(section_index)):
            if instr.startswith("TOC") and "\\o" in instr:
                return True
        return False

    def has_object_list_in_section(self, section_index: int) -> bool:
        for instr in self._find_instr_texts(self.section(section_index)):
            if instr.startswith("TOC") and "\\t" in instr:
                return True
        return False

    # ==========================================================
    # TEXT
    # ==========================================================

    def has_text_in_section(self, section_index: int) -> bool:
        for el in self.section(section_index):
            for t in el.findall(".//w:t", NS):
                if t.text and t.text.strip():
                    return True
        return False

    # ==========================================================
    # BIBLIOGRAFIE (SPRÁVNĚ PŘES SDT)
    # ==========================================================

    def has_bibliography_in_section(self, section_index: int) -> bool:
        for el in self.section(section_index):
            for sdt in el.findall(".//w:sdt", NS):
                sdt_pr = sdt.find("w:sdtPr", NS)
                if sdt_pr is not None and sdt_pr.find("w:bibliography", NS) is not None:
                    return True
        return False
    
    def debug_sections(self):
        for i in range(self.section_count()):
            texts = []
            for el in self.section(i):
                for t in el.findall(".//w:t", NS):
                    if t.text and t.text.strip():
                        texts.append(t.text.strip())
            print(f"SECTION {i+1}: {len(texts)} text nodes")
            print("  sample:", texts[:10])



        import re

    def get_field_instructions(self, section):
        instrs = []

        # 1) fldSimple (nejdůležitější!)
        for el in section:
            for fld in el.findall(".//w:fldSimple", self.NS):
                instr = fld.attrib.get(f"{{{self.NS['w']}}}instr")
                if instr:
                    instrs.append(instr.strip())

        # 2) instrText (méně časté, ale ponecháme)
        for el in section:
            for instr in el.findall(".//w:instrText", self.NS):
                if instr.text:
                    instrs.append(instr.text.strip())

        return instrs


    def _extract_toc_labels(self, instr: str) -> list[str]:
        labels = []

        if "\\t" not in instr:
            return labels

        # rozděl podle \t
        parts = instr.split("\\t")

        for part in parts[1:]:
            part = part.strip()

            # očekáváme něco jako: "Obrázek,1"
            if part.startswith('"') and '"' in part[1:]:
                value = part.split('"', 2)[1]  # obsah mezi uvozovkami
                label = value.split(",")[0].strip()
                if label:
                    labels.append(label)

        return labels
    

    def has_list_of_figures_in_section(self, section_index: int) -> bool:
        section = self.section(section_index)

        for instr in self.get_field_instructions(section):
            instr_u = instr.upper()

            if not instr_u.startswith("TOC"):
                continue

            # Word může použít \c "Obrázek" nebo \t "Figure,1"
            if "\\C" in instr_u or "\\T" in instr_u:
                if "OBRÁZEK" in instr_u or "FIGURE" in instr_u:
                    return True

        return False


    def has_list_of_tables_in_section(self, section_index: int) -> bool:
        section = self.section(section_index)

        for instr in self.get_field_instructions(section):
            instr_u = instr.upper()

            if not instr_u.startswith("TOC"):
                continue

            if "\\C" in instr_u or "\\T" in instr_u:
                if "TABULKA" in instr_u or "TABLE" in instr_u:
                    return True

        return False
    
    def _normalize_font(self, font: str | None) -> str | None:
        if not font:
            return None
        # odstraní cokoliv od "(" až do konce
        font = re.sub(r"\s*\(.*$", "", font)
        return font.strip()
    
    def _resolve_font(self, style):
        while style is not None:
            rpr = style.find("w:rPr", self.NS)
            if rpr is not None:
                fonts = rpr.find("w:rFonts", self.NS)
                if fonts is not None:
                    font = (
                        fonts.attrib.get(f"{{{self.NS['w']}}}ascii")
                        or fonts.attrib.get(f"{{{self.NS['w']}}}hAnsi")
                        or fonts.attrib.get(f"{{{self.NS['w']}}}cs")
                    )
                    if font:
                        return self._normalize_font(font)

            based = style.find("w:basedOn", self.NS)
            style = (
                self._find_style_by_id(based.attrib.get(f"{{{self.NS['w']}}}val"))
                if based is not None
                else None
            )

        return None
    

    def _resolve_color(self, style) -> str | None:
        while style is not None:
            rpr = style.find("w:rPr", self.NS)
            if rpr is not None:
                col = rpr.find("w:color", self.NS)
                if col is not None:
                    val = col.attrib.get(f"{{{self.NS['w']}}}val")
                    if val:
                        return val.upper()

            # jdi na parent styl
            based = style.find("w:basedOn", self.NS)
            if based is None:
                break

            style = self._find_style_by_id(
                based.attrib.get(f"{{{self.NS['w']}}}val")
            )

        return None
    
    def _find_style_by_id(self, style_id: str):
        if not style_id:
            return None

        for style in self._styles_xml.findall(".//w:style", self.NS):
            if style.attrib.get(f"{{{self.NS['w']}}}styleId") == style_id:
                return style

        return None
    
    def _resolve_size(self, style) -> int | None:
        while style is not None:
            rpr = style.find("w:rPr", self.NS)
            if rpr is not None:
                sz = rpr.find("w:sz", self.NS)
                if sz is not None:
                    return int(sz.attrib[f"{{{self.NS['w']}}}val"]) // 2
                
            based = style.find("w:basedOn", self.NS)
            if based is None:
                break

            style = self._find_style_by_id(based.attrib.get(f"{{{self.NS['w']}}}val"))

        return self.get_doc_default_font_size()
    
    def _resolve_underline(self, style) -> bool | None:
        visited = set()

        while style is not None and id(style) not in visited:
            visited.add(id(style))

            rpr = style.find("w:rPr", self.NS)
            if rpr is not None:
                u = rpr.find("w:u", self.NS)
                if u is not None:
                    val = u.attrib.get(f"{{{self.NS['w']}}}val", "single")
                    return val != "none"

            # linked character style
            linked = self._get_linked_style(style)
            if linked is not None and id(linked) not in visited:
                style = linked
                continue

            # basedOn
            based = style.find("w:basedOn", self.NS)
            if based is not None:
                style = self._find_style_by_id(
                    based.attrib.get(f"{{{self.NS['w']}}}val")
                )
                continue

            break

        return None
    
    def _resolve_page_break_before(self, style) -> bool:
        visited = set()
        while style is not None and id(style) not in visited:
            visited.add(id(style))

            ppr = style.find("w:pPr", self.NS)
            if ppr is not None:
                if ppr.find("w:pageBreakBefore", self.NS) is not None:
                    return True

            # linked / basedOn
            linked = self._get_linked_style(style)
            if linked is not None and id(linked) not in visited:
                style = linked
                continue

            based = style.find("w:basedOn", self.NS)
            style = self._find_style_by_id(based.attrib.get(f"{{{self.NS['w']}}}val")) if based is not None else None

        return False
    
    def get_doc_default_font_size(self) -> int | None:
        dd = self._styles_xml.find(".//w:docDefaults/w:rPrDefault/w:rPr", self.NS)
        if dd is None:
            return None
        sz = dd.find("w:sz", self.NS)
        if sz is None:
            return None
        return int(sz.attrib[f"{{{self.NS['w']}}}val"]) // 2
    
    def _resolve_bool(self, style, tag: str) -> bool | None:
        visited = set()

        while style is not None and id(style) not in visited:
            visited.add(id(style))

            rpr = style.find("w:rPr", self.NS)
            if rpr is not None:
                el = rpr.find(f"w:{tag}", self.NS)
                if el is not None:
                    if el.attrib.get(f"{{{self.NS['w']}}}val") == "0":
                        return False
                    return True

            # 1️⃣ linked character style
            linked = self._get_linked_style(style)
            if linked is not None and id(linked) not in visited:
                style = linked
                continue

            # 2️⃣ basedOn
            based = style.find("w:basedOn", self.NS)
            if based is not None:
                style = self._find_style_by_id(
                    based.attrib.get(f"{{{self.NS['w']}}}val")
                )
                continue

            break

        return None

    def _get_linked_style(self, style):
        link = style.find("w:link", self.NS)
        if link is None:
            return None
        return self._find_style_by_id(
            link.attrib.get(f"{{{self.NS['w']}}}val")
        )

    def _build_style_spec(self, style, *, default_alignment=None) -> StyleSpec:
        font = self._resolve_font(style)
        color = self._resolve_color(style)
        size = self._resolve_size(style)
        underline = self._resolve_underline(style)
        bold = self._resolve_bool(style, "b")
        italic = self._resolve_bool(style, "i")
        
        all_caps = self._resolve_bool(style, "caps")
        alignment = default_alignment
        page_break = self._resolve_page_break_before(style)
        num_level = None
        is_numbered = None
        line_height = None
        based_on = None

        indent_left = None
        indent_right = None
        indent_first = None
        indent_hanging = None


        ppr = style.find("w:pPr", self.NS)
        if ppr is not None:
            jc = ppr.find("w:jc", self.NS)
            if jc is not None:
                alignment = jc.attrib.get(f"{{{self.NS['w']}}}val") or alignment

            spacing = ppr.find("w:spacing", self.NS)
            if spacing is not None:
                line = spacing.attrib.get(f"{{{self.NS['w']}}}line")
                if line:
                    line_height = int(line) / 240

            # ppr = style.find("w:pPr", self.NS)
            if ppr is not None:
                numpr = ppr.find("w:numPr", NS)
                if numpr is not None:
                    num_id = numpr.find("w:numId", self.NS)
                    if num_id is not None:
                        val = int(num_id.attrib.get(f"{{{self.NS['w']}}}val", 0))
                        if val > 0:
                            is_numbered = True
                            ilvl = numpr.find("w:ilvl", self.NS)
                            num_level = int(ilvl.attrib.get(f"{{{self.NS['w']}}}val")) if ilvl is not None else 0
            
            ind = ppr.find("w:ind", self.NS)
            if ind is not None:
                indent_left = ind.attrib.get(f"{{{self.NS['w']}}}left")
                indent_right = ind.attrib.get(f"{{{self.NS['w']}}}right")
                indent_first = ind.attrib.get(f"{{{self.NS['w']}}}firstLine")
                indent_hanging = ind.attrib.get(f"{{{self.NS['w']}}}hanging")

        based_el = style.find("w:basedOn", self.NS)
        if based_el is not None:
            based_on = based_el.attrib.get(f"{{{self.NS['w']}}}val")

        name_el = style.find("w:name", self.NS)
        name = name_el.attrib.get(f"{{{self.NS['w']}}}val") if name_el is not None else None

        return StyleSpec(
            name=name,
            font=font,
            size=size,
            bold=bold,
            italic=italic,
            underline=underline,
            allCaps=all_caps,
            color=color,
            alignment=alignment,
            lineHeight=line_height,
            pageBreakBefore=page_break,
            isNumbered=is_numbered,
            numLevel=num_level,
            basedOn=based_on,
            indentLeft=int(indent_left) if indent_left else None,
            indentRight=int(indent_right) if indent_right else None,
            indentFirstLine=int(indent_first) if indent_first else None,
            indentHanging=int(indent_hanging) if indent_hanging else None,
        )
    
    def _find_style(self, *, name: str | None = None, default: bool = False):
        for style in self._styles_xml.findall(".//w:style", self.NS):
            if default:
                if style.attrib.get(f"{{{self.NS['w']}}}default") == "1":
                    return style
                continue

            if not name:
                continue

            # 1) match podle styleId
            style_id = style.attrib.get(f"{{{self.NS['w']}}}styleId")
            if style_id and style_id.strip().lower() == name.strip().lower():
                return style

            # 2) match podle zobrazovaného názvu w:name
            name_el = style.find("w:name", self.NS)
            if name_el is None:
                continue
            actual_name = name_el.attrib.get(f"{{{self.NS['w']}}}val")
            if actual_name and actual_name.strip().lower() == name.strip().lower():
                return style

        return None
    
    def get_normal_style(self) -> StyleSpec | None:
        style = self._find_style(default=True)
        if style is None:
            return None
        return self._build_style_spec(style, default_alignment="both")

    def get_heading_style(self, level: int) -> StyleSpec | None:
        style = self._find_style(name=f"heading {level}")
        if style is None:
            return None
        return self._build_style_spec(style, default_alignment="start")

    def get_custom_style(self, style_name: str) -> StyleSpec | None:
        style = self._find_style(name=style_name)
        if style is None:
            return None
        return self._build_style_spec(style)
    

    def get_style_by_any_name(self, names: list[str], *, default_alignment: str | None = None) -> StyleSpec | None:
        for n in names:
            style = self._find_style(name=n)
            if style is not None:
                return self._build_style_spec(style, default_alignment=default_alignment)
        return None
    
    # Test pro spravne pouziti nadpisu
    def _paragraph_text(self, p: ET.Element) -> str:
        parts = []
        for t in p.findall(".//w:t", self.NS):
            if t.text:
                parts.append(t.text)
        # Word občas láme text do více runů → spojíme
        txt = "".join(parts)
        # normalizace mezer
        txt = re.sub(r"\s+", " ", txt).strip()
        return txt

    def _paragraph_style_id(self, p: ET.Element) -> str | None:
        ppr = p.find("w:pPr", self.NS)
        if ppr is None:
            return None
        ps = ppr.find("w:pStyle", self.NS)
        if ps is None:
            return None
        return ps.attrib.get(f"{{{self.NS['w']}}}val")

    # def _style_level_from_styles_xml(self, style_id: str) -> int | None:
    #     """
    #     Zkusí odvodit level nadpisu ze styles.xml:
    #     - podle w:name (Heading 1 / heading 1 / Nadpis 1 / Nadpis1…)
    #     - nebo podle w:outlineLvl (0 -> H1, 1 -> H2, 2 -> H3…)
    #     """
    #     style = self._find_style_by_id(style_id)
    #     if style is None:
    #         return None

    #     # 1) podle názvu stylu
    #     name_el = style.find("w:name", self.NS)
    #     if name_el is not None:
    #         nm = (name_el.attrib.get(f"{{{self.NS['w']}}}val") or "").strip().lower()
    #         # typicky: "heading 2", "nadpis 3", "nadpis3"
    #         m = re.search(r"(heading|nadpis)\s*([0-9]+)", nm)
    #         if m:
    #             return int(m.group(2))

    #     # 2) podle outlineLvl (0=H1, 1=H2, 2=H3...)
    #     ppr = style.find("w:pPr", self.NS)
    #     if ppr is not None:
    #         out = ppr.find("w:outlineLvl", self.NS)
    #         if out is not None:
    #             v = out.attrib.get(f"{{{self.NS['w']}}}val")
    #             if v is not None:
    #                 return int(v) + 1

    #     return None

    def _style_level_from_styles_xml(self, style_id: str) -> int | None:
        style = self._find_style_by_id(style_id)
        if style is None:
            return None

        visited = set()

        while style is not None and id(style) not in visited:
            visited.add(id(style))

            # 1️⃣ zkus název stylu (nejspolehlivější)
            name_el = style.find("w:name", self.NS)
            if name_el is not None:
                name = (name_el.attrib.get(f"{{{self.NS['w']}}}val") or "").lower()
                m = re.search(r"(heading|nadpis)\s*([0-9]+)", name)
                if m:
                    return int(m.group(2))

            # 2️⃣ outlineLvl
            ppr = style.find("w:pPr", self.NS)
            if ppr is not None:
                out = ppr.find("w:outlineLvl", self.NS)
                if out is not None:
                    return int(out.attrib.get(f"{{{self.NS['w']}}}val")) + 1

            # 3️⃣ basedOn → jdi výš
            based = style.find("w:basedOn", self.NS)
            if based is None:
                break

            style = self._find_style_by_id(
                based.attrib.get(f"{{{self.NS['w']}}}val")
            )

        return None

    def iter_headings(self) -> list[tuple[str, int]]:
        """
        Vrátí seznam (text, level) pro odstavce, které vypadají jako nadpisy.
        """
        items: list[tuple[str, int]] = []

        for p in self._xml.findall(".//w:body/w:p", self.NS):
            txt = self._paragraph_text(p)
            if not txt:
                continue

            sid = self._paragraph_style_id(p)
            if not sid:
                continue

            lvl = self._style_level_from_styles_xml(sid)
            if lvl is None:
                continue

            items.append((txt, lvl))

        return items
    
    # kontrola původního formátování
    def has_manual_formatting(self) -> bool:
        for p in self._xml.findall(".//w:p", self.NS):
            ppr = p.find("w:pPr", self.NS)
            style = ppr.find("w:pStyle", self.NS) if ppr is not None else None

            # žádný styl → podezření
            if style is None:
                for r in p.findall(".//w:r", self.NS):
                    rpr = r.find("w:rPr", self.NS)
                    if rpr is not None:
                        return True
        return False
    
    def has_inline_font_changes(self) -> bool:
        for r in self._xml.findall(".//w:r", self.NS):
            rpr = r.find("w:rPr", self.NS)
            if rpr is not None:
                if (
                    rpr.find("w:sz", self.NS) is not None or
                    rpr.find("w:rFonts", self.NS) is not None or
                    rpr.find("w:color", self.NS) is not None
                ):
                    return True
        return False
    
    def has_html_artifacts(self) -> bool:
        for t in self.iter_texts():
            if any(x in t for x in ["&nbsp;", "&#160;", "<", ">"]):
                return True
        return False
    
    # číslování položek v obsahu
    def toc_shows_numbers(self) -> bool | None:
        """
        True  = obsah zobrazuje čísla
        False = obsah čísla NEzobrazuje
        None  = žádný obsah
        """
        for instr in self._xml.findall(".//w:instrText", self.NS):
            txt = instr.text or ""
            if txt.strip().startswith("TOC"):
                # \n = suppress numbering
                return "\\n" not in txt
        return None
    
    # číslovaní kapitol, které nemají
    def _paragraph_is_numbered(self, p: ET.Element) -> bool:
        # 1️⃣ nejdřív zkus přímé číslování odstavce
        ppr = p.find("w:pPr", self.NS)
        if ppr is not None:
            numpr = ppr.find("w:numPr", self.NS)
            if numpr is not None:
                num_id = numpr.find("w:numId", self.NS)
                if num_id is not None:
                    val = int(num_id.attrib.get(f"{{{self.NS['w']}}}val", "0"))
                    if val > 0:
                        return True
                    # numId=0 → výslovně vypnuté → dál NEHLEDEJ
                    return False

        # 2️⃣ pokud odstavec nemá vlastní numPr, zkus styl
        style_id = self._paragraph_style_id(p)
        if not style_id:
            return False

        style = self._find_style_by_id(style_id)
        if style is None:
            return False

        ppr = style.find("w:pPr", self.NS)
        if ppr is None:
            return False

        numpr = ppr.find("w:numPr", self.NS)
        if numpr is None:
            return False

        num_id = numpr.find("w:numId", self.NS)
        if num_id is None:
            return False

        val = int(num_id.attrib.get(f"{{{self.NS['w']}}}val", "0"))
        return val > 0


    # Kontrola, zda hůavní kapitola začíná na nové straně
    def paragraph_has_page_break(self, p):
        ppr = p.find("w:pPr", self.NS)
        if ppr is None:
            return False
        return ppr.find("w:pageBreakBefore", self.NS) is not None
    
    def style_has_page_break(self, style_id):
        style = self._find_style_by_id(style_id)
        if style is None:
            return False

        ppr = style.find("w:pPr", self.NS)
        if ppr is None:
            return False

        return ppr.find("w:pageBreakBefore", self.NS) is not None
    
    # Horizontální ruční formátování
    def paragraph_text_raw(self, p: ET.Element) -> str:
        parts = []

        # texty
        for t in p.findall(".//w:t", self.NS):
            if t.text is not None:
                parts.append(t.text)

        # tabulátory z xml (jsou to reálné „zarovnávací“ taby)
        for tab in p.findall(".//w:tab", self.NS):
            parts.append("\t")

        return "".join(parts)
    
    # Kontrola veci co do obsahu nepatri
    def paragraph_has_fld_char(self, p: ET.Element, kind: str) -> bool:
        """
        kind = "begin" | "end"
        """
        for fc in p.findall(".//w:fldChar", self.NS):
            if fc.attrib.get(f"{{{self.NS['w']}}}fldCharType") == kind:
                return True
        return False
    
    # Kontrola, zda 2. sekce ma cislovani od 1
    def section_properties(self, index: int) -> ET.Element | None:
        """
        Vrátí <w:sectPr> pro daný oddíl.
        V tvém splitu může být sectPr:
        - jako samostatný element v body
        - nebo uvnitř posledního odstavce přes w:pPr/w:sectPr
        """
        sec = self.section(index)
        if not sec:
            return None

        # projdi odzadu, většinou je sectPr na konci oddílu
        for el in reversed(sec):
            # 1) sectPr jako přímý element
            if el.tag.endswith("}sectPr"):
                return el

            # 2) sectPr uvnitř odstavce
            if el.tag.endswith("}p"):
                ppr = el.find("w:pPr", self.NS)
                if ppr is not None:
                    sect = ppr.find("w:sectPr", self.NS)
                    if sect is not None:
                        return sect

        return None
    
    def section_has_page_number_field(self, section_index: int) -> bool:
        """
        Vrátí True, pokud má oddíl v hlavičce nebo zápatí pole PAGE.
        """
        sect_pr = self.section_properties(section_index)
        if sect_pr is None:
            return False

        # header / footer reference
        refs = sect_pr.findall("w:headerReference", self.NS) + \
            sect_pr.findall("w:footerReference", self.NS)

        for ref in refs:
            r_type = ref.attrib.get(f"{{{self.NS['w']}}}type")
            r_id = ref.attrib.get(f"{{{self.NS['w']}}}id")
            if not r_id:
                continue

            part_name = f"word/{'header' if 'header' in ref.tag else 'footer'}{r_id}.xml"
            try:
                xml = self._load(part_name)
            except KeyError:
                continue

            # hledáme PAGE field
            for instr in xml.findall(".//w:instrText", self.NS):
                if instr.text and "PAGE" in instr.text.upper():
                    return True

        return False