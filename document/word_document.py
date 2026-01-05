from typing import Iterable
import zipfile
import xml.etree.ElementTree as ET
import re

from assignment.word.word_assignment_model import StyleSpec
from pathlib import Path
import xml.dom.minidom as minidom

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
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

    def _is_builtin_style_name(self, name: str) -> bool:
        return name.strip().lower() in BUILTIN_STYLE_NAMES
    
    def split_assignment_styles(self, assignment):
        custom = {}
        builtin = {}

        for name, spec in assignment.styles.items():
            if self._is_builtin_style_name(name):
                builtin[name] = spec
            else:
                custom[name] = spec

        return custom, builtin

    # oddíly
    def _split_into_sections(self):
        sections = []
        current = []

        body = self._xml.find("w:body", self.NS)

        for el in body:
            current.append(el)

            if el.tag.endswith("}sectPr"):
                sections.append(current)
                current = []
                continue

            if el.tag.endswith("}p"):
                ppr = el.find("w:pPr", self.NS)
                if ppr is not None and ppr.find("w:sectPr", self.NS) is not None:
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

    def _find_instr_texts(self, section):
        texts = []
        for el in section:
            for instr in el.findall(".//w:instrText", self.NS):
                if instr.text:
                    texts.append(instr.text.strip())
        return texts
    
    def has_toc_in_section(self, section_index: int) -> bool:
        for instr in self._find_instr_texts(self.section(section_index)):
            if instr.startswith("TOC") and "\\o" in instr:
                return True
        return False

    def has_text_in_section(self, section_index: int) -> bool:
        for el in self.section(section_index):
            for t in el.findall(".//w:t", self.NS):
                if t.text and t.text.strip():
                    return True
        return False

    # bibliography
    def has_bibliography_in_section(self, section_index: int) -> bool:
        for el in self.section(section_index):
            for sdt in el.findall(".//w:sdt", self.NS):
                sdt_pr = sdt.find("w:sdtPr", self.NS)
                if sdt_pr is not None and sdt_pr.find("w:bibliography", self.NS) is not None:
                    return True
        return False
    
    def get_field_instructions(self, section):
        instrs = []

        for el in section:
            for fld in el.findall(".//w:fldSimple", self.NS):
                instr = fld.attrib.get(f"{{{self.NS['w']}}}instr")
                if instr:
                    instrs.append(instr.strip())

        for el in section:
            for instr in el.findall(".//w:instrText", self.NS):
                if instr.text:
                    instrs.append(instr.text.strip())

        return instrs

    def has_list_of_figures_in_section(self, section_index: int) -> bool:
        section = self.section(section_index)

        for instr in self.get_field_instructions(section):
            instr_u = instr.upper()

            if not instr_u.startswith("TOC"):
                continue

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
    
    def _resolve_alignment(self, style):
        while style is not None:
            ppr = style.find("w:pPr", self.NS)
            if ppr is not None:
                jc = ppr.find("w:jc", self.NS)
                if jc is not None:
                    return jc.attrib.get(f"{{{self.NS['w']}}}val")

            based = style.find("w:basedOn", self.NS)
            style = (
                self._find_style_by_id(based.attrib.get(f"{{{self.NS['w']}}}val"))
                if based is not None
                else None
            )

        return None

    def _resolve_space_before(self, style) -> int | None:
        while style is not None:
            ppr = style.find("w:pPr", self.NS)
            if ppr is not None:
                spacing = ppr.find("w:spacing", self.NS)
                if spacing is not None:
                    val = spacing.attrib.get(f"{{{self.NS['w']}}}before")
                    if val is not None:
                        return int(val)

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

            linked = self._get_linked_style(style)
            if linked is not None and id(linked) not in visited:
                style = linked
                continue

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

            linked = self._get_linked_style(style)
            if linked is not None and id(linked) not in visited:
                style = linked
                continue

            based = style.find("w:basedOn", self.NS)
            style = self._find_style_by_id(based.attrib.get(f"{{{self.NS['w']}}}val")) if based is not None else None

        return False
    
    def _resolve_tabs(self, style) -> list[tuple[str, int]] | None:
        tabs = []

        while style is not None:
            ppr = style.find("w:pPr", self.NS)
            if ppr is not None:
                tabs_el = ppr.find("w:tabs", self.NS)
                if tabs_el is not None:
                    for tab in tabs_el.findall("w:tab", self.NS):
                        val = tab.attrib.get(f"{{{self.NS['w']}}}val")
                        pos = tab.attrib.get(f"{{{self.NS['w']}}}pos")
                        if val and pos:
                            tabs.append((val, int(pos)))

                    if tabs:
                        return tabs

            based = style.find("w:basedOn", self.NS)
            style = (
                self._find_style_by_id(based.attrib.get(f"{{{self.NS['w']}}}val"))
                if based is not None
                else None
            )

        return None
    
    def _resolve_line_height(self, style) -> float | None:
        while style is not None:
            ppr = style.find("w:pPr", self.NS)
            if ppr is not None:
                spacing = ppr.find("w:spacing", self.NS)
                if spacing is not None:
                    line = spacing.attrib.get(f"{{{self.NS['w']}}}line")
                    rule = spacing.attrib.get(f"{{{self.NS['w']}}}lineRule", "auto")

                    if line and rule == "auto":
                        return int(line) / 240

            based = style.find("w:basedOn", self.NS)
            style = (
                self._find_style_by_id(based.attrib.get(f"{{{self.NS['w']}}}val"))
                if based is not None
                else None
            )

        return None
    
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

            linked = self._get_linked_style(style)
            if linked is not None and id(linked) not in visited:
                style = linked
                continue

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
        alignment = self._resolve_alignment(style) or default_alignment
        line_height = self._resolve_line_height(style)
        page_break = self._resolve_page_break_before(style)
        num_level = None
        is_numbered = None
        before = self._resolve_space_before(style)
        based_on = None

        indent_left = None
        indent_right = None
        indent_first = None
        indent_hanging = None
        tabs = self._resolve_tabs(style)


        ppr = style.find("w:pPr", self.NS)
        if ppr is not None:
            jc = ppr.find("w:jc", self.NS)
            if jc is not None:
                alignment = jc.attrib.get(f"{{{self.NS['w']}}}val") or alignment

            if ppr is not None:
                numpr = ppr.find("w:numPr", self.NS)
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
            spaceBefore=before,
            indentLeft=int(indent_left) if indent_left else None,
            indentRight=int(indent_right) if indent_right else None,
            indentFirstLine=int(indent_first) if indent_first else None,
            indentHanging=int(indent_hanging) if indent_hanging else None,
            tabs=tabs,
        )
    
    def _find_style(self, *, name: str | None = None, default: bool = False):
        for style in self._styles_xml.findall(".//w:style", self.NS):
            if default:
                if style.attrib.get(f"{{{self.NS['w']}}}default") == "1":
                    return style
                continue

            if not name:
                continue

            style_id = style.attrib.get(f"{{{self.NS['w']}}}styleId")
            if style_id and style_id.strip().lower() == name.strip().lower():
                return style

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
    
    def _paragraph_text(self, p: ET.Element) -> str:
        parts = []
        for t in p.findall(".//w:t", self.NS):
            if t.text:
                parts.append(t.text)

        txt = "".join(parts)

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

    def _style_level_from_styles_xml(self, style_id: str) -> int | None:
        style = self._find_style_by_id(style_id)
        if style is None:
            return None

        visited = set()

        while style is not None and id(style) not in visited:
            visited.add(id(style))

            name_el = style.find("w:name", self.NS)
            if name_el is not None:
                name = (name_el.attrib.get(f"{{{self.NS['w']}}}val") or "").lower()
                m = re.search(r"(heading|nadpis)\s*([0-9]+)", name)
                if m:
                    return int(m.group(2))

            ppr = style.find("w:pPr", self.NS)
            if ppr is not None:
                out = ppr.find("w:outlineLvl", self.NS)
                if out is not None:
                    return int(out.attrib.get(f"{{{self.NS['w']}}}val")) + 1

            based = style.find("w:basedOn", self.NS)
            if based is None:
                break

            style = self._find_style_by_id(
                based.attrib.get(f"{{{self.NS['w']}}}val")
            )

        return None

    def iter_headings(self) -> list[tuple[str, int]]:
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
        
    def has_manual_formatting(self) -> bool:
        return bool(self.find_manual_formatting())

    def has_inline_font_changes(self) -> bool:
        return bool(self.find_inline_font_changes())
    
    def has_html_artifacts(self) -> bool:
        return bool(self.find_html_artifacts())
    
    def find_html_artifacts(self) -> list[tuple[int, str]]:
        results = []

        paragraphs = list(self._xml.findall(".//w:body/w:p", self.NS))

        for i, p in enumerate(paragraphs, start=1):

            if self._paragraph_is_toc_or_object_list(p):
                continue

            text = self._paragraph_text(p)
            if not text:
                continue

            if any(x in text for x in ["&nbsp;", "&#160;", "<", ">"]):
                results.append((i, text))

        return results

    def find_manual_formatting(self) -> list[tuple[int, str]]:
        results = []

        paragraphs = list(self._xml.findall(".//w:body/w:p", self.NS))

        for i, p in enumerate(paragraphs, start=1):
            if self._paragraph_is_toc_or_object_list(p):
                continue
            if self.paragraph_is_caption(p):
                continue

            if (
                p.findall(".//w:drawing", self.NS) or
                p.findall(".//m:oMath", self.NS) or
                p.findall(".//m:oMathPara", self.NS)
            ):
                continue

            text = self._paragraph_text(p)
            if not text:
                continue

            has_ref_field = False

            for instr in p.findall(".//w:instrText", self.NS):
                text = (instr.text or "").strip().upper()
                if text.startswith("REF "):
                    has_ref_field = True
                    break

            if has_ref_field:
                continue

            if p.findall(".//w:hyperlink", self.NS):
                continue

            for r in p.findall(".//w:r", self.NS):
                rpr = r.find("w:rPr", self.NS)
                if rpr is None:
                    continue

                rs = rpr.find("w:rStyle", self.NS)
                if rs is not None:
                    continue

                results.append((i, text))
                break

        return results
    
    def find_inline_font_changes(self) -> list[tuple[int, str]]:
        results = []

        paragraphs = list(self._xml.findall(".//w:body/w:p", self.NS))

        for i, p in enumerate(paragraphs, start=1):
            if self._paragraph_is_toc_or_object_list(p):
                continue
            if self.paragraph_is_caption(p):
                continue

            has_object = False

            if p.findall(".//w:drawing", self.NS):
                has_object = True

            if p.findall(".//m:oMath", self.NS):
                has_object = True

            if has_object:
                continue

            for r in p.findall(".//w:r", self.NS):
                rpr = r.find("w:rPr", self.NS)
                if rpr is None:
                    continue

                if (
                    rpr.find("w:sz", self.NS) is not None or
                    rpr.find("w:rFonts", self.NS) is not None or
                    rpr.find("w:color", self.NS) is not None
                ):
                    text = self._paragraph_text(p)
                    if text:
                        results.append((i, text))
                    break

        return results

    def toc_shows_numbers(self) -> bool | None:
        for instr in self._xml.findall(".//w:instrText", self.NS):
            txt = instr.text or ""
            if txt.strip().startswith("TOC"):
                return "\\n" not in txt
        return None
    
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
    
    def paragraph_text_raw(self, p: ET.Element) -> str:
        parts = []

        for t in p.findall(".//w:t", self.NS):
            if t.text is not None:
                parts.append(t.text)

        for tab in p.findall(".//w:tab", self.NS):
            parts.append("\t")

        return "".join(parts)
        
    def section_properties(self, index: int) -> ET.Element | None:
        sec = self.section(index)
        if not sec:
            return None

        for el in reversed(sec):
            if el.tag.endswith("}sectPr"):
                return el

            if el.tag.endswith("}p"):
                ppr = el.find("w:pPr", self.NS)
                if ppr is not None:
                    sect = ppr.find("w:sectPr", self.NS)
                    if sect is not None:
                        return sect

        return None
        
    def section_has_header_or_footer_content(self, section_index: int) -> bool:
        sect_pr = self.section_properties(section_index)
        if sect_pr is None:
            return False

        refs = (
            sect_pr.findall("w:headerReference", self.NS)
            + sect_pr.findall("w:footerReference", self.NS)
        )

        for ref in refs:
            r_id = ref.attrib.get(f"{{{self.NS['r']}}}id")
            if not r_id:
                continue

            xml = self.load_part_by_rid(r_id)
            if xml is None:
                continue

            # viditelný text
            for t in xml.findall(".//w:t", self.NS):
                if t.text and t.text.strip():
                    return True

            # pole (PAGE, DATE..)
            for instr in xml.findall(".//w:instrText", self.NS):
                if instr.text and instr.text.strip():
                    return True

        return False
        
    def get_heading_num_id(self, p: ET.Element) -> str | None:
        ppr = p.find("w:pPr", self.NS)
        if ppr is not None:
            numpr = ppr.find("w:numPr", self.NS)
            if numpr is not None:
                ilvl = numpr.find("w:ilvl", self.NS)
                if ilvl is None or ilvl.attrib.get(f"{{{self.NS['w']}}}val") == "0":
                    num_id = numpr.find("w:numId", self.NS)
                    if num_id is not None:
                        return num_id.attrib.get(f"{{{self.NS['w']}}}val")

        # ze stylu
        style_id = self._paragraph_style_id(p)
        if not style_id:
            return None

        style = self._find_style_by_id(style_id)
        if style is None:
            return None

        ppr = style.find("w:pPr", self.NS)
        if ppr is None:
            return None

        numpr = ppr.find("w:numPr", self.NS)
        if numpr is None:
            return None

        ilvl = numpr.find("w:ilvl", self.NS)
        if ilvl is None or ilvl.attrib.get(f"{{{self.NS['w']}}}val") == "0":
            num_id = numpr.find("w:numId", self.NS)
            if num_id is not None:
                return num_id.attrib.get(f"{{{self.NS['w']}}}val")

        return None
    #-------------------------------
    # Objekty
    def iter_objects(self):
        objects = []

        for p in self._xml.findall(".//w:p", self.NS):
            for drawing in p.findall(".//w:drawing", self.NS):
                gd = drawing.find(".//a:graphicData", self.NS)
                if gd is None:
                    continue

                uri = gd.attrib.get("uri", "")

                if "picture" in uri:
                    obj_type = "image"
                elif "chart" in uri:
                    obj_type = "chart"
                else:
                    continue

                objects.append({
                    "type": obj_type,
                    "element": p
                })

            if p.findall(".//m:oMath", self.NS) or p.findall(".//m:oMathPara", self.NS):
                objects.append({
                    "type": "equation",
                    "element": p
                })

        for tbl in self._xml.findall(".//w:tbl", self.NS):
            objects.append({
                "type": "table",
                "element": tbl
            })

        return objects
        
    def iter_figure_caption_texts(self) -> list[str]:
        captions = []

        for p in self._xml.findall(".//w:p", self.NS):
            if self.paragraph_has_seq_caption(p) != "Obrázek":
                continue

            text = self._paragraph_text(p)
            if text:
                captions.append(text.strip())

        return captions
    
    def iter_list_of_figures_texts(self) -> list[str]:
        items: list[str] = []

        paragraphs = list(self._xml.findall(".//w:body/w:p", self.NS))

        is_inside_figures_toc = False

        for p in paragraphs:
            for instr in p.findall(".//w:instrText", self.NS):
                if instr.text:
                    txt = instr.text.upper()
                    if "TOC" in txt and "\\C" in txt and "OBRÁZEK" in txt:
                        is_inside_figures_toc = True
                        break

            if not is_inside_figures_toc:
                continue

            if p.find("w:pPr/w:sectPr", self.NS) is not None:
                break

            hl = p.find("w:hyperlink", self.NS)
            if hl is None:
                continue

            text = self._paragraph_text(hl)
            if text:
                items.append(text.strip())

        return items
            
    def object_image_rids(self, element) -> list[str]:
        rids = []

        for blip in element.findall(".//a:blip", self.NS):
            rid = blip.attrib.get(f"{{{self.NS['r']}}}embed")
            if rid:
                rids.append(rid)

        return rids
    
    def get_image_bytes(self, r_id: str) -> bytes | None:
        try:
            rels = self._load("word/_rels/document.xml.rels")
        except KeyError:
            return None

        for rel in rels.findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
            if rel.attrib.get("Id") == r_id:
                target = rel.attrib.get("Target")

                if not target.startswith("media/"):
                    return None

                media_path = f"word/{target}"

                try:
                    with self._zip.open(media_path) as f:
                        return f.read()
                except KeyError:
                    return None

        return None
    
    def paragraph_has_seq_caption(self, p: ET.Element) -> str | None:
        if p is None:
            return None

        for fld in p.findall(".//w:fldSimple", self.NS):
            instr = fld.attrib.get(f"{{{self.NS['w']}}}instr") or ""
            m = re.search(r"\bSEQ\s+([^\s\\]+)", instr, re.IGNORECASE)
            if m:
                return m.group(1)

        instr_parts = []

        for instr in p.findall(".//w:instrText", self.NS):
            if instr.text:
                instr_parts.append(instr.text)

        if instr_parts:
            joined = " ".join(instr_parts)
            m = re.search(r"\bSEQ\s+([^\s\\]+)", joined, re.IGNORECASE)
            if m:
                return m.group(1)

        return None

    def paragraph_before(self, element):
        paragraphs = self._xml.findall(".//w:p", self.NS)

        try:
            idx = paragraphs.index(element)
        except ValueError:
            return None

        for p in reversed(paragraphs[:idx]):
            if self._paragraph_text(p):
                return p

        return None


    def paragraph_after(self, element):
        paragraphs = self._xml.findall(".//w:p", self.NS)

        try:
            idx = paragraphs.index(element)
        except ValueError:
            return None

        for p in paragraphs[idx + 1:]:
            if self._paragraph_text(p):
                return p

        return None
    

    def paragraph_is_caption(self, p: ET.Element) -> bool:
        return self.paragraph_has_seq_caption(p) is not None

    
    def _paragraph_is_toc_or_object_list(self, p):
        ppr = p.find("w:pPr", self.NS)
        if ppr is not None:
            ps = ppr.find("w:pStyle", self.NS)
            if ps is not None:
                style = (ps.attrib.get(f"{{{self.NS['w']}}}val") or "").lower()
                if any(x in style for x in ("toc", "obsah", "seznam")):
                    return True

        for instr in p.findall(".//w:instrText", self.NS):
            if instr.text:
                txt = instr.text.upper()
                if txt.startswith("TOC") or "PAGEREF" in txt:
                    return True

        return False
    
    def iter_crossref_anchors_in_body_text(self) -> set[str]:
        anchors = set()

        for p in self._xml.findall(".//w:body/w:p", self.NS):
            if self._paragraph_is_toc_or_object_list(p):
                continue
            if self.paragraph_is_caption(p):
                continue

            for hl in p.findall(".//w:hyperlink", self.NS):
                a = hl.attrib.get(f"{{{self.NS['w']}}}anchor")
                if a:
                    anchors.add(a)

            for instr in p.findall(".//w:instrText", self.NS):
                if not instr.text:
                    continue
                txt = instr.text.strip()
                if txt.upper().startswith("REF "):
                    parts = txt.split()
                    if len(parts) >= 2:
                        anchors.add(parts[1])

        return anchors
    #-------------------------------------
    # Liteatura
    def has_word_bibliography(self) -> bool:
        for sdt in self._xml.findall(".//w:sdt", self.NS):
            sdt_pr = sdt.find("w:sdtPr", self.NS)
            if sdt_pr is not None and sdt_pr.find("w:bibliography", self.NS) is not None:
                return True
        return False
    
    def count_word_citations(self) -> int:
        count = 0
        for sdt in self._xml.findall(".//w:sdt", self.NS):
            sdt_pr = sdt.find("w:sdtPr", self.NS)
            if sdt_pr is not None and sdt_pr.find("w:citation", self.NS) is not None:
                count += 1
        return count
    
    def count_bibliography_items(self) -> int:
        count = 0

        for sdt in self._xml.findall(".//w:sdt", self.NS):
            sdt_pr = sdt.find("w:sdtPr", self.NS)
            if sdt_pr is None or sdt_pr.find("w:bibliography", self.NS) is None:
                continue

            for p in sdt.findall(".//w:p", self.NS):
                ppr = p.find("w:pPr", self.NS)
                if ppr is None:
                    continue

                ps = ppr.find("w:pStyle", self.NS)
                if ps is None:
                    continue

                style = ps.attrib.get(f"{{{self.NS['w']}}}val", "").lower()
                if style in ("bibliografie", "bibliography"):
                    if self._paragraph_text(p):
                        count += 1

        return count
    #----------------------------------
    # Header/Footer
    def load_part_by_rid(self, r_id: str):
        try:
            rels = self._load("word/_rels/document.xml.rels")
        except KeyError:
            return None

        for rel in rels.findall(
            ".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
        ):
            if rel.attrib.get("Id") == r_id:
                target = rel.attrib.get("Target")
                if not target:
                    return None
                return self._load(f"word/{target}")

        return None

    def resolve_part_target(self, r_id: str) -> str | None:
        try:
            rels = self._load("word/_rels/document.xml.rels")
        except KeyError:
            return None

        for rel in rels.findall(".//rel:Relationship", self.NS):
            if rel.attrib.get("Id") != r_id:
                continue

            target = rel.attrib.get("Target")
            if not target:
                return None

            if not target.startswith("word/"):
                return f"word/{target}"

            return target

        return None
    
    def has_text_after_paragraph(self, paragraphs, index: int) -> bool:
        for p in paragraphs[index + 1:]:
            if self._paragraph_text(p):
                return True
        return False
    
    def paragraph_is_generated_by_field(self, p) -> bool:
        if p.findall(".//w:fldChar", self.NS):
            return True
        if p.findall(".//w:instrText", self.NS):
            return True
        return False
    
    def _visible_text(self, element) -> str:
        parts = []
        for r in element.findall(".//w:r", self.NS):
            rpr = r.find("w:rPr", self.NS)
            if rpr is not None and rpr.find("w:webHidden", self.NS) is not None:
                continue

            for t in r.findall("w:t", self.NS):
                if t.text:
                    parts.append(t.text)

        txt = "".join(parts)
        txt = re.sub(r"\s+", " ", txt).strip()
        return txt
    
    def first_heading_in_section(self, section_index: int, level: int = 1):
        for el in self.section(section_index):
            if not el.tag.endswith("}p"):
                continue

            if not self._paragraph_text(el):
                continue

            style_id = self._paragraph_style_id(el)
            if not style_id:
                continue

            lvl = self._style_level_from_styles_xml(style_id)
            if lvl == level:
                return el

        return None
  

    def iter_section_paragraphs(self, section_index: int) -> Iterable[ET.Element]:
        for el in self.section(section_index):
            if el.tag.endswith("}p"):
                yield el
            yield from el.findall(".//w:p", self.NS)