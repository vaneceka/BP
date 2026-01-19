from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET
from assignment.word.word_assignment_model import StyleSpec
import xml.dom.minidom as minidom
import xml.etree.ElementTree as ET

from document.text_document import TextDocument

NS = {
    "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "fo": "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
    "loext": "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0",
}
COVER_STYLE_ALIASES = {
            "desky-fakulta": [
                "desky-fakulta",
                "Faculty",
                "University",
                "Title",
            ],
            "desky-nazev-prace": [
                "desky-nazev-prace",
                "Title",
            ],
            "desky-rok-a-jmeno": [
                "desky-rok-a-jmeno",
                "Subtitle",
                "Author",
            ],
        }

class WriterDocument(TextDocument):
    def __init__(self, path: str):
        self.path = path
        self._zip = zipfile.ZipFile(path)
        self.content = self._load("content.xml")
        self.styles = self._load("styles.xml")
    
        

    def _load(self, name):
        with self._zip.open(name) as f:
            return ET.fromstring(f.read())
        
    def save_xml(self, out_dir: str | Path = "debug_writer_xml"):
        out_dir = Path(out_dir)
        out_dir.mkdir(parents=True, exist_ok=True)

        self._save_pretty_xml(self.content, out_dir / "content.xml")
        self._save_pretty_xml(self.styles, out_dir / "styles.xml")

    def _save_pretty_xml(self, root: ET.Element, path: Path):
        rough_string = ET.tostring(root, encoding="utf-8")
        reparsed = minidom.parseString(rough_string)
        pretty = reparsed.toprettyxml(indent="  ", encoding="utf-8")

        with open(path, "wb") as f:
            f.write(pretty)
        
    def _find_style(self, name: str):
        for root in (self.styles, self.content):
            for style in root.findall(".//style:style", NS):

                internal = style.attrib.get(f"{{{NS['style']}}}name", "").lower()
                if internal == name.lower():
                    return style

                display = style.attrib.get(f"{{{NS['style']}}}display-name", "")
                if display.lower() == name.lower():
                    return style

        return None
        
    def _build_style_spec(self, style, *, default_alignment=None) -> StyleSpec:
        para_props = style.find("style:paragraph-properties", NS)

        font = self._resolve_font_name(style)

        size = self._resolve_font_size_pt(style)
        bold = self._resolve_bold(style)

        italic = False
        fs = self._resolve_text_prop(style, f"{{{NS['fo']}}}font-style")
        if fs is not None:
            italic = (fs.lower() == "italic")

        # alignment (tady zatím bez dědičnosti; kdyby bylo potřeba, uděláme stejně)
        alignment = default_alignment
        if para_props is not None:
            alignment = para_props.attrib.get(f"{{{NS['fo']}}}text-align", default_alignment)

        # color: pokud není, ber černou (správně)
        color_val = self._resolve_text_prop(style, f"{{{NS['fo']}}}color")
        if color_val is None:
            color = "000000"
        else:
            color = color_val.lstrip("#").upper()

        # ALL CAPS (fo:text-transform="uppercase")
        text_transform = self._resolve_text_prop(
            style,
            f"{{{NS['fo']}}}text-transform"
        )

        all_caps = (text_transform == "uppercase")

        based_on = style.attrib.get(f"{{{NS['style']}}}parent-style-name")

        mt = self._resolve_paragraph_prop(
            style,
            f"{{{NS['fo']}}}margin-top"
        )
        space_before = self._cm_to_twips(mt) if mt else None

        tabs = self._resolve_tabs(style)

        page_break_before = None
        if para_props is not None:
            br = para_props.attrib.get(f"{{{NS['fo']}}}break-before")
            if br == "page":
                page_break_before = True
            elif br is not None:
                page_break_before = False

        # === OUTLINE LEVEL (ODT experiment) ===
        outline_level = style.attrib.get(
            f"{{{NS['style']}}}default-outline-level"
        )

        num_level = None
        if outline_level:
            try:
                num_level = int(outline_level) - 1
            except ValueError:
                num_level = None

        return StyleSpec(
            name=style.attrib.get(f"{{{NS['style']}}}name"),
            font=font,
            size=size,
            bold=bold,
            italic=italic,
            alignment=alignment,
            color=color,
            allCaps=all_caps,
            basedOn=based_on,
            spaceBefore=space_before,
            tabs=tabs,
            pageBreakBefore=page_break_before,
            numLevel=num_level,
        )
    
    def get_style_by_any_name(self, names, *, default_alignment=None):
        for name in names:
            style = self._find_style(name)
            if style is not None:
                return self._build_style_spec(
                    style,
                    default_alignment=default_alignment
                )
        return None
    
    def get_doc_default_font_size(self) -> int | None:
        for root in (self.styles, self.content):
            default = root.find(".//style:default-style", NS)
            if default is not None:
                tp = default.find("style:text-properties", NS)
                if tp is not None:
                    fs = tp.attrib.get(f"{{{NS['fo']}}}font-size")
                    if fs and fs.endswith("pt"):
                        return int(float(fs.replace("pt", "")))
        return None
    
    def get_cover_style(self, key: str) -> StyleSpec | None:
        names = COVER_STYLE_ALIASES.get(key, [])
        return self.get_style_by_any_name(
            names,
            default_alignment="center"
        )
    
    def _get_parent_style_name(self, style_el: ET.Element) -> str | None:
        # atribut style:parent-style-name
        return style_el.attrib.get(f"{{{NS['style']}}}parent-style-name")
    
    def _resolve_text_prop(self, style_el: ET.Element, attr_qname: str) -> str | None:
        visited = set()

        while style_el is not None and id(style_el) not in visited:
            visited.add(id(style_el))

            tp = style_el.find("style:text-properties", NS)
            if tp is not None:
                val = tp.attrib.get(attr_qname)
                if val is not None:
                    return val

            parent_name = self._get_parent_style_name(style_el)
            if not parent_name:
                break

            style_el = self._find_style(parent_name)

        return None
    
    def _resolve_bold(self, style_el: ET.Element) -> bool:
        visited = set()

        while style_el is not None and id(style_el) not in visited:
            visited.add(id(style_el))

            tp = style_el.find("style:text-properties", NS)
            if tp is not None:
                # klasika
                fw = tp.attrib.get(f"{{{NS['fo']}}}font-weight")
                if fw is not None:
                    return fw.lower() == "bold"

                # LibreOffice často: style:font-style-name="Bold"
                fsn = tp.attrib.get(f"{{{NS['style']}}}font-style-name")
                if fsn is not None and "bold" in fsn.lower():
                    return True

            parent_name = self._get_parent_style_name(style_el)
            if not parent_name:
                break
            style_el = self._find_style(parent_name)

        return False
    
    def _resolve_font_size_pt(self, style_el: ET.Element) -> int | None:
        val = self._resolve_text_prop(style_el, f"{{{NS['fo']}}}font-size")
        if not val:
            return None

        v = val.strip().lower()
        if v.endswith("pt"):
            try:
                return int(float(v[:-2]))
            except Exception:
                return None

        return None
    
    def _resolve_paragraph_prop(self, style_el, attr_qname):
        visited = set()

        while style_el is not None and id(style_el) not in visited:
            visited.add(id(style_el))

            pp = style_el.find("style:paragraph-properties", NS)
            if pp is not None:
                val = pp.attrib.get(attr_qname)
                if val is not None:
                    return val

            parent = self._get_parent_style_name(style_el)
            if not parent:
                break
            style_el = self._find_style(parent)

        return None
    
    def _cm_to_twips(self, cm: str) -> int | None:
        try:
            return int(float(cm.replace("cm", "")) * 567)
        except Exception:
            return None
        
    def _resolve_font_name(self, style_el: ET.Element) -> str | None:
        val = self._resolve_text_prop(
            style_el,
            f"{{{NS['style']}}}font-name"
        )
        return val
    
    def _resolve_tabs(self, style_el: ET.Element) -> list[tuple[str, int]] | None:
        visited = set()

        while style_el is not None and id(style_el) not in visited:
            visited.add(id(style_el))

            para = style_el.find("style:paragraph-properties", NS)
            if para is not None:
                tabs_el = para.find("style:tab-stops", NS)
                if tabs_el is not None:
                    tabs = []
                    for t in tabs_el.findall("style:tab-stop", NS):
                        pos = t.attrib.get(f"{{{NS['style']}}}position")
                        typ = t.attrib.get(f"{{{NS['style']}}}type", "left")

                        if pos and pos.endswith("cm"):
                            cm = float(pos.replace("cm", ""))
                            twips = int(round(cm * 567))
                            tabs.append((typ, twips))

                    return tabs if tabs else None

            parent = self._get_parent_style_name(style_el)
            if not parent:
                break
            style_el = self._find_style(parent)

        return None
    
    def get_style_parent(self, style_name: str) -> str | None:
        style_el = self._find_style(style_name)
        if style_el is None:
            return None

        parent = style_el.attrib.get(
            "{urn:oasis:names:tc:opendocument:xmlns:style:1.0}parent-style-name"
        )

        return parent
    
    def get_used_paragraph_styles(self) -> set[str]:
        used = set()

        for p in self.content.findall(".//text:p", NS):
            style_name = p.attrib.get(f"{{{NS['text']}}}style-name")
            if style_name:
                used.add(style_name)

        return used


    def style_exists(self, style_name: str) -> bool:
        return self._find_style(style_name) is not None
    
    def get_custom_style(self, style_name: str) -> StyleSpec | None:
        style = self._find_style(style_name)
        if style is None:
            return None
        return self._build_style_spec(style)
    
    def get_heading_style(self, level: int) -> StyleSpec | None:
        style = self._find_style(f"Heading {level}")
        if style is None:
            return None

        return self._build_style_spec(
            style,
            default_alignment="start"
        )
 
    def _find_outline_style(self):
        roots = [self.styles, self.content]

        for root in roots:
            if root is None:
                continue

            outline = root.find(".//text:outline-style", NS)
            if outline is not None:
                return outline

            outline = root.find(
                ".//{urn:oasis:names:tc:opendocument:xmlns:text:1.0}outline-style"
            )
            if outline is not None:
                return outline

        return None


    def get_heading_numbering_info(self, level: int):
        outline = self._find_outline_style()
        if outline is None:
            return False, False, None

        lvl = outline.find(
            f".//{{{NS['text']}}}outline-level-style[@{{{NS['text']}}}level='{level}']"
        )
        if lvl is None:
            return False, False, None

        num_format = lvl.attrib.get(f"{{{NS['style']}}}num-format")
        if not num_format:
            return False, False, None

        num_list = (
            lvl.attrib.get(f"{{{NS['loext']}}}num-list-format")
            or lvl.attrib.get(f"{{{NS['text']}}}num-list-format")
            or ""
        )

        required = [f"%{i}%" for i in range(1, level + 1)]
        is_hierarchical = all(r in num_list for r in required)

        return True, is_hierarchical, level - 1
    
    def get_heading_outline_level(self, level: int) -> int | None:
        """
        Vrátí outline level (0-based) pro Heading X
        """
        style = self._find_style(f"Heading {level}")
        if style is None:
            return None

        val = style.attrib.get(f"{{{NS['style']}}}default-outline-level")
        if not val:
            return None

        try:
            return int(val) - 1
        except ValueError:
            return None
        

    def iter_headings(self) -> list[tuple[str, int]]:
        items = []

        for p in self.content.findall(".//text:h", NS):
            txt = "".join(p.itertext()).strip()
            if not txt:
                continue

            lvl = p.attrib.get(f"{{{NS['text']}}}outline-level")
            if not lvl:
                continue

            try:
                level = int(lvl)
            except ValueError:
                continue

            items.append((txt, level))

        return items
    
    def find_inline_formatting(self) -> list[dict]:
        results = []

        for p in self.content.findall(".//text:p", NS):

            for span in p.findall(".//text:span", NS):
                span_text = "".join(span.itertext()).strip()
                if not span_text:
                    continue

                span_style_name = span.attrib.get(f"{{{NS['text']}}}style-name")
                if not span_style_name:
                    continue

                span_style = self._find_style(span_style_name)
                if span_style is None:
                    continue

                tp = span_style.find("style:text-properties", NS)
                if tp is None:
                    continue

                if tp.attrib.get(f"{{{NS['fo']}}}font-weight") == "bold":
                    results.append({
                        "text": span_text,
                        "problem": "tučné písmo",
                    })

                if tp.attrib.get(f"{{{NS['fo']}}}font-style") == "italic":
                    results.append({
                        "text": span_text,
                        "problem": "kurzíva",
                    })

                if f"{{{NS['fo']}}}font-size" in tp.attrib:
                    results.append({
                        "text": span_text,
                        "problem": "změna velikosti písma",
                    })

                if f"{{{NS['fo']}}}color" in tp.attrib:
                    results.append({
                        "text": span_text,
                        "problem": "změna barvy",
                    })

        return results
    

    def iter_main_headings(self):
        for h in self.content.findall(".//text:h", NS):
            lvl = h.attrib.get(f"{{{NS['text']}}}outline-level")
            if lvl == "1":
                yield h


    def heading_starts_on_new_page(self, h) -> bool:
        style_name = h.attrib.get(f"{{{NS['text']}}}style-name")
        if not style_name:
            return False

        style = self._find_style(style_name)
        if style is None:
            return False

        pprops = style.find("style:paragraph-properties", NS)
        if pprops is None:
            return False

        return pprops.attrib.get(f"{{{NS['fo']}}}break-before") == "page"
    
    def get_visible_text(self, element) -> str:
        return "".join(element.itertext()).strip()
    
    

    def paragraph_text_raw(self, p) -> str:
        parts = []

        for el in p.iter():
            # normální text
            if el.text:
                parts.append(el.text)

            # <text:s text:c="N"/>
            if el.tag == f"{{{NS['text']}}}s":
                c = el.attrib.get(f"{{{NS['text']}}}c")
                parts.append(" " * (int(c) if c and c.isdigit() else 1))

            # <text:tab/>
            elif el.tag == f"{{{NS['text']}}}tab":
                parts.append("\t")

        return "".join(parts)

    def paragraph_is_toc(self, p) -> bool:
        style = p.attrib.get(f"{{{NS['text']}}}style-name", "").lower()
        return "toc" in style or "obsah" in style
    
    def iter_paragraphs(self):
        return self.content.findall(".//text:p", NS)
    
    def paragraph_is_empty(self, p) -> bool:
        return not self.paragraph_text(p).strip()


    def paragraph_has_text(self, p) -> bool:
        return bool(self.paragraph_text(p).strip())


    def paragraph_text(self, p) -> str:
        return self._odt_text_with_specials(p).strip()


    def paragraph_style_name(self, p) -> str:
        return p.attrib.get(f"{{{NS['text']}}}style-name", "bez stylu")


    def paragraph_has_spacing_before(self, p) -> bool:
        style_name = self.paragraph_style_name(p)
        style = self._find_style(style_name)
        if style is None:
            return False

        pp = style.find("style:paragraph-properties", NS)
        if pp is None:
            return False

        mt = pp.attrib.get(f"{{{NS['fo']}}}margin-top")
        return mt is not None and mt != "0cm"
    
    def _odt_text_with_specials(self, el) -> str:
        out = []

        def walk(node):
            if node.text:
                out.append(node.text)

            for ch in list(node):
                if ch.tag == f"{{{NS['text']}}}s":
                    c = ch.attrib.get(f"{{{NS['text']}}}c")
                    out.append(" " * (int(c) if c and c.isdigit() else 1))

                elif ch.tag == f"{{{NS['text']}}}tab":
                    out.append("\t")

                elif ch.tag == f"{{{NS['text']}}}line-break":
                    out.append("\n")

                else:
                    walk(ch)

                if ch.tail:
                    out.append(ch.tail)

        walk(el)
        return "".join(out)
    
    def _paragraph_is_toc_or_object_list(self, p) -> bool:
        style = p.attrib.get(f"{{{NS['text']}}}style-name", "").lower()

        if any(x in style for x in ("toc", "obsah", "seznam")):
            return True

        return False
    
    def paragraph_is_generated(self, p) -> bool:
        return False
    
    def paragraph_has_page_break(self, p) -> bool:
        return False