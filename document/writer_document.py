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
                if style.attrib.get(f"{{{NS['style']}}}name", "").lower() == name.lower():
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
