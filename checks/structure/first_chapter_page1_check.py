from checks.base_check import BaseCheck, CheckResult
import re


class FirstChapterStartsOnPageOneCheck(BaseCheck):
    name = "První kapitola ve druhém oddílu nezačíná na straně 1"
    penalty = -5

    def _clean_text(self, text: str) -> str:
        return re.sub(r"\s+", " ", text or "").strip()

    def _paragraph_visible_text(self, p, document) -> str:
        parts = []
        for t in p.findall(".//w:t", document.NS):
            if t.text:
                parts.append(t.text)
        return self._clean_text("".join(parts))

    def _paragraph_style_id(self, p, document) -> str | None:
        ppr = p.find("w:pPr", document.NS)
        if ppr is None:
            return None
        ps = ppr.find("w:pStyle", document.NS)
        if ps is None:
            return None
        return ps.attrib.get(f"{{{document.NS['w']}}}val")

    def _is_visible_content_before_heading(self, el, document) -> bool:
        """
        Co považujeme za 'obsah' před první kapitolou:
        - viditelný text
        - tabulka
        - obrázek/graf (drawing/pict)
        - rovnice (oMath)
        """
        # tabulka
        if el.tag.endswith("}tbl"):
            return True

        # obrázky/grafy
        if el.findall(".//w:drawing", document.NS) or el.findall(".//w:pict", document.NS):
            return True

        # rovnice
        if el.findall(".//m:oMath", document.NS) or el.findall(".//m:oMathPara", document.NS):
            return True

        # text
        if el.tag.endswith("}p"):
            txt = self._paragraph_visible_text(el, document)
            if txt:
                return True

        return False

    def run(self, document, assignment=None):

        # Musí existovat alespoň 2 oddíly
        if document.section_count() < 2:
            return CheckResult(
                False,
                "Dokument nemá druhý oddíl – nelze ověřit začátek číslování.",
                self.penalty,
            )

        second_section = document.section(1)

        # 1) ověř restart číslování stránek ve 2. oddílu
        sect_pr = document.section_properties(1)
        if sect_pr is None:
            return CheckResult(
                False,
                "Ve druhém oddílu nebyly nalezeny vlastnosti oddílu (sectPr).",
                self.penalty,
            )
        
        # 0️⃣ MUSÍ existovat číslování stránek
        if not document.section_has_page_number_field(1):
            return CheckResult(
                False,
                "Ve druhém oddílu není zapnuté číslování stránek.",
                self.penalty,
            )

        pg_num = sect_pr.find("w:pgNumType", document.NS)
        if pg_num is None:
            return CheckResult(
                False,
                "Ve druhém oddílu není nastavené číslování stránek (pgNumType).",
                self.penalty,
            )

        start = pg_num.attrib.get(f"{{{document.NS['w']}}}start")
        if start != "1":
            return CheckResult(
                False,
                "Číslování stránek ve druhém oddílu nezačíná od 1.",
                self.penalty,
            )

        # 2) najdi první nadpis úrovně 1 ve 2. oddílu (NE podle stringu "Heading 1")
        first_h1 = None
        for el in second_section:
            if not el.tag.endswith("}p"):
                continue

            sid = self._paragraph_style_id(el, document)
            if not sid:
                continue

            lvl = document._style_level_from_styles_xml(sid)
            if lvl == 1:
                # ignoruj prázdné nadpisy
                if self._paragraph_visible_text(el, document):
                    first_h1 = el
                    break

        if first_h1 is None:
            return CheckResult(
                False,
                "Ve druhém oddílu nebyla nalezena žádná kapitola (nadpis úrovně 1).",
                self.penalty,
            )

        # 3) ověř, že před první kapitolou není žádný viditelný obsah
        for el in second_section:
            if el is first_h1:
                break

            if self._is_visible_content_before_heading(el, document):
                return CheckResult(
                    False,
                    "První kapitola není první obsah v druhém oddílu (není na straně 1).",
                    self.penalty,
                )

        return CheckResult(
            True,
            "První kapitola ve druhém oddílu začíná na straně 1.",
            0,
        )
    
    