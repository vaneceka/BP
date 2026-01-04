from checks.base_check import BaseCheck, CheckResult
import re


class TOCIllegalContentCheck(BaseCheck):
    name = "Ruční text nebo nepovolená položka v obsahu"
    penalty = -10

    def _clean_text(self, text: str) -> str:
        return re.sub(r"\s+", " ", text or "").strip()

    def run(self, document, assignment=None):
        toc_section = None
        for i in range(document.section_count()):
            if document.has_toc_in_section(i):
                toc_section = i
                break

        if toc_section is None:
            return CheckResult(True, "Obsah dokumentu neexistuje.", 0)

        toc_sdt = None
        for el in document.section(toc_section):
            if el.tag.endswith("}sdt"):
                pr = el.find("w:sdtPr", document.NS)
                if pr is not None and pr.find("w:docPartObj", document.NS) is not None:
                    toc_sdt = el.find("w:sdtContent", document.NS)
                    break

        if toc_sdt is None:
            return CheckResult(
                False,
                "Obsah je poškozený – chybí sdtContent.",
                self.penalty
            )

        heading_bookmarks = set()

        for p in document._xml.findall(".//w:body/w:p", document.NS):
            sid = document._paragraph_style_id(p)
            if not sid:
                continue

            lvl = document._style_level_from_styles_xml(sid)
            if lvl is None or not (1 <= lvl <= 3):
                continue

            for bm in p.findall(".//w:bookmarkStart", document.NS):
                name = bm.attrib.get(f"{{{document.NS['w']}}}name")
                if name:
                    heading_bookmarks.add(name)

        errors = []

        for el in toc_sdt:
            if el.tag.endswith("}tbl"):
                errors.append("V obsahu je vložená tabulka.")
                continue

            if el.findall(".//w:drawing", document.NS):
                errors.append("V obsahu je vložený obrázek nebo graf.")
                continue

            if (
                el.findall(".//m:oMath", document.NS)
                or el.findall(".//m:oMathPara", document.NS)
            ):
                errors.append("V obsahu je vložená rovnice.")
                continue

            if not el.tag.endswith("}p"):
                errors.append("V obsahu je nepovolený objekt.")
                continue

            link = el.find("w:hyperlink", document.NS)

            if link is None:
                text = document._visible_text(el)
                if text and text.lower() != "obsah":
                    errors.append(f"Ručně vložený text v obsahu: „{text}“")
                continue

            text = document._visible_text(link)

            has_pageref = False

            for instr in link.findall(".//w:instrText", document.NS):
                if instr.text and "PAGEREF" in instr.text:
                    has_pageref = True
                    break

            if not has_pageref:
                errors.append(f"Položka obsahu bez odkazu na stránku: „{text}“")
                continue

            anchor = link.attrib.get(f"{{{document.NS['w']}}}anchor")

            if not anchor:
                errors.append(f"Položka obsahu bez anchor odkazu: „{text}“")
                continue

            clean_text = self._clean_text(text).lower()

            if clean_text in {"bibliografie", "literatura", "references"}:
                if document.has_word_bibliography():
                    continue  

                if anchor not in heading_bookmarks:
                    errors.append(
                        f"Položka obsahu bez odpovídajícího nadpisu: „{text}“"
                    )

        if errors:
            return CheckResult(
                False,
                "V obsahu jsou neplatné položky:\n"
                + "\n".join(f"– {e}" for e in errors),
                self.penalty * len(errors),
            )

        return CheckResult(
            True,
            "Obsah obsahuje pouze platné položky.",
            0,
        )