from checks.base_check import BaseCheck, CheckResult
import re


class TOCIllegalContentCheck(BaseCheck):
    name = "Ruční text nebo nepovolená položka v obsahu"
    penalty = -10

    # ==================================================
    # Pomocné funkce
    # ==================================================

    def _clean_text(self, text: str) -> str:
        return re.sub(r"\s+", " ", text or "").strip()

    def _extract_visible_text(self, element, document) -> str:
        parts = []

        for r in element.findall(".//w:r", document.NS):
            rpr = r.find("w:rPr", document.NS)
            if rpr is not None and rpr.find("w:webHidden", document.NS) is not None:
                continue

            for t in r.findall("w:t", document.NS):
                if t.text:
                    parts.append(t.text)

        return self._clean_text("".join(parts))

    # ==================================================
    # Hlavní logika
    # ==================================================

    def run(self, document, assignment=None):

        # 1️⃣ Najdi sekci s obsahem
        toc_section = None
        for i in range(document.section_count()):
            if document.has_toc_in_section(i):
                toc_section = i
                break

        if toc_section is None:
            return CheckResult(True, "Obsah dokumentu neexistuje.", 0)

        # 2️⃣ Najdi sdtContent (TOC)
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

        # 3️⃣ Bookmarky skutečných nadpisů (H1–H3)
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

        # 4️⃣ Kontrola položek obsahu
        errors = []

        for el in toc_sdt:

            # ❌ tabulka v obsahu
            if el.tag.endswith("}tbl"):
                errors.append("V obsahu je vložená tabulka.")
                continue

            # ❌ obrázek / graf
            if el.findall(".//w:drawing", document.NS):
                errors.append("V obsahu je vložený obrázek nebo graf.")
                continue

            # ❌ rovnice
            if (
                el.findall(".//m:oMath", document.NS)
                or el.findall(".//m:oMathPara", document.NS)
            ):
                errors.append("V obsahu je vložená rovnice.")
                continue

            # ❌ nepovolený element
            if not el.tag.endswith("}p"):
                errors.append("V obsahu je nepovolený objekt.")
                continue

            # ⬇️ je to <w:p>
            link = el.find("w:hyperlink", document.NS)

            # ❌ ručně vložený text
            if link is None:
                text = self._extract_visible_text(el, document)
                if text and text.lower() != "obsah":
                    errors.append(f"Ručně vložený text v obsahu: „{text}“")
                continue

            # ✔ text položky (jen pro hlášení)
            text = self._extract_visible_text(link, document)

            # ❌ chybí PAGEREF
            if not any(
                instr.text and "PAGEREF" in instr.text
                for instr in link.findall(".//w:instrText", document.NS)
            ):
                errors.append(f"Položka obsahu bez odkazu na stránku: „{text}“")
                continue

            # ✔ anchor
            anchor = link.attrib.get(f"{{{document.NS['w']}}}anchor")

            if not anchor:
                errors.append(f"Položka obsahu bez anchor odkazu: „{text}“")
                continue

            # ⛔ POVOLENÁ VÝJIMKA – Bibliografie
            clean_text = self._clean_text(text).lower()

            if clean_text in {"bibliografie", "literatura", "references"}:
                # ověř, že v dokumentu skutečně existuje Word bibliografie
                if document.has_word_bibliography():
                    continue  # OK – Wordem generovaná bibliografie

                if anchor not in heading_bookmarks:
                    errors.append(
                        f"Položka obsahu bez odpovídajícího nadpisu: „{text}“"
                    )

        # 5️⃣ Výsledek
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