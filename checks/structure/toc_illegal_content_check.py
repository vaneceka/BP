from checks.base_check import BaseCheck, CheckResult
import re


class TOCIllegalContentCheck(BaseCheck):
    name = "Ruční text nebo nepovolená položka v obsahu"
    penalty = -10

    # ==================================================
    # Pomocné funkce
    # ==================================================

    def _clean_text(self, text: str) -> str:
        return re.sub(r"\s+", " ", text).strip()

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
            return CheckResult(False, "Obsah je poškozený – chybí sdtContent.", self.penalty)

        # 3️⃣ Povolené nadpisy (H1–H3)
        headings = {
            self._clean_text(text)
            for (text, lvl) in document.iter_headings()
            if 1 <= lvl <= 3
        }

        # 3️⃣b Captiony objektů (zakázané v obsahu)
        captions = set()

        if assignment and hasattr(assignment, "objects"):
            for obj in assignment.objects:
                cap = obj.get("caption")
                if cap:
                    captions.add(self._clean_text(cap))

        # 4️⃣ Whitelist kontrola
        errors = []

        for el in toc_sdt:

            # ❌ TABULKA
            if el.tag.endswith("}tbl"):
                errors.append("V obsahu je vložená tabulka.")
                continue

            # ❌ OBRÁZEK / GRAF
            if el.findall(".//w:drawing", document.NS):
                errors.append("V obsahu je vložený obrázek nebo graf.")
                continue

            # ❌ ROVNICE
            if (
                el.findall(".//m:oMath", document.NS)
                or el.findall(".//m:oMathPara", document.NS)
            ):
                errors.append("V obsahu je vložená rovnice.")
                continue

            # ❌ NEPOVOLENÝ ELEMENT
            if not el.tag.endswith("}p"):
                errors.append("V obsahu je nepovolený objekt.")
                continue

            # ⬇️ odtud víme, že je to <w:p>
            link = el.find("w:hyperlink", document.NS)

            # ❌ RUČNÍ TEXT
            if link is None:
                text = self._extract_visible_text(el, document)

                if not text:
                    continue

                if text.lower() == "obsah":
                    continue

                errors.append(f"Ručně vložený text v obsahu: „{text}“")
                continue

            # ❌ CHYBÍ PAGEREF
            if not any(
                instr.text and "PAGEREF" in instr.text
                for instr in link.findall(".//w:instrText", document.NS)
            ):
                errors.append("Položka obsahu bez odkazu na stránku (PAGEREF).")
                continue

            # ✔ TEXT POLOŽKY
            text = self._extract_visible_text(link, document)
            text = re.sub(r"^\d+\.\s*", "", text)

            if not text:
                continue
            
            if text in captions:
                errors.append(
                    f"Položka obsahu odpovídá popisku obrázku/tabulky: „{text}“"
                )
                continue

            if text not in headings:
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

        return CheckResult(True, "Obsah obsahuje pouze platné položky.", 0)