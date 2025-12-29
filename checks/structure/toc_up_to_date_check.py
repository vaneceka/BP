import re
from checks.base_check import BaseCheck, CheckResult


class TOCUpToDateCheck(BaseCheck):
    name = "Obsah není aktuální"
    penalty = -5

    ALLOWED_EXTRA_TOC_ITEMS = {
        "bibliografie",
        "seznam literatury",
        "literatura",
    }

    # --------------------------------------------------
    # Pomocné funkce
    # --------------------------------------------------

    def _norm(self, text: str) -> str:
        """
        Normalizuje text nadpisu / položky obsahu:
        - odstraní číslování (1., 1.1., 2.3.4.)
        - odstraní koncová čísla (čísla stránek)
        - sjednotí mezery
        """
        if not text:
            return ""

        text = re.sub(r"\s+", " ", text).strip()
        text = re.sub(r"^\s*\d+(?:\.\d+)*\.?\s*", "", text)
        text = re.sub(r"\s+\d+\s*$", "", text)

        return text.strip()

    def _visible_text(self, element, document) -> str:
        """
        Vrátí pouze viditelný text (ignoruje webHidden).
        """
        parts = []

        for r in element.findall(".//w:r", document.NS):
            rpr = r.find("w:rPr", document.NS)
            if rpr is not None and rpr.find("w:webHidden", document.NS) is not None:
                continue

            for t in r.findall("w:t", document.NS):
                if t.text:
                    parts.append(t.text)

        return self._norm("".join(parts))

    # --------------------------------------------------
    # Hlavní logika
    # --------------------------------------------------

    def run(self, document, assignment=None):

        # 1️⃣ Najdi oddíl s obsahem
        toc_section = None
        for i in range(document.section_count()):
            if document.has_toc_in_section(i):
                toc_section = i
                break

        if toc_section is None:
            return CheckResult(True, "Obsah neexistuje – nelze ověřit aktuálnost.", 0)

        # 2️⃣ Nadpisy v dokumentu (H1–H3)
        headings = {
            self._norm(t)
            for (t, lvl) in document.iter_headings()
            if 1 <= lvl <= 3
        }

        if not headings:
            return CheckResult(True, "Dokument nemá nadpisy – obsah nelze posoudit.", 0)

        # 3️⃣ Položky obsahu
        toc_items = set()

        paragraphs = document.section(toc_section)
        inside_toc = False

        for el in paragraphs:
            for p in el.findall(".//w:p", document.NS):

                # detekce TOC pole
                for instr in p.findall(".//w:instrText", document.NS):
                    if instr.text and instr.text.strip().upper().startswith("TOC"):
                        inside_toc = True
                        break

                if not inside_toc:
                    continue

                # konec obsahu
                if p.find("w:pPr/w:sectPr", document.NS) is not None:
                    break

                hl = p.find("w:hyperlink", document.NS)
                if hl is None:
                    continue

                text = self._visible_text(hl, document)
                if text and text.lower() != "obsah":
                    toc_items.add(text)

        # 4️⃣ Porovnání
        allowed = {t.lower() for t in self.ALLOWED_EXTRA_TOC_ITEMS}

        missing = sorted(
            h for h in headings
            if h not in toc_items
        )

        extra = sorted(
            t for t in toc_items
            if t not in headings and t.lower() not in allowed
        )

        if not missing and not extra:
            return CheckResult(True, "Obsah je aktuální.", 0)

        # 5️⃣ Výpis chyb
        lines = []

        if missing:
            lines.append("V obsahu chybí nadpisy:")
            lines += [f"- {t}" for t in missing]

        if extra:
            lines.append("\nV obsahu jsou nadpisy navíc:")
            lines += [f"- {t}" for t in extra]

        return CheckResult(
            False,
            "Obsah pravděpodobně není aktuální:\n" + "\n".join(lines),
            self.penalty,
        )