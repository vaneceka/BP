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

    def _norm(self, text: str) -> str:
        if not text:
            return ""

        text = re.sub(r"\s+", " ", text).strip()
        text = re.sub(r"^\s*\d+(?:\.\d+)*\.?\s*", "", text)
        text = re.sub(r"\s+\d+\s*$", "", text)

        return text.strip()
        
    def run(self, document, assignment=None):
        toc_section = None
        for i in range(document.section_count()):
            if document.has_toc_in_section(i):
                toc_section = i
                break

        if toc_section is None:
            return CheckResult(True, "Obsah neexistuje – nelze ověřit aktuálnost.", 0)

        headings = {
            self._norm(t)
            for (t, lvl) in document.iter_headings()
            if 1 <= lvl <= 3
        }

        if not headings:
            return CheckResult(True, "Dokument nemá nadpisy – obsah nelze posoudit.", 0)

        toc_items = set()

        paragraphs = document.section(toc_section)
        inside_toc = False

        for el in paragraphs:
            for p in el.findall(".//w:p", document.NS):
                for instr in p.findall(".//w:instrText", document.NS):
                    if instr.text and instr.text.strip().upper().startswith("TOC"):
                        inside_toc = True
                        break

                if not inside_toc:
                    continue

                if p.find("w:pPr/w:sectPr", document.NS) is not None:
                    break

                hl = p.find("w:hyperlink", document.NS)
                if hl is None:
                    continue

                text = self._norm(document._visible_text(hl))
                if text and text.lower() != "obsah":
                    toc_items.add(text)

        allowed = {t.lower() for t in self.ALLOWED_EXTRA_TOC_ITEMS}

        missing = []

        for h in headings:
            if h not in toc_items:
                missing.append(h)

        missing.sort()


        extra = []

        for t in toc_items:
            if t not in headings and t.lower() not in allowed:
                extra.append(t)

        extra.sort()

        if not missing and not extra:
            return CheckResult(True, "Obsah je aktuální.", 0)

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