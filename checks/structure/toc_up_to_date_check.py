from checks.base_check import BaseCheck, CheckResult


class TOCUpToDateCheck(BaseCheck):
    name = "Obsah není aktuální"
    penalty = -5

    def run(self, document, assignment=None):
        # 1️⃣ najdi oddíl s obsahem
        toc_section = None
        for i in range(document.section_count()):
            if document.has_toc_in_section(i):
                toc_section = i
                break

        if toc_section is None:
            return CheckResult(True, "Obsah neexistuje – nelze ověřit aktuálnost.", 0)

        # 2️⃣ spočítej nadpisy (H1–H3)
        headings = [
            (t, lvl) for (t, lvl) in document.iter_headings()
            if 1 <= lvl <= 3
        ]

        if not headings:
            return CheckResult(True, "Dokument nemá nadpisy – obsah nelze posoudit.", 0)

        # 3️⃣ spočítej odstavce s textem v oddílu obsahu
        toc_paragraphs = 0
        for el in document.section(toc_section):
            for p in el.findall(".//w:p", document.NS):
                for t in p.findall(".//w:t", document.NS):
                    if t.text and t.text.strip():
                        toc_paragraphs += 1
                        break

        # obvykle je první odstavec jen nadpis "Obsah"
        toc_items = max(0, toc_paragraphs - 1)

        if toc_items == 0:
            return CheckResult(
                False,
                "Obsah je prázdný nebo nebyl aktualizován.",
                self.penalty,
            )

        if toc_items < len(headings):
            return CheckResult(
                False,
                "Obsah pravděpodobně není aktuální – neobsahuje všechny nadpisy.",
                self.penalty,
            )

        return CheckResult(True, "Obsah je pravděpodobně aktuální.", 0)