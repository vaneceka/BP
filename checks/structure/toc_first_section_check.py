from checks.base_check import BaseCheck, CheckResult


class TOCFirstSectionContentCheck(BaseCheck):
    name = "Obsah obsahuje text z prvního oddílu"
    penalty = -10  # násobí se

    def run(self, document, assignment=None):
        # 1️⃣ najdi oddíl s obsahem
        toc_section = None
        for i in range(document.section_count()):
            if document.has_toc_in_section(i):
                toc_section = i
                break

        if toc_section is None:
            return CheckResult(True, "Obsah neexistuje.", 0)

        # 2️⃣ nadpisy v dokumentu (TOC je z nich)
        headings = document.iter_headings()

        # 3️⃣ nadpisy, které patří do 1. oddílu
        first_section = document.section(0)

        illegal = []

        for text, level in headings:
            # jednoduchá heuristika:
            # pokud se text nadpisu nachází v 1. oddílu → je to špatně
            for el in first_section:
                if text in document._paragraph_text(el):
                    illegal.append(text)
                    break

        if illegal:
            return CheckResult(
                False,
                "Obsah obsahuje nadpisy z prvního oddílu:\n"
                + "\n".join(f"– {t}" for t in illegal[:5]),
                self.penalty * len(illegal),
            )

        return CheckResult(
            True,
            "Obsah neobsahuje žádný text z prvního oddílu.",
            0,
        )