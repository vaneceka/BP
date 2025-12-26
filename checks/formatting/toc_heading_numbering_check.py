from checks.base_check import BaseCheck, CheckResult


class TocHeadingNumberingCheck(BaseCheck):
    name = "Čísla nadpisů se nezobrazují v obsahu (nebo naopak)"
    penalty = -5

    def run(self, document, assignment=None):
        toc_numbers = document.toc_shows_numbers()
        if toc_numbers is None:
            return CheckResult(True, "Obsah v dokumentu není.", 0)

        headings_numbered = False
        for lvl in (1, 2, 3):
            st = document.get_heading_style(lvl)
            if st and st.isNumbered:
                headings_numbered = True
                break

        if toc_numbers != headings_numbered:
            return CheckResult(
                False,
                "Číslování nadpisů a nastavení obsahu nejsou v souladu.",
                self.penalty,
            )

        return CheckResult(
            True,
            "Číslování nadpisů odpovídá nastavení obsahu.",
            0,
        )