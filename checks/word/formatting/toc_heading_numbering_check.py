# from checks.base_check import BaseCheck, CheckResult


# class TocHeadingNumberingCheck(BaseCheck):
#     name = "Čísla nadpisů se nezobrazují v obsahu (nebo naopak)"
#     penalty = -5

#     def run(self, document, assignment=None):
#         toc_numbers = document.toc_shows_numbers()
#         if toc_numbers is None:
#             return CheckResult(True, "Obsah v dokumentu není.", 0)

#         headings_numbered = False
#         for lvl in (1, 2, 3):
#             st = document.get_heading_style(lvl)
#             if st and st.isNumbered:
#                 headings_numbered = True
#                 break

#         if toc_numbers != headings_numbered:
#             return CheckResult(
#                 False,
#                 "Číslování nadpisů a nastavení obsahu nejsou v souladu.",
#                 self.penalty,
#             )

#         return CheckResult(
#             True,
#             "Číslování nadpisů odpovídá nastavení obsahu.",
#             0,
#         )

from checks.base_check import BaseCheck, CheckResult

class TocHeadingNumberingCheck(BaseCheck):
    name = "Čísla nadpisů se nezobrazují v obsahu (nebo naopak)"
    penalty = -5

    def run(self, document, assignment=None):
        any_toc = False
        mismatches = []

        for level in (1, 2, 3):
            toc_has = document.toc_level_contains_numbers(level)
            if toc_has is None:
                continue  # pro tuhle úroveň TOC nemáme položky

            any_toc = True
            head_num = document.heading_level_is_numbered(level)

            if toc_has:
                any_numbered_toc = True
            if head_num:
                any_numbered_heading = True

            if not toc_has and not head_num:
                mismatches.append((level, False, False))
                continue

            if toc_has != head_num:
                mismatches.append((level, toc_has, head_num))

        if not any_toc:
            return CheckResult(True, "Obsah v dokumentu není.", 0)
        
        if mismatches:
            lines = []
            for lvl, toc_has, head_num in mismatches:
                if lvl == "ALL":
                    lines.append(
                        "- Nadpisy nejsou číslované a obsah neobsahuje číslování"
                    )
                else:
                    lines.append(
                        f"- Úroveň {lvl}: TOC čísla={'ANO' if toc_has else 'NE'}, "
                        f"nadpisy číslované={'ANO' if head_num else 'NE'}"
                    )

            return CheckResult(
                False,
                "Číslování nadpisů a obsah nejsou v souladu:\n" + "\n".join(lines),
                self.penalty,
            )

        return CheckResult(True, "Číslování nadpisů odpovídá obsahu.", 0)