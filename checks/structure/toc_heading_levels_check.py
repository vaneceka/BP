import re
from checks.base_check import BaseCheck, CheckResult


class TOCHeadingLevelsCheck(BaseCheck):
    name = "Nadpisy prvních tří úrovní v obsahu chybí"
    penalty = -5

    TOC_LEVEL_RE = re.compile(r'\\o\s*"(\d+)\s*-\s*(\d+)"')

    def run(self, document, assignment=None):

        toc_found = False

        for i in range(document.section_count()):
            section = document.section(i)

            for instr in document.get_field_instructions(section):
                instr = instr.strip()

                if not instr.upper().startswith("TOC"):
                    continue

                toc_found = True

                # ⬇️ pokud \o není → implicitně OK
                match = self.TOC_LEVEL_RE.search(instr)
                if not match:
                    return CheckResult(
                        True,
                        "Obsah zahrnuje nadpisy prvních tří úrovní.",
                        0,
                    )

                start = int(match.group(1))
                end = int(match.group(2))

                if start <= 1 and end >= 3:
                    return CheckResult(
                        True,
                        "Obsah zahrnuje nadpisy prvních tří úrovní.",
                        0,
                    )

                return CheckResult(
                    False,
                    "V obsahu nejsou zahrnuty nadpisy prvních tří úrovní (H1–H3).",
                    self.penalty,
                )

        if not toc_found:
            return CheckResult(
                False,
                "V dokumentu nebyl nalezen obsah.",
                self.penalty,
            )

        return CheckResult(
            True,
            "Obsah zahrnuje nadpisy prvních tří úrovní.",
            0,
        )