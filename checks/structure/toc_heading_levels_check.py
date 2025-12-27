import re
from checks.base_check import BaseCheck, CheckResult


class TOCHeadingLevelsCheck(BaseCheck):
    name = "Nadpisy prvních tří úrovní v obsahu chybí"
    penalty = -5  # násobí se

    TOC_LEVEL_RE = re.compile(r'\\o\s*"(\d+)\s*-\s*(\d+)"')

    def run(self, document, assignment=None):
        errors = 0

        for i in range(document.section_count()):
            section = document.section(i)

            for instr in document.get_field_instructions(section):
                instr = instr.strip()

                if not instr.upper().startswith("TOC"):
                    continue

                match = self.TOC_LEVEL_RE.search(instr)
                if not match:
                    # TOC existuje, ale nemá \o → špatně
                    errors += 1
                    continue

                start = int(match.group(1))
                end = int(match.group(2))

                if start > 1 or end < 3:
                    errors += 1

        if errors:
            return CheckResult(
                False,
                "V obsahu nejsou zahrnuty nadpisy prvních tří úrovní (H1–H3).",
                self.penalty * errors,
            )

        return CheckResult(
            True,
            "Obsah zahrnuje nadpisy prvních tří úrovní.",
            0,
        )