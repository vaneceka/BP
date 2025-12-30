from checks.word.base_check import BaseCheck, CheckResult


class TOCHeadingLevelsCheck(BaseCheck):
    name = "Nadpisy prvních tří úrovní v obsahu chybí"
    penalty = -5

    def _parse_toc_levels(self, instr: str):
        if "\\o" not in instr:
            return None

        try:
            quoted = instr.split("\\o", 1)[1].split('"')[1]
            start, end = quoted.split("-")
            return int(start), int(end)
        except (IndexError, ValueError):
            return None

    def run(self, document, assignment=None):

        for i in range(document.section_count()):
            section = document.section(i)

            for instr in document.get_field_instructions(section):
                instr = instr.strip()

                if not instr.upper().startswith("TOC"):
                    continue

                levels = self._parse_toc_levels(instr)

                # \o není -> Word implicitně H1–H3
                if levels is None:
                    return CheckResult(
                        True,
                        "Obsah zahrnuje nadpisy prvních tří úrovní.",
                        0,
                    )

                start, end = levels

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

        return CheckResult(
            False,
            "V dokumentu nebyl nalezen obsah.",
            self.penalty,
        )