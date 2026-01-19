from checks.base_check import BaseCheck, CheckResult

#NOTE funguje pro word i writer
class HeadingStyleCheck(BaseCheck):
    def __init__(self, level: int):
        self.level = level
        self.name = f"Styl Nadpis {level} není změněn dle zadání"
        self.penalty = -5

    def run(self, document, assignment=None):
        expected = assignment.styles.get(f"Heading {self.level}")
        
        if expected is None:
            return CheckResult(True, "Zadání styl neřeší.", 0)

        actual = document.get_heading_style(self.level)
        if actual is None:
            return CheckResult(
                False,
                f"Styl Heading {self.level} v dokumentu neexistuje.",
                self.penalty,
            )

        diffs = actual.diff(expected)

        is_numbered, is_hierarchical, actual_num_level = \
            document.get_heading_numbering_info(self.level)

        expected_num_level = getattr(expected, "numLevel", None)

        if expected_num_level is not None:
            if not is_numbered:
                diffs.append("numLevel: nadpis není číslovaný")
            elif not is_hierarchical:
                diffs.append("numLevel: číslování není hierarchické")
            elif actual_num_level != expected_num_level:
                diffs.append(
                    f"numLevel: očekáváno {expected_num_level}, nalezeno {actual_num_level}"
                )

        if diffs:
            message = (
                f"Styl Heading {self.level} neodpovídá zadání:\n"
                + "\n".join(f"- {d}" for d in diffs)
            )

            return CheckResult(
                False,
                message,
                self.penalty,
            )

        return CheckResult(
            True,
            f"Styl Heading {self.level} odpovídá zadání.",
            0,
        )