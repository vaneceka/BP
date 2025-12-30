from checks.word.base_check import BaseCheck, CheckResult


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

        diffs = actual.diff(expected, strict=True)
        if diffs:
            # složíme čitelnou zprávu
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