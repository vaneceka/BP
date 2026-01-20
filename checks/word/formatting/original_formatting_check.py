from checks.base_check import BaseCheck, CheckResult

#NOTE pridat PDF
class OriginalFormattingCheck(BaseCheck):
    name = "Text obsahuje původní formátování (HTML, TXT, PDF)"
    penalty = -100

    def run(self, document, assignment=None):
        problems = []

        if document.has_html_artifacts():
            problems.append("Text obsahuje znaky typické pro HTML.")

        if problems:
            return CheckResult(
                False,
                "Bylo detekováno původní formátování:\n" +
                "\n".join(f"- {p}" for p in problems),
                self.penalty
            )

        return CheckResult(
            True,
            "Text neobsahuje původní formátování.",
            0
        )