from checks.word.base_check import BaseCheck, CheckResult


class OriginalFormattingCheck(BaseCheck):
    name = "Text obsahuje původní formátování (HTML, TXT, PDF)"
    penalty = -100

    def run(self, document, assignment=None):
        problems = []

        manual = document.find_manual_formatting()
        if manual:
            lines = [
                f"- ruční formátování v odstavci {idx}: „{txt[:80]}…“"
                for idx, txt in manual[:5]
            ]
            problems.append(
                "Ruční formátování místo stylů:\n" + "\n".join(lines)
            )

        if document.has_inline_font_changes():
            problems.append("Text obsahuje inline změny fontu/velikosti.")

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