from ..base_check import BaseCheck, CheckResult


class Section2TextCheck(BaseCheck):
    name = "Text ve 2. oddílu"
    penalty = -5

    def run(self, document, assignment=None):
        if not document.has_text_in_section(1):
            return CheckResult(
                False,
                "Ve druhém oddílu chybí text dokumentu.",
                self.penalty
            )

        return CheckResult(True, "Text ve 2. oddílu OK", 0)