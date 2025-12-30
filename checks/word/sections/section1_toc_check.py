from ...base_check import BaseCheck, CheckResult


class Section1TOCCheck(BaseCheck):
    name = "Obsah v 1. oddílu"
    penalty = -5

    def run(self, document, assignment=None):
        if not document.has_toc_in_section(0):
            return CheckResult(
                False,
                "V prvním oddílu chybí obsah dokumentu.",
                self.penalty
            )

        return CheckResult(True, "Obsah v 1. oddílu OK", 0)