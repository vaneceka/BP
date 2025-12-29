from ..base_check import BaseCheck, CheckResult

class BibliographyCheck(BaseCheck):
    name = "Seznam literatury"
    penalty = -5

    def run(self, document):
        if not document.has_bibliography():
            return CheckResult(
                False,
                "Chybí seznam literatury",
                self.penalty
            )

        return CheckResult(True, "Seznam literatury OK", 0)