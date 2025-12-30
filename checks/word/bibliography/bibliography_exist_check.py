from checks.word.base_check import BaseCheck, CheckResult


class MissingBibliographyCheck(BaseCheck):
    name = "Chybí seznam literatury"
    penalty = -100

    def run(self, document, assignment=None):

        if document.section_count() == 0:
            return CheckResult(
                False,
                "Dokument neobsahuje žádné oddíly.",
                self.penalty,
            )

        last_section = document.section_count() - 1

        if not document.has_bibliography_in_section(last_section):
            return CheckResult(
                False,
                "Seznam literatury chybí.",
                self.penalty,
            )

        return CheckResult(
            True,
            "Seznam literatury je v dokumentu přítomen.",
            0,
        )