from checks.base_check import BaseCheck, CheckResult


class HeaderFooterMissingCheck(BaseCheck):
    name = "Záhlaví nebo zápatí v dokumentu není řešeno"
    penalty = -30

    def run(self, document, assignment=None):

        for i in range(document.section_count()):
            if document.section_has_header_or_footer_content(i):
                return CheckResult(
                    True,
                    "Záhlaví nebo zápatí je v dokumentu řešeno.",
                    0,
                )

        return CheckResult(
            False,
            "Dokument neobsahuje žádné záhlaví ani zápatí.",
            self.penalty,
        )