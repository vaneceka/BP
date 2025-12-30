from checks.word.base_check import BaseCheck, CheckResult

class BibliographyNotUpdatedCheck(BaseCheck):
    name = "Seznam literatury není aktuální"
    penalty = -5

    def run(self, document, assignment=None):

        citations = document.count_word_citations()

        # žádné citace → nemá smysl kontrolovat
        if citations == 0:
            return CheckResult(True, "Dokument neobsahuje citace.", 0)

        # bibliografie chybí → řeší jiný check
        if not document.has_word_bibliography():
            return CheckResult(True, "Bibliografie chybí – řeší jiný check.", 0)

        items = document.count_bibliography_items()

        if items < citations:
            return CheckResult(
                False,
                f"Seznam literatury není aktuální (citací: {citations}, položek: {items}).",
                self.penalty,
            )

        return CheckResult(
            True,
            "Seznam literatury odpovídá počtu citací.",
            0,
        )