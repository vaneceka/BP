from checks.word.base_check import BaseCheck, CheckResult

class TOCExistsCheck(BaseCheck):
    name = "Obsah zcela chybí"
    penalty = -100

    def run(self, document, assignment=None):
        for i in range(document.section_count()):
            if document.has_toc_in_section(i):
                return CheckResult(True, "Obsah dokumentu existuje.", 0)

        return CheckResult(
            False,
            "V dokumentu zcela chybí obsah.",
            self.penalty,
            fatal=True,
        )