from ..base_check import BaseCheck, CheckResult


class Section3BibliographyCheck(BaseCheck):
    name = "Seznam literatury ve 3. oddílu"
    penalty = -5

    def run(self, document, assignment=None):
        section = document.section(2)

        for el in section:
            for sdt in el.findall(".//w:sdt", document.NS):
                sdt_pr = sdt.find("w:sdtPr", document.NS)
                if sdt_pr is not None and sdt_pr.find("w:bibliography", document.NS) is not None:
                    return CheckResult(True, "Seznam literatury nalezen.", 0)

        return CheckResult(
            False,
            "Ve třetím oddílu chybí seznam literatury.",
            self.penalty
        )