from checks.base_check import BaseCheck, CheckResult


class SecondSectionPageNumberStartsAtOneCheck(BaseCheck):
    name = "Druhý oddíl má číslování stránek od 1"
    penalty = -5

    def run(self, document, assignment=None):
        if document.section_count() < 2:
            return CheckResult(True, "Dokument má méně než dva oddíly.", 0)

        sect_pr = document.section_properties(1)
        if sect_pr is None:
            return CheckResult(
                False,
                "Druhý oddíl nemá nastavení oddílu.",
                self.penalty
            )

        pg_num = sect_pr.find("w:pgNumType", document.NS)

        if pg_num is None:
            return CheckResult(
                False,
                "Číslování stránek druhého oddílu nezačíná od 1.",
                self.penalty
            )

        start = pg_num.attrib.get(f"{{{document.NS['w']}}}start")

        if start != "1":
            return CheckResult(
                False,
                "Číslování stránek druhého oddílu nezačíná od 1.",
                self.penalty
            )

        return CheckResult(
            True,
            "Číslování stránek druhého oddílu začíná od 1.",
            0
        )