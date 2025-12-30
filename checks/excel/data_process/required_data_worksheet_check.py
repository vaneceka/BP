from checks.base_check import BaseCheck, CheckResult


class RequiredDataWorksheetCheck(BaseCheck):
    name = "Požadovaný list „zdroj“ zcela chybí"
    penalty = -100

    def run(self, document, assignment=None):
        sheet_names = document.sheet_names()

        if not any(name.lower() == "data" for name in sheet_names):
            return CheckResult(
                False,
                "Požadovaný list „data“ zcela chybí.",
                self.penalty,
            )

        return CheckResult(
            True,
            "Požadovaný list „data“ je v sešitu přítomen.",
            0,
        )