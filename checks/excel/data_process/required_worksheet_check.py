from checks.base_check import BaseCheck, CheckResult


class RequiredWorksheetCheck(BaseCheck):
    name = "Požadovaný list „zdroj“ zcela chybí"
    penalty = -100

    def run(self, document, assignment=None):
        sheet_names = document.sheet_names()

        if not any(name.lower() == "zdroj" for name in sheet_names):
            return CheckResult(
                False,
                "Požadovaný list „zdroj“ zcela chybí.",
                self.penalty,
            )

        return CheckResult(
            True,
            "Požadovaný list „zdroj“ je v sešitu přítomen.",
            0,
        )