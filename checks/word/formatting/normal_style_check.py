from checks.base_check import BaseCheck, CheckResult


class NormalStyleCheck(BaseCheck):
    name = "Styl Normální není změněn dle zadání"
    penalty = -5

    def run(self, document, assignment=None):
        expected = assignment.styles.get("Normal")
        actual = document.get_normal_style()

        if actual is None:
            return CheckResult(
                False,
                "Definice stylu Normal nebyla nalezena.",
                self.penalty
            )
        
        diffs = actual.diff(expected)
        if diffs:
            message = (
                "Styl Normal neodpovídá zadání:\n"
                + "\n".join(f"- {d}" for d in diffs)
            )

        if not actual.matches(expected):
            return CheckResult(
                False,
                message,
                self.penalty
            )

        return CheckResult(
            True,
            "Styl Normal je nastaven správně dle zadání.",
            0
        )