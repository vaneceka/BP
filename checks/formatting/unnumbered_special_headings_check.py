from checks.base_check import BaseCheck, CheckResult


class UnnumberedSpecialHeadingsCheck(BaseCheck):
    name = "Styly speciálních kapitol nesmí být číslované"
    penalty = -1

    SPECIAL_STYLES = [
        "obsah",
        "Nadpisobsahu",
        "content heading",
        "Seznamobrzk",
        "seznam obrázků",
        "seznam tabulek",
        "bibliografie",
        "Vrazncitt"
    ]

    def run(self, document, assignment=None):
        errors = []

        for style_name in self.SPECIAL_STYLES:
            style = document.get_style_by_any_name([style_name])
            if style and style.isNumbered:
                errors.append(style.name)

        if errors:
            return CheckResult(
                False,
                "Následující styly mají zapnuté číslování:\n"
                + "\n".join(f"- {s}" for s in errors),
                self.penalty,
            )

        return CheckResult(True, "Styly speciálních kapitol nejsou číslované.", 0)