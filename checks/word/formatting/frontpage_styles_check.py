from checks.base_check import BaseCheck, CheckResult

#NOTE funguje pro word i odt
class FrontpageStylesCheck(BaseCheck):
    name = "Styly pro úvodní list"
    penalty = -5

    def run(self, document, assignment=None):
        errors = []

        required_styles = {
            "desky-fakulta": assignment.styles.get("desky-fakulta"),
            "uvodni-tema": assignment.styles.get("uvodni-tema"),
            "uvodni-autor": assignment.styles.get("uvodni-autor"),
        }

        for style_name, expected in required_styles.items():
            if expected is None:
                continue

            actual = document.get_custom_style(style_name)
            if actual is None:
                errors.append(f"Styl „{style_name}“ v dokumentu neexistuje.")
                continue

            diffs = actual.diff(expected)
            if diffs:
                errors.append(
                    f"Styl „{style_name}“ neodpovídá zadání:\n"
                    + "\n".join(f"  - {d}" for d in diffs)
                )

        if errors:
            return CheckResult(
                False,
                "Chyby ve stylech úvodního listu:\n" + "\n".join(errors),
                self.penalty,
            )

        return CheckResult(True, "Styly úvodního listu odpovídají zadání.", 0)

