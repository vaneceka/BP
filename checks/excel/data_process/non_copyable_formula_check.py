from checks.base_check import BaseCheck, CheckResult
import re


class NonCopyableFormulasCheck(BaseCheck):
    name = "Vzorce nelze kopírovat"
    penalty = -100

    CELL_REF_RE = re.compile(r"\b[A-Z]{1,3}\d+\b")

    def run(self, document, assignment=None):
        bad_cells = []

        for cell in document.cells_with_formulas():
            formula = cell["formula"]

            if not isinstance(formula, str):
                bad_cells.append(
                    (cell["sheet"], cell["address"], "maticový nebo interní vzorec")
                )
                continue

            has_cell_ref = bool(self.CELL_REF_RE.search(formula))

            has_any_absolute = "$" in formula

            if has_cell_ref and not has_any_absolute:
                bad_cells.append(
                    (cell["sheet"], cell["address"], formula)
                )

        if bad_cells:
            lines = [
                f"- {sheet}!{addr}: {info}"
                for sheet, addr, info in bad_cells[:5]
            ]

            return CheckResult(
                False,
                "Vzorce nelze kopírovat:\n" + "\n".join(lines),
                self.penalty,
                fatal=True,
            )

        return CheckResult(
            True,
            "Vzorce jsou kopírovatelné.",
            0,
        )