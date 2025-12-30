from checks.base_check import BaseCheck, CheckResult
import re


class NonCopyableFormulasCheck(BaseCheck):
    name = "Vzorce nelze kopírovat"
    penalty = -100

    CELL_REF_RE = re.compile(r"\b[A-Z]{1,3}\d+\b")
    ABS_REF_RE = re.compile(r"\$[A-Z]{1,3}\$\d+")

    def run(self, document, assignment=None):

        bad_cells = []

        for cell in document.cells_with_formulas():
            formula = cell["formula"]

            if not formula:
                continue

            # maticový vzorec
            if cell.get("is_array"):
                bad_cells.append(
                    (cell["sheet"], cell["address"], "maticový vzorec")
                )
                continue

            # relativní odkazy
            has_rel = bool(self.CELL_REF_RE.search(formula))
            has_abs = bool(self.ABS_REF_RE.search(formula))

            if has_rel and not has_abs:
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