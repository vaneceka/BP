from checks.base_check import BaseCheck, CheckResult
import re


# class NonCopyableFormulasCheck(BaseCheck):
#     name = "Vzorce nelze kopírovat"
#     penalty = -100

#     CELL_REF_RE = re.compile(r"\b[A-Z]{1,3}\d+\b")

#     def run(self, document, assignment=None):
#         bad_cells = []

#         for cell in document.cells_with_formulas():
#             formula = cell["formula"]

#             if not isinstance(formula, str):
#                 bad_cells.append(
#                     (cell["sheet"], cell["address"], "maticový nebo interní vzorec")
#                 )
#                 continue

#             has_cell_ref = bool(self.CELL_REF_RE.search(formula))

#             has_any_absolute = "$" in formula

#             if has_cell_ref and not has_any_absolute:
#                 bad_cells.append(
#                     (cell["sheet"], cell["address"], formula)
#                 )

#         if bad_cells:
#             lines = [
#                 f"- {sheet}!{addr}: {info}"
#                 for sheet, addr, info in bad_cells[:5]
#             ]

#             return CheckResult(
#                 False,
#                 "Vzorce nelze kopírovat:\n" + "\n".join(lines),
#                 self.penalty,
#                 fatal=True,
#             )

#         return CheckResult(
#             True,
#             "Vzorce jsou kopírovatelné.",
#             0,
#         )


# NOTE mozna se to bude jeste upravoat
class NonCopyableFormulasCheck(BaseCheck):
    name = "Vzorce nelze kopírovat"
    penalty = -100

    CELL_REF_RE = re.compile(r"\b[A-Z]{1,3}\$?\d+\b")

    def _normalize_formula(self, f: str | None) -> str:
        if not f:
            return ""

        f = f.strip()

        if f.startswith("of:="):
            f = "=" + f[4:]

        f = re.sub(r"\[\.(\w+\$?\d+)\]", r"\1", f)
        f = re.sub(r"\[\.(\w+):\.(\w+)\]", r"\1:\2", f)
        f = f.replace(";", ",")
        f = re.sub(r"\s+", "", f)

        return f.upper()

    def run(self, document, assignment=None):
        bad_cells = []

        for cell in document.cells_with_formulas():
            raw_formula = cell["formula"]
            formula = self._normalize_formula(raw_formula)

            if not formula:
                bad_cells.append(
                    (cell["sheet"], cell["address"], "nečitelný vzorec")
                )
                continue

            has_cell_ref = bool(self.CELL_REF_RE.search(formula))
            has_absolute = "$" in formula

            if has_cell_ref and not has_absolute:
                bad_cells.append(
                    (cell["sheet"], cell["address"], raw_formula)
                )

        if bad_cells:
            lines = [
                f"- {sheet}!{addr}: {info}"
                for sheet, addr, info in bad_cells
            ]

            return CheckResult(
                False,
                "Vzorce nelze kopírovat (chybí absolutní odkazy):\n"
                + "\n".join(lines),
                self.penalty,
                fatal=True,
            )

        return CheckResult(
            True,
            "Vzorce jsou kopírovatelné.",
            0,
        )