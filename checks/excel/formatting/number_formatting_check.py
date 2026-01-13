# from checks.base_check import BaseCheck, CheckResult

# class NumberFormattingCheck(BaseCheck):
#     name = "Chybné formátování číselných hodnot"
#     penalty = -2
#     SHEET = "data"

#     def run(self, document, assignment=None):
#         if assignment is None or not hasattr(assignment, "cells"):
#             return CheckResult(True, "Chybí assignment – check přeskočen.", 0)

#         problems = []

#         for addr, spec in assignment.cells.items():
#             style = spec.style or {}
#             expected_fmt = style.get("numberFormat")

#             cell = document.get_cell(f"{self.SHEET}!{addr}")
#             if cell is None:
#                 continue

#             raw = cell["raw_cell"]

#             if expected_fmt is not None:
#                 if raw.number_format != expected_fmt:
#                     problems.append(
#                         f"{addr}: špatný formát čísla (oček. {expected_fmt}, nalezen {raw.number_format})"
#                     )

#         if problems:
#             return CheckResult(
#                 False,
#                 "Chybné formátování:\n" + "\n".join("– " + p for p in problems[:50]),
#                 self.penalty,
#                 fatal=True
#             )

#         return CheckResult(True, "Formátování odpovídá zadání.", 0)

from checks.base_check import BaseCheck, CheckResult
from document.calc_document import CalcDocument

class NumberFormattingCheck(BaseCheck):
    name = "Chybné formátování číselných hodnot"
    penalty = -2
    SHEET = "data"

    def expected_decimal_places(self, fmt: str) -> int:
        if "." not in fmt:
            return 0
        return len(fmt.split(".")[1])

    def run(self, document, assignment=None):
        if assignment is None or not hasattr(assignment, "cells"):
            return CheckResult(True, "Chybí assignment – check přeskočen.", 0)

        # ✅ univerzální kontrola listu
        if hasattr(document, "has_sheet") and not document.has_sheet(self.SHEET):
            return CheckResult(
                True,
                f'List "{self.SHEET}" neexistuje – kontrola se přeskakuje.',
                0
            )

        problems = []

        for addr, spec in assignment.cells.items():
            style_req = spec.style or {}
            expected_fmt = style_req.get("numberFormat")

            if not expected_fmt:
                continue

            style = document.get_cell_style(self.SHEET, addr)
            if not style:
                continue

            # 🔹 CALC
            if isinstance(document, CalcDocument):
                expected_dp = self.expected_decimal_places(expected_fmt)
                found_dp = style.get("decimal_places")

                if found_dp is None:
                    cell = document._find_cell(self.SHEET, addr)
                    if not cell:
                        continue

                    text = "".join(cell["raw_cell"].itertext()).strip()
                    if "," in text:
                        found_dp = len(text.split(",")[1])
                    elif "." in text:
                        found_dp = len(text.split(".")[1])
                    else:
                        found_dp = 0

                if found_dp != expected_dp:
                    problems.append(
                        f"{addr}: špatný počet desetinných míst "
                        f"(oček. {expected_dp}, nalezen {found_dp})"
                    )

            # 🔹 EXCEL
            else:
                number_fmt = style.get("number_format")
                if number_fmt != expected_fmt:
                    problems.append(
                        f"{addr}: špatný formát čísla "
                        f"(oček. {expected_fmt}, nalezen {number_fmt})"
                    )

        if problems:
            return CheckResult(
                False,
                "Chybné formátování:\n" + "\n".join("– " + p for p in problems[:50]),
                self.penalty,
                fatal=True,
            )

        return CheckResult(True, "Formátování odpovídá zadání.", 0)