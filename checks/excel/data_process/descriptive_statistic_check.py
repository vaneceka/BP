import re
from checks.base_check import BaseCheck, CheckResult

# class DescriptiveStatisticsCheck(BaseCheck):
#     name = "Chybí popisná charakteristika pro datovou řadu"
#     penalty = -5
#     SHEET = "data"

#     REQUIRED_CELLS = {
#         "Výška": ["B28", "C28", "D28", "E28"],
#         "Váha":  ["B29", "C29", "D29", "E29"],
#         "BMI":   ["B30", "C30", "D30", "E30"],
#     }

#     def run(self, document, assignment=None):
#         problems = []

#         if self.SHEET not in document.sheet_names():
#             return CheckResult(
#                 False,
#                 f'Chybí list "{self.SHEET}".',
#                 self.penalty,
#                 fatal=True
#             )

#         for series, cells in self.REQUIRED_CELLS.items():
#             for addr in cells:
#                 cell = document.get_cell(f"{self.SHEET}!{addr}")

#                 if cell is None or not cell["formula"]:
#                     problems.append(f"{series}: {addr} chybí vzorec")
#                     continue

#                 if cell["value_cached"] is None:
#                     problems.append(
#                         f"{series}: {addr} nemá uložený výsledek (nevypočteno / neuloženo)"
#                     )

#         if problems:
#             return CheckResult(
#                 False,
#                 "Popisná charakteristika není kompletní:\n"
#                 + "\n".join("– " + p for p in problems),
#                 self.penalty,
#                 fatal=True 
#             )

#         return CheckResult(
#             True,
#             "Popisná charakteristika je kompletní.",
#             0
#         )

from checks.base_check import BaseCheck, CheckResult

class DescriptiveStatisticsCheck(BaseCheck):
    name = "Chybí popisná charakteristika pro datovou řadu"
    penalty = -5
    SHEET = "data"

    def row_from_addr(self, addr: str) -> int:
        return int(re.match(r"[A-Z]+(\d+)", addr).group(1))


    def run(self, document, assignment=None):

        if assignment is None or not hasattr(assignment, "cells"):
            return CheckResult(True, "Chybí assignment – check přeskočen.", 0)

        if self.SHEET not in document.sheet_names():
            return CheckResult(
                False,
                f'Chybí list "{self.SHEET}".',
                self.penalty,
                fatal=True,
            )

        rows = sorted({self.row_from_addr(a) for a in assignment.cells})
        start_of_stats = None

        for a, b in zip(rows, rows[1:]):
            if b > a + 1:
                start_of_stats = b
                break

        if start_of_stats is None:
            return CheckResult(
                False,
                "Nelze určit začátek popisné charakteristiky.",
                self.penalty,
                fatal=True,
            )

        problems = []

        for addr, spec in assignment.cells.items():
            if not getattr(spec, "expression", None):
                continue

            if self.row_from_addr(addr) <= start_of_stats:
                continue

            info = document.get_cell_info(self.SHEET, addr)

            if info is None or not info["formula"]:
                problems.append(f"{self.SHEET}!{addr}: chybí vzorec")

        if problems:
            return CheckResult(
                False,
                "Chybí popisná charakteristika:\n"
                + "\n".join("– " + p for p in problems),
                self.penalty,
                fatal=True,
            )

        return CheckResult(
            True,
            "Popisná charakteristika je přítomna.",
            0,
        )