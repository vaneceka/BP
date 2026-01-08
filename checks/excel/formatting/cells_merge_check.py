import re
from checks.base_check import BaseCheck, CheckResult
from openpyxl.utils import range_boundaries

# class MergedCellsCheck(BaseCheck):
#     name = "Chybné sloučení buněk"
#     penalty = -1
#     SHEET = "data"

#     FORBIDDEN_RANGES = [
#         "A2:F23",  
#         "A28:E30",
#     ]

#     def _overlaps(self, r1, r2):
#         c1_min, r1_min, c1_max, r1_max = r1
#         c2_min, r2_min, c2_max, r2_max = r2

#         return not (
#             c1_max < c2_min or
#             c1_min > c2_max or
#             r1_max < r2_min or
#             r1_min > r2_max
#         )

#     def run(self, document, assignment=None):
#         if self.SHEET not in document.wb.sheetnames:
#             return CheckResult(
#                 False,
#                 f'Chybí list "{self.SHEET}".',
#                 self.penalty,
#                 fatal=True
#             )

#         ws = document.wb[self.SHEET]
#         problems = []

#         forbidden = [range_boundaries(r) for r in self.FORBIDDEN_RANGES]

#         for merged in ws.merged_cells.ranges:
#             merged_bounds = range_boundaries(str(merged))

#             for forb in forbidden:
#                 if self._overlaps(merged_bounds, forb):
#                     problems.append(
#                         f"Sloučené buňky {merged} zasahují do datové oblasti"
#                     )

#         if problems:
#             return CheckResult(
#                 False,
#                 "Chybné sloučení buněk:\n" +
#                 "\n".join("– " + p for p in problems),
#                 self.penalty,
#                 fatal=True 
#             )

#         return CheckResult(
#             True,
#             "Sloučení buněk je použito správně.",
#             0
#         )

class MergedCellsCheck(BaseCheck):
    name = "Chybné sloučení buněk"
    penalty = -1
    SHEET = "data"


    def split_addr(self, addr: str) -> tuple[int, int]:
        m = re.fullmatch(r"([A-Z]+)(\d+)", addr.upper())
        if not m:
            raise ValueError(f"Neplatná adresa buňky: {addr}")

        col_letters, row = m.groups()
        row = int(row)

        col = 0
        for c in col_letters:
            col = col * 26 + (ord(c) - ord("A") + 1)

        return col, row

    def data_table_bounds(self, assignment):
        rows = []
        cols = []

        for addr in assignment.cells:
            col, row = self.split_addr(addr)
            rows.append(row)
            cols.append(col)

        return min(cols), min(rows), max(cols), max(rows)

    def _overlaps(self, a, b):
        c1, r1, c2, r2 = a
        C1, R1, C2, R2 = b
        return not (
            c2 < C1 or c1 > C2 or
            r2 < R1 or r1 > R2
        )

    def run(self, document, assignment=None):

        if self.SHEET not in document.sheet_names():
            return CheckResult(
                False,
                f'Chybí list "{self.SHEET}".',
                self.penalty,
                fatal=True
            )

        if assignment is None:
            return CheckResult(True, "Chybí assignment – přeskočeno.", 0)

        forbidden = self.data_table_bounds(assignment)
        problems = []

        for merged in document.merged_ranges(self.SHEET):
            if self._overlaps(merged, forbidden):
                problems.append(
                    f"Sloučené buňky zasahují do datové oblasti"
                )

        if problems:
            return CheckResult(
                False,
                "Chybné sloučení buněk:\n" +
                "\n".join("– " + p for p in problems),
                self.penalty,
                fatal=True
            )

        return CheckResult(
            True,
            "Sloučení buněk je použito správně.",
            0
        )