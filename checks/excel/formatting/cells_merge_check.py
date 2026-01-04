from checks.base_check import BaseCheck, CheckResult
from openpyxl.utils import range_boundaries

class MergedCellsCheck(BaseCheck):
    name = "Chybné sloučení buněk"
    penalty = -1
    SHEET = "data"

    FORBIDDEN_RANGES = [
        "A2:F23",  
        "A28:E30",
    ]

    def _overlaps(self, r1, r2):
        c1_min, r1_min, c1_max, r1_max = r1
        c2_min, r2_min, c2_max, r2_max = r2

        return not (
            c1_max < c2_min or
            c1_min > c2_max or
            r1_max < r2_min or
            r1_min > r2_max
        )

    def run(self, document, assignment=None):
        if self.SHEET not in document.wb.sheetnames:
            return CheckResult(
                False,
                f'Chybí list "{self.SHEET}".',
                self.penalty,
                fatal=True
            )

        ws = document.wb[self.SHEET]
        problems = []

        forbidden = [range_boundaries(r) for r in self.FORBIDDEN_RANGES]

        for merged in ws.merged_cells.ranges:
            merged_bounds = range_boundaries(str(merged))

            for forb in forbidden:
                if self._overlaps(merged_bounds, forb):
                    problems.append(
                        f"Sloučené buňky {merged} zasahují do datové oblasti"
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