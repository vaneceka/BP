from checks.base_check import BaseCheck, CheckResult

class DescriptiveStatisticsCheck(BaseCheck):
    name = "Chybí popisná charakteristika pro datovou řadu"
    penalty = -5
    SHEET = "data"

    REQUIRED_CELLS = {
        "Výška": ["B28", "C28", "D28", "E28"],
        "Váha":  ["B29", "C29", "D29", "E29"],
        "BMI":   ["B30", "C30", "D30", "E30"],
    }

    def run(self, document, assignment=None):
        problems = []

        if self.SHEET not in document.sheet_names():
            return CheckResult(
                False,
                f'Chybí list "{self.SHEET}".',
                self.penalty,
                fatal=True
            )

        for series, cells in self.REQUIRED_CELLS.items():
            for addr in cells:
                cell = document.get_cell(f"{self.SHEET}!{addr}")

                if cell is None or not cell["formula"]:
                    problems.append(f"{series}: {addr} chybí vzorec")
                    continue

                if cell["value_cached"] is None:
                    problems.append(
                        f"{series}: {addr} nemá uložený výsledek (nevypočteno / neuloženo)"
                    )

        if problems:
            return CheckResult(
                False,
                "Popisná charakteristika není kompletní:\n"
                + "\n".join("– " + p for p in problems),
                self.penalty,
                fatal=True 
            )

        return CheckResult(
            True,
            "Popisná charakteristika je kompletní.",
            0
        )