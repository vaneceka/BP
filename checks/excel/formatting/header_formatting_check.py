from checks.base_check import BaseCheck, CheckResult

class HeaderFormattingCheck(BaseCheck):
    name = "Není formátováno záhlaví tabulky"
    penalty = -1

    SHEET = "data"

    def run(self, document, assignment=None):
        if assignment is None or not hasattr(assignment, "cells"):
            return CheckResult(True, "Chybí assignment – check přeskočen.", 0)

        problems = []

        for addr, spec in assignment.cells.items():
            style = spec.style
            if not style:
                continue

            if not style.get("bold", False):
                continue

            expected_alignment = style.get("alignment", False)

            cell = document.get_cell(f"{self.SHEET}!{addr}")
            if cell is None:
                problems.append(f"{addr}: záhlaví chybí")
                continue

            raw = cell["raw_cell"]

            if not raw.font or not raw.font.bold:
                problems.append(f"{addr}: záhlaví není tučné")

            if expected_alignment:
                h = raw.alignment.horizontal
                v = raw.alignment.vertical

                if h != "center" or v != "center":
                    problems.append(
                        f"{addr}: záhlaví není zarovnáno na střed (nalezeno {h}/{v})"
                    )

        if problems:
            return CheckResult(
                False,
                "Záhlaví tabulky není správně formátováno:\n"
                + "\n".join("– " + p for p in problems),
                self.penalty,
                fatal=True
            )

        return CheckResult(
            True,
            "Záhlaví tabulek jsou správně formátována.",
            0
        )