from checks.base_check import BaseCheck, CheckResult

class NumberFormattingCheck(BaseCheck):
    name = "Chybné formátování číselných hodnot"
    penalty = -2
    SHEET = "data"

    def run(self, document, assignment=None):
        if assignment is None or not hasattr(assignment, "cells"):
            return CheckResult(True, "Chybí assignment – check přeskočen.", 0)

        problems = []

        for addr, spec in assignment.cells.items():
            style = spec.style or {}
            expected_fmt = style.get("numberFormat")
            expected_align = style.get("alignment", False)

            cell = document.get_cell(f"{self.SHEET}!{addr}")
            if cell is None:
                continue

            raw = cell["raw_cell"]

            if expected_fmt is not None:
                if raw.number_format != expected_fmt:
                    problems.append(
                        f"{addr}: špatný formát čísla (oček. {expected_fmt}, nalezen {raw.number_format})"
                    )

            if expected_align is True:
                h = raw.alignment.horizontal
                v = raw.alignment.vertical

                if h != "center" or v != "center":
                    problems.append(
                        f"{addr}: špatné zarovnání (oček. center/center, nalezen {h}/{v})"
                    )

        if problems:
            return CheckResult(
                False,
                "Chybné formátování:\n" + "\n".join("– " + p for p in problems[:50]),
                self.penalty,
                fatal=True
            )

        return CheckResult(True, "Formátování odpovídá zadání.", 0)