from checks.base_check import BaseCheck, CheckResult

class NumberFormattingCheck(BaseCheck):
    name = "Chybné formátování číselných hodnot"
    penalty = -2

    def run(self, document, assignment=None):
        if assignment is None or not hasattr(assignment, "cells"):
            return CheckResult(True, "Chybí assignment – check přeskočen.", 0)

        problems = []

        for addr, spec in assignment.cells.items():
            expected_fmt = spec.style.get("numberFormat") if hasattr(spec, "style") else None
            expected_align = spec.style.get("alignment") if hasattr(spec, "style") else None

            if expected_fmt is None:
                continue  # neřešíme textové buňky

            cell = document.get_cell(f"data!{addr}")
            if cell is None:
                continue

            raw = cell["raw_cell"]

            # number format
            if raw.number_format != expected_fmt:
                problems.append(
                    f"{addr}: špatný formát čísla (oček. {expected_fmt}, nalezen {raw.number_format})"
                )

            # zarovnání (pokud je požadováno)
            if expected_align:
                actual_align = raw.alignment.horizontal
                if actual_align not in ("right", "center"):
                    problems.append(
                        f"{addr}: špatné zarovnání ({actual_align})"
                    )

        if problems:
            return CheckResult(
                False,
                "Chybné formátování číselných hodnot:\n"
                + "\n".join("– " + p for p in problems),
                self.penalty,
                fatal=True  # hvězdička v zadání
            )

        return CheckResult(
            True,
            "Číselné hodnoty jsou formátovány správně.",
            0
        )