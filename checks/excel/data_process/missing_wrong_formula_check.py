import re
from checks.base_check import BaseCheck, CheckResult


class MissingOrWrongFormulaOrNotCalculatedCheck(BaseCheck):
    name = "Zcela chybí výpočet/vzorec, je chybně nebo hodnota vypočtena není."
    penalty = -10 

    def _norm_formula(self, f: str | None) -> str:
        if not f:
            return ""
        f = re.sub(r"\s+", "", f)
        return f.upper()

    def _is_error_value(self, v) -> bool:
        return isinstance(v, str) and v.startswith("#")

    def run(self, document, assignment=None) -> CheckResult:

        if assignment is None or not hasattr(assignment, "cells"):
            return CheckResult(
                True,
                "Chybí excel assignment – check přeskočen.",
                0
            )

        sheet = "data"

        if sheet not in document.sheet_names():
            return CheckResult(
                False,
                f'List "{sheet}" chybí – nelze ověřit vzorce.',
                self.penalty,
                fatal=True
            )

        errors = []

        for addr, spec in assignment.cells.items():
            expected = getattr(spec, "expression", None)
            if not expected:
                continue

            cell = document.get_cell(f"{sheet}!{addr}")

            if cell is None:
                errors.append(f"{sheet}!{addr}: buňka neexistuje")
                continue

            actual_formula = cell["formula"]

            if not (isinstance(actual_formula, str) and actual_formula.startswith("=")):
                errors.append(f"{sheet}!{addr}: chybí vzorec")
                continue

            if self._norm_formula(actual_formula) != self._norm_formula(expected):
                errors.append(
                    f"{sheet}!{addr}: špatný vzorec "
                    f"(oček. {expected}, nalezeno {actual_formula})"
                )
                continue

            cached = cell.get("value_cached")

            if cached is None:
                errors.append(f"{sheet}!{addr}: vzorec nemá uložený výsledek")
                continue

            if self._is_error_value(cached):
                errors.append(f"{sheet}!{addr}: výsledek je chyba {cached}")
                continue

        if errors:
            return CheckResult(
                False,
                "Problémy se vzorci / výpočtem:\n"
                + "\n".join(f"– {e}" for e in errors),
                self.penalty,
                fatal=True
            )

        return CheckResult(
            True,
            "Všechny požadované vzorce existují a mají uložený výsledek.",
            0
        )