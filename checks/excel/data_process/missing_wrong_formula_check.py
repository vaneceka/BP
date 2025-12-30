import re
from checks.base_check import BaseCheck, CheckResult

class MissingOrWrongFormulaOrNotCalculatedCheck(BaseCheck):
    name = "Zcela chybí výpočet/vzorec, je chybně nebo hodnota vypočtena není."
    penalty = -10  # u tebe je [-10*] => klidně nastav fatal=True

    def _norm_formula(self, f: str | None) -> str:
        if not f:
            return ""
        # odstraníme whitespace a sjednotíme velikost písmen (funkce)
        f = re.sub(r"\s+", "", f)
        return f.upper()

    def _is_error_value(self, v) -> bool:
        # Excel chyby bývají stringy typu "#DIV/0!" apod.
        return isinstance(v, str) and v.startswith("#")

    def run(self, document, assignment=None) -> CheckResult:
        # očekávám, že do runneru předáš excel assignment (ne word)
        if assignment is None or not hasattr(assignment, "cells"):
            return CheckResult(True, "Chybí excel assignment – check přeskočen.", 0)

        # ⚠️ pokud máš ve více listech, rozšiř model o sheet u každé buňky
        # pro teď předpokládám fixní list "data"
        sheet = "data"

        errors = []

        # projdi všechny buňky, které mají mít vzorec
        for addr, spec in assignment.cells.items():
            expected = getattr(spec, "expression", None)
            if not expected:
                continue

            # 1) existuje list?
            if sheet not in document.sheet_names():
                return CheckResult(False, f'List "{sheet}" chybí – nelze ověřit vzorce.', self.penalty, fatal=True)

            cell = document.get_cell(sheet, addr)
            actual_formula = cell.value if isinstance(cell.value, str) else None

            # openpyxl vrací vzorec jako string začínající "="
            if not (isinstance(actual_formula, str) and actual_formula.startswith("=")):
                errors.append(f"{sheet}!{addr}: chybí vzorec")
                continue

            # 2) porovnání vzorce (normalizované)
            if self._norm_formula(actual_formula) != self._norm_formula(expected):
                errors.append(
                    f"{sheet}!{addr}: špatný vzorec (oček. {expected}, nalezeno {actual_formula})"
                )
                continue

            # 3) kontrola „vypočteno není“ => cached value v data_only=True
            cached = document.get_cell_value_cached(sheet, addr)

            if cached is None:
                errors.append(f"{sheet}!{addr}: vzorec nemá uložený výsledek (nevypočteno / neuloženo)")
                continue

            if self._is_error_value(cached):
                errors.append(f"{sheet}!{addr}: výsledek je chyba {cached}")
                continue

        if errors:
            return CheckResult(
                False,
                "Problémy se vzorci / výpočtem:\n" + "\n".join(f"– {e}" for e in errors),
                self.penalty,   # pokud chceš penalizovat za každou buňku, dej self.penalty * len(errors)
                fatal=True      # pokud to [-10*] má být „hvězdička = fatální“
            )

        return CheckResult(True, "Všechny požadované vzorce existují a mají uložený výsledek.", 0)