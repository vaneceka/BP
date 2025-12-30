from checks.base_check import BaseCheck, CheckResult
import re


class NamedRangeUsageCheck(BaseCheck):
    name = "Použity pojmenované oblasti místo adres buněk"
    penalty = -10

    TOKEN_RE = re.compile(r"\b[A-Za-z_][A-Za-z0-9_]*\b")
    CELL_RE = re.compile(r"\$?[A-Z]{1,3}\$?\d+")

    def run(self, document, assignment=None):

        defined_names = set(document.wb.defined_names.keys())
        bad_cells = []

        for cell in document.cells_with_formulas():
            formula = cell["formula"]
            if not isinstance(formula, str):
                continue

            tokens = self.TOKEN_RE.findall(formula)

            for token in tokens:
                # funkce (IF, SUM, AVERAGE…)
                if token.upper() == token:
                    continue

                # adresa buňky
                if self.CELL_RE.fullmatch(token):
                    continue

                # pojmenovaná oblast
                if token in defined_names:
                    bad_cells.append(
                        f"{cell['sheet']}!{cell['address']}: {token}"
                    )
                    break

        if bad_cells:
            return CheckResult(
                False,
                "Použity pojmenované oblasti místo adres buněk:\n"
                + "\n".join(f"- {c}" for c in bad_cells[:5]),
                self.penalty,
            )

        return CheckResult(
            True,
            "Vzorce používají pouze adresy buněk.",
            0,
        )