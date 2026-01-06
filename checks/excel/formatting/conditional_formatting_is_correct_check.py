from checks.base_check import BaseCheck, CheckResult
import re

# NOTE - kontroluje se pouze, zda barva existuje.
class ConditionalFormattingCorrectnessCheck(BaseCheck):
    name = "Podmíněné formátování nefunguje správně"
    penalty = -2
    SHEET = "data"

    NUMBER_RE = re.compile(r"-?\d+(\.\d+)?")

    def run(self, document, assignment=None):

        if assignment is None or not hasattr(assignment, "cells"):
            return CheckResult(True, "Chybí assignment – check přeskočen.", 0)

        if self.SHEET not in document.wb.sheetnames:
            return CheckResult(
                False,
                f'Chybí list "{self.SHEET}".',
                self.penalty,
                fatal=True
            )

        ws = document.wb[self.SHEET]

        expected = []

        for addr, spec in assignment.cells.items():
            if not spec.conditionalFormat:
                continue

            column = re.sub(r"\d+", "", addr)  

            for rule in spec.conditionalFormat:
                expected.append({
                    "operator": rule["operator"],
                    "value": float(rule["value"]),
                    "column": column,
                    "needs_fill": rule.get("fillColor") is not None,
                    "needs_font": rule.get("textColor") is not None,
                })

        if not expected:
            return CheckResult(
                False,
                "V assignmentu není definováno žádné podmíněné formátování.",
                self.penalty,
                fatal=True
            )

        found = [False] * len(expected)

        for cf, rules in ws.conditional_formatting._cf_rules.items():
            range_str = str(cf.sqref)  

            for r in rules:
                if r.type != "cellIs" or not r.formula:
                    continue

                m = self.NUMBER_RE.search(r.formula[0])
                if not m:
                    continue

                try:
                    actual_value = float(m.group())
                except ValueError:
                    continue

                for i, exp in enumerate(expected):
                    if found[i]:
                        continue

                    if exp["column"] not in range_str:
                        continue

                    if r.operator != exp["operator"]:
                        continue

                    if abs(actual_value - exp["value"]) > 0.01:
                        continue

                    if exp["needs_fill"]:
                        if not (r.dxf and r.dxf.fill and r.dxf.fill.fgColor):
                            continue

                    if exp["needs_font"]:
                        if not (r.dxf and r.dxf.font and r.dxf.font.color):
                            continue

                    found[i] = True

        missing = []

        for i in range(len(expected)):
            if not found[i]:
                e = expected[i]
                missing.append(
                    f'{e["operator"]} {e["value"]} (sloupec {e["column"]})'
                )

        if missing:
            return CheckResult(
                False,
                "Podmíněné formátování neodpovídá zadání:\n– "
                + "\n– ".join(missing),
                self.penalty * len(missing),
                fatal=True
            )

        return CheckResult(
            True,
            "Podmíněné formátování odpovídá zadání.",
            0
        )