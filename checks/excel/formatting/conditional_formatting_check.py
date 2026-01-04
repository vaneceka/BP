from checks.base_check import BaseCheck, CheckResult

class ConditionalFormattingCheck(BaseCheck):
    name = "Chybí podmíněné formátování"
    penalty = -5
    SHEET = "data"

    def _get(self, obj, key):
        if isinstance(obj, dict):
            return obj.get(key)
        return getattr(obj, key, None)

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

        # 1️⃣ expected CF z assignmentu (deduplikované)
        expected = {}

        for spec in assignment.cells.values():
            print(vars(spec))
            cf = self._get(spec, "conditional_format")
            if not cf:
                continue

            for rule in cf:
                key = (
                    rule["type"],
                    rule["operator"],
                    str(rule["value"]),
                    rule.get("fillColor") is not None,
                    rule.get("textColor") is not None,
                )
                expected[key] = {
                    "operator": rule["operator"],
                    "value": str(rule["value"]),
                    "needs_fill": rule.get("fillColor") is not None,
                    "needs_font": rule.get("textColor") is not None,
                }

        if not expected:
            return CheckResult(
                False,
                "V assignmentu není definováno žádné podmíněné formátování.",
                self.penalty,
                fatal=True
            )

        # 2️⃣ skutečná CF v Excelu
        found = {key: False for key in expected}

        for rules in ws.conditional_formatting._cf_rules.values():
            for r in rules:
                if r.type != "cellIs":
                    continue

                for key, exp in expected.items():
                    if found[key]:
                        continue

                    if r.operator != exp["operator"]:
                        continue

                    if not r.formula or exp["value"] not in r.formula[0]:
                        continue

                    if exp["needs_fill"] and not (r.dxf and r.dxf.fill):
                        continue

                    if exp["needs_font"] and not (r.dxf and r.dxf.font):
                        continue

                    found[key] = True

        # 3️⃣ vyhodnocení – SČÍTÁNÍ CHYB
        missing = []

        for key, ok in found.items():
            if not ok:
                exp = expected[key]
                missing.append(f'{exp["operator"]} {exp["value"]}')

        if missing:
            return CheckResult(
                False,
                "Chybí podmíněné formátování:\n– " + "\n– ".join(missing),
                self.penalty * len(missing),
                fatal=True
            )

        return CheckResult(
            True,
            "Podmíněné formátování odpovídá assignmentu.",
            0
        )