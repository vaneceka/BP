from checks.word.base_check import BaseCheck, CheckResult

TOLERANCE = 10  # twips


class CustomStyleWithTabsCheck(BaseCheck):
    name = "Vlastní styl s definovanými tabulátory"
    penalty = -2  # násobí se

    def run(self, document, assignment=None):
        if assignment is None:
            return CheckResult(True, "Zadání nebylo předáno.", 0)

        errors = []
        total_penalty = 0

        for style_name, spec in assignment.styles.items():
            if not spec.tabs:
                continue  # styl nemá mít tabulátory

            style_el = document._find_style(name=style_name)
            if style_el is None:
                errors.append(f"Styl „{style_name}“ v dokumentu neexistuje.")
                total_penalty += self.penalty
                continue

            # tabulátory musí být explicitně ve stylu
            ppr = style_el.find("w:pPr", document.NS)
            tabs_el = ppr.find("w:tabs", document.NS) if ppr is not None else None

            if tabs_el is None:
                errors.append(
                    f"Styl „{style_name}“ nemá explicitně definované tabulátory."
                )
                total_penalty += self.penalty
                continue

            # skutečné tabulátory
            actual_tabs = {}
            for tab in tabs_el.findall("w:tab", document.NS):
                val = tab.attrib.get(f"{{{document.NS['w']}}}val")
                pos = tab.attrib.get(f"{{{document.NS['w']}}}pos")

                if not val or not pos or val.lower() == "clear":
                    continue

                actual_tabs[val.lower()] = int(pos)

            # očekávané tabulátory (twips)
            expected_tabs = {}

            for align, pos in spec.tabs:
                expected_tabs[align.lower()] = int(pos)

            # kontrola počtu
            if set(actual_tabs.keys()) != set(expected_tabs.keys()):
                errors.append(
                    f"Styl „{style_name}“ má špatné typy tabulátorů:\n"
                    f"  očekáváno: {expected_tabs}\n"
                    f"  nalezeno:  {actual_tabs}"
                )
                total_penalty += self.penalty
                continue

            for align, expected_pos in expected_tabs.items():
                actual_pos = actual_tabs[align]
                if abs(expected_pos - actual_pos) > TOLERANCE:
                    errors.append(
                        f"Styl „{style_name}“ má špatné tabulátory:\n"
                        f"  očekáváno: {expected_tabs}\n"
                        f"  nalezeno:  {actual_tabs}"
                    )
                    total_penalty += self.penalty
                    break

        if errors:
            return CheckResult(
                False,
                "Chyby v nastavení tabulátorů:\n" + "\n".join(errors),
                total_penalty,
            )

        return CheckResult(
            True,
            "Všechny styly s tabulátory jsou nastaveny správně dle zadání.",
            0,
        )