from checks.base_check import BaseCheck, CheckResult

class CustomStyleInheritanceCheck(BaseCheck):
    name = "Vlastní styl s dědičností"
    penalty = -5

    # z různých heading 1, Heading 1 -> heading1
    def _norm(self, s: str | None) -> str:
        return (s or "").lower().replace("-", "").replace("_", "").replace(" ", "")
    
    def run(self, document, assignment=None):
        if assignment is None:
            return CheckResult(True, "Zadání nebylo předáno.", 0)

        errors = []

        for style_name, expected in assignment.styles.items():
            if not expected or not expected.basedOn:
                continue 

            style_el = document._find_style(name=style_name)
            if style_el is None:
                errors.append(
                    f'Styl „{style_name}“ neexistuje (nelze ověřit dědičnost).'
                )
                continue

            based_el = style_el.find("w:basedOn", document.NS)
            if based_el is None:
                errors.append(
                    f'Styl „{style_name}“ nemá nastaveno basedOn (má dědit z „{expected.basedOn}“).'
                )
                continue

            actual_parent_id = based_el.attrib.get(f"{{{document.NS['w']}}}val")
            parent_style = document._find_style_by_id(actual_parent_id)

            if parent_style is None:
                errors.append(
                    f'Styl „{style_name}“ dědí z neexistujícího stylu („{actual_parent_id}“).'
                )
                continue

            parent_name_el = parent_style.find("w:name", document.NS)
            parent_name = parent_name_el.attrib.get(f"{{{document.NS['w']}}}val") if parent_name_el is not None else ""

            if self._norm(parent_name) != self._norm(expected.basedOn):
                errors.append(
                    f'Styl „{style_name}“ dědí z „{parent_name}“, '
                    f'ale má dědit z „{expected.basedOn}“.'
                )

        if errors:
            return CheckResult(
                False,
                "Chyby v dědičnosti stylů:\n" + "\n".join(errors),
                self.penalty,
            )

        return CheckResult(
            True,
            "Požadovaná dědičnost stylů je správně nastavena.",
            0,
        )