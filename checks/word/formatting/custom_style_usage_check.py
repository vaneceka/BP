from checks.base_check import BaseCheck, CheckResult


class RequiredCustomStylesUsageCheck(BaseCheck):
    name = "Použití požadovaných vlastních stylů"
    penalty = -2  # násobí se

    def run(self, document, assignment=None):
        if assignment is None:
            return CheckResult(True, "Zadání nebylo předáno.", 0)

        errors = []
        total_penalty = 0

        # rozdělíme styly
        custom_styles, _ = document.split_assignment_styles(assignment)

        # zjistíme base styly
        base_styles = set()

        for spec in custom_styles.values():
            if spec.basedOn:
                base_styles.add(spec.basedOn)

        used_style_ids = set()
        for p in document.iter_paragraphs():
            ppr = p.find("w:pPr", document.NS)
            if ppr is None:
                continue
            ps = ppr.find("w:pStyle", document.NS)
            if ps is not None:
                used_style_ids.add(
                    ps.attrib.get(f"{{{document.NS['w']}}}val")
                )

        for style_name, spec in custom_styles.items():

            style_el = document._find_style(name=style_name)
            if style_el is None:
                errors.append(f"Styl „{style_name}“ v dokumentu neexistuje.")
                total_penalty += self.penalty
                continue

            if style_name in base_styles:
                continue

            style_id = style_el.attrib.get(f"{{{document.NS['w']}}}styleId")
            if style_id not in used_style_ids:
                errors.append(f"Styl „{style_name}“ existuje, ale není použit.")
                total_penalty += self.penalty

        if errors:
            return CheckResult(
                False,
                "Problémy s použitím vlastních stylů:\n" + "\n".join(errors),
                total_penalty,
            )

        return CheckResult(
            True,
            "Všechny požadované vlastní styly existují a jsou použity.",
            0,
        )