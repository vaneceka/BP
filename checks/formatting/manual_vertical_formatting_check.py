from checks.base_check import BaseCheck, CheckResult


class ManualVerticalSpacingCheck(BaseCheck):
    name = "Ruční vertikální odsazení pomocí prázdných řádků"
    penalty = -5  # násobí se

    def run(self, document, assignment=None):
        paragraphs = list(document.iter_paragraphs())
        errors = 0

        for i, p in enumerate(paragraphs):
            # 1️⃣ má odstavec text?
            has_text = any(
                t.text and t.text.strip()
                for t in p.findall(".//w:t", document.NS)
            )
            if has_text:
                continue

            # 2️⃣ poslední odstavec dokumentu → ignoruj
            if i == len(paragraphs) - 1:
                continue

            # 3️⃣ sectPr (oddíly) → OK
            if p.find("w:pPr/w:sectPr", document.NS) is not None:
                continue

            next_p = paragraphs[i + 1]

            # 4️⃣ následující odstavec má page break (přímo nebo ve stylu)
            if document.paragraph_has_page_break(next_p):
                continue

            next_style_id = document._paragraph_style_id(next_p)
            if next_style_id and document.style_has_page_break(next_style_id):
                continue

            # 5️⃣ následující odstavec má spaceBefore → OK
            next_ppr = next_p.find("w:pPr", document.NS)
            if next_ppr is not None:
                spacing = next_ppr.find("w:spacing", document.NS)
                if spacing is not None:
                    before = spacing.attrib.get(f"{{{document.NS['w']}}}before")
                    if before and int(before) > 0:
                        continue

            # ❌ jinak je to ruční vertikální odsazení
            errors += 1

        if errors:
            return CheckResult(
                False,
                "V dokumentu je text vertikálně formátován pomocí prázdných řádků "
                "(místo použití stylů nebo konce stránky).",
                self.penalty * errors,
            )

        return CheckResult(
            True,
            "Nenalezeno ruční vertikální odsazení pomocí prázdných řádků.",
            0,
        )