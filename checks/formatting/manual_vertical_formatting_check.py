from checks.base_check import BaseCheck, CheckResult


class ManualVerticalSpacingCheck(BaseCheck):
    name = "Ruční vertikální odsazení pomocí prázdných řádků"
    penalty = -5  # násobí se

    def run(self, document, assignment=None):
        paragraphs = list(document.iter_paragraphs())
        errors: list[tuple[int, str]] = []

        for i, p in enumerate(paragraphs):
            # ⛔ ignoruj odstavec, který obsahuje objekt (obrázek, graf, rovnice)
            if (
                p.findall(".//w:drawing", document.NS) or
                p.findall(".//m:oMath", document.NS) or
                p.findall(".//m:oMathPara", document.NS)
            ):
                continue
            # 1️⃣ má odstavec text?
            has_text = any(
                t.text and t.text.strip()
                for t in p.findall(".//w:t", document.NS)
            )
            if has_text:
                continue

            # 2️⃣ poslední odstavec dokumentu → ignoruj
            # ⛔ prázdné řádky na konci dokumentu ignoruj
            if not document.has_text_after_paragraph(paragraphs, i):
                continue

            # 3️⃣ sectPr (oddíly) → OK
            if p.find("w:pPr/w:sectPr", document.NS) is not None:
                continue

            next_p = paragraphs[i + 1]
            # ⛔ ignoruj, pokud následující odstavec patří obsahu / seznamům
            if document._paragraph_is_toc_or_object_list(next_p):
                continue

            # 4️⃣ page break
            if document.paragraph_has_page_break(next_p):
                continue

            next_style_id = document._paragraph_style_id(next_p)
            if next_style_id and document.style_has_page_break(next_style_id):
                continue

            # 5️⃣ spaceBefore
            next_ppr = next_p.find("w:pPr", document.NS)
            if next_ppr is not None:
                spacing = next_ppr.find("w:spacing", document.NS)
                if spacing is not None:
                    before = spacing.attrib.get(f"{{{document.NS['w']}}}before")
                    if before and int(before) > 0:
                        continue

            # ❌ ruční odsazení
            context = document._paragraph_text(next_p)
            style = document._paragraph_style_id(next_p) or "bez stylu"
            errors.append((i + 1, i + 2, style, context))

        if errors:
            lines = [
                f"- prázdný řádek (odstavec {empty_idx}) před odstavcem {next_idx} "
                f"(styl: {style}): „{txt[:80]}…“"
                for empty_idx, next_idx, style, txt in errors[:5]
            ]

            return CheckResult(
                False,
                "V dokumentu je text vertikálně formátován pomocí prázdných řádků:\n"
                + "\n".join(lines),
                self.penalty * len(errors),
            )

        return CheckResult(
            True,
            "Nenalezeno ruční vertikální odsazení pomocí prázdných řádků.",
            0,
        )