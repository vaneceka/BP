from checks.base_check import BaseCheck, CheckResult


class ManualVerticalSpacingCheck(BaseCheck):
    name = "Ruční vertikální odsazení pomocí prázdných řádků"
    penalty = -5  # násobí se

    def run(self, document, assignment=None):
        paragraphs = list(document.iter_paragraphs())
        errors: list[tuple[int, str]] = []

        for i, p in enumerate(paragraphs):
            # ignoruj odstavec, který obsahuje objekt (obrázek, graf, rovnice)
            if (
                p.findall(".//w:drawing", document.NS) or
                p.findall(".//m:oMath", document.NS) or
                p.findall(".//m:oMathPara", document.NS)
            ):
                continue
            # má odstavec text?
            has_text = False

            for t in p.findall(".//w:t", document.NS):
                if t.text is not None and t.text.strip() != "":
                    has_text = True
                    break
            
            if has_text:
                continue

            # poslední odstavec dokumentu se ignoruje
            if not document.has_text_after_paragraph(paragraphs, i):
                continue

            if p.find("w:pPr/w:sectPr", document.NS) is not None:
                continue

            next_p = paragraphs[i + 1]

            # ignoruj, pokud: prázdný řádek je součástí TOC / seznamu, nebo sousedí s TOC / seznamem, nebo je generovaný polem
            if (
                document._paragraph_is_toc_or_object_list(p)
                or document._paragraph_is_toc_or_object_list(next_p)
                or document.paragraph_is_generated_by_field(p)
                or document.paragraph_is_generated_by_field(next_p)
            ):
                continue

            if document.paragraph_has_page_break(next_p):
                continue

            next_style_id = document._paragraph_style_id(next_p)
            if next_style_id and document.style_has_page_break(next_style_id):
                continue

            next_ppr = next_p.find("w:pPr", document.NS)
            if next_ppr is not None:
                spacing = next_ppr.find("w:spacing", document.NS)
                if spacing is not None:
                    before = spacing.attrib.get(f"{{{document.NS['w']}}}before")
                    if before and int(before) > 0:
                        continue

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