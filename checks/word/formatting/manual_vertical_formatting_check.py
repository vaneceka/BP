# from checks.base_check import BaseCheck, CheckResult


# class ManualVerticalSpacingCheck(BaseCheck):
#     name = "Ruční vertikální odsazení pomocí prázdných řádků"
#     penalty = -5 

#     def run(self, document, assignment=None):
#         paragraphs = list(document.iter_paragraphs())
#         errors: list[tuple[int, str]] = []

#         for i, p in enumerate(paragraphs):
#             if (
#                 p.findall(".//w:drawing", document.NS) or
#                 p.findall(".//m:oMath", document.NS) or
#                 p.findall(".//m:oMathPara", document.NS)
#             ):
#                 continue
#             has_text = False

#             for t in p.findall(".//w:t", document.NS):
#                 if t.text is not None and t.text.strip() != "":
#                     has_text = True
#                     break
            
#             if has_text:
#                 continue

#             if not document.has_text_after_paragraph(paragraphs, i):
#                 continue

#             if p.find("w:pPr/w:sectPr", document.NS) is not None:
#                 continue

#             next_p = paragraphs[i + 1]

#             if (
#                 document._paragraph_is_toc_or_object_list(p)
#                 or document._paragraph_is_toc_or_object_list(next_p)
#                 or document.paragraph_is_generated_by_field(p)
#                 or document.paragraph_is_generated_by_field(next_p)
#             ):
#                 continue

#             if document.paragraph_has_page_break(next_p):
#                 continue

#             next_style_id = document._paragraph_style_id(next_p)
#             if next_style_id and document.style_has_page_break(next_style_id):
#                 continue

#             next_ppr = next_p.find("w:pPr", document.NS)
#             if next_ppr is not None:
#                 spacing = next_ppr.find("w:spacing", document.NS)
#                 if spacing is not None:
#                     before = spacing.attrib.get(f"{{{document.NS['w']}}}before")
#                     if before and int(before) > 0:
#                         continue

#             context = document._paragraph_text(next_p)
#             style = document._paragraph_style_id(next_p) or "bez stylu"
#             errors.append((i + 1, i + 2, style, context))

#         if errors:
#             lines = [
#                 f"- prázdný řádek (odstavec {empty_idx}) před odstavcem {next_idx} "
#                 f"(styl: {style}): „{txt[:80]}…“"
#                 for empty_idx, next_idx, style, txt in errors[:5]
#             ]

#             return CheckResult(
#                 False,
#                 "V dokumentu je text vertikálně formátován pomocí prázdných řádků:\n"
#                 + "\n".join(lines),
#                 self.penalty * len(errors),
#             )

#         return CheckResult(
#             True,
#             "Nenalezeno ruční vertikální odsazení pomocí prázdných řádků.",
#             0,
#         )

from checks.base_check import BaseCheck, CheckResult


class ManualVerticalSpacingCheck(BaseCheck):
    name = "Ruční vertikální odsazení pomocí prázdných řádků"
    penalty = -5

    def run(self, document, assignment=None):
        paragraphs = list(document.iter_paragraphs())
        errors = []

        i = 0
        while i < len(paragraphs) - 1:
            p = paragraphs[i]

            if not document.paragraph_is_empty(p):
                i += 1
                continue

            empty_count = 1
            j = i + 1
            while j < len(paragraphs) and document.paragraph_is_empty(paragraphs[j]):
                empty_count += 1
                j += 1

            if empty_count == 1:
                i = j
                continue

            if j >= len(paragraphs):
                break

            next_p = paragraphs[j]

            if not document.paragraph_has_text(next_p):
                i = j
                continue

            if document._paragraph_is_toc_or_object_list(next_p):
                i = j
                continue

            if document.paragraph_is_generated(next_p):
                i = j
                continue

            if document.paragraph_has_page_break(next_p):
                i = j
                continue

            if document.paragraph_has_spacing_before(next_p):
                i = j
                continue

            errors.append({
                "empty_from": i + 1,
                "empty_to": j,
                "style": document.paragraph_style_name(next_p) or "bez stylu",
                "text": document.paragraph_text(next_p),
                "count": empty_count,
            })

            i = j

        if errors:
            lines = []
            for e in errors[:5]:
                lines.append(
                    f"- {e['count']} prázdné řádky (odstavce {e['empty_from']}–{e['empty_to']}) "
                    f"před odstavcem (styl: {e['style']}):\n"
                    f"  „{e['text'][:80]}…“"
                )

            return CheckResult(
                False,
                "V dokumentu je text vertikálně formátován pomocí více prázdných řádků:\n"
                + "\n".join(lines),
                self.penalty * len(errors),
            )

        return CheckResult(
            True,
            "Nenalezeno ruční vertikální odsazení pomocí více prázdných řádků.",
            0,
        )