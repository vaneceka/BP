from checks.word.base_check import BaseCheck, CheckResult

class InconsistentFormattingCheck(BaseCheck):
    name = "Nekonzistentní formátování textu"
    penalty = -5

    def run(self, document, assignment=None):
        errors = []

        for p in document.iter_paragraphs():

            # ignoruj obsah / seznamy
            if document._paragraph_is_toc_or_object_list(p):
                continue

            ppr = p.find("w:pPr", document.NS)
            if ppr is None:
                continue

            pstyle = ppr.find("w:pStyle", document.NS)
            if pstyle is None:
                continue

            style_name = pstyle.attrib.get(f"{{{document.NS['w']}}}val")

            preview = document._paragraph_text(p)[:60]

            for r in p.findall(".//w:r", document.NS):

                if r.find("w:instrText", document.NS) is not None:
                    continue
                if r.find("w:fldChar", document.NS) is not None:
                    continue

                # musí mít skutečný text
                has_text = False

                for t in r.findall(".//w:t", document.NS):
                    if t.text is not None and t.text.strip() != "":
                        has_text = True
                        break
                if not has_text:
                    continue

                rpr = r.find("w:rPr", document.NS)
                if rpr is None:
                    continue

                problems = []

                if rpr.find("w:b", document.NS) is not None:
                    problems.append("tučné písmo")
                if rpr.find("w:i", document.NS) is not None:
                    problems.append("kurzíva")
                if rpr.find("w:sz", document.NS) is not None:
                    problems.append("změna velikosti písma")
                if rpr.find("w:rFonts", document.NS) is not None:
                    problems.append("změna fontu")
                if rpr.find("w:color", document.NS) is not None:
                    problems.append("změna barvy")

                if problems:
                    errors.append(
                        f"- Styl „{style_name}“, ruční zásah: {', '.join(problems)}"
                        f"\n  Text: „{preview}“"
                    )
                    break

        if errors:
            return CheckResult(
                False,
                "V dokumentu je použito nekonzistentní formátování:\n"
                + "\n".join(errors),
                self.penalty * len(errors),
            )

        return CheckResult(True, "Nenalezeno nekonzistentní formátování.", 0)