from checks.base_check import BaseCheck, CheckResult
import re


class InconsistentFormattingCheck(BaseCheck):
    name = "Nekonzistentní formátování textu"
    penalty = -5  # násobí se

    def run(self, document, assignment=None):
        errors = []

        for p in document.iter_paragraphs():
            ppr = p.find("w:pPr", document.NS)
            if ppr is None:
                continue

            pstyle = ppr.find("w:pStyle", document.NS)
            if pstyle is None:
                continue  # bez stylu → neřešíme tady

            style_name = pstyle.attrib.get(f"{{{document.NS['w']}}}val")

            # text odstavce (zkrácený)
            texts = []
            for t in p.findall(".//w:t", document.NS):
                if t.text:
                    texts.append(t.text)
            preview = re.sub(r"\s+", " ", "".join(texts)).strip()[:60]

            for r in p.findall(".//w:r", document.NS):
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
                    break  # jeden odstavec = jedna chyba

        if errors:
            return CheckResult(
                False,
                "V dokumentu je použito nekonzistentní formátování:\n"
                + "\n".join(errors),
                self.penalty * len(errors),
            )

        return CheckResult(
            True,
            "Nenalezeno nekonzistentní formátování.",
            0,
        )