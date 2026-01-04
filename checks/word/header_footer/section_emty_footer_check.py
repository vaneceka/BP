from checks.base_check import BaseCheck, CheckResult


class SectionFooterEmptyCheck(BaseCheck):
    penalty = -2

    def __init__(self, section_number: int):
        self.section_number = section_number
        self.section_index = section_number - 1
        self.name = f"{section_number}. oddíl nemá prázdné zápatí"

    def run(self, document, assignment=None):
        if document.section_count() <= self.section_index:
            return CheckResult(True, "Oddíl v dokumentu neexistuje.", 0)

        sect_pr = document.section_properties(self.section_index)
        if sect_pr is None:
            return CheckResult(True, "Oddíl nemá nastavení.", 0)

        footer_refs = sect_pr.findall("w:footerReference", document.NS)

        if not footer_refs:
            return CheckResult(True, "Zápatí oddílu je prázdné.", 0)

        for ref in footer_refs:
            r_id = ref.attrib.get(f"{{{document.NS['r']}}}id")
            if not r_id:
                continue

            part_name = document.resolve_part_target(r_id)
            if not part_name:
                continue

            try:
                xml = document._load(part_name)
            except KeyError:
                continue
            
            for t in xml.findall(".//w:t", document.NS):
                if t.text and t.text.strip():
                    return CheckResult(
                        False,
                        f"{self.section_number}. oddíl obsahuje neprázdné zápatí (text).",
                        self.penalty,
                    )

            if (
                xml.findall(".//w:drawing", document.NS)
                or xml.findall(".//m:oMath", document.NS)
                or xml.findall(".//m:oMathPara", document.NS)
            ):
                return CheckResult(
                    False,
                    f"{self.section_number}. oddíl obsahuje neprázdné zápatí (objekt).",
                    self.penalty,
                )

        return CheckResult(True, "Zápatí oddílu je prázdné.", 0)