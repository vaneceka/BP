from checks.base_check import BaseCheck, CheckResult


class SectionFooterHasPageNumberCheck(BaseCheck):
    penalty = -2

    def __init__(self, section_number: int):
        self.section_number = section_number
        self.section_index = section_number - 1

        self.name = f"{section_number}. oddíl má číslo stránky v zápatí"

    def run(self, document, assignment=None):

        if document.section_count() <= self.section_index:
            return CheckResult(True, "Oddíl v dokumentu neexistuje.", 0)

        sect_pr = document.section_properties(self.section_index)
        if sect_pr is None:
            return CheckResult(
                False,
                f"{self.section_number}. oddíl nemá nastavení oddílu.",
                self.penalty,
            )

        footer_refs = sect_pr.findall("w:footerReference", document.NS)

        if not footer_refs:
            return CheckResult(
                False,
                f"{self.section_number}. oddíl nemá vlastní zápatí s číslem stránky.",
                self.penalty,
            )

        for ref in footer_refs:
            if ref.find("w:linkToPrevious", document.NS) is not None:
                continue

            r_id = ref.attrib.get(f"{{{document.NS['r']}}}id")
            if not r_id:
                continue

            part_path = document.resolve_part_target(r_id)
            if not part_path:
                continue

            try:
                footer_xml = document._load(part_path)
            except KeyError:
                continue

            for fld in footer_xml.findall(".//w:fldSimple", document.NS):
                instr = fld.attrib.get(f"{{{document.NS['w']}}}instr", "")
                if "PAGE" in instr.upper():
                    return CheckResult(
                        True,
                        f"Zápatí {self.section_number}. oddílu obsahuje číslo stránky.",
                        0,
                    )

            for instr in footer_xml.findall(".//w:instrText", document.NS):
                if instr.text and "PAGE" in instr.text.upper():
                    return CheckResult(
                        True,
                        f"Zápatí {self.section_number}. oddílu obsahuje číslo stránky.",
                        0,
                    )

        return CheckResult(
            False,
            f"{self.section_number}. oddíl nemá vlastní zápatí s číslem stránky.",
            self.penalty,
        )