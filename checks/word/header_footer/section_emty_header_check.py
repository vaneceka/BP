from checks.base_check import BaseCheck, CheckResult


class SectionHeaderEmptyCheck(BaseCheck):
    penalty = -2

    def __init__(self, section_number: int):
        self.section_number = section_number
        self.section_index = section_number - 1

        self.name = f"{section_number}. oddíl nemá prázdné záhlaví"

    def run(self, document, assignment=None):

        if document.section_count() <= self.section_index:
            return CheckResult(True, "Oddíl v dokumentu neexistuje.", 0)

        sect_pr = document.section_properties(self.section_index)
        if sect_pr is None:
            return CheckResult(True, "Oddíl nemá nastavení.", 0)

        header_refs = sect_pr.findall("w:headerReference", document.NS)

        # žádný headerReference -> implicitně prázdné
        if not header_refs:
            return CheckResult(
                True,
                "Záhlaví oddílu je prázdné.",
                0,
            )

        for ref in header_refs:
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

            # text v záhlaví
            for t in xml.findall(".//w:t", document.NS):
                if t.text and t.text.strip():
                    return CheckResult(
                        False,
                        f"{self.section_number}. oddíl obsahuje neprázdné záhlaví.",
                        self.penalty,
                    )

            # pole (PAGE, DATE, apod.)
            for instr in xml.findall(".//w:instrText", document.NS):
                if instr.text and instr.text.strip():
                    return CheckResult(
                        False,
                        f"{self.section_number}. oddíl obsahuje neprázdné záhlaví.",
                        self.penalty,
                    )

        return CheckResult(
            True,
            "Záhlaví oddílu je prázdné.",
            0,
        )