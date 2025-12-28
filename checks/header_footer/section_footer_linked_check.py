from checks.base_check import BaseCheck, CheckResult


class FooterLinkedToPreviousCheck(BaseCheck):
    penalty = -2

    def __init__(self, section_number: int):
        """
        section_number = číslo oddílu (1-based, lidské číslování)
        """
        self.section_number = section_number
        self.section_index = section_number - 1

        self.name = (
            f"{section_number}. oddíl je propojen s předchozím oddílem v zápatí"
        )

    def run(self, document, assignment=None):

        # první oddíl nemá předchozí → OK
        if self.section_index == 0:
            return CheckResult(True, "První oddíl nemá předchozí oddíl.", 0)

        # musí existovat aktuální i předchozí oddíl
        if document.section_count() <= self.section_index:
            return CheckResult(True, "Oddíl v dokumentu neexistuje.", 0)

        sect_prev = document.section_properties(self.section_index - 1)
        sect_curr = document.section_properties(self.section_index)

        if sect_prev is None or sect_curr is None:
            return CheckResult(True, "Oddíly nemají nastavení.", 0)

        # zápatí obou oddílů
        footers_prev = sect_prev.findall("w:footerReference", document.NS)
        footers_curr = sect_curr.findall("w:footerReference", document.NS)

        # ❌ žádný footerReference → implicitně zděděné
        if not footers_curr:
            return CheckResult(
                False,
                "Zápatí oddílu je převzaté z předchozího oddílu.",
                self.penalty,
            )

        R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

        def footer_map(footers):
            result = {}
            for f in footers:
                f_type = f.attrib.get(f"{{{document.NS['w']}}}type", "default")
                r_id = f.attrib.get(f"{{{R_NS}}}id")
                if r_id:
                    result[f_type] = r_id
            return result

        map_prev = footer_map(footers_prev)
        map_curr = footer_map(footers_curr)

        # stejný r:id = propojené
        for f_type in map_prev:
            if f_type in map_curr and map_prev[f_type] == map_curr[f_type]:
                return CheckResult(
                    False,
                    "Zápatí oddílu je propojené s předchozím oddílem.",
                    self.penalty,
                )

        return CheckResult(
            True,
            "Zápatí oddílu není propojené s předchozím oddílem.",
            0,
        )