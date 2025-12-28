from checks.base_check import BaseCheck, CheckResult


class FirstSectionHeaderEmptyCheck(BaseCheck):
    name = "První oddíl nemá prázdné záhlaví"
    penalty = -2

    def run(self, document, assignment=None):

        if document.section_count() == 0:
            return CheckResult(True, "Dokument nemá oddíly.", 0)

        sect_pr = document.section_properties(0)
        if sect_pr is None:
            return CheckResult(True, "První oddíl nemá nastavení.", 0)

        header_refs = sect_pr.findall("w:headerReference", document.NS)

        for ref in header_refs:
            # ✅ POZOR: je to r:id, ne w:id
            r_id = ref.attrib.get(f"{{{document.NS['r']}}}id")
            if not r_id:
                continue

            # ✅ rId -> reálný soubor
            part_name = document.resolve_part_target(r_id)
            if not part_name:
                continue

            try:
                xml = document._load(part_name)
            except KeyError:
                continue

            # obsah = text
            for t in xml.findall(".//w:t", document.NS):
                if t.text and t.text.strip():
                    return CheckResult(
                        False,
                        "První oddíl obsahuje neprázdné záhlaví.",
                        self.penalty,
                    )

            # obsah = pole (PAGE apod.)
            for instr in xml.findall(".//w:instrText", document.NS):
                if instr.text and instr.text.strip():
                    return CheckResult(
                        False,
                        "První oddíl obsahuje neprázdné záhlaví.",
                        self.penalty,
                    )

        return CheckResult(True, "Záhlaví prvního oddílu je prázdné.", 0)