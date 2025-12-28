from checks.base_check import BaseCheck, CheckResult


class FirstSectionFooterEmptyCheck(BaseCheck):
    name = "První oddíl nemá prázdné zápatí"
    penalty = -2

    def run(self, document, assignment=None):

        # 1️⃣ musí existovat aspoň jeden oddíl
        if document.section_count() == 0:
            return CheckResult(True, "Dokument nemá oddíly.", 0)

        # 2️⃣ první oddíl
        sect_pr = document.section_properties(0)
        if sect_pr is None:
            return CheckResult(True, "První oddíl nemá nastavení.", 0)

        # 3️⃣ reference na zápatí v 1. oddílu
        footer_refs = sect_pr.findall("w:footerReference", document.NS)

        for ref in footer_refs:
            # ✅ footerReference má r:id (relationships), ne w:id
            r_id = ref.attrib.get(f"{{{document.NS['r']}}}id")
            if not r_id:
                continue

            # ✅ rId -> reálný soubor (např. word/footer1.xml)
            part_name = document.resolve_part_target(r_id)
            if not part_name:
                continue

            try:
                xml = document._load(part_name)
            except KeyError:
                continue

            # 4️⃣ hledej skutečný obsah (text)
            for t in xml.findall(".//w:t", document.NS):
                if t.text and t.text.strip():
                    return CheckResult(
                        False,
                        "První oddíl obsahuje neprázdné zápatí.",
                        self.penalty,
                    )

            # 5️⃣ hledej pole (PAGE, NUMPAGES, DATE, ...)
            for instr in xml.findall(".//w:instrText", document.NS):
                if instr.text and instr.text.strip():
                    return CheckResult(
                        False,
                        "První oddíl obsahuje neprázdné zápatí.",
                        self.penalty,
                    )

        return CheckResult(True, "Zápatí prvního oddílu je prázdné.", 0)