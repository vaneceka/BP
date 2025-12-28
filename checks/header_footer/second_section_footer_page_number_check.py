from checks.base_check import BaseCheck, CheckResult


class SecondSectionFooterHasPageNumberCheck(BaseCheck):
    name = "Druhý oddíl má číslo stránky v zápatí"
    penalty = -2

    def run(self, document, assignment=None):

        if document.section_count() < 2:
            return CheckResult(True, "Dokument má méně než dva oddíly.", 0)

        sect_pr = document.section_properties(1)
        if sect_pr is None:
            return CheckResult(False, "Druhý oddíl nemá nastavení oddílu.", self.penalty)

        footer_refs = sect_pr.findall("w:footerReference", document.NS)
        if not footer_refs:
            return CheckResult(False, "Druhý oddíl nemá žádné zápatí.", self.penalty)

        for ref in footer_refs:

            # ❌ zděděné zápatí → bereme jako NEEXISTUJÍCÍ
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

            # 1️⃣ fldSimple
            for fld in footer_xml.findall(".//w:fldSimple", document.NS):
                instr = fld.attrib.get(f"{{{document.NS['w']}}}instr", "")
                if "PAGE" in instr.upper():
                    return CheckResult(True, "Zápatí druhého oddílu obsahuje číslo stránky.", 0)

            # 2️⃣ instrText
            for instr in footer_xml.findall(".//w:instrText", document.NS):
                if instr.text and "PAGE" in instr.text.upper():
                    return CheckResult(True, "Zápatí druhého oddílu obsahuje číslo stránky.", 0)

        return CheckResult(
            False,
            "Druhý oddíl nemá vlastní zápatí s číslem stránky.",
            self.penalty
        )