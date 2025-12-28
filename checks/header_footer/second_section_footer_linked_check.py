from checks.base_check import BaseCheck, CheckResult


class SecondSectionFooterLinkedCheck(BaseCheck):
    name = "Druhý oddíl je propojen s předchozím oddílem v zápatí"
    penalty = -2

    def run(self, document, assignment=None):

        # 1️⃣ Musí existovat alespoň 2 oddíly
        if document.section_count() < 2:
            return CheckResult(True, "Dokument má méně než dva oddíly.", 0)

        sect1 = document.section_properties(0)
        sect2 = document.section_properties(1)

        if sect1 is None or sect2 is None:
            return CheckResult(True, "Oddíly nemají nastavení.", 0)

        # 2️⃣ Zápatí obou oddílů
        footers1 = sect1.findall("w:footerReference", document.NS)
        footers2 = sect2.findall("w:footerReference", document.NS)

        # ⚠️ KLÍČOVÉ:
        # Pokud druhý oddíl NEMÁ footerReference → je propojen s předchozím
        if not footers2:
            return CheckResult(
                False,
                "Zápatí druhého oddílu je propojeno s předchozím oddílem.",
                self.penalty,
            )

        # 3️⃣ Mapování typ → r:id
        R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

        def footer_map(footers):
            result = {}
            for f in footers:
                f_type = f.attrib.get(f"{{{document.NS['w']}}}type", "default")
                r_id = f.attrib.get(f"{{{R_NS}}}id")
                if r_id:
                    result[f_type] = r_id
            return result

        map1 = footer_map(footers1)
        map2 = footer_map(footers2)

        # 4️⃣ Stejný r:id = propojeno
        for f_type in map1:
            if f_type in map2 and map1[f_type] == map2[f_type]:
                return CheckResult(
                    False,
                    "Zápatí druhého oddílu je propojeno s předchozím oddílem.",
                    self.penalty,
                )

        return CheckResult(
            True,
            "Zápatí druhého oddílu není propojeno s předchozím oddílem.",
            0,
        )